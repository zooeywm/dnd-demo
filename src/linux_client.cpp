#include "vfd_protocol.h"

#include <QApplication>
#include <QCommandLineParser>
#include <QDateTime>
#include <QDragEnterEvent>
#include <QDropEvent>
#include <QFile>
#include <QFileInfo>
#include <QMimeData>
#include <QPainter>
#include <QTcpSocket>
#include <QUrl>
#include <QWidget>

class VfdClient final : public QObject {
    Q_OBJECT

public:
    explicit VfdClient(QObject *parent = nullptr)
        : QObject(parent)
    {
        connect(&m_socket, &QTcpSocket::connected, this, [] {
            qInfo().noquote() << "connected to Windows agent";
        });
        connect(&m_socket, &QTcpSocket::disconnected, this, [] {
            qWarning().noquote() << "disconnected from Windows agent";
        });
        connect(&m_socket, &QTcpSocket::errorOccurred, this, [this](QAbstractSocket::SocketError) {
            qWarning().noquote() << "socket error:" << m_socket.errorString();
        });
        connect(&m_socket, &QTcpSocket::readyRead, this, &VfdClient::onReadyRead);
    }

    void connectToAgent(const QString &host, quint16 port)
    {
        m_socket.connectToHost(host, port);
    }

    bool hasRemoteSize() const
    {
        return m_remoteWidth > 0 && m_remoteHeight > 0;
    }

    QSize remoteSize() const
    {
        return {m_remoteWidth, m_remoteHeight};
    }

    bool sendFileDescriptor(const QString &path)
    {
        const QFileInfo info(path);
        if (!info.exists() || !info.isFile()) {
            qWarning().noquote() << "dragEnter rejected non-file path:" << path;
            return false;
        }

        const QByteArray nameUtf8 = info.fileName().toUtf8();
        QByteArray payload;
        vfd::appendLe64(payload, static_cast<quint64>(info.size()));
        vfd::appendLe64s(payload, info.lastModified().toMSecsSinceEpoch());
        vfd::appendLe32(payload, static_cast<quint32>(nameUtf8.size()));
        payload.append(nameUtf8);

        m_filePath = info.absoluteFilePath();
        m_fileSize = static_cast<quint64>(info.size());

        qInfo().noquote() << "dragEnter file:" << info.fileName()
                          << "size:" << info.size();
        return send(vfd::MsgType::FileDescriptor, payload);
    }

    void sendMove(const QPoint &remotePos)
    {
        QByteArray payload;
        vfd::appendLe32s(payload, remotePos.x());
        vfd::appendLe32s(payload, remotePos.y());
        qInfo().noquote() << "MOVE x:" << remotePos.x() << "y:" << remotePos.y();
        send(vfd::MsgType::Move, payload);
    }

    void sendDrop(const QPoint &remotePos)
    {
        QByteArray payload;
        vfd::appendLe32s(payload, remotePos.x());
        vfd::appendLe32s(payload, remotePos.y());
        qInfo().noquote() << "DROP x:" << remotePos.x() << "y:" << remotePos.y();
        send(vfd::MsgType::Drop, payload);
    }

    void sendCancel()
    {
        qInfo().noquote() << "CANCEL";
        send(vfd::MsgType::Cancel);
    }

private slots:
    void onReadyRead()
    {
        m_buffer.append(m_socket.readAll());

        while (true) {
            vfd::MsgType type = vfd::MsgType::Error;
            QByteArray payload;
            QString errorText;
            if (!vfd::tryTakeMessage(m_buffer, type, payload, &errorText)) {
                if (!errorText.isEmpty()) {
                    qWarning().noquote() << "protocol error:" << errorText;
                    m_socket.disconnectFromHost();
                }
                return;
            }

            switch (type) {
            case vfd::MsgType::Hello:
                handleHello(payload);
                break;
            case vfd::MsgType::ReadRequest:
                handleReadRequest(payload);
                break;
            case vfd::MsgType::Ok:
                qInfo().noquote() << "OK from Windows agent";
                break;
            case vfd::MsgType::Error:
                qWarning().noquote() << "ERROR from Windows agent:"
                                     << QString::fromUtf8(payload);
                break;
            default:
                qWarning().noquote() << "unexpected message type from Windows agent:"
                                     << static_cast<quint16>(type);
                break;
            }
        }
    }

private:
    bool send(vfd::MsgType type, const QByteArray &payload = {})
    {
        if (m_socket.state() != QAbstractSocket::ConnectedState) {
            qWarning().noquote() << "cannot send message; Windows agent is not connected";
            return false;
        }

        const QByteArray message = vfd::makeMessage(type, payload);
        return m_socket.write(message) == message.size();
    }

    void handleHello(const QByteArray &payload)
    {
        if (payload.size() != 8) {
            qWarning().noquote() << "invalid HELLO payload size:" << payload.size();
            return;
        }

        m_remoteWidth = static_cast<int>(vfd::readLe32(payload.constData()));
        m_remoteHeight = static_cast<int>(vfd::readLe32(payload.constData() + 4));
        qInfo().noquote() << "Windows desktop size:" << m_remoteWidth << "x" << m_remoteHeight;
    }

    void handleReadRequest(const QByteArray &payload)
    {
        if (payload.size() != 16) {
            qWarning().noquote() << "invalid READ_REQUEST payload size:" << payload.size();
            return;
        }

        const quint32 seq = vfd::readLe32(payload.constData());
        const quint64 offset = vfd::readLe64(payload.constData() + 4);
        const quint32 requestedSize = vfd::readLe32(payload.constData() + 12);
        qInfo().noquote() << "READ_REQUEST offset:" << offset
                          << "size:" << requestedSize;

        QFile file(m_filePath);
        if (m_filePath.isEmpty() || !file.open(QIODevice::ReadOnly)) {
            sendError(QStringLiteral("failed to open local file for READ_REQUEST"));
            return;
        }

        QByteArray data;
        if (offset < m_fileSize) {
            if (!file.seek(static_cast<qint64>(offset))) {
                sendError(QStringLiteral("failed to seek local file"));
                return;
            }
            data = file.read(static_cast<qint64>(requestedSize));
            if (data.size() < 0) {
                sendError(QStringLiteral("failed to read local file"));
                return;
            }
        }

        QByteArray response;
        vfd::appendLe32(response, seq);
        vfd::appendLe64(response, offset);
        vfd::appendLe32(response, static_cast<quint32>(data.size()));
        response.append(data);

        qInfo().noquote() << "READ_RESPONSE size:" << data.size();
        send(vfd::MsgType::ReadResponse, response);
    }

    void sendError(const QString &message)
    {
        qWarning().noquote() << message;
        send(vfd::MsgType::Error, message.toUtf8());
    }

    QTcpSocket m_socket;
    QByteArray m_buffer;
    QString m_filePath;
    quint64 m_fileSize = 0;
    int m_remoteWidth = 0;
    int m_remoteHeight = 0;
};

class DragWindow final : public QWidget {
    Q_OBJECT

public:
    explicit DragWindow(VfdClient *client)
        : m_client(client)
    {
        setAcceptDrops(true);
        setMinimumSize(320, 240);
        resize(800, 600);
        setWindowTitle(QStringLiteral("VFD Linux Client"));
    }

protected:
    void paintEvent(QPaintEvent *) override
    {
        QPainter painter(this);
        painter.fillRect(rect(), Qt::black);
    }

    void dragEnterEvent(QDragEnterEvent *event) override
    {
        const QString path = singleLocalFilePath(event->mimeData());
        if (path.isEmpty()) {
            event->ignore();
            return;
        }

        if (!m_client->sendFileDescriptor(path)) {
            event->ignore();
            return;
        }

        event->acceptProposedAction();
    }

    void dragMoveEvent(QDragMoveEvent *event) override
    {
        m_client->sendMove(mapToRemote(event->position().toPoint()));
        event->acceptProposedAction();
    }

    void dropEvent(QDropEvent *event) override
    {
        m_client->sendDrop(mapToRemote(event->position().toPoint()));
        event->acceptProposedAction();
    }

    void dragLeaveEvent(QDragLeaveEvent *event) override
    {
        m_client->sendCancel();
        QWidget::dragLeaveEvent(event);
    }

private:
    static QString singleLocalFilePath(const QMimeData *mimeData)
    {
        if (!mimeData || !mimeData->hasUrls() || mimeData->urls().size() != 1) {
            return {};
        }

        const QUrl url = mimeData->urls().first();
        if (!url.isLocalFile()) {
            return {};
        }

        const QFileInfo info(url.toLocalFile());
        if (!info.exists() || !info.isFile()) {
            return {};
        }

        return info.absoluteFilePath();
    }

    QPoint mapToRemote(const QPoint &localPos) const
    {
        if (!m_client->hasRemoteSize()) {
            return localPos;
        }

        const QSize remote = m_client->remoteSize();
        const int w = qMax(1, width());
        const int h = qMax(1, height());
        return {
            static_cast<int>((static_cast<qint64>(localPos.x()) * remote.width()) / w),
            static_cast<int>((static_cast<qint64>(localPos.y()) * remote.height()) / h),
        };
    }

    VfdClient *m_client = nullptr;
};

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);
    QCoreApplication::setApplicationName(QStringLiteral("vfd-linux-client"));

    QCommandLineParser parser;
    parser.setApplicationDescription(QStringLiteral("Linux client for virtual file drag prototype."));
    parser.addHelpOption();
    QCommandLineOption hostOption(QStringLiteral("host"),
                                  QStringLiteral("Windows agent host."),
                                  QStringLiteral("host"),
                                  QStringLiteral("127.0.0.1"));
    QCommandLineOption portOption(QStringLiteral("port"),
                                  QStringLiteral("Windows agent TCP port."),
                                  QStringLiteral("port"),
                                  QStringLiteral("45454"));
    parser.addOption(hostOption);
    parser.addOption(portOption);
    parser.process(app);

    bool ok = false;
    const quint16 port = parser.value(portOption).toUShort(&ok);
    if (!ok || port == 0) {
        qCritical().noquote() << "invalid --port value";
        return 2;
    }

    VfdClient client;
    client.connectToAgent(parser.value(hostOption), port);

    DragWindow window(&client);
    window.show();

    return app.exec();
}

#include "linux_client.moc"
