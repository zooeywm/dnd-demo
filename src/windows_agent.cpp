#include "vfd_protocol.h"

#include <QCommandLineParser>
#include <QCoreApplication>
#include <QHash>
#include <QMutex>
#include <QMutexLocker>
#include <QThread>
#include <QPoint>
#include <QTcpServer>
#include <QTcpSocket>
#include <QVector>
#include <QWaitCondition>

#include <atomic>
#include <memory>
#include <utility>

#include <objidl.h>
#include <ole2.h>
#include <shlobj.h>
#include <windows.h>

struct FileMetadata {
    QString name;
    quint64 size = 0;
    qint64 mtimeUnixMs = 0;
};

Q_DECLARE_METATYPE(FileMetadata)

static void sendMouseButton(bool down)
{
    INPUT input {};
    input.type = INPUT_MOUSE;
    input.mi.dwFlags = down ? MOUSEEVENTF_LEFTDOWN : MOUSEEVENTF_LEFTUP;
    const UINT sent = SendInput(1, &input, sizeof(INPUT));
    if (sent != 1) {
        qWarning().noquote() << "SendInput" << (down ? "LEFTDOWN" : "LEFTUP") << "failed:" << GetLastError();
    } else {
        qInfo().noquote() << "SendInput" << (down ? "LEFTDOWN" : "LEFTUP");
    }
}

static void moveMouseToVirtualDesktop(int x, int y)
{
    const int left = GetSystemMetrics(SM_XVIRTUALSCREEN);
    const int top = GetSystemMetrics(SM_YVIRTUALSCREEN);
    const int width = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    const int height = GetSystemMetrics(SM_CYVIRTUALSCREEN);
    if (width <= 1 || height <= 1) {
        return;
    }

    const int clampedX = qBound(left, left + x, left + width - 1);
    const int clampedY = qBound(top, top + y, top + height - 1);

    INPUT input {};
    input.type = INPUT_MOUSE;
    input.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK;
    input.mi.dx = static_cast<LONG>((static_cast<qint64>(clampedX - left) * 65535) / (width - 1));
    input.mi.dy = static_cast<LONG>((static_cast<qint64>(clampedY - top) * 65535) / (height - 1));
    const UINT sent = SendInput(1, &input, sizeof(INPUT));
    if (sent != 1) {
        qWarning().noquote() << "SendInput MOVE failed:" << GetLastError();
    } else {
        qInfo().noquote() << "SendInput MOVE x:" << x << "y:" << y;
    }
}

class DragState final {
public:
    enum Value {
        Active = 0,
        Drop = 1,
        Cancel = 2,
    };

    void reset()
    {
        m_value.store(Active, std::memory_order_release);
    }

    void setDrop()
    {
        qInfo().noquote() << "network received DROP";
        m_value.store(Drop, std::memory_order_release);
    }

    void setCancel()
    {
        qInfo().noquote() << "network received CANCEL";
        m_value.store(Cancel, std::memory_order_release);
    }

    Value value() const
    {
        return static_cast<Value>(m_value.load(std::memory_order_acquire));
    }

private:
    std::atomic<int> m_value {Active};
};

class NetworkWorker final : public QObject {
    Q_OBJECT

public:
    explicit NetworkWorker(DragState *dragState)
        : m_dragState(dragState)
    {
    }

public slots:
    void start(quint16 port)
    {
        m_server = new QTcpServer(this);
        connect(m_server, &QTcpServer::newConnection, this, &NetworkWorker::onNewConnection);
        if (!m_server->listen(QHostAddress::Any, port)) {
            qCritical().noquote() << "failed to listen on port" << port << ":" << m_server->errorString();
            QCoreApplication::exit(2);
            return;
        }

        m_remoteWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
        m_remoteHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);
        qInfo().noquote() << "Windows agent listening on port" << port
                          << "desktop:" << m_remoteWidth << "x" << m_remoteHeight;
    }

    void sendReadRequest(quint32 seq, quint64 offset, quint32 size)
    {
        if (!m_socket || m_socket->state() != QAbstractSocket::ConnectedState) {
            emit readFailed(seq, QStringLiteral("Linux client is not connected"));
            return;
        }

        QByteArray payload;
        vfd::appendLe32(payload, seq);
        vfd::appendLe64(payload, offset);
        vfd::appendLe32(payload, size);
        if (m_socket->write(vfd::makeMessage(vfd::MsgType::ReadRequest, payload)) < 0) {
            emit readFailed(seq, QStringLiteral("failed to send READ_REQUEST"));
        }
    }

signals:
    void fileDescriptorReceived(const FileMetadata &metadata, const QPoint &initialPos);
    void readResponseReceived(quint32 seq, quint64 offset, const QByteArray &data);
    void readFailed(quint32 seq, const QString &message);
    void allReadsFailed(const QString &message);

private slots:
    void onNewConnection()
    {
        if (m_socket) {
            QTcpSocket *extraSocket = m_server->nextPendingConnection();
            extraSocket->disconnectFromHost();
            extraSocket->deleteLater();
            qWarning().noquote() << "rejected extra Linux client connection";
            return;
        }

        m_socket = m_server->nextPendingConnection();
        m_socket->setParent(this);
        connect(m_socket, &QTcpSocket::readyRead, this, &NetworkWorker::onReadyRead);
        connect(m_socket, &QTcpSocket::disconnected, this, [this] {
            qWarning().noquote() << "Linux client disconnected";
            emit allReadsFailed(QStringLiteral("Linux client disconnected"));
            m_socket->deleteLater();
            m_socket = nullptr;
        });

        QByteArray payload;
        vfd::appendLe32(payload, static_cast<quint32>(m_remoteWidth));
        vfd::appendLe32(payload, static_cast<quint32>(m_remoteHeight));
        m_socket->write(vfd::makeMessage(vfd::MsgType::Hello, payload));
        qInfo().noquote() << "Linux client connected; sent HELLO";
    }

    void onReadyRead()
    {
        m_buffer.append(m_socket->readAll());

        while (true) {
            vfd::MsgType type = vfd::MsgType::Error;
            QByteArray payload;
            QString errorText;
            if (!vfd::tryTakeMessage(m_buffer, type, payload, &errorText)) {
                if (!errorText.isEmpty()) {
                    qWarning().noquote() << "protocol error:" << errorText;
                    m_socket->disconnectFromHost();
                }
                return;
            }

            switch (type) {
            case vfd::MsgType::FileDescriptor:
                handleFileDescriptor(payload);
                break;
            case vfd::MsgType::Move:
                handleMove(payload);
                break;
            case vfd::MsgType::Drop:
                handleDrop(payload);
                break;
            case vfd::MsgType::Cancel:
                handleCancel();
                break;
            case vfd::MsgType::ReadResponse:
                handleReadResponse(payload);
                break;
            case vfd::MsgType::Error:
                qWarning().noquote() << "ERROR from Linux client:" << QString::fromUtf8(payload);
                break;
            default:
                qWarning().noquote() << "unexpected message type from Linux client:"
                                     << static_cast<quint16>(type);
                break;
            }
        }
    }

private:
    void handleFileDescriptor(const QByteArray &payload)
    {
        if (payload.size() < 28) {
            qWarning().noquote() << "invalid FILE_DESCRIPTOR payload size:" << payload.size();
            return;
        }

        const char *data = payload.constData();
        FileMetadata metadata;
        metadata.size = vfd::readLe64(data);
        metadata.mtimeUnixMs = vfd::readLe64s(data + 8);
        const int remoteX = vfd::readLe32s(data + 16);
        const int remoteY = vfd::readLe32s(data + 20);
        const quint32 nameLen = vfd::readLe32(data + 24);
        if (payload.size() != 28 + static_cast<int>(nameLen)) {
            qWarning().noquote() << "invalid FILE_DESCRIPTOR name length:" << nameLen;
            return;
        }
        metadata.name = QString::fromUtf8(payload.constData() + 28, static_cast<int>(nameLen));
        qInfo().noquote() << "FILE_DESCRIPTOR name:" << metadata.name
                          << "size:" << metadata.size
                          << "mtime_unix_ms:" << metadata.mtimeUnixMs
                          << "initial x:" << remoteX
                          << "y:" << remoteY;
        m_dragState->reset();
        emit fileDescriptorReceived(metadata, QPoint(remoteX, remoteY));
    }

    void handleMove(const QByteArray &payload)
    {
        if (payload.size() != 8) {
            qWarning().noquote() << "invalid MOVE payload size:" << payload.size();
            return;
        }

        const int x = vfd::readLe32s(payload.constData());
        const int y = vfd::readLe32s(payload.constData() + 4);
        qInfo().noquote() << "MOVE x:" << x << "y:" << y;
        moveMouse(x, y);
    }

    void handleDrop(const QByteArray &payload)
    {
        if (payload.size() == 8) {
            const int x = vfd::readLe32s(payload.constData());
            const int y = vfd::readLe32s(payload.constData() + 4);
            qInfo().noquote() << "DROP x:" << x << "y:" << y;
            moveMouse(x, y);
        } else {
            qWarning().noquote() << "invalid DROP payload size:" << payload.size();
        }
        m_dragState->setDrop();
        sendMouseButton(false);
    }

    void handleCancel()
    {
        m_dragState->setCancel();
        sendMouseButton(false);
    }

    void handleReadResponse(const QByteArray &payload)
    {
        if (payload.size() < 16) {
            qWarning().noquote() << "invalid READ_RESPONSE payload size:" << payload.size();
            return;
        }

        const quint32 seq = vfd::readLe32(payload.constData());
        const quint64 offset = vfd::readLe64(payload.constData() + 4);
        const quint32 size = vfd::readLe32(payload.constData() + 12);
        if (payload.size() != 16 + static_cast<int>(size)) {
            emit readFailed(seq, QStringLiteral("invalid READ_RESPONSE size"));
            return;
        }

        qInfo().noquote() << "READ_RESPONSE seq:" << seq
                          << "offset:" << offset
                          << "size:" << size;
        emit readResponseReceived(seq, offset, payload.sliced(16, size));
    }

    void moveMouse(int x, int y) const
    {
        Q_UNUSED(m_remoteWidth);
        Q_UNUSED(m_remoteHeight);
        moveMouseToVirtualDesktop(x, y);
    }

    DragState *m_dragState = nullptr;
    QTcpServer *m_server = nullptr;
    QTcpSocket *m_socket = nullptr;
    QByteArray m_buffer;
    int m_remoteWidth = 0;
    int m_remoteHeight = 0;
};

class ReadBroker final : public QObject {
    Q_OBJECT

public:
    explicit ReadBroker(NetworkWorker *worker)
        : m_worker(worker)
    {
        connect(worker,
                &NetworkWorker::readResponseReceived,
                this,
                &ReadBroker::completeRead,
                Qt::DirectConnection);
        connect(worker,
                &NetworkWorker::readFailed,
                this,
                &ReadBroker::failRead,
                Qt::DirectConnection);
        connect(worker,
                &NetworkWorker::allReadsFailed,
                this,
                &ReadBroker::failAllReads,
                Qt::DirectConnection);
    }

    HRESULT read(quint64 offset, quint32 size, QByteArray &data)
    {
        const quint32 seq = m_nextSeq.fetch_add(1, std::memory_order_relaxed);

        auto pending = std::make_shared<PendingRead>();
        {
            QMutexLocker locker(&m_mutex);
            m_pending.insert(seq, pending);
        }

        QMetaObject::invokeMethod(m_worker,
                                  "sendReadRequest",
                                  Qt::QueuedConnection,
                                  Q_ARG(quint32, seq),
                                  Q_ARG(quint64, offset),
                                  Q_ARG(quint32, size));

        QMutexLocker pendingLocker(&pending->mutex);
        while (!pending->done) {
            pending->condition.wait(&pending->mutex);
        }

        {
            QMutexLocker locker(&m_mutex);
            m_pending.remove(seq);
        }

        if (!pending->error.isEmpty()) {
            qWarning().noquote() << "IStream::Read failed:" << pending->error;
            return STG_E_READFAULT;
        }

        data = pending->data;
        return S_OK;
    }

public slots:
    void completeRead(quint32 seq, quint64 offset, const QByteArray &data)
    {
        Q_UNUSED(offset);
        const std::shared_ptr<PendingRead> pending = takePending(seq, false);
        if (!pending) {
            qWarning().noquote() << "READ_RESPONSE for unknown seq:" << seq;
            return;
        }

        QMutexLocker locker(&pending->mutex);
        pending->data = data;
        pending->done = true;
        pending->condition.wakeAll();
    }

    void failRead(quint32 seq, const QString &message)
    {
        const std::shared_ptr<PendingRead> pending = takePending(seq, false);
        if (!pending) {
            return;
        }

        QMutexLocker locker(&pending->mutex);
        pending->error = message;
        pending->done = true;
        pending->condition.wakeAll();
    }

    void failAllReads(const QString &message)
    {
        QList<std::shared_ptr<PendingRead>> pendingReads;
        {
            QMutexLocker locker(&m_mutex);
            pendingReads = m_pending.values();
        }

        for (const std::shared_ptr<PendingRead> &pending : pendingReads) {
            QMutexLocker locker(&pending->mutex);
            if (!pending->done) {
                pending->error = message;
                pending->done = true;
                pending->condition.wakeAll();
            }
        }
    }

private:
    struct PendingRead {
        QMutex mutex;
        QWaitCondition condition;
        bool done = false;
        QByteArray data;
        QString error;
    };

    std::shared_ptr<PendingRead> takePending(quint32 seq, bool remove)
    {
        QMutexLocker locker(&m_mutex);
        const auto pending = m_pending.value(seq);
        if (remove) {
            m_pending.remove(seq);
        }
        return pending;
    }

    NetworkWorker *m_worker = nullptr;
    std::atomic<quint32> m_nextSeq {1};
    QMutex m_mutex;
    QHash<quint32, std::shared_ptr<PendingRead>> m_pending;
};

static FILETIME unixMsToFileTime(qint64 unixMs)
{
    constexpr quint64 fileTimeUnixEpoch = 116444736000000000ULL;
    const quint64 value = fileTimeUnixEpoch + static_cast<quint64>(unixMs) * 10000ULL;
    FILETIME fileTime {};
    fileTime.dwLowDateTime = static_cast<DWORD>(value & 0xffffffffULL);
    fileTime.dwHighDateTime = static_cast<DWORD>(value >> 32);
    return fileTime;
}

class RemoteFileStream final : public IStream {
public:
    RemoteFileStream(FileMetadata metadata, ReadBroker *readBroker)
        : m_metadata(std::move(metadata))
        , m_readBroker(readBroker)
    {
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **object) override
    {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_ISequentialStream || riid == IID_IStream) {
            *object = static_cast<IStream *>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++m_refCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const ULONG ref = --m_refCount;
        if (ref == 0) {
            delete this;
        }
        return ref;
    }

    HRESULT STDMETHODCALLTYPE Read(void *pv, ULONG cb, ULONG *pcbRead) override
    {
        if (!pv) {
            return STG_E_INVALIDPOINTER;
        }
        if (pcbRead) {
            *pcbRead = 0;
        }
        if (cb == 0) {
            return S_OK;
        }
        if (m_position >= m_metadata.size) {
            qInfo().noquote() << "IStream::Read offset:" << m_position
                              << "size:" << cb
                              << "result: EOF";
            return S_FALSE;
        }

        const quint32 requestSize = static_cast<quint32>(qMin<quint64>(cb, m_metadata.size - m_position));
        QByteArray data;
        const HRESULT hr = m_readBroker->read(m_position, requestSize, data);
        if (FAILED(hr)) {
            qInfo().noquote() << "IStream::Read offset:" << m_position
                              << "size:" << cb
                              << "result:" << Qt::hex << hr << Qt::dec;
            return hr;
        }

        const qsizetype copySize = qMin(data.size(), static_cast<qsizetype>(cb));
        memcpy(pv, data.constData(), static_cast<size_t>(copySize));
        m_position += static_cast<quint64>(copySize);
        if (pcbRead) {
            *pcbRead = static_cast<ULONG>(copySize);
        }

        const HRESULT result = copySize == static_cast<qsizetype>(cb) ? S_OK : S_FALSE;
        qInfo().noquote() << "IStream::Read offset:" << (m_position - static_cast<quint64>(copySize))
                          << "size:" << cb
                          << "result:" << (result == S_OK ? "S_OK" : "S_FALSE");
        return result;
    }

    HRESULT STDMETHODCALLTYPE Write(const void *, ULONG, ULONG *) override
    {
        return STG_E_ACCESSDENIED;
    }

    HRESULT STDMETHODCALLTYPE Seek(LARGE_INTEGER dlibMove,
                                   DWORD dwOrigin,
                                   ULARGE_INTEGER *plibNewPosition) override
    {
        qint64 base = 0;
        switch (dwOrigin) {
        case STREAM_SEEK_SET:
            base = 0;
            break;
        case STREAM_SEEK_CUR:
            base = static_cast<qint64>(m_position);
            break;
        case STREAM_SEEK_END:
            base = static_cast<qint64>(m_metadata.size);
            break;
        default:
            return STG_E_INVALIDFUNCTION;
        }

        const qint64 next = base + dlibMove.QuadPart;
        if (next < 0) {
            return STG_E_INVALIDFUNCTION;
        }
        m_position = static_cast<quint64>(next);
        if (plibNewPosition) {
            plibNewPosition->QuadPart = m_position;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE SetSize(ULARGE_INTEGER) override
    {
        return STG_E_ACCESSDENIED;
    }

    HRESULT STDMETHODCALLTYPE CopyTo(IStream *, ULARGE_INTEGER, ULARGE_INTEGER *, ULARGE_INTEGER *) override
    {
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE Commit(DWORD) override
    {
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Revert() override
    {
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE LockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override
    {
        return STG_E_INVALIDFUNCTION;
    }

    HRESULT STDMETHODCALLTYPE UnlockRegion(ULARGE_INTEGER, ULARGE_INTEGER, DWORD) override
    {
        return STG_E_INVALIDFUNCTION;
    }

    HRESULT STDMETHODCALLTYPE Stat(STATSTG *pstatstg, DWORD) override
    {
        if (!pstatstg) {
            return STG_E_INVALIDPOINTER;
        }
        memset(pstatstg, 0, sizeof(STATSTG));
        pstatstg->type = STGTY_STREAM;
        pstatstg->cbSize.QuadPart = m_metadata.size;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Clone(IStream **) override
    {
        return E_NOTIMPL;
    }

private:
    std::atomic<ULONG> m_refCount {1};
    FileMetadata m_metadata;
    ReadBroker *m_readBroker = nullptr;
    quint64 m_position = 0;
};

class FormatEtcEnumerator final : public IEnumFORMATETC {
public:
    explicit FormatEtcEnumerator(QVector<FORMATETC> formats)
        : m_formats(std::move(formats))
    {
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **object) override
    {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IEnumFORMATETC) {
            *object = static_cast<IEnumFORMATETC *>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++m_refCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const ULONG ref = --m_refCount;
        if (ref == 0) {
            delete this;
        }
        return ref;
    }

    HRESULT STDMETHODCALLTYPE Next(ULONG celt, FORMATETC *rgelt, ULONG *pceltFetched) override
    {
        if (!rgelt) {
            return E_POINTER;
        }

        ULONG fetched = 0;
        while (fetched < celt && m_index < m_formats.size()) {
            rgelt[fetched++] = m_formats[m_index++];
        }
        if (pceltFetched) {
            *pceltFetched = fetched;
        }
        return fetched == celt ? S_OK : S_FALSE;
    }

    HRESULT STDMETHODCALLTYPE Skip(ULONG celt) override
    {
        m_index = qMin<int>(m_index + static_cast<int>(celt), m_formats.size());
        return m_index < m_formats.size() ? S_OK : S_FALSE;
    }

    HRESULT STDMETHODCALLTYPE Reset() override
    {
        m_index = 0;
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE Clone(IEnumFORMATETC **ppenum) override
    {
        if (!ppenum) {
            return E_POINTER;
        }
        auto *clone = new FormatEtcEnumerator(m_formats);
        clone->m_index = m_index;
        *ppenum = clone;
        return S_OK;
    }

private:
    std::atomic<ULONG> m_refCount {1};
    QVector<FORMATETC> m_formats;
    int m_index = 0;
};

class VirtualFileDataObject final : public IDataObject {
public:
    VirtualFileDataObject(FileMetadata metadata, ReadBroker *readBroker)
        : m_metadata(std::move(metadata))
        , m_readBroker(readBroker)
    {
        m_fileDescriptorFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormatW(CFSTR_FILEDESCRIPTORW));
        m_fileContentsFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormatW(CFSTR_FILECONTENTS));
        m_preferredDropEffectFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormatW(CFSTR_PREFERREDDROPEFFECT));
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **object) override
    {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IDataObject) {
            *object = static_cast<IDataObject *>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++m_refCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const ULONG ref = --m_refCount;
        if (ref == 0) {
            delete this;
        }
        return ref;
    }

    HRESULT STDMETHODCALLTYPE GetData(FORMATETC *pformatetcIn, STGMEDIUM *pmedium) override
    {
        if (!pformatetcIn || !pmedium) {
            return E_POINTER;
        }
        memset(pmedium, 0, sizeof(STGMEDIUM));

        logFormat(QStringLiteral("GetData"), *pformatetcIn);
        if (isFileDescriptorRequest(*pformatetcIn)) {
            qInfo().noquote() << "GetData request: FILEDESCRIPTOR";
            return getFileDescriptor(pmedium);
        }
        if (isFileContentsRequest(*pformatetcIn)) {
            qInfo().noquote() << "GetData request: FILECONTENTS";
            pmedium->tymed = TYMED_ISTREAM;
            pmedium->pstm = new RemoteFileStream(m_metadata, m_readBroker);
            pmedium->pUnkForRelease = nullptr;
            return S_OK;
        }
        if (isPreferredDropEffectRequest(*pformatetcIn)) {
            qInfo().noquote() << "GetData request: PREFERREDDROPEFFECT";
            return getDropEffect(pmedium, DROPEFFECT_COPY);
        }

        return DV_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE GetDataHere(FORMATETC *, STGMEDIUM *) override
    {
        return DATA_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE QueryGetData(FORMATETC *pformatetc) override
    {
        if (!pformatetc) {
            return E_POINTER;
        }
        logFormat(QStringLiteral("QueryGetData"), *pformatetc);
        return isFileDescriptorRequest(*pformatetc)
                || isFileContentsRequest(*pformatetc)
                || isPreferredDropEffectRequest(*pformatetc)
            ? S_OK
            : DV_E_FORMATETC;
    }

    HRESULT STDMETHODCALLTYPE GetCanonicalFormatEtc(FORMATETC *, FORMATETC *pformatetcOut) override
    {
        if (pformatetcOut) {
            pformatetcOut->ptd = nullptr;
        }
        return DATA_S_SAMEFORMATETC;
    }

    HRESULT STDMETHODCALLTYPE SetData(FORMATETC *, STGMEDIUM *, BOOL) override
    {
        return E_NOTIMPL;
    }

    HRESULT STDMETHODCALLTYPE EnumFormatEtc(DWORD dwDirection, IEnumFORMATETC **ppenumFormatEtc) override
    {
        if (!ppenumFormatEtc) {
            return E_POINTER;
        }
        if (dwDirection != DATADIR_GET) {
            return E_NOTIMPL;
        }

        QVector<FORMATETC> formats;
        FORMATETC descriptor {};
        descriptor.cfFormat = m_fileDescriptorFormat;
        descriptor.dwAspect = DVASPECT_CONTENT;
        descriptor.lindex = -1;
        descriptor.tymed = TYMED_HGLOBAL;
        formats.push_back(descriptor);

        FORMATETC contents {};
        contents.cfFormat = m_fileContentsFormat;
        contents.dwAspect = DVASPECT_CONTENT;
        contents.lindex = 0;
        contents.tymed = TYMED_ISTREAM;
        formats.push_back(contents);

        FORMATETC preferred {};
        preferred.cfFormat = m_preferredDropEffectFormat;
        preferred.dwAspect = DVASPECT_CONTENT;
        preferred.lindex = -1;
        preferred.tymed = TYMED_HGLOBAL;
        formats.push_back(preferred);

        *ppenumFormatEtc = new FormatEtcEnumerator(formats);
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE DAdvise(FORMATETC *, DWORD, IAdviseSink *, DWORD *) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }

    HRESULT STDMETHODCALLTYPE DUnadvise(DWORD) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }

    HRESULT STDMETHODCALLTYPE EnumDAdvise(IEnumSTATDATA **) override
    {
        return OLE_E_ADVISENOTSUPPORTED;
    }

private:
    bool isFileDescriptorRequest(const FORMATETC &format) const
    {
        return format.cfFormat == m_fileDescriptorFormat
            && (format.tymed & TYMED_HGLOBAL)
            && format.dwAspect == DVASPECT_CONTENT
            && format.lindex == -1;
    }

    bool isFileContentsRequest(const FORMATETC &format) const
    {
        return format.cfFormat == m_fileContentsFormat
            && (format.tymed & TYMED_ISTREAM)
            && format.dwAspect == DVASPECT_CONTENT
            && format.lindex == 0;
    }

    bool isPreferredDropEffectRequest(const FORMATETC &format) const
    {
        return format.cfFormat == m_preferredDropEffectFormat
            && (format.tymed & TYMED_HGLOBAL)
            && format.dwAspect == DVASPECT_CONTENT;
    }

    HRESULT getDropEffect(STGMEDIUM *medium, DWORD effect) const
    {
        HGLOBAL global = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, sizeof(DWORD));
        if (!global) {
            return STG_E_MEDIUMFULL;
        }
        auto *value = static_cast<DWORD *>(GlobalLock(global));
        if (!value) {
            GlobalFree(global);
            return STG_E_MEDIUMFULL;
        }
        *value = effect;
        GlobalUnlock(global);
        medium->tymed = TYMED_HGLOBAL;
        medium->hGlobal = global;
        medium->pUnkForRelease = nullptr;
        return S_OK;
    }

    HRESULT getFileDescriptor(STGMEDIUM *medium) const
    {
        const SIZE_T size = sizeof(FILEGROUPDESCRIPTORW);
        HGLOBAL global = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, size);
        if (!global) {
            return STG_E_MEDIUMFULL;
        }

        auto *group = static_cast<FILEGROUPDESCRIPTORW *>(GlobalLock(global));
        if (!group) {
            GlobalFree(global);
            return STG_E_MEDIUMFULL;
        }

        group->cItems = 1;
        FILEDESCRIPTORW &descriptor = group->fgd[0];
        descriptor.dwFlags = FD_FILESIZE | FD_WRITESTIME | FD_ATTRIBUTES;
        descriptor.nFileSizeHigh = static_cast<DWORD>(m_metadata.size >> 32);
        descriptor.nFileSizeLow = static_cast<DWORD>(m_metadata.size & 0xffffffffULL);
        descriptor.ftLastWriteTime = unixMsToFileTime(m_metadata.mtimeUnixMs);
        descriptor.dwFileAttributes = FILE_ATTRIBUTE_NORMAL;

        const std::wstring fileName = m_metadata.name.toStdWString();
        wcsncpy_s(descriptor.cFileName, MAX_PATH, fileName.c_str(), _TRUNCATE);

        GlobalUnlock(global);
        medium->tymed = TYMED_HGLOBAL;
        medium->hGlobal = global;
        medium->pUnkForRelease = nullptr;
        return S_OK;
    }

    void logFormat(const QString &prefix, const FORMATETC &format) const
    {
        qInfo().noquote() << prefix
                          << "format:" << formatName(format.cfFormat)
                          << "cfFormat:" << format.cfFormat
                          << "tymed:" << format.tymed
                          << "lindex:" << format.lindex;
    }

    QString formatName(CLIPFORMAT cfFormat) const
    {
        if (cfFormat == m_fileDescriptorFormat) {
            return QStringLiteral("CFSTR_FILEDESCRIPTORW");
        }
        if (cfFormat == m_fileContentsFormat) {
            return QStringLiteral("CFSTR_FILECONTENTS");
        }
        if (cfFormat == m_preferredDropEffectFormat) {
            return QStringLiteral("CFSTR_PREFERREDDROPEFFECT");
        }
        if (cfFormat == CF_HDROP) {
            return QStringLiteral("CF_HDROP");
        }
        return QStringLiteral("unknown");
    }

    std::atomic<ULONG> m_refCount {1};
    FileMetadata m_metadata;
    ReadBroker *m_readBroker = nullptr;
    CLIPFORMAT m_fileDescriptorFormat = 0;
    CLIPFORMAT m_fileContentsFormat = 0;
    CLIPFORMAT m_preferredDropEffectFormat = 0;
};

class DropSource final : public IDropSource {
public:
    explicit DropSource(DragState *state)
        : m_state(state)
    {
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void **object) override
    {
        if (!object) {
            return E_POINTER;
        }
        if (riid == IID_IUnknown || riid == IID_IDropSource) {
            *object = static_cast<IDropSource *>(this);
            AddRef();
            return S_OK;
        }
        *object = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++m_refCount;
    }

    ULONG STDMETHODCALLTYPE Release() override
    {
        const ULONG ref = --m_refCount;
        if (ref == 0) {
            delete this;
        }
        return ref;
    }

    HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL escapePressed, DWORD keyState) override
    {
        if (escapePressed) {
            qInfo().noquote() << "DropSource received ESC";
            return DRAGDROP_S_CANCEL;
        }
        if (keyState != m_lastKeyState) {
            qInfo().noquote() << "DropSource keyState:" << keyState;
            m_lastKeyState = keyState;
        }
        switch (m_state->value()) {
        case DragState::Cancel:
            qInfo().noquote() << "DropSource received CANCEL";
            return DRAGDROP_S_CANCEL;
        case DragState::Drop:
            qInfo().noquote() << "DropSource received DROP";
            return DRAGDROP_S_DROP;
        case DragState::Active:
            if (!m_loggedActive) {
                qInfo().noquote() << "DropSource active";
                m_loggedActive = true;
            }
            return S_OK;
        }
        return S_OK;
    }

    HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD) override
    {
        return DRAGDROP_S_USEDEFAULTCURSORS;
    }

private:
    std::atomic<ULONG> m_refCount {1};
    DragState *m_state = nullptr;
    bool m_loggedActive = false;
    DWORD m_lastKeyState = 0xffffffff;
};

class OleDragController final : public QObject {
    Q_OBJECT

public:
    OleDragController(DragState *dragState, ReadBroker *readBroker)
        : m_dragState(dragState)
        , m_readBroker(readBroker)
    {
    }

public slots:
    void startDrag(const FileMetadata &metadata, const QPoint &initialPos)
    {
        IDataObject *dataObject = new VirtualFileDataObject(metadata, m_readBroker);
        IDropSource *dropSource = new DropSource(m_dragState);

        POINT cursor {};
        GetCursorPos(&cursor);
        qInfo().noquote() << "starting DoDragDrop for" << metadata.name
                          << "initial x:" << initialPos.x()
                          << "y:" << initialPos.y()
                          << "current cursor x:" << cursor.x
                          << "y:" << cursor.y;

        moveMouseToVirtualDesktop(initialPos.x(), initialPos.y());
        sendMouseButton(true);

        DWORD effect = DROPEFFECT_NONE;
        const HRESULT hr = DoDragDrop(dataObject, dropSource, DROPEFFECT_COPY, &effect);
        qInfo().noquote() << "DoDragDrop returned hr:" << Qt::hex << hr << Qt::dec
                          << "effect:" << effect;

        sendMouseButton(false);

        dropSource->Release();
        dataObject->Release();
    }

private:
    DragState *m_dragState = nullptr;
    ReadBroker *m_readBroker = nullptr;
};

int main(int argc, char *argv[])
{
    QCoreApplication app(argc, argv);
    QCoreApplication::setApplicationName(QStringLiteral("vfd-windows-agent"));
    qRegisterMetaType<FileMetadata>("FileMetadata");

    QCommandLineParser parser;
    parser.setApplicationDescription(QStringLiteral("Windows user-session OLE virtual file drag agent."));
    parser.addHelpOption();
    QCommandLineOption portOption(QStringLiteral("listen-port"),
                                  QStringLiteral("TCP listen port."),
                                  QStringLiteral("port"),
                                  QStringLiteral("45454"));
    parser.addOption(portOption);
    parser.process(app);

    bool ok = false;
    const quint16 port = parser.value(portOption).toUShort(&ok);
    if (!ok || port == 0) {
        qCritical().noquote() << "invalid --listen-port value";
        return 2;
    }

    HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) {
        qCritical().noquote() << "CoInitializeEx failed:" << Qt::hex << hr << Qt::dec;
        return 3;
    }
    hr = OleInitialize(nullptr);
    if (FAILED(hr)) {
        qCritical().noquote() << "OleInitialize failed:" << Qt::hex << hr << Qt::dec;
        CoUninitialize();
        return 3;
    }

    DragState dragState;
    QThread networkThread;
    auto *worker = new NetworkWorker(&dragState);
    worker->moveToThread(&networkThread);
    QObject::connect(&networkThread, &QThread::finished, worker, &QObject::deleteLater);
    networkThread.start();
    QMetaObject::invokeMethod(worker, "start", Qt::QueuedConnection, Q_ARG(quint16, port));

    ReadBroker readBroker(worker);
    OleDragController dragController(&dragState, &readBroker);
    QObject::connect(worker,
                     &NetworkWorker::fileDescriptorReceived,
                     &dragController,
                     &OleDragController::startDrag,
                     Qt::QueuedConnection);

    const int result = app.exec();

    networkThread.quit();
    networkThread.wait();
    OleUninitialize();
    CoUninitialize();
    return result;
}

#include "windows_agent.moc"
