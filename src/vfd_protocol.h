#pragma once

#include <QByteArray>
#include <QDataStream>
#include <QtEndian>
#include <QtGlobal>

namespace vfd {

constexpr quint32 Magic = 0x31444656; // "VFD1" on little-endian wire order.
constexpr qsizetype HeaderSize = 12;

enum class MsgType : quint16 {
    Hello = 1,
    FileDescriptor = 2,
    Move = 3,
    Drop = 4,
    Cancel = 5,
    ReadRequest = 6,
    ReadResponse = 7,
    Error = 8,
    Ok = 9,
};

inline void appendLe16(QByteArray &out, quint16 value)
{
    const qsizetype oldSize = out.size();
    out.resize(oldSize + 2);
    qToLittleEndian(value, reinterpret_cast<uchar *>(out.data() + oldSize));
}

inline void appendLe32(QByteArray &out, quint32 value)
{
    const qsizetype oldSize = out.size();
    out.resize(oldSize + 4);
    qToLittleEndian(value, reinterpret_cast<uchar *>(out.data() + oldSize));
}

inline void appendLe64(QByteArray &out, quint64 value)
{
    const qsizetype oldSize = out.size();
    out.resize(oldSize + 8);
    qToLittleEndian(value, reinterpret_cast<uchar *>(out.data() + oldSize));
}

inline void appendLe32s(QByteArray &out, qint32 value)
{
    appendLe32(out, static_cast<quint32>(value));
}

inline void appendLe64s(QByteArray &out, qint64 value)
{
    appendLe64(out, static_cast<quint64>(value));
}

inline quint16 readLe16(const char *data)
{
    return qFromLittleEndian<quint16>(reinterpret_cast<const uchar *>(data));
}

inline quint32 readLe32(const char *data)
{
    return qFromLittleEndian<quint32>(reinterpret_cast<const uchar *>(data));
}

inline quint64 readLe64(const char *data)
{
    return qFromLittleEndian<quint64>(reinterpret_cast<const uchar *>(data));
}

inline qint32 readLe32s(const char *data)
{
    return static_cast<qint32>(readLe32(data));
}

inline qint64 readLe64s(const char *data)
{
    return static_cast<qint64>(readLe64(data));
}

inline QByteArray makeMessage(MsgType type, const QByteArray &payload = {}, quint16 flags = 0)
{
    QByteArray out;
    out.reserve(HeaderSize + payload.size());
    appendLe32(out, Magic);
    appendLe16(out, static_cast<quint16>(type));
    appendLe16(out, flags);
    appendLe32(out, static_cast<quint32>(payload.size()));
    out.append(payload);
    return out;
}

inline bool tryTakeMessage(QByteArray &buffer,
                           MsgType &type,
                           QByteArray &payload,
                           QString *errorText = nullptr)
{
    if (buffer.size() < HeaderSize) {
        return false;
    }

    const char *header = buffer.constData();
    const quint32 magic = readLe32(header);
    if (magic != Magic) {
        if (errorText) {
            *errorText = QStringLiteral("invalid message magic");
        }
        buffer.clear();
        return false;
    }

    const quint32 length = readLe32(header + 8);
    if (length > 64 * 1024 * 1024) {
        if (errorText) {
            *errorText = QStringLiteral("message too large");
        }
        buffer.clear();
        return false;
    }

    const qsizetype totalSize = HeaderSize + static_cast<qsizetype>(length);
    if (buffer.size() < totalSize) {
        return false;
    }

    type = static_cast<MsgType>(readLe16(header + 4));
    payload = buffer.sliced(HeaderSize, length);
    buffer.remove(0, totalSize);
    return true;
}

} // namespace vfd
