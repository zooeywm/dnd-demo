# dnd-demo

Qt prototype for Linux to Windows virtual file drag and drop.

The Linux client accepts a single local file dragged into a black Qt window. It sends only file metadata and pointer events to the Windows user-session agent. File bytes are transferred lazily only when the Windows drop target reads the `IStream` returned for `CFSTR_FILECONTENTS`.

## Targets

- `vfd-linux-client`: Linux Qt client window and file-content responder.
- `vfd-windows-agent`: Windows current-user session agent that exposes a virtual file through OLE drag and drop.

The Windows agent must run in the active desktop user session. It is not designed to run as a Session 0 service.

## Protocol

The two processes communicate over TCP with a little-endian binary protocol.

Message header:

```cpp
struct MsgHeader {
    uint32_t magic;   // "VFD1"
    uint16_t type;
    uint16_t flags;
    uint32_t length;
};
```

Implemented message types:

- `HELLO`
- `FILE_DESCRIPTOR`
- `MOVE`
- `DROP`
- `CANCEL`
- `READ_REQUEST`
- `READ_RESPONSE`
- `ERROR`
- `OK`

`READ_REQUEST` and `READ_RESPONSE` include a `seq` field so blocking stream reads can wait for the matching response without depending on packet order.

## Windows Virtual File Data

The Windows agent implements:

- `IDataObject`
- `IDataObjectAsyncCapability`
- `IDropSource`
- `IStream`

Supported clipboard formats:

- `CFSTR_FILEDESCRIPTORW` with `TYMED_HGLOBAL`
- `CFSTR_FILECONTENTS` with `TYMED_ISTREAM`
- `CFSTR_PREFERREDDROPEFFECT` with `DROPEFFECT_COPY`

The prototype intentionally does not use `CF_HDROP`, `WM_DROPFILES`, Windows temporary files, or pre-created complete files on the Windows side.

## Build

Requirements:

- CMake 3.16 or newer
- Qt 6.5 or newer
- C++17 compiler

Linux client:

```bash
cmake -S . -B build
cmake --build build --target vfd-linux-client
```

Windows agent:

```bat
cmake -S . -B build
cmake --build build --target vfd-windows-agent
```

## Run

Start the Windows agent first:

```bat
vfd-windows-agent --listen-port 45454
```

Start the Linux client and connect it to the Windows host:

```bash
vfd-linux-client --host <windows-ip> --port 45454
```

Drag one local file into the Linux client window, move over the target area, then drop. The Windows side exposes a virtual file. The Linux side starts reading and sending file bytes only after the target application requests `CFSTR_FILECONTENTS` and reads from the returned stream.

## Current Scope

- Single file only.
- No directory support.
- No multi-file support.
- No resume support.
- Intended test targets: Explorer and browsers.
- This is a prototype and logs OLE format queries, stream reads, drag state changes, pointer movement, and read requests.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE).
