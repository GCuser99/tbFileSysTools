# tbFileSysTools

A modern replacement for Scripting Runtime's `FileSystemObject`, written in [twinBASIC](https://twinbasic.com).

It does everything FSO does, plus the things FSO has never been able to do: **full text-encoding support**, **files larger than 2 GB**, and **line-ending detection and normalization**.

```vb
' Read a UTF-8 file with a BOM, a UTF-16LE file, and a Shift-JIS file.
' You don't have to know which is which.
Dim text As String
text = FileSystemTools.TextFileToString("data.txt")     ' encoding auto-detected
```

\---

## Comparison with Scripting FSO

`Scripting.FileSystemObject` shipped in 1996 and shows it:

|Feature|FSO|tbFileSysTools|
|-|-|-|
|Text encodings|ASCII and UTF-16 only|Any Windows code page, with BOM handling and auto-detection|
|Files > 2 GB|Reads them as **empty**, silently|Streams them, in both directions|
|Appending to a > 2 GB file|Not possible|Works|
|Line endings|No support|Detect, preserve, normalize|
|Junction in a folder tree|`Folder.Size` double-counts, or recurses forever|Skipped|
|Unreadable folder|Reports it as empty|Raises|

Everything above is measured, not assumed — see [Verification](#verification).

\---

## Install

> \*\*TODO:\*\* installation instructions — package/`.twinpack` reference, or "add the `src/` modules to your project."

**Requires:** twinBASIC. *(TODO: minimum build number.)*

\---

## Two ways in

The library has one implementation and two doorways onto it.

### `FileSystemTools` — the standard module

Preferred for twinBASIC code. Call it directly; there's no object to create.

```vb
If FileSystemTools.FileExists(path) Then
    Debug.Print FileSystemTools.GetFile(path).Size
End If
```

### `FileSystemObject` — the COM class

A drop-in replacement for the Scripting Runtime object. Use it for late binding, for VBA/VBScript hosts that need an object, or when porting existing FSO code that you'd rather not rewrite.

```vb
Dim fso As New FileSystemObject          ' or CreateObject("tbFileSysTools.FileSystemObject")
Debug.Print fso.GetFile(path).Size
```

The class is a one-line delegation to the module for every member — same behaviour, same defaults. It uses FSO's argument names (`fileSpec`, `folderSpec`) so named arguments in ported code keep working.

> \*\*Note:\*\* the class is deliberately named `FileSystemObject`, the same as Scripting's. In a project referencing both, qualify it — `tbFileSysTools.FileSystemObject` — or drop the Scripting Runtime reference.

\---

## Text encodings

The headline feature. FSO can read ASCII and UTF-16. This reads anything Windows has a code page for.

```vb
' Auto-detect: BOM first, then UTF-16/32 and UTF-8 heuristics, then system ANSI.
Dim s As String
s = FileSystemTools.TextFileToString("mystery.txt")

' Or be explicit. A BOM in the file still wins.
s = FileSystemTools.TextFileToString("legacy.txt", encGB18030)

' Write with a BOM by choosing a BOM-bearing encoding.
FileSystemTools.StringToTextFile s, "out.txt", encUtf8Bom

' What encoding IS this file?
Debug.Print FileSystemTools.GetFileEncoding("mystery.txt")   ' e.g. encUtf16Bom
```

Supported: UTF-8, UTF-16 LE/BE, UTF-32 LE/BE, UTF-7, GB2312, GB18030, Big5, Latin-1, Latin-9, US-ASCII, system ANSI — each with and without a BOM (if applicable) — plus any other code page installed on the machine.

**Auto-detection on Read.** Defaults to auto-detect using a reliable heuristical algorithm but user can specify a code-page if known.

### Line endings

```vb
Debug.Print FileSystemTools.GetFileLineEnding("script.sh")   ' nlUnix

' Rewrite in place: UTF-8, CRLF, attributes preserved. Skips the write if the
' file is already in that form.
FileSystemTools.NormalizeTextFile "script.sh", encUtf8, nlWindows
```

Appending to an existing file adopts **that file's** newline style, so you can't accidentally turn a clean file into a mixed one.

\---

## Large files

FSO cannot read a file larger than 2 GB. It doesn't error — it returns an empty string. This library streams them.

```vb
Dim ts As TextStream
Set ts = FileSystemTools.OpenTextFile("huge.log", ForReading)
Do Until ts.AtEndOfStream
    ProcessLine ts.ReadLine()       ' no size limit
Loop
ts.Close

' Appending to a 4 GB file also works.
Set ts = FileSystemTools.OpenTextFile("huge.log", ForAppending)
ts.WriteLine "another line"
ts.Close
```

`ReadLine`, `Read`, `Write` and append are **unbounded**. `ReadAll` is capped at 2 GB by the size of a VB `String` and raises a clear error rather than misbehaving — use `ReadLine` for anything larger.

Multi-byte encodings are handled correctly across chunk boundaries, including surrogate pairs split by a 64 KB read. This is verified against 54 million lines of mixed 1-, 2-, 3- and 4-byte characters (see below).

\---

## Objects

`File`, `Folder`, `Drive` and their collections work as they do in FSO.

```vb
Dim f As Folder
Set f = FileSystemTools.GetFolder("C:\\Projects")

Debug.Print f.Size                          ' total bytes, subtree
For Each fl In f.Files
    Debug.Print fl.Name, fl.Size, fl.DateLastModified
Next
```

`File` and `Folder` objects are **live**: every property read re-stats the path, so values are never stale — and a deleted file raises rather than reporting fossils. The trade-off is that each property read costs a round trip, so hoist values out of tight loops.

`Folder.Files` and `Folder.SubFolders` return a **snapshot** of the membership. The objects inside are live; the list is not.

\---

## Deviations from FSO

Parity is the goal, but not at any price. Each of these was checked against the `Scripting.FileSystemObject`, and broken deliberately:

**`Folder.Size` does not follow directory reparse points.** FSO does. A junction pointing into its own subtree makes FSO double-count (measured: 2500 vs 1500), and a junction pointing at an ancestor makes it recurse until it dies. This library reports what physically lives in the tree.

**An unreadable folder raises.** FSO silently reports it as empty. A size or file list that quietly omits a subtree that can't be read is worse than an error.

**`FileAttribute.Volume` and `.Alias` are not provided.** `Volume` has no Win32 equivalent (use `Drive.VolumeName`). `Alias` is `FILE\_ATTRIBUTE\_REPARSE\_POINT` under a misleading name — and collides with VBA's `vbAlias` (64 vs 1024). Use `ReparsePoint`.

\---

## Beyond FSO...

|Member||
|-|-|
|`CreateTextFile`|Create any format - not just ascii/utf16|
|`OpenTextFile`|format auto-detection or user-specified|
|`GetFileEncoding`|Detect a file's encoding|
|`GetFileLineEnding`|Detect CRLF / LF / CR / mixed|
|`NormalizeTextFile`|Rewrite encoding + newlines in place, idempotently|
|`GetFilePaths`|Enumerate with a wildcard, recursion, hidden/system filters|
|`GetRelativePath`|Path from A to B|
|`CleanFileName`|Strip reserved characters and device names|
|`Rename`|In-place rename (FSO makes you assign to `.Name`)|
|`GetFileType`|Shell type description ("Text Document")|
|`GetCurrentDir` / `SetCurrentDir`||
|`GetSpecialFolder`|25 known folders, not FSO's 3|
|`GetStandardStream`|stdin / stdout / stderr as a `TextStream`|
|`ReadStream` / `WriteStream`|Raw bytes, Unicode-safe|
|`TextFileToString` / `StringToTextFile`|Whole-file text I/O|
|`TextFileToArray` / `ArrayToTextFile`|Whole-file line I/O|
|`File.Encoding`|Detect a file's encoding|
|`File.HasAttribute`|Determines is an attribute is set|
|`File.LineEnding`|Detect CRLF / LF / CR / mixed|
|`File.Normalize`|Rewrite encoding + newlines in place, idempotently|
|`File.OpenAsTextStream`|format auto-detection or user-specified|
|`File.SetAttribute`|Sets a single attribute|
|`File.ToStream`|Reads file to byte array|
|`File.ToString`|Reads file to string|
|`File.Version`|Gets the file version string|
|`Folder.HasAttribute`|Determines is an attribute is set|
|`Folder.SetAttribute`|Sets a single attribute|
|`TextStream.IsStreaming`|byte-streaming or buffered access|
|`TextStream.Encoding`|Detect a file's encoding|

\---

## Errors

Error numbers follow FSO, so existing handlers keep working:

|||
|-|-|
|**5**|Invalid argument, or content isn't decodable text|
|**52**|Bad file name|
|**53**|File not found|
|**58**|File already exists|
|**61**|Disk full|
|**68**|Device unavailable|
|**70**|Permission denied|
|**71**|Disk not ready|
|**76**|Path not found|

\---

## Verification

The large-file and encoding claims above are measured, not asserted. The probe modules are in the repo *(TODO: path)* and can be re-run.

|Claim|Evidence|
|-|-|
|Streams past 2 GB and 4 GB, read and write|4.5 GB file written and read back block-by-block; every block verified in place|
|Append to a > 2 GB file is correct and non-destructive|Head, size and tail all verified after appending to a 2.4 GB file|
|Multi-byte encodings survive chunk boundaries|54 M lines of UTF-16LE and UTF-8 containing 1-, 2-, 3- and 4-byte characters and surrogate pairs; line length chosen so chunk boundaries sweep every character position. Zero errors|
|Line/column tracking is exact|14 CR/LF edge cases, plus 35 M lines end to end|
|FSO parity|Differential tests against the real `Scripting.FileSystemObject`|

\---

## Structure

|||
|-|-|
|`FileSystemTools`|The API. All the logic lives here|
|`FileSystemObject`|COM-creatable thin wrapper over the above|
|`TextStream`|Streaming text reader/writer|
|`TextCodec`|Encoding detection, encode/decode, BOM handling|
|`Shared`|shared file/folder/drive procs|
|`WinAPI`|WinDevLib's Win32 declarations|
|`File`, `Folder`, `Drive`|Objects|
|`Files`, `Folders`, `Drives`|Collections|

\---

## License

MIT © 2026 GCUser99

