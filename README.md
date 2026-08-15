
![vbaXray logo](logo_vbaxray.png)
# vbaXray

![License: MIT](https://img.shields.io/badge/License-MIT-darkgreen.svg)
![Platform: Windows](https://img.shields.io/badge/Platform-Windows-blue.svg)
![VBA](https://img.shields.io/badge/VBA-32bit%20%7C%2064bit-purple.svg)
![Dependencies](https://img.shields.io/badge/Dependencies-none%20(mostly)-teal.svg)

vbaXray is a single-class VBA module for reading VBA source code straight out of Office files - Excel, Word, PowerPoint, Publisher, Outlook, and Access - without opening them.

## How it works

Office macro-enabled files (eg: `.xlsm`, `.docm`, etc.) are ZIP archives. Inside, `xl/vbaProject.bin` (or `word/vbaProject.bin`, `ppt/vbaProject.bin`) is a "Compound Document" file — the same binary format as the older `.xls` / `.doc` files.

Depending on the file type, vbaXray then:

1. Uses the Windows ZIP folder handler (`ZipFldr.dll`) and reads `vbaProject.bin` straight out of the archive as an `IStorage`. Nothing touches the disk. No temporary copies, no polling loops, no `Shell.Application` waiting impatiently for the file system to catch up.
2. For legacy compound files (`.xls`, `.doc`, `.pub`, `.otm`), opens the file directly with `StgOpenStorage` and navigates to the VBA storage using the host-specific path (`_VBA_PROJECT_CUR\VBA` for Excel, `Macros\VBA` for Word, etc.).
3. For legacy PowerPoint (`.ppt`), scans the `PowerPoint Document` stream for `VbaProjectStg` records. These are sometimes raw CFB, sometimes zlib-compressed — the compressed ones get wrapped in a gzip shell and fed through `archiveint.dll` because PowerPoint's DEFLATE payloads end with a sync flush rather than a conventional terminator, which is a delightful thing to discover empirically (I'm lying...)
4. For Access (`.accdb`, `.mdb`), there is no `vbaProject.bin` at all. Access shreds the VBA across LVAL database pages. vbaXray walks the pages, follows the row chains, decompresses each candidate blob, and keeps whatever comes out looking module-ish. The only reason I know any of this is thanks to WilliamSmithEdward's pyOpenVBA, which is a fantastic resource for anyone ~~unhinged enough~~ interested in Access internals.
5. Once the VBA storage is open, reads the `VBA/dir` stream and decompresses it using the MS-OVBA LZ77 variant (see MS-OVBA 2.4.1).
6. Parses the `dir` stream records to extract each module's name, stream name, and compressed-source offset.
7. Reads the `PROJECT` stream (plain text) to refine module types — standard, class, document, designer.
8. For each module, reads the corresponding VBA stream, slices from the stored offset, and decompresses the source.

All COM interfaces (`IStorage`, `IStream`, `IEnumSTATSTG`) use `DispCallFunc` vtable dispatch ... because I'm into that sort of thing...

## Basic Usage

### Reading a project

```vba
Dim xray As New clsVBAXray

If xray.LoadFromFile("C:\Projects\MyMacroFile.xlsm") Then
  Debug.Print "Project:  " & xray.ProjectName
  Debug.Print "Modules:  " & xray.ModuleCount

  Dim i As Long
  For i = 1 To xray.ModuleCount
    Debug.Print i, xray.ModuleName(i)
  Next i
Else
  Debug.Print "Load failed: " & xray.LastError
End If
```

### Extracting source code

```vba
Dim xray As New clsVBAXray

If xray.LoadFromFile("C:\Projects\MyMacroFile.xlsm") Then
  Debug.Print xray.SourceCode(1)                        ' one module (by index)
  Debug.Print xray.FullSourceCode("--- {0} ---")        ' the lot
  xray.ExportAll "C:\Export\MyProject\"                 ' to disk, sorted by type
  xray.ExportModuleByName "Module1", "C:\Export\"       ' just the one
End If
```

### Poking about

```vba
Dim xray As New clsVBAXray
xray.LogLevel = xrayDebug                               ' show the raw hresults

If xray.LoadFromFile("C:\Suspicious\thing.xls") Then
  xray.DebugDumpStorageTree                             ' whole tree, to the Immediate window
  xray.DebugDumpStorageTree "C:\Temp\tree.txt"          ' or to a file
End If
```

> [!CAUTION]
> Legal note: this is covered off in the MIT License (see below/attached/to the side/over there), but it is worth reiterating: vbaXray is provided entirely "as is". No warranty, express or implied, is given. If you use this in production, you do so at your own risk and with my deepest sympathy.

## What it reads

There are two different meanings of "supported" used here: one is the truth, and the other is a guess.

**Implemented** means that vbaXray contains a code path intended to handle the format. **Tested** means that I have actually had a real file of that type through it and inspected the result. Those are not the same thing.

| Family | Extensions | Implementation | Testing |
| :--- | :--- | :--- | :--- |
| **OOXML** | `.xlsm` `.xlam` `.xlsb` `.xltm` `.docm` `.dotm` `.pptm` `.potm` `.ppsm` `.ppam` `.sldm` | Implemented | Main formats tested; not every variant |
| **Legacy Excel** | `.xls` `.xla` `.xlt` | Implemented | Tested |
| **Legacy Word** | `.doc` `.dot` | Implemented | Tested |
| **Publisher / Outlook** | `.pub` `.otm` | Implemented | Tested |
| **Legacy PowerPoint** | `.ppt` `.pps` `.pot` `.ppa` | Implemented | Partially tested; `.pps`, `.pot` and `.ppa` are untested - in fact, I have never encountered any |
| **Access** | `.accdb` and `.mdb` | Implemented for the modern version | Tested |
| **Raw** | `vbaProject.bin` | Implemented | Tested |

`.accde` and `.mde` will not work - you're forgiven for thinking that this is just me being lazy, but it appears that Access strips the source when it compiles those, leaving only p-code, so there is genuinely nothing there to recover. 

## A 'quick' note about Access

Every other Office application stores its VBA the same 'civilised' way: a compound document called `vbaProject.bin`, sitting either as a file inside the OOXML zip or as a storage inside the old binary format. Point any OLE parser at it and the modules are right there.

Access does not do this. Access takes the VBA project, chops it into pieces, and stores those pieces as rows in hidden system tables. There is no `vbaProject.bin` to find. This is why olevba doesn't support it, why mdbtools has an open feature request for it that nobody has picked up, and why the canonical advice for "how do I get the code out of this 2003 MDB" is still "install an old copy of Access".

vbaXray takes the other route. It scans the database for LVAL pages, follows the row chains to reassemble anything that spans pages, decompresses each candidate blob, and keeps whatever comes out looking like a module. No system tables, no catalog parsing, no reassembling a synthetic compound document to feed to a parser that expects one. If it decompresses into something that starts `Attribute VB_Name = `, it's a module.

> [!NOTE]
> Access support is experimental and currently assumes 4096-byte ACE/Jet 4 pages. That covers the newer ACE files and Jet 4-era databases, but I have **not** exhaustively tested the various Access generations and file variants. Access 97 used 2048-byte pages and a different storage scheme again, so it is not currently handled.

## Requirements

Windows, and a VBA host. 32-bit and 64-bit are both supported, and VBA6/VB6 hosts compile through the `#Else` branch. 

Legacy PowerPoint is the one exception, and needs Microsoft's `archiveint.dll` (the Windows build of libarchive - see ChibiArc) when it encounters compressed VBA records. Raw/uncompressed records do not need the inflater.

> [!NOTE]
> `archiveint.dll` has shipped as a Windows component since Windows 10 version 1803. On supported modern Windows installations it should therefore already be present. If you are running something older, or a particularly stripped-down Windows installation, legacy PowerPoint support may not work. 

> [!IMPORTANT]
> As ever, any bugs, blunders, oversights, and general acts of coding inelegance are entirely my own. Any sparks of coding brilliance very likely belong to other people.

## API Overview

### Loading

| Method | Description |
| :--- | :--- |
| `LoadFromFile(Path)` | Load from any supported file. Returns `True` only if something actually loaded. |
| `LoadFromByteArray(Data())` | Load from a `vbaProject.bin` you already have in memory. |
| `Reset()` | Clear everything and start again. |

### Reading

| Property | Description |
| :--- | :--- |
| `IsLoaded` | `True` when a project with at least one module is in memory |
| `FilePath` | Where the current project was loaded from |
| `ProjectName` | The `Name=` value from the PROJECT stream |
| `ProjectDescription` | The `Description=` value from the PROJECT stream |
| `CodePage` | The ANSI code page the source is encoded in, e.g. 1252. Needs testing with CJK projects. |
| `IsPasswordProtected` | `True` if the project has a password set |
| `ModuleCount` | How many modules were found |
| `ModuleName(Index)` | Name of the module at 1-based `Index` |
| `ModuleType(Index)` | `modStandard`, `modClass`, `modDocument`, `modDesigner`, `modOther` |
| `IsForm(Index)` | `True` if it's a UserForm designer module |
| `SourceCode(Index)` | The full decompressed source |

### Exporting

| Method | Description |
| :--- | :--- |
| `ExportAll(Path)` | Everything, into numbered sub-folders by module type |
| `ExportModule(Index, Path)` | One module, by index |
| `ExportModuleByName(Name, Path)` | One module, by name, case-insensitive |
| `FullSourceCode([Separator])` | The whole project as one string. `{0}` in the separator becomes the module name |

### Diagnostics

| Member | Description |
| :--- | :--- |
| `LastError` | Message from the most recent failure. Cleared at the start of each public call |
| `LogLevel` | Threshold for Immediate window output: `xrayDebug`, `xrayInfo` (default), `xrayWarning`, `xrayError` |
| `LastOperationTime` | Elapsed seconds for the most recent load or `ExportAll` |
| `Version` | Class version as a `Single` |
| `IsCompoundFile(Path)` | `True` if the file starts with the CFB signature. Handy for routing before you load |
| `DebugDumpStorageTree([Path])` | Geniunely useful! The full storage tree with stream sizes, to the Immediate window or to a file. Non-printable characters in stream names come out as `\xNN`, which is how you find out what's actually in there |

### Export layout

```
ExportPath\
    01_Standard\     .bas
    02_Class\        .cls  
    03_Document\     .cls  (ThisWorkbook, Sheet1, and so on)
    04_Designer\     .frm  (UserForms)
    05_Other\        .txt  (when all else fails)
```

Classes get the `VERSION 1.0 CLASS` etc bolted onto the start of module, because the VBE refuses to import them otherwise. Files are written in the project's own code page rather than UTF-8, so non-ASCII identifiers **should** survive the round trip.

## Limitations

* **Read-only.** vbaXray reads. It does not write. Injecting VBA back into a container is a whole other headache that I have not got around to yet. Office will gleefully reject output that complies with the specification but not byte-identical to what their own compressor would have (should have?) produced, which is a delightful thing to discover empirically.

* **No FRX extraction.** UserForm binary data - images, control layout, embedded objects - lives in a parallel ~~universe~~  `.frx` stream that is not extracted. The `.frm` text comes out in full (with some curious though benign extras in the header portion), so the code is all there, but an exported form will not re-import at all as is without the FRX data. I don't know that anyone has managed to reverse-engineer the FRX format thus far, but I live in hope.

* **No p-code.** This is yet another headache, and I will not claim to know enough about the VBA flavour of p-code to anything sensible with it as yet. Only the original source is recovered.

* **No embedded OLE extraction and recursion.** ... yet! Stay tuned!

* **No references list.** The `dir` stream carries the project's library references and vbaXray currently walks straight past them. This one is genuinely on the to-do list.

* **No Visio support.** Frankly, I've never used Visio, and while I genuinely did try to add support, I ended up removing  the Visio-specific extraction path because it just did not work on the single Visio file I had available. But I'm an adorable and naievely trusting sort-of-person, so if you have a few non-malware-riddled Visio files that you would be happy to share or can otherwise direct me to, please get in touch.

## Changes in 2.0

Version 1.0 reads OOXML files via `Shell.Application` and not much else. Version 2.0 is, in parts, a rewrite rather than a polish, hence the number.

* **OOXML extraction rewritten** onto the `ZipFldr.dll` `IStorage` route. Nothing touches the disk, and the polling loop is gone.
* **Legacy compound files** - `.xls`, `.xla`, `.xlt`, `.doc`, `.dot`, `.pub`, and Outlook `.otm`.
* **Legacy PowerPoint** - `.ppt`, `.pps`, `.pot`, `.ppa`. This one fought back; see `PPT_Deflate_Memo.md` if you enjoy that sort of thing.
* **Access** - `.accdb`, `.accdt`, `.accdr`, `.mdb`. This remains experimental and is currently aimed at 4096-byte ACE/Jet 4 page layouts.
* **Diagnostics** - `LastError`, `LogLevel`, `LastOperationTime` and `Version`, so failures are inspectable from code instead of only visible in the Immediate window.
* **COM lifetime fix.** Intermediate storage objects are now held alive while navigating a path. Previously they could be released early, and the child storage would come back `STG_E_REVERTED` (`0x80030102`), which is not an error message that tells you very much.
* **`LoadFromFile` now tells the truth.** It used to return `True` whether or not anything had loaded.
* **No ADODB dependency.** File writing goes through `CreateFileW` / `WriteFile` directly.
* **New:** `IsLoaded`, `FilePath`, `ProjectDescription`, `IsCompoundFile`, `DebugDumpStorageTree`.

## Future work

There are a few things I'd quite like to investigate, although none of them should be taken as any assurance that I have any idea what I'm doing.

* **Project references.** The `dir` stream contains them; vbaXray currently doesn't expose them.
* **FRX extraction.** The `.frm` source comes out, but the accompanying binary `.frx` data does not.
* **Older MDB format.**
* **Embedded files and OLE objects.** The storage tree already exposes where these things live; actually extracting and/or following them is another job (mostly complete).
* **Writing/editing.** Eventually I'd like to see whether a VBA project can be modified and successfully written back into its container.

---

## Credits

vbaXray was created by Kallun Willock (me).

* **WQWeto**, without whom I would not know `archiveint.dll` existed at all: <https://www.vbforums.com/showthread.php?894163-VB6-Decompress-gzip-stream-with-libarchive-on-Win10>
* **fafalone and The trick**, for the ZipFldr IStorage technique the OOXML path is built on: <https://www.vbforums.com/showthread.php?804893>
* **Beakerboy**, for a great deal of careful MS-OVBA work: <https://github.com/Beakerboy/>
* **WilliamSmithEdward** for solving the ACCDB/MDB formats: <https://github.com/WilliamSmithEdward/pyOpenVBA/>
* **OLEVBA** — generally, but also the source reference for PPT record layout: <https://github.com/decalage2/oletools>

---

## References / resources

* MS-OVBA (VBA File Format Structure) - the canonical reference for how any of this works: <https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/>
* MS-CFB (Compound File Binary File Format): <https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/>
* MS-PPT (PowerPoint 97-2003 Binary File Format), for where the VBA hides in a `.ppt`: <https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/>
* ChibiArc: <https://github.com/KallunWillock/ChibiArc> (`archiveint.dll` is the Windows build of libarchive)
* pyOpenVBA - MS Access Lessons Learned: <https://github.com/WilliamSmithEdward/pyOpenVBA/blob/main/docs/msaccess_lessons_learned.md>
* Cristian Buse - Excel-ZipTools: <https://github.com/cristianbuse/excel-ziptools>

---

## Changelog

* **2.0** - Extended the original engine: OOXML via ZipFldr IStorage, legacy compound files, legacy PowerPoint, Access, and diagnostic tools.
* **1.0** - Initial public release. OOXML only, via `Shell.Application`.

## License

MIT — see [LICENSE](LICENSE).

## Author

Kallun Willock — [https://github.com/KallunWillock](https://github.com/KallunWillock)
