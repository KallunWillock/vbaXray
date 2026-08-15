VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBAXray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'  |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'  |||||||||||||||||||||||||||    vbaXray (v2.0)    ||||||||||||||||||||||||||||||
'  |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'
'  AUTHOR:   Kallun Willock
'
'  URL:      https://github.com/KallunWillock/vbaXray
'
'  NOTES:    * Reads VBA source code straight out of Office files.
'            * OOXML (.xlsm, .xlam, .xlsb, .xltm, .docm, .dotm, .pptm, .potm, .ppsm,
'              .ppam, .sldm) via the ZipFldr IStorage route. Nothing written to disk.
'            * Legacy compound files (.xls/.xla/.xlt, .doc/.dot, .pub, .otm) read
'              directly through StgOpenStorage.
'            * Legacy PowerPoint (.ppt/.pps/.pot/.ppa), including zlib-compressed
'              VbaProjectStg records.
'            * Access (.accdb/.mdb) by walking LVAL database pages. Experimental,
'              assumes 4096-byte ACE/Jet 4 pages.
'
'  CREDITS:  wqweto, without whom I would not know archiveint.dll existed at all:
'            https://www.vbforums.com/showthread.php?894163-VB6-Decompress-gzip-stream-with-libarchive-on-Win10
'
'            fafalone and The trick, for the ZipFldr IStorage technique the OOXML path
'            is built on: https://www.vbforums.com/showthread.php?804893
'
'            Beakerboy, for a great deal of careful MS-OVBA work:
'            https://github.com/Beakerboy/
'
'            WilliamSmithEdward for solving the ACCDB/MDB formats:
'            https://github.com/WilliamSmithEdward/pyOpenVBA
'
'            OLEVBA - generally, but is also the source reference for PPT record layout:
'            https://github.com/decalage2/oletools
'
'  LICENSE:  MIT
'
'  VERSION:  1.0    23/03/2026    Initial release. OOXML only, via Shell.Application.
'            2.0    15/08/2026    OOXML via ZipFldr IStorage, legacy compound
'                                 files, legacy PowerPoint, Access, diagnostics.

Option Explicit

#If Win64 Then
  Private Const PTR_SIZE As Long = 8
#Else
  Private Const PTR_SIZE As Long = 4
#End If

#If VBA7 Then
  #If Win64 Then
    Private Declare PtrSafe Function archive_read_data_block Lib "archiveint.dll" (ByVal hRead As LongPtr, ByRef p As LongPtr, ByRef sz As LongPtr, ByRef off As LongLong) As Long
  #Else
    Private Declare PtrSafe Function archive_read_data_block Lib "archiveint.dll" (ByVal hRead As LongPtr, ByRef p As LongPtr, ByRef sz As LongPtr, ByRef off As Currency) As Long
  #End If
  Private Declare PtrSafe Function archive_read_new Lib "archiveint.dll" () As LongPtr
  Private Declare PtrSafe Function archive_read_support_format_raw Lib "archiveint.dll" (ByVal hRead As LongPtr) As Long
  Private Declare PtrSafe Function archive_read_support_filter_all Lib "archiveint.dll" (ByVal hRead As LongPtr) As Long
  Private Declare PtrSafe Function archive_read_open_memory Lib "archiveint.dll" (ByVal hRead As LongPtr, ByVal pBuffer As LongPtr, ByVal bufferSize As LongPtr) As Long
  Private Declare PtrSafe Function archive_read_next_header Lib "archiveint.dll" (ByVal hRead As LongPtr, ByRef hEntry As LongPtr) As Long
  Private Declare PtrSafe Function archive_read_close Lib "archiveint.dll" (ByVal hRead As LongPtr) As Long
  Private Declare PtrSafe Function archive_read_free Lib "archiveint.dll" (ByVal hRead As LongPtr) As Long

  #If Win64 Then
    ' Write side, used only to deflate vbaProject.bin when repackaging an OOXML
    ' container. Declared for 64-bit alone: archive_entry_set_size and
    ' archive_write_data take 64-bit integers, and getting that marshalling right
    ' under 32-bit stdcall is not something this has been tested against. Off
    ' 64-bit the entry is stored uncompressed instead, which is valid but larger.
    Private Declare PtrSafe Function archive_write_new Lib "archiveint.dll" () As LongPtr
    Private Declare PtrSafe Function archive_write_add_filter_gzip Lib "archiveint.dll" (ByVal hWrite As LongPtr) As Long
    Private Declare PtrSafe Function archive_write_set_format_raw Lib "archiveint.dll" (ByVal hWrite As LongPtr) As Long
    Private Declare PtrSafe Function archive_write_set_options Lib "archiveint.dll" (ByVal hWrite As LongPtr, ByVal opts As LongPtr) As Long
    Private Declare PtrSafe Function archive_write_open_memory Lib "archiveint.dll" (ByVal hWrite As LongPtr, ByVal pBuffer As LongPtr, ByVal bufferSize As LongPtr, ByRef used As LongPtr) As Long
    Private Declare PtrSafe Function archive_write_header Lib "archiveint.dll" (ByVal hWrite As LongPtr, ByVal hEntry As LongPtr) As Long
    Private Declare PtrSafe Function archive_write_data Lib "archiveint.dll" (ByVal hWrite As LongPtr, ByVal pData As LongPtr, ByVal Size As LongPtr) As LongLong
    Private Declare PtrSafe Function archive_write_close Lib "archiveint.dll" (ByVal hWrite As LongPtr) As Long
    Private Declare PtrSafe Function archive_write_free Lib "archiveint.dll" (ByVal hWrite As LongPtr) As Long
    Private Declare PtrSafe Function archive_entry_new Lib "archiveint.dll" () As LongPtr
    Private Declare PtrSafe Function archive_entry_free Lib "archiveint.dll" (ByVal hEntry As LongPtr) As Long
    Private Declare PtrSafe Sub archive_entry_set_size Lib "archiveint.dll" (ByVal hEntry As LongPtr, ByVal Size As LongLong)
    Private Declare PtrSafe Sub archive_entry_set_filetype Lib "archiveint.dll" (ByVal hEntry As LongPtr, ByVal fileType As Long)
  #End If

  Private Declare PtrSafe Function apiCreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
  Private Declare PtrSafe Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
  Private Declare PtrSafe Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
  Private Declare PtrSafe Function apiGetFileSizeEx Lib "kernel32" Alias "GetFileSizeEx" (ByVal hFile As LongPtr, ByRef lpFileSize As LongPtr) As Long
  Private Declare PtrSafe Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function apiCreateDirectoryW Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As LongPtr, ByVal lpSecurityAttributes As LongPtr) As Long
  Private Declare PtrSafe Function apiGetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As LongPtr) As Long

  Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
  Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
  Private Declare PtrSafe Function StgOpenStorage Lib "ole32" (ByVal pwcsName As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ByRef ppstgOpen As stdole.IUnknown) As Long
  Private Declare PtrSafe Function CreateILockBytesOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ByRef ppLkbyt As stdole.IUnknown) As Long
  Private Declare PtrSafe Function StgOpenStorageOnILockBytes Lib "ole32" (ByVal pLkbyt As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ppstgOpen As stdole.IUnknown) As Long
  Private Declare PtrSafe Function GetHGlobalFromILockBytes Lib "ole32" (ByVal pLkbyt As LongPtr, ByRef phglobal As LongPtr) As Long
  Private Declare PtrSafe Function ILCreateFromPathW Lib "shell32" (ByVal pszPath As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

  Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
  Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
  Private Declare PtrSafe Function CoCreateInstance Lib "ole32" (rclsid As UUID, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, riid As UUID, ppv As stdole.IUnknown) As Long
  Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As LongPtr, pclsid As UUID) As Long
  Private Declare PtrSafe Function ILFree Lib "shell32" (ByVal pidl As LongPtr) As Long
  Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
  Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
#Else
  Private Enum LongPtr
    [_]
  End Enum
  Private Declare Function archive_read_new Lib "archiveint.dll" Alias "_archive_read_new@0" () As Long
  Private Declare Function archive_read_free Lib "archiveint.dll" Alias "_archive_read_free@4" (ByVal hArchive As Long) As Long
  Private Declare Function archive_read_close Lib "archiveint.dll" Alias "_archive_read_close@4" (ByVal hArchive As Long) As Long
  Private Declare Function archive_read_support_format_raw Lib "archiveint.dll" Alias "_archive_read_support_format_raw@4" (ByVal h As Long) As Long
  Private Declare Function archive_read_support_filter_all Lib "archiveint.dll" Alias "_archive_read_support_filter_all@4" (ByVal h As Long) As Long
  Private Declare Function archive_read_open_memory Lib "archiveint.dll" Alias "_archive_read_open_memory@12" (ByVal h As Long, ByVal pBuffer As Long, ByVal bufferSize As Long) As Long
  Private Declare Function archive_read_next_header Lib "archiveint.dll" Alias "_archive_read_next_header@8" (ByVal h As Long, ByRef hEntry As Long) As Long
  Private Declare Function archive_read_data_block Lib "archiveint.dll" Alias "_archive_read_data_block@16" (ByVal h As Long, ByRef pBuff As Long, ByRef size As Long, ByRef offset As Currency) As Long

  Private Declare Function apiCreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
  Private Declare Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
  Private Declare Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
  Private Declare Function apiGetFileSizeEx Lib "kernel32" Alias "GetFileSizeEx" (ByVal hFile As LongPtr, ByRef lpFileSize As LongPtr) As Long
  Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As LongPtr) As Long
  Private Declare Function ILCreateFromPathW Lib "shell32" (ByVal pszPath As LongPtr) As LongPtr
  Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
  Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
  Private Declare Function apiCreateDirectoryW Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As LongPtr, ByVal lpSecurityAttributes As LongPtr) As Long
  Private Declare Function apiGetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As LongPtr) As Long
  Private Declare Function CoCreateInstance Lib "ole32" (rclsid As UUID, ByVal pUnkOuter As LongPtr, ByVal dwClsContext As Long, riid As UUID, ppv As stdole.IUnknown) As Long
  Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As LongPtr, pclsid As UUID) As Long
  Private Declare Function ILFree Lib "shell32" (ByVal pidl As LongPtr) As Long
  Private Declare Function StgOpenStorage Lib "ole32" (ByVal pwcsName As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ByRef ppstgOpen As stdole.IUnknown) As Long
  Private Declare Function CreateILockBytesOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ByRef ppLkbyt As stdole.IUnknown) As Long
  Private Declare Function StgOpenStorageOnILockBytes Lib "ole32" (ByVal pLkbyt As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ppstgOpen As stdole.IUnknown) As Long
  Private Declare Function GetHGlobalFromILockBytes Lib "ole32" (ByVal pLkbyt As LongPtr, ByRef phglobal As LongPtr) As Long

  Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
  Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

  Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
  Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
  Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
#End If

Private Type UUID
  Data1             As Long
  Data2             As Integer
  Data3             As Integer
  Data4(0 To 7)     As Byte
End Type

Private Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type

Private Type STATSTG
  pwcsName          As LongPtr
  dwType            As Long
  cbSize            As Currency
  mtime             As FILETIME
  ctime             As FILETIME
  atime             As FILETIME
  grfMode           As Long
  grfLocksSupported As Long
  clsid             As UUID
  grfStateBits      As Long
  reserved          As Long
End Type

Private Enum IStorageInterface
  IStg_CreateStream = 3
  IStg_OpenStream = 4
  IStg_CreateStorage = 5
  IStg_OpenStorage = 6
  IStg_CopyTo = 7
  IStg_MoveElementTo = 8
  IStg_Commit = 9
  IStg_Revert = 10
  IStg_EnumElements = 11
  IStg_DestroyElement = 12
  IStg_RenameElement = 13
  IStg_SetElementTimes = 14
  IStg_SetClass = 15
  IStg_SetStateBits = 16
  IStg_stat = 17
End Enum

Private Enum IEnumSTATSTGInterface
  IEnum_Next = 3
  IEnum_Skip = 4
  IEnum_Reset = 5
  IEnum_Clone = 6
End Enum

Private Enum IStreamInterface
  IStm_Read = 3
  IStm_Write = 4
  IStm_Seek = 5
  IStm_SetSize = 6
  IStm_CopyTo = 7
  IStm_Commit = 8
  IStm_Revert = 9
  IStm_LockRegion = 10
  IStm_UnlockRegion = 11
  IStm_stat = 12
  IStm_Clone = 13
End Enum

Private Enum ILockBytesInterface
  ILkb_ReadAt = 3
  ILkb_WriteAt = 4
  ILkb_Flush = 5
  ILkb_SetSize = 6
  ILkb_LockRegion = 7
  ILkb_UnlockRegion = 8
  ILkb_stat = 9
End Enum

Private Enum STGTY
  STGTY_STORAGE = 1
  STGTY_STREAM = 2
  STGTY_LOCKBYTES = 3
  STGTY_PROPERTY = 4
End Enum

Private Const STGM_READ                 As Long = &H0
Private Const STGM_WRITE                As Long = &H1
Private Const STGM_READWRITE            As Long = &H2
Private Const STGM_SHARE_DENY_WRITE     As Long = &H20
Private Const STGM_SHARE_EXCLUSIVE      As Long = &H10
Private Const STATFLAG_NONAME           As Long = &H1
Private Const CC_STDCALL                As Long = 4
Private Const INVALID_FILE_ATTRIBUTES   As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10
Private Const GMEM_MOVEABLE             As Long = &H2

Private Const GENERIC_WRITE             As Long = &H40000000
Private Const GENERIC_READ              As Long = &H80000000
Private Const OPEN_EXISTING             As Long = 3
Private Const FILE_SHARE_READ           As Long = &H1
Private Const FILE_SHARE_WRITE          As Long = &H2
Private Const FILE_SHARE_DELETE         As Long = &H4
Private Const CREATE_ALWAYS             As Long = 2
Private Const INVALID_HANDLE_VALUE      As LongPtr = -1
Private Const CLSCTX_INPROC_SERVER      As Long = 1
Private Const ARCHIVE_OK                As Long = 0
Private Const ARCHIVE_EOF               As Long = 1
Private Const PPT_VBA_PROJECT_STG       As Long = &H1011&
Private Const MAX_EMBEDDED_PROJECT_SIZE As Long = 268435456  ' 256 Mb safety ceiling
Private Const GZIP_HEADER_SIZE          As Long = 10
Private Const GZIP_TRAILER_SIZE         As Long = 8
Private Const AE_IFREG                  As Long = &H8000&
Private Const ZIP_METHOD_STORE          As Long = 0
Private Const ZIP_METHOD_DEFLATE        As Long = 8
Private Const VBAXRAY_VERSION           As Single = 2#

Private Const IID_IPersistFolder        As String = "{000214EA-0000-0000-C000-000000000046}"
Private Const IID_IStorage              As String = "{0000000B-0000-0000-C000-000000000046}"
Private Const CLSID_ZipFolder           As String = "{E88DCCE0-B7B3-11d1-A9F0-00AA0060FA31}"

Public Enum XrayLogLevel
  xrayDebug = 0
  xrayInfo = 1
  xrayWarning = 2
  xrayError = 3
End Enum

Public Enum ModuleType
  modOther
  modStandard
  modClass
  modDesigner
  modDocument
End Enum

' Where the loaded project came from, and therefore how it can be written back.
Private Enum XraySourceKind
  srcNone = 0
  srcProjectBin = 1     ' a standalone vbaProject.bin
  srcCompoundFile = 2   ' a legacy compound container (.xls, .doc, .pub, .otm)
  srcOOXML = 3          ' vbaProject.bin lifted out of a zip container
  srcAccess = 4         ' read-only: the compiled p-code, not the source, is authoritative
  srcLegacyPPT = 5      ' read-only: writing legacy PowerPoint is out of scope here
End Enum

Private Type ModuleRecord
  Name        As String
  streamName  As String
  offset      As Long
  ModType     As ModuleType
  IsForm      As Boolean
  SourceCode  As String
  IsDirty     As Boolean
End Type

Private Type ProjectRecord
  Name                As String
  Description         As String
  FilePath            As String
  CodePage            As Long
  IsPasswordProtected As Boolean
  ModuleCount         As Long
  Modules()           As ModuleRecord
End Type

Private Type CFBLayout
  VBAPath     As String
  projectPath As String
End Type

Private m_Project           As ProjectRecord
Private m_StorageTree       As Collection
Private m_LastError         As String
Private m_LogLevel          As XrayLogLevel
Private m_LastOperationTime As Double
Private m_StartTimer        As Double

' Write-path state. m_SourceBytes is the compound file that gets rewritten: the
' whole container for a legacy compound file, or the extracted vbaProject.bin for
' a zip container. m_ContainerPath is only set for OOXML, where the rebuilt bin
' has to be put back into the archive it came from.
Private m_SourceKind        As XraySourceKind
Private m_SourceBytes()     As Byte
Private m_SourceExt         As String
Private m_ContainerPath     As String

' CRC-32 table for the zip writer, built on first use.
Private m_CrcTable()        As Long
Private m_CrcReady          As Boolean

Public Property Get Version() As Single
  Version = VBAXRAY_VERSION
End Property

Public Property Get LastError() As String
  LastError = m_LastError
End Property

Public Property Get LastOperationTime() As Double
  LastOperationTime = m_LastOperationTime
End Property

Public Property Let LogLevel(ByVal val As XrayLogLevel)
  m_LogLevel = val
End Property

Public Property Get LogLevel() As XrayLogLevel
  LogLevel = m_LogLevel
End Property

Public Property Get ProjectName() As String
  ProjectName = m_Project.Name
End Property

Public Property Get ProjectDescription() As String
  ProjectDescription = m_Project.Description
End Property

Public Property Get FilePath() As String
  FilePath = m_Project.FilePath
End Property

Public Property Get CodePage() As Long
  CodePage = m_Project.CodePage
End Property

Public Property Get IsPasswordProtected() As Boolean
  IsPasswordProtected = m_Project.IsPasswordProtected
End Property

Public Property Get IsLoaded() As Boolean
  IsLoaded = (m_Project.ModuleCount > 0)
End Property

Public Property Get ModuleCount() As Long
  ModuleCount = m_Project.ModuleCount
End Property

Public Property Get ModuleName(ByVal index As Long) As String
  If index < 1 Or index > m_Project.ModuleCount Then
    SetError "[ModuleName] Index out of range": Exit Property
  End If
  ModuleName = m_Project.Modules(index).Name
End Property

Public Property Get ModuleType(ByVal index As Long) As ModuleType
  If index < 1 Or index > m_Project.ModuleCount Then
    SetError "[ModuleType] Index out of range": Exit Property
  End If
  ModuleType = m_Project.Modules(index).ModType
End Property

Public Property Get IsForm(ByVal index As Long) As Boolean
  If index < 1 Or index > m_Project.ModuleCount Then
    SetError "[IsForm] Index out of range": Exit Property
  End If
  IsForm = m_Project.Modules(index).IsForm
End Property

Public Property Get SourceCode(ByVal index As Long) As String
  If index < 1 Or index > m_Project.ModuleCount Then
    SetError "[SourceCode] Index out of range": Exit Property
  End If
  SourceCode = m_Project.Modules(index).SourceCode
End Property

Public Function LoadFromFile(ByVal FilePath As String) As Boolean
  ClearError
  Reset
  LoadFromFile = False
  m_StartTimer = Timer
  If FileExists(FilePath) Then
    Dim ext As String
    ext = LCase$(VBA.Mid$(FilePath, InStrRev(FilePath, ".") + 1))
    If ext = "accde" Or ext = "mde" Then
      SetError "[LoadFromFile] " & UCase$(ext) & " files have VBA source removed at compile time - only p-code remains. Source extraction is not possible."
    ElseIf ext = "accdb" Or ext = "accdt" Or ext = "accdr" Or ext = "mdb" Then
      LoadFromFile = LoadFromAccessDB(FilePath)
      If LoadFromFile Then
        m_Project.FilePath = FilePath
        m_SourceKind = srcAccess
      End If
    ElseIf IsCompoundFile(FilePath) Then
      Select Case ext
      Case "ppt", "pps", "pot", "ppa"
        LoadFromFile = LoadFromLegacyPowerPoint(FilePath)
        If LoadFromFile Then
          m_Project.FilePath = FilePath
          m_SourceKind = srcLegacyPPT
        End If
      Case Else
        Dim fileStorage As stdole.IUnknown
        Set fileStorage = StorageOpenFile(FilePath)
        LoadFromStorage fileStorage, ext
        LoadFromFile = IsLoaded
        If LoadFromFile Then
          m_Project.FilePath = FilePath
          ' Only the path is kept. The container bytes are read when saving, so
          ' loading a large legacy file to read it costs nothing extra.
          m_ContainerPath = FilePath
          m_SourceExt = ext
          m_SourceKind = srcCompoundFile
        End If
      End Select
    Else
      Select Case ext
      Case "xlsm", "xlam", "xlsb", "xltm", "docm", "dotm", "pptm", "potm", "ppsm", "ppam", "sldm"
        Dim vbaBytes() As Byte
        vbaBytes = GetVBAProjectBinBytes(FilePath)
        If ByteCount(vbaBytes) > 0 Then
          LoadFromFile = LoadFromByteArray(vbaBytes)
          If LoadFromFile Then
            m_Project.FilePath = FilePath
            ' LoadFromByteArray recorded srcProjectBin; the bin actually came out
            ' of a zip, so saving has to put it back.
            m_ContainerPath = FilePath
            m_SourceKind = srcOOXML
          End If
        End If
      Case Else
        SetError "[LoadFromFile] Unrecognised file type '." & ext & "'"
      End Select
    End If
  Else
    SetError "[LoadFromFile] File not found: " & FilePath
  End If
  m_LastOperationTime = Timer - m_StartTimer
End Function

Public Function LoadFromByteArray(ByRef VBAProjectBinData() As Byte) As Boolean
    ClearError
    Reset
    LoadFromByteArray = False
    m_StartTimer = Timer
    If ByteCount(VBAProjectBinData) >= 8 Then
      If IsCompoundFileBytes(VBAProjectBinData) Then
        Dim byteStorage As stdole.IUnknown
        Set byteStorage = StorageOpenBytes(VBAProjectBinData)
        If Not byteStorage Is Nothing Then
          LoadFromStorage byteStorage, "bin"
          LoadFromByteArray = IsLoaded
          If LoadFromByteArray Then
            ' Set after the load: LoadFromStorage calls Reset on entry.
            m_SourceBytes = VBAProjectBinData
            m_SourceExt = "bin"
            m_SourceKind = srcProjectBin
          End If
        End If
      Else
        SetError "[LoadFromByteArray] Data does not have a valid Compound Document signature"
      End If
    Else
      SetError "[LoadFromByteArray] Data is too small to be a valid Compound Document"
    End If
    m_LastOperationTime = Timer - m_StartTimer
End Function

Public Sub Reset()
  Dim EmptyProject As ProjectRecord
  m_Project = EmptyProject
  m_Project.CodePage = 1252
  Set m_StorageTree = Nothing
  m_SourceKind = srcNone
  m_SourceExt = vbNullString
  m_ContainerPath = vbNullString
  Erase m_SourceBytes
End Sub

Public Sub ExportAll(ByVal ExportPath As String)
  ClearError
  If Not PrepareExport("ExportAll", ExportPath) Then Exit Sub

  m_StartTimer = Timer
  Dim i         As Long
  Dim succeeded As Long
  Dim Failed    As Long
  For i = 1 To m_Project.ModuleCount
    If WriteModule(i, ExportPath) Then
      succeeded = succeeded + 1
    Else
      Failed = Failed + 1
    End If
  Next i
  m_LastOperationTime = Timer - m_StartTimer

  LogMessage xrayInfo, "Export complete. " & succeeded & " module(s) exported; " & Failed & " failed. Destination: " & ExportPath
End Sub

Public Sub ExportModule(ByVal index As Long, ByVal ExportPath As String)
  ClearError
  If Not PrepareExport("ExportModule", ExportPath) Then Exit Sub
  If index < 1 Or index > m_Project.ModuleCount Then
    SetError "[ExportModule] Index " & index & " out of range"
    Exit Sub
  End If
  If Not WriteModule(index, ExportPath) Then
    SetError "[ExportModule] Export failed for module " & m_Project.Modules(index).Name
  End If
End Sub

Public Sub ExportModuleByName(ByVal modName As String, ByVal ExportPath As String)
  ClearError
  Dim i As Long
  For i = 1 To m_Project.ModuleCount
    If StrComp(m_Project.Modules(i).Name, modName, vbTextCompare) = 0 Then
      ExportModule i, ExportPath
      Exit Sub
    End If
  Next i
  SetError "[ExportModuleByName] Module not found: " & modName
End Sub

Private Function PrepareExport(ByVal caller As String, ByRef ExportPath As String) As Boolean
  If Not IsLoaded Then
    SetError "[" & caller & "] No project loaded - call LoadFromFile or LoadFromByteArray first"
    Exit Function
  End If

  ExportPath = VBA.Trim$(ExportPath)
  If Right$(ExportPath, 1) <> "\" Then ExportPath = ExportPath & "\"
  If Not FolderExists(ExportPath) Then
    If Not CreateFolderRecursive(ExportPath) Then
      SetError "[" & caller & "] Could not create export folder: " & ExportPath
      Exit Function
    End If
  End If

  PrepareExport = True
End Function

Public Function FullSourceCode(Optional ByVal Separator As String = vbCrLf & vbCrLf & _
                               "' ------------------------------" & vbCrLf & _
                               "' Module: {0}" & vbCrLf & _
                               "' ------------------------------" & vbCrLf) As String
  If m_Project.ModuleCount = 0 Then Exit Function
  Dim parts() As String, i As Long
  ReDim parts(1 To m_Project.ModuleCount)
  For i = 1 To m_Project.ModuleCount
    parts(i) = Replace(Separator, "{0}", m_Project.Modules(i).Name) & m_Project.Modules(i).SourceCode
  Next i
  FullSourceCode = Join(parts, vbNullString)
End Function

' ---------------------------------------------------------------------------
'  Writing
'
'  Replacing a module's source leaves the compiled p-code in that module's
'  performance cache untouched, so the cache no longer agrees with the source.
'  Excel, Word and PowerPoint notice the mismatch on open and recompile from
'  source, which is what makes writing possible without a p-code assembler.
'  Access does not recompile from the cache, which is why Access stays read-only.
' ---------------------------------------------------------------------------

' True when the loaded project came from a container this class can write back.
Public Property Get CanWrite() As Boolean
  Select Case m_SourceKind
    Case srcProjectBin, srcCompoundFile, srcOOXML: CanWrite = True
  End Select
End Property

' True when any module's source has been replaced since the project was loaded.
Public Property Get IsModified() As Boolean
  Dim i As Long
  For i = 1 To m_Project.ModuleCount
    If m_Project.Modules(i).IsDirty Then
      IsModified = True
      Exit Property
    End If
  Next i
End Property

' Replace one module's source. Nothing is written until Save or SaveAs.
'
' Line endings are normalised to CRLF, which is what a module stream stores. A
' class source may be supplied either as a bare body or as a .cls exported by
' ExportModule: the "VERSION 1.0 CLASS" preamble that export adds is not part of
' the stream and is stripped back off here. The Attribute lines are part of the
' stream and are kept.
Public Function SetModuleSource(ByVal index As Long, ByVal SourceCode As String) As Boolean
  ClearError

  If Not IsLoaded Then
    SetError "[SetModuleSource] No project is loaded"
    Exit Function
  End If
  If index < 1 Or index > m_Project.ModuleCount Then
    SetError "[SetModuleSource] Module index " & index & " is out of range (1 to " & m_Project.ModuleCount & ")"
    Exit Function
  End If
  If Not CanWrite Then
    SetError "[SetModuleSource] " & WriteBlockedReason()
    Exit Function
  End If

  With m_Project.Modules(index)
    If .ModType = modDesigner Then
      LogMessage xrayWarning, "[SetModuleSource] '" & .Name & "' is a designer module. " & _
                              "Its code-behind is replaced; the form's binary FRX data is left as it is."
    End If
    .SourceCode = StripClassPreamble(NormaliseLineEndings(SourceCode))
    .IsDirty = True
  End With

  SetModuleSource = True
End Function

Public Function SetModuleSourceByName(ByVal modName As String, ByVal SourceCode As String) As Boolean
  Dim i As Long
  ClearError
  For i = 1 To m_Project.ModuleCount
    If LCase$(m_Project.Modules(i).Name) = LCase$(modName) Then
      SetModuleSourceByName = SetModuleSource(i, SourceCode)
      Exit Function
    End If
  Next i
  SetError "[SetModuleSourceByName] No module named '" & modName & "'"
End Function

' Write the project, with every replaced module in place, to OutputPath.
'
' Edits are applied to the container as it was loaded, every time, so saving to
' several paths in turn produces the same result each time and IsModified keeps
' reporting how the project differs from what was loaded.
'
' A password-locked project needs AllowProtected. The lock bytes are carried
' over untouched, but rewriting the source of a project someone deliberately
' locked is worth being explicit about. A digitally signed project has its
' signature streams dropped, because any change to the modules invalidates it
' and a signature that no longer verifies is worse than none; pass
' AllowInvalidateSignature to accept that without the warning.
Public Function SaveAs(ByVal OutputPath As String, _
                       Optional ByVal AllowProtected As Boolean = False, _
                       Optional ByVal AllowInvalidateSignature As Boolean = False) As Boolean
  Dim sourceCFB() As Byte
  Dim rebuilt()   As Byte

  On Error GoTo Failed
  ClearError
  m_StartTimer = Timer

  If Not IsLoaded Then
    SetError "[SaveAs] No project is loaded"
    GoTo Done
  End If
  If Not CanWrite Then
    SetError "[SaveAs] " & WriteBlockedReason()
    GoTo Done
  End If
  If Len(VBA.Trim$(OutputPath)) = 0 Then
    SetError "[SaveAs] No output path given"
    GoTo Done
  End If
  If m_Project.IsPasswordProtected And Not AllowProtected Then
    SetError "[SaveAs] The VBA project is password-protected. Its lock bytes are " & _
             "preserved as they are, so the result still needs the original password. " & _
             "Pass AllowProtected:=True to confirm that is intended."
    GoTo Done
  End If

  ' Outlook keeps its macros in one store, VbaProject.OTM, which it holds open
  ' while running and writes back out when it closes. An edit made underneath a
  ' running Outlook is overwritten on exit.
  If LCase$(m_SourceExt) = "otm" Then
    LogMessage xrayWarning, "Outlook keeps VbaProject.OTM open while it runs and rewrites " & _
                            "it on exit. Close Outlook before writing, or the change will be lost."
  End If

  sourceCFB = GetWritableCFB()
  If ByteCount(sourceCFB) = 0 Then GoTo Done

  rebuilt = RewriteCFB(sourceCFB, m_SourceExt, AllowInvalidateSignature)
  If ByteCount(rebuilt) = 0 Then GoTo Done

  If m_SourceKind = srcOOXML Then
    If Not RepackageOOXML(m_ContainerPath, OutputPath, rebuilt) Then GoTo Done
  Else
    If Not WriteFile(OutputPath, rebuilt) Then
      SetError "[SaveAs] Could not write " & OutputPath
      GoTo Done
    End If
  End If

  SaveAs = True
  LogMessage xrayInfo, "Saved: " & OutputPath

Done:
  m_LastOperationTime = Timer - m_StartTimer
  Exit Function

Failed:
  ' Nothing has been written at this point unless WriteFile itself failed, so the
  ' original file is intact either way.
  SaveAs = False
  SetError "[SaveAs] Error " & Err.Number & " - " & Err.Description
  Resume Done
End Function

' Write back over the file the project was loaded from.
Public Function Save(Optional ByVal AllowProtected As Boolean = False, _
                     Optional ByVal AllowInvalidateSignature As Boolean = False) As Boolean
  Dim target As String
  ClearError

  Select Case m_SourceKind
    Case srcOOXML, srcCompoundFile: target = m_ContainerPath
    Case srcProjectBin:             target = m_Project.FilePath
  End Select

  If Len(target) = 0 Then
    SetError "[Save] This project was loaded from memory and has no path of its own; use SaveAs"
    Exit Function
  End If

  Save = SaveAs(target, AllowProtected, AllowInvalidateSignature)
End Function

Private Function WriteBlockedReason() As String
  Select Case m_SourceKind
    Case srcNone
      WriteBlockedReason = "No project is loaded"
    Case srcAccess
      WriteBlockedReason = "Access stores compiled p-code separately from the source cache and " & _
                           "never recompiles from it, so edits to the source would not survive a reload"
    Case srcLegacyPPT
      WriteBlockedReason = "Writing legacy PowerPoint files is not supported"
    Case Else
      WriteBlockedReason = "This project cannot be written back"
  End Select
End Function

' The compound file the write path rewrites: the extracted vbaProject.bin for a
' zip container, or the container itself for a legacy compound file. Loading a
' legacy container keeps only its path, so its bytes are re-read here.
Private Function GetWritableCFB() As Byte()
  Select Case m_SourceKind
    Case srcProjectBin, srcOOXML
      GetWritableCFB = m_SourceBytes
      If ByteCount(GetWritableCFB) = 0 Then
        SetError "[GetWritableCFB] The loaded project has no retained source bytes"
      End If
    Case srcCompoundFile
      GetWritableCFB = ReadFile(m_ContainerPath)
      If ByteCount(GetWritableCFB) = 0 Then
        SetError "[GetWritableCFB] Could not re-read " & m_ContainerPath
      End If
  End Select
End Function

' Rebuild a compound file with every modified module's source replaced.
Private Function RewriteCFB(ByRef cfbBytes() As Byte, ByVal ext As String, _
                            ByVal AllowInvalidateSignature As Boolean) As Byte()
  Dim lockBytes As stdole.IUnknown
  Dim rootStg   As stdole.IUnknown
  Dim vbaStg    As stdole.IUnknown
  Dim keepAlive As New Collection
  Dim layout    As CFBLayout
  Dim i         As Long
  Dim rewritten As Long

  Set rootStg = StorageOpenBytesCore(cfbBytes, STGM_READWRITE, lockBytes)
  If rootStg Is Nothing Then Exit Function

  layout = GetCFBLayout(ext)
  Set vbaStg = StorageNavigate(rootStg, layout.VBAPath, keepAlive, True)
  If vbaStg Is Nothing Then
    SetError "[RewriteCFB] VBA storage not found at '" & layout.VBAPath & "' for writing"
    Exit Function
  End If

  For i = 1 To m_Project.ModuleCount
    If m_Project.Modules(i).IsDirty Then
      If Not RewriteModuleStream(vbaStg, i) Then Exit Function
      rewritten = rewritten + 1
    End If
  Next i

  If rewritten > 0 Then DropSignatureStreams rootStg, vbaStg, AllowInvalidateSignature

  ' Commit inward-out: the module streams were committed as they were written,
  ' then the VBA storage, then the root.
  If Not StorageCommit(vbaStg) Then Exit Function
  Set vbaStg = Nothing
  Set keepAlive = Nothing
  If Not StorageCommit(rootStg) Then Exit Function
  Set rootStg = Nothing

  RewriteCFB = LockBytesToArray(lockBytes)
  LogMessage xrayInfo, "Rebuilt project with " & rewritten & " rewritten module stream(s)"
End Function

' Replace the compressed source in one module stream, keeping the performance
' cache that precedes it. MODULEOFFSET does not move, so the dir stream needs no
' change for a source-only edit.
Private Function RewriteModuleStream(ByRef vbaStg As stdole.IUnknown, ByVal index As Long) As Boolean
  Dim pStm         As stdole.IUnknown
  Dim modBytes()   As Byte
  Dim sourceBytes() As Byte
  Dim compressed() As Byte
  Dim newStream()  As Byte
  Dim prefixLen    As Long
  Dim compLen      As Long

  With m_Project.Modules(index)
    Set pStm = StorageOpenStreamRW(vbaStg, .streamName)
    If pStm Is Nothing Then
      SetError "[RewriteModuleStream] Could not open stream '" & .streamName & "' for writing"
      Exit Function
    End If

    modBytes = StreamReadBytes(pStm)
    If ByteCount(modBytes) = 0 Then
      SetError "[RewriteModuleStream] Stream '" & .streamName & "' is empty"
      Exit Function
    End If
    If .offset > UBound(modBytes) Then
      SetError "[RewriteModuleStream] Offset (" & .offset & ") exceeds stream length for module " & .Name
      Exit Function
    End If

    prefixLen = .offset
    sourceBytes = StringToBytes(.SourceCode, m_Project.CodePage)
    compressed = CompressStream(sourceBytes)
    compLen = ByteCount(compressed)
    If compLen = 0 Then Exit Function

    ReDim newStream(0 To prefixLen + compLen - 1)
    If prefixLen > 0 Then
      CopyMemory VarPtr(newStream(0)), VarPtr(modBytes(LBound(modBytes))), prefixLen
    End If
    CopyMemory VarPtr(newStream(prefixLen)), VarPtr(compressed(LBound(compressed))), compLen

    If Not StreamWriteBytes(pStm, newStream) Then Exit Function
    LogMessage xrayInfo, "[Rewrote] " & .Name & " - " & Len(.SourceCode) & _
                         " chars into " & compLen & " compressed bytes"
  End With

  RewriteModuleStream = True
End Function

' A digital signature covers the modules, so any edit invalidates it. Leaving a
' signature that no longer verifies behind is worse than leaving none, so the
' stale streams go. They sit either inside the VBA storage or beside it at the
' root depending on how the project was signed, so both are tried.
Private Sub DropSignatureStreams(ByRef rootStg As stdole.IUnknown, ByRef vbaStg As stdole.IUnknown, _
                                 ByVal Silent As Boolean)
  Dim names   As Variant
  Dim i       As Long
  Dim dropped As Long

  names = Array("_VBA_PROJECT_SIGNATURE", "_VBA_PROJECT_SIGNATURE_AGILE", "_VBA_PROJECT_SIGNATURE_V3")

  For i = LBound(names) To UBound(names)
    If DestroyElement(vbaStg, CStr(names(i))) Then dropped = dropped + 1
    If DestroyElement(rootStg, CStr(names(i))) Then dropped = dropped + 1
  Next i

  If dropped > 0 And Not Silent Then
    LogMessage xrayWarning, "Dropped " & dropped & " stale digital signature stream(s). " & _
                            "The project was modified, so the signature no longer verifies. " & _
                            "Re-sign the file to restore trust."
  End If
End Sub

' False when the element was not there, which is the common case and not an error.
Private Function DestroyElement(ByRef pStg As stdole.IUnknown, ByVal Name As String) As Boolean
  If pStg Is Nothing Then Exit Function
  DestroyElement = (DispCallByVtbl(pStg, IStorageInterface.IStg_DestroyElement, StrPtr(Name)) >= 0)
End Function

' Module streams store CRLF. Bare LF, from a file written on another platform or
' a here-doc in a script, is converted on the way in, because the VBE treats a
' lone LF as part of the line and the whole module ends up on one line.
Private Function NormaliseLineEndings(ByVal s As String) As String
  NormaliseLineEndings = Replace(Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf), vbLf, vbCrLf)
End Function

' Drop the "VERSION 1.0 CLASS / BEGIN ... END" block that ExportModule writes
' ahead of a class module. It is a VBE import convention, not part of the stream.
Private Function StripClassPreamble(ByVal s As String) As String
  Dim beginPos As Long
  Dim endPos   As Long
  Const END_MARKER As String = vbCrLf & "END" & vbCrLf

  StripClassPreamble = s
  If UCase$(Left$(LTrim$(s), 8)) <> "VERSION " Then Exit Function

  beginPos = InStr(1, s, vbCrLf & "BEGIN", vbTextCompare)
  If beginPos = 0 Then Exit Function
  endPos = InStr(beginPos, s, END_MARKER, vbTextCompare)
  If endPos = 0 Then Exit Function

  StripClassPreamble = Mid$(s, endPos + Len(END_MARKER))
  ' ExportModule leaves a blank line after the preamble.
  If Left$(StripClassPreamble, 2) = vbCrLf Then StripClassPreamble = Mid$(StripClassPreamble, 3)
End Function

' Put a rebuilt vbaProject.bin back into the OOXML package it came from.
'
' Every other entry is copied across with its compressed bytes exactly as Office
' wrote them: the method, CRC and sizes come straight from the original central
' directory and the payload is moved verbatim. Nothing is inflated and nothing is
' deflated, so no compressor is needed and the rest of the package is unchanged
' byte for byte. Only the replaced entry is rewritten, and it is stored rather
' than deflated, which the format permits and Office reads.
'
' The read path mounts the package through ZipFldr as an IStorage. That route is
' not used here: rebuilding an archive needs the raw records, not a shell view of
' them. See GetVBAProjectBinBytes for the read side.
Private Function RepackageOOXML(ByVal ContainerPath As String, ByVal OutputPath As String, _
                                ByRef newBin() As Byte) As Boolean
  Dim zipBytes()  As Byte
  Dim outBytes()  As Byte
  Dim eocd        As Long
  Dim entryCount  As Long
  Dim dirStart    As Long
  Dim at          As Long
  Dim i           As Long

  Dim names()     As Long   ' offset of each entry's name bytes within zipBytes
  Dim nameLens()  As Long
  Dim flags()     As Long
  Dim methods()   As Long
  Dim dosTimes()  As Long
  Dim dosDates()  As Long
  Dim crcs()      As Long
  Dim compSizes() As Long
  Dim rawSizes()  As Long
  Dim dataAt()    As Long
  Dim replaceIdx  As Long

  Dim outPos      As Long
  Dim dirOutStart As Long
  Dim localAt()   As Long
  Dim total       As Long
  Dim newLen      As Long

  Dim repPayload() As Byte    ' the replacement entry as it will be stored
  Dim repMethod    As Long
  Dim repCompLen   As Long
  Dim repCrc       As Long    ' always over the uncompressed bytes

  zipBytes = ReadFile(ContainerPath)
  If ByteCount(zipBytes) = 0 Then
    SetError "[RepackageOOXML] Could not read " & ContainerPath
    Exit Function
  End If

  eocd = FindEndOfCentralDirectory(zipBytes)
  If eocd < 0 Then
    SetError "[RepackageOOXML] No end-of-central-directory record in " & ContainerPath
    Exit Function
  End If
  If ReadUInt16LE(zipBytes, eocd + 4) <> 0 Or ReadUInt16LE(zipBytes, eocd + 6) <> 0 Then
    SetError "[RepackageOOXML] Multi-disk archives are not supported"
    Exit Function
  End If

  entryCount = ReadUInt16LE(zipBytes, eocd + 10)
  If entryCount = 65535 Then
    SetError "[RepackageOOXML] Zip64 archives are not supported"
    Exit Function
  End If
  If entryCount = 0 Then
    SetError "[RepackageOOXML] Archive contains no entries"
    Exit Function
  End If

  ReDim names(1 To entryCount):     ReDim nameLens(1 To entryCount)
  ReDim flags(1 To entryCount):     ReDim methods(1 To entryCount)
  ReDim dosTimes(1 To entryCount):  ReDim dosDates(1 To entryCount)
  ReDim crcs(1 To entryCount):      ReDim compSizes(1 To entryCount)
  ReDim rawSizes(1 To entryCount):  ReDim dataAt(1 To entryCount)
  ReDim localAt(1 To entryCount)

  at = UInt32ToLong(zipBytes, eocd + 16)
  If at < 0 Then
    SetError "[RepackageOOXML] Central directory offset is out of range"
    Exit Function
  End If

  ' Deflate the replacement if libarchive will; otherwise store it. A stored
  ' entry is valid and Office reads it, but vbaProject.bin is the bulk of a
  ' macro-enabled file and storing it roughly doubles the package.
  newLen = ByteCount(newBin)
  repCrc = Crc32(newBin)

  If TryDeflateRaw(newBin, repPayload) Then
    repMethod = ZIP_METHOD_DEFLATE
    repCompLen = ByteCount(repPayload)
  End If

  ' No deflate available, or one that did not pay for itself.
  If repMethod <> ZIP_METHOD_DEFLATE Or repCompLen >= newLen Then
    repMethod = ZIP_METHOD_STORE
    repPayload = newBin
    repCompLen = newLen
  End If

  For i = 1 To entryCount
    If at + 46 > UBound(zipBytes) Then
      SetError "[RepackageOOXML] Central directory is truncated at entry " & i
      Exit Function
    End If
    If UInt32ToLong(zipBytes, at) <> &H2014B50 Then
      SetError "[RepackageOOXML] Central directory record " & i & " has a bad signature"
      Exit Function
    End If

    flags(i) = ReadUInt16LE(zipBytes, at + 8)
    If (flags(i) And 1) <> 0 Then
      SetError "[RepackageOOXML] Encrypted entries are not supported"
      Exit Function
    End If

    methods(i) = ReadUInt16LE(zipBytes, at + 10)
    dosTimes(i) = ReadUInt16LE(zipBytes, at + 12)
    dosDates(i) = ReadUInt16LE(zipBytes, at + 14)
    crcs(i) = UInt32ToRaw(zipBytes, at + 16)
    compSizes(i) = UInt32ToLong(zipBytes, at + 20)
    rawSizes(i) = UInt32ToLong(zipBytes, at + 24)
    nameLens(i) = ReadUInt16LE(zipBytes, at + 28)
    names(i) = at + 46
    localAt(i) = UInt32ToLong(zipBytes, at + 42)

    If compSizes(i) < 0 Or rawSizes(i) < 0 Or localAt(i) < 0 Then
      SetError "[RepackageOOXML] Entry " & i & " needs Zip64, which is not supported"
      Exit Function
    End If

    dataAt(i) = LocalEntryDataStart(zipBytes, localAt(i))
    If dataAt(i) < 0 Then
      SetError "[RepackageOOXML] Local header for entry " & i & " is not where the directory says"
      Exit Function
    End If
    If dataAt(i) + compSizes(i) - 1 > UBound(zipBytes) Then
      SetError "[RepackageOOXML] Entry " & i & " runs past the end of the archive"
      Exit Function
    End If

    If replaceIdx = 0 Then
      If EntryNameIsVBAProject(zipBytes, names(i), nameLens(i)) Then replaceIdx = i
    End If

    ' A local header without its sizes (bit 3) puts them in a trailing data
    ' descriptor. The rebuilt headers carry real sizes, so that flag is cleared.
    total = total + 30 + nameLens(i) + 46 + nameLens(i)
    If i = replaceIdx Then
      total = total + repCompLen
    Else
      total = total + compSizes(i)
    End If

    at = at + 46 + nameLens(i) + ReadUInt16LE(zipBytes, at + 30) + ReadUInt16LE(zipBytes, at + 32)
  Next i

  If replaceIdx = 0 Then
    SetError "[RepackageOOXML] No vbaProject.bin entry found in " & ContainerPath
    Exit Function
  End If

  ReDim outBytes(0 To total + 22 - 1)

  ' Local headers and payloads.
  For i = 1 To entryCount
    Dim thisFlags As Long
    Dim thisCrc   As Long
    Dim thisComp  As Long
    Dim thisRaw   As Long

    thisFlags = flags(i) And Not 8&
    If i = replaceIdx Then
      thisCrc = repCrc
      thisComp = repCompLen
      thisRaw = newLen
    Else
      thisCrc = crcs(i)
      thisComp = compSizes(i)
      thisRaw = rawSizes(i)
    End If

    localAt(i) = outPos                     ' reused: now the offset in the output
    PutUInt32LE outBytes, outPos, &H4034B50
    PutUInt16LE outBytes, outPos + 4, 20
    PutUInt16LE outBytes, outPos + 6, thisFlags
    PutUInt16LE outBytes, outPos + 8, IIf(i = replaceIdx, repMethod, methods(i))
    PutUInt16LE outBytes, outPos + 10, dosTimes(i)
    PutUInt16LE outBytes, outPos + 12, dosDates(i)
    PutUInt32LE outBytes, outPos + 14, thisCrc
    PutUInt32LE outBytes, outPos + 18, thisComp
    PutUInt32LE outBytes, outPos + 22, thisRaw
    PutUInt16LE outBytes, outPos + 26, nameLens(i)
    PutUInt16LE outBytes, outPos + 28, 0        ' extra field dropped
    outPos = outPos + 30

    CopyMemory VarPtr(outBytes(outPos)), VarPtr(zipBytes(names(i))), nameLens(i)
    outPos = outPos + nameLens(i)

    If i = replaceIdx Then
      If repCompLen > 0 Then
        CopyMemory VarPtr(outBytes(outPos)), VarPtr(repPayload(LBound(repPayload))), repCompLen
        outPos = outPos + repCompLen
      End If
    ElseIf compSizes(i) > 0 Then
      CopyMemory VarPtr(outBytes(outPos)), VarPtr(zipBytes(dataAt(i))), compSizes(i)
      outPos = outPos + compSizes(i)
    End If
  Next i

  ' Central directory.
  dirOutStart = outPos
  For i = 1 To entryCount
    PutUInt32LE outBytes, outPos, &H2014B50
    PutUInt16LE outBytes, outPos + 4, 20
    PutUInt16LE outBytes, outPos + 6, 20
    PutUInt16LE outBytes, outPos + 8, flags(i) And Not 8&
    PutUInt16LE outBytes, outPos + 10, IIf(i = replaceIdx, repMethod, methods(i))
    PutUInt16LE outBytes, outPos + 12, dosTimes(i)
    PutUInt16LE outBytes, outPos + 14, dosDates(i)
    If i = replaceIdx Then
      PutUInt32LE outBytes, outPos + 16, repCrc
      PutUInt32LE outBytes, outPos + 20, repCompLen
      PutUInt32LE outBytes, outPos + 24, newLen
    Else
      PutUInt32LE outBytes, outPos + 16, crcs(i)
      PutUInt32LE outBytes, outPos + 20, compSizes(i)
      PutUInt32LE outBytes, outPos + 24, rawSizes(i)
    End If
    PutUInt16LE outBytes, outPos + 28, nameLens(i)
    PutUInt16LE outBytes, outPos + 30, 0    ' extra
    PutUInt16LE outBytes, outPos + 32, 0    ' comment
    PutUInt16LE outBytes, outPos + 34, 0    ' disk number
    PutUInt16LE outBytes, outPos + 36, 0    ' internal attributes
    PutUInt32LE outBytes, outPos + 38, 0    ' external attributes
    PutUInt32LE outBytes, outPos + 42, localAt(i)
    outPos = outPos + 46

    CopyMemory VarPtr(outBytes(outPos)), VarPtr(zipBytes(names(i))), nameLens(i)
    outPos = outPos + nameLens(i)
  Next i

  PutUInt32LE outBytes, outPos, &H6054B50
  PutUInt16LE outBytes, outPos + 4, 0
  PutUInt16LE outBytes, outPos + 6, 0
  PutUInt16LE outBytes, outPos + 8, entryCount
  PutUInt16LE outBytes, outPos + 10, entryCount
  PutUInt32LE outBytes, outPos + 12, outPos - dirOutStart
  PutUInt32LE outBytes, outPos + 16, dirOutStart
  PutUInt16LE outBytes, outPos + 20, 0
  outPos = outPos + 22

  If outPos <> ByteCount(outBytes) Then ReDim Preserve outBytes(0 To outPos - 1)

  If Not WriteFile(OutputPath, outBytes) Then
    SetError "[RepackageOOXML] Could not write " & OutputPath
    Exit Function
  End If

  LogMessage xrayInfo, "Repackaged " & entryCount & " zip entries, replacing entry " & replaceIdx
  RepackageOOXML = True
End Function

' Scan back from the end for the end-of-central-directory signature. The record
' is 22 bytes plus a comment of up to 65535, so that is as far back as it can be.
Private Function FindEndOfCentralDirectory(ByRef data() As Byte) As Long
  Dim size   As Long
  Dim lowest As Long
  Dim at     As Long

  FindEndOfCentralDirectory = -1
  size = ByteCount(data)
  If size < 22 Then Exit Function

  lowest = size - 22 - 65535
  If lowest < 0 Then lowest = 0

  For at = size - 22 To lowest Step -1
    If data(at) = &H50 Then
      If data(at + 1) = &H4B And data(at + 2) = &H5 And data(at + 3) = &H6 Then
        FindEndOfCentralDirectory = at
        Exit Function
      End If
    End If
  Next at
End Function

Private Function LocalEntryDataStart(ByRef data() As Byte, ByVal localOffset As Long) As Long
  LocalEntryDataStart = -1
  If localOffset < 0 Or localOffset + 30 > UBound(data) Then Exit Function
  If UInt32ToLong(data, localOffset) <> &H4034B50 Then Exit Function
  LocalEntryDataStart = localOffset + 30 + ReadUInt16LE(data, localOffset + 26) + _
                        ReadUInt16LE(data, localOffset + 28)
End Function

' True when the entry name ends with "/vbaProject.bin", whichever host folder it
' sits in (xl, word or ppt).
Private Function EntryNameIsVBAProject(ByRef data() As Byte, ByVal at As Long, ByVal nameLen As Long) As Boolean
  Const TARGET As String = "/vbaproject.bin"
  Dim name As String
  Dim i    As Long

  If nameLen < Len(TARGET) Then Exit Function
  For i = 0 To nameLen - 1
    name = name & Chr$(data(at + i))
  Next i
  EntryNameIsVBAProject = (LCase$(Right$(name, Len(TARGET))) = TARGET)
End Function

' CRC-32 as the zip format defines it. ROneCOne's version uses LongLong, which is
' 64-bit VBA only; vbaXray still builds under VBA6, so the shifts here are done on
' a signed Long with the sign bit handled separately.
Private Function Crc32(ByRef data() As Byte) As Long
  Dim crc As Long
  Dim i   As Long

  EnsureCrcTable
  crc = -1                                ' &HFFFFFFFF
  If ByteCount(data) > 0 Then
    For i = LBound(data) To UBound(data)
      crc = ShiftRight8(crc) Xor m_CrcTable((crc Xor CLng(data(i))) And &HFF&)
    Next i
  End If
  Crc32 = crc Xor -1
End Function

Private Sub EnsureCrcTable()
  Dim i     As Long
  Dim bit   As Long
  Dim entry As Long

  If m_CrcReady Then Exit Sub
  ReDim m_CrcTable(0 To 255)
  For i = 0 To 255
    entry = i
    For bit = 1 To 8
      If (entry And 1&) <> 0 Then
        entry = &HEDB88320 Xor ShiftRight1(entry)
      Else
        entry = ShiftRight1(entry)
      End If
    Next bit
    m_CrcTable(i) = entry
  Next i
  m_CrcReady = True
End Sub

' Logical, not arithmetic: VBA's \ operator would carry the sign bit down.
Private Function ShiftRight1(ByVal v As Long) As Long
  ShiftRight1 = (v And &H7FFFFFFF) \ 2&
  If (v And &H80000000) <> 0 Then ShiftRight1 = ShiftRight1 Or &H40000000
End Function

Private Function ShiftRight8(ByVal v As Long) As Long
  ShiftRight8 = (v And &H7FFFFFFF) \ &H100&
  If (v And &H80000000) <> 0 Then ShiftRight8 = ShiftRight8 Or &H800000
End Function

' A zip 32-bit field as a Long, or -1 when it will not fit one (Zip64 territory).
Private Function UInt32ToLong(ByRef data() As Byte, ByVal pos As Long) As Long
  Dim v As Double
  v = ReadUInt32LE(data, pos)
  If v > 2147483647# Then
    UInt32ToLong = -1
  Else
    UInt32ToLong = CLng(v)
  End If
End Function

' The same four bytes reinterpreted, keeping all 32 bits. The CRC needs this: its
' top bit carries data, and UInt32ToLong would report that as out of range.
Private Function UInt32ToRaw(ByRef data() As Byte, ByVal pos As Long) As Long
  CopyMemory VarPtr(UInt32ToRaw), VarPtr(data(pos)), 4
End Function

Private Sub PutUInt16LE(ByRef buf() As Byte, ByVal pos As Long, ByVal value As Long)
  buf(pos) = value And &HFF&
  buf(pos + 1) = (value \ 256&) And &HFF&
End Sub

Private Sub PutUInt32LE(ByRef buf() As Byte, ByVal pos As Long, ByVal value As Long)
  CopyMemory VarPtr(buf(pos)), VarPtr(value), 4
End Sub

' Deflate a buffer into the raw deflate stream a zip method-8 entry stores.
'
' libarchive exposes no raw-deflate writer, but it does have a gzip filter, and a
' gzip member is exactly a raw deflate stream inside a small header and an 8-byte
' trailer. This is the relationship TryZlibInflate already relies on, run the
' other way.
'
' Any failure returns False and the caller stores the entry uncompressed instead,
' which is valid but larger.
Private Function TryDeflateRaw(ByRef data() As Byte, ByRef result() As Byte) As Boolean
#If Win64 Then
  Dim hWrite     As LongPtr
  Dim hEntry     As LongPtr
  Dim buffer()   As Byte
  Dim optBytes() As Byte
  Dim used       As LongPtr
  Dim capacity   As Long
  Dim total      As Long
  Dim gzipLen    As Long
  Dim rawLen     As Long
  Dim headerLen  As Long
  Dim flg        As Long

  On Error GoTo Cleanup   ' a missing DLL or entry point lands here, still False

  total = ByteCount(data)
  If total = 0 Then Exit Function

  ' Deflate can expand incompressible input slightly; leave room for that plus
  ' the gzip framing.
  capacity = total + (total \ 1000&) + 4096
  ReDim buffer(0 To capacity - 1)

  hWrite = archive_write_new()
  If hWrite = 0 Then GoTo Cleanup
  If archive_write_add_filter_gzip(hWrite) <> ARCHIVE_OK Then GoTo Cleanup
  If archive_write_set_format_raw(hWrite) <> ARCHIVE_OK Then GoTo Cleanup

  ' Level 9, and no timestamp, so identical input always deflates to identical
  ' bytes. StringToBytes does not terminate, hence the explicit null.
  optBytes = StringToBytes("gzip:compression-level=9,gzip:timestamp=0" & vbNullChar, 1252)
  If archive_write_set_options(hWrite, VarPtr(optBytes(0))) <> ARCHIVE_OK Then GoTo Cleanup
  If archive_write_open_memory(hWrite, VarPtr(buffer(0)), capacity, used) <> ARCHIVE_OK Then GoTo Cleanup

  hEntry = archive_entry_new()
  If hEntry = 0 Then GoTo Cleanup
  archive_entry_set_size hEntry, CLngLng(total)
  archive_entry_set_filetype hEntry, AE_IFREG
  If archive_write_header(hWrite, hEntry) <> ARCHIVE_OK Then GoTo Cleanup
  If archive_write_data(hWrite, VarPtr(data(LBound(data))), total) <> total Then GoTo Cleanup
  If archive_write_close(hWrite) <> ARCHIVE_OK Then GoTo Cleanup

  gzipLen = CLng(used)
  If gzipLen <= GZIP_HEADER_SIZE + GZIP_TRAILER_SIZE Then GoTo Cleanup
  If buffer(0) <> &H1F Or buffer(1) <> &H8B Or buffer(2) <> 8 Then GoTo Cleanup

  ' Optional fields named by FLG sit between the fixed header and the payload.
  ' libarchive writes none of them here, but the format allows them.
  flg = buffer(3)
  headerLen = GZIP_HEADER_SIZE
  If (flg And &H4) <> 0 Then      ' FEXTRA
    headerLen = headerLen + 2 + (CLng(buffer(headerLen)) Or (CLng(buffer(headerLen + 1)) * 256&))
  End If
  If (flg And &H8) <> 0 Then      ' FNAME, null terminated
    Do While headerLen < gzipLen
      If buffer(headerLen) = 0 Then Exit Do
      headerLen = headerLen + 1
    Loop
    headerLen = headerLen + 1
  End If
  If (flg And &H10) <> 0 Then     ' FCOMMENT, null terminated
    Do While headerLen < gzipLen
      If buffer(headerLen) = 0 Then Exit Do
      headerLen = headerLen + 1
    Loop
    headerLen = headerLen + 1
  End If
  If (flg And &H2) <> 0 Then headerLen = headerLen + 2   ' FHCRC

  rawLen = gzipLen - headerLen - GZIP_TRAILER_SIZE
  If rawLen <= 0 Or headerLen >= gzipLen Then GoTo Cleanup

  ReDim result(0 To rawLen - 1)
  CopyMemory VarPtr(result(0)), VarPtr(buffer(headerLen)), rawLen
  TryDeflateRaw = True

Cleanup:
  On Error Resume Next
  If hEntry <> 0 Then archive_entry_free hEntry
  If hWrite <> 0 Then archive_write_free hWrite
  If Not TryDeflateRaw Then
    LogMessage xrayDebug, "[TryDeflateRaw] Deflate unavailable; the entry will be stored instead"
  End If
#End If
End Function

Public Function IsCompoundFile(ByVal FilePath As String) As Boolean
  Dim header() As Byte
  header = ReadFile(FilePath, 8)
  If ByteCount(header) < 8 Then Exit Function
  IsCompoundFile = IsCompoundFileBytes(header)
End Function

Public Sub DebugDumpStorageTree(Optional ByVal OutputFilePath As String = vbNullString)
  If m_StorageTree Is Nothing Then
    SetError "[DebugDumpStorageTree] No storage loaded."
    Exit Sub
  End If
  Dim output As String
  output = "." & vbCrLf
  TreeNodeRender m_StorageTree, "", output
  If Len(OutputFilePath) > 0 Then
    Dim rawOutput() As Byte
    rawOutput = StrConv(output, vbFromUnicode)
    If WriteFile(OutputFilePath, rawOutput) Then
      LogMessage xrayInfo, "Storage tree written to: " & OutputFilePath
    Else
      SetError "[DebugDumpStorageTree] Failed to write to: " & OutputFilePath
    End If
  Else
    Debug.Print output
  End If
End Sub

Private Function EscapeName(ByVal s As String) As String
  Dim i  As Long
  Dim ch As Long
  Dim out As String
  For i = 1 To Len(s)
    ch = AscW(VBA.Mid$(s, i, 1)) And &HFFFF&
    If ch < 32 Or ch = 127 Then
      out = out & "\x" & Right$("0" & Hex$(ch), 2)
    Else
      out = out & VBA.Mid$(s, i, 1)
    End If
  Next i
  EscapeName = out
End Function

Private Function GetVBAProjectBinBytes(ByVal OfficePath As String) As Byte()
  Dim target    As String
  Dim folderPath As String

  Select Case LCase$(Right$(OfficePath, 4))
    Case "xlsm", "xlam", "xlsb", "xltm"
      folderPath = "xl":   target = "vbaProject.bin"
    Case "docm", "dotm"
      folderPath = "word": target = "vbaProject.bin"
    Case "pptm", "potm", "ppsm", "ppam", "sldm"
      folderPath = "ppt":  target = "vbaProject.bin"
    Case Else
      If LCase$(OfficePath) Like "*vbaproject.bin" Then
        GetVBAProjectBinBytes = ReadFile(OfficePath)
      Else
        SetError "[GetVBAProjectBinBytes] Unrecognised extension"
      End If
      Exit Function
  End Select

  Dim zipStg    As stdole.IUnknown
  Dim folderStg As stdole.IUnknown
  Dim folderRaw As stdole.IUnknown
  Dim pStm      As stdole.IUnknown
  Dim hr As Long
  Set zipStg = StorageOpenZip(OfficePath)
  If Not zipStg Is Nothing Then
    hr = DispCallByVtbl(zipStg, IStorageInterface.IStg_OpenStorage, StrPtr(folderPath), 0&, STGM_READ Or STGM_SHARE_EXCLUSIVE, 0&, 0&, VarPtr(folderRaw))
    If hr = 0 And Not folderRaw Is Nothing Then
      Set folderStg = COMQueryInterface(folderRaw, IID_IStorage)
      If Not folderStg Is Nothing Then
        Set pStm = StorageOpenStream(folderStg, target)
        If Not pStm Is Nothing Then
          GetVBAProjectBinBytes = StreamReadBytes(pStm)
          Exit Function
        End If
      End If
    End If
    SetError "[GetVBAProjectBinBytes] ZipFldr failed for: " & OfficePath
  End If

End Function

' NB: PowerPoint stores VBA inside a VbaProjectStg record (recType &H1011) in the
' "PowerPoint Document" stream. recInstance 0 is raw CFB and instance 1 is zlib.
Private Function LoadFromLegacyPowerPoint(ByVal FilePath As String) As Boolean
  On Error GoTo Failed
  LoadFromLegacyPowerPoint = False

  Dim outerStorage As stdole.IUnknown
  Dim pptStream    As stdole.IUnknown
  Set outerStorage = StorageOpenFile(FilePath)
  If outerStorage Is Nothing Then Exit Function

  Set pptStream = StorageOpenStream(outerStorage, "PowerPoint Document")
  If pptStream Is Nothing Then
    SetError "[LoadFromLegacyPowerPoint] Missing 'PowerPoint Document' stream"
    Exit Function
  End If

  Dim pptData() As Byte
  pptData = StreamReadBytes(pptStream)
  If ByteCount(pptData) < 8 Then
    SetError "[LoadFromLegacyPowerPoint] 'PowerPoint Document' stream is empty or invalid"
    Exit Function
  End If

  Dim lb As Long
  Dim ub As Long
  lb = LBound(pptData)
  ub = UBound(pptData)

  Dim pos            As Long
  Dim recVerInstance As Long
  Dim recVer         As Long
  Dim recInstance    As Long
  Dim recLenValue    As Double
  Dim recLen         As Long
  Dim payloadPos     As Long
  Dim expectedValue  As Double
  Dim ExpectedSize   As Long
  Dim compressedSize As Long
  Dim candidateReady As Boolean
  Dim candidate()    As Byte
  Dim zlibData()     As Byte

  For pos = lb To ub - 7
    If pptData(pos + 2) = (PPT_VBA_PROJECT_STG And &HFF&) And pptData(pos + 3) = ((PPT_VBA_PROJECT_STG \ 256&) And &HFF&) Then
      recVerInstance = ReadUInt16LE(pptData, pos)
      recVer = recVerInstance And &HF&
      recInstance = recVerInstance \ &H10&

      If recVer = 0 And (recInstance = 0 Or recInstance = 1) Then
        recLenValue = ReadUInt32LE(pptData, pos + 4)
        payloadPos = pos + 8

        If recLenValue >= 0# And recLenValue <= MAX_EMBEDDED_PROJECT_SIZE And recLenValue <= CDbl(ub - payloadPos + 1) Then

          recLen = CLng(recLenValue)
          candidateReady = False
          Erase candidate

          Select Case recInstance
            Case 0
              If recLen >= 8 Then
                If HasCompoundFileSignatureAt(pptData, payloadPos) Then
                  candidate = SliceBytes(pptData, payloadPos, recLen)
                  candidateReady = (ByteCount(candidate) = recLen)
                End If
              End If

            Case 1
              ' Four-byte decompressed-size prefix followed by PowerPoint's
              ' zlib-compatible DEFLATE payload.
              If recLen >= 12 Then
                expectedValue = ReadUInt32LE(pptData, payloadPos)
                compressedSize = recLen - 4

                If expectedValue >= 8# And expectedValue <= MAX_EMBEDDED_PROJECT_SIZE And compressedSize >= 8 Then
                  ExpectedSize = CLng(expectedValue)
                  zlibData = SliceBytes(pptData, payloadPos + 4, compressedSize)
                  If ByteCount(zlibData) = compressedSize Then
                    candidateReady = TryZlibInflate(zlibData, ExpectedSize, candidate)
                  End If
                  Erase zlibData
                End If
              End If
          End Select

          If candidateReady Then
            If IsCompoundFileBytes(candidate) Then
              If LoadFromByteArray(candidate) Then
                LoadFromLegacyPowerPoint = True
                LogMessage xrayInfo, "Legacy PowerPoint VBA project found (" & IIf(recInstance = 0, "raw", "compressed") & ") at stream offset " & pos
                Exit Function
              End If
            End If
          End If
        End If
      End If
    End If
  Next pos

  Reset
  SetError "[LoadFromLegacyPowerPoint] No valid VBA project found"
  Exit Function

Failed:
  Dim errorNumber As Long
  Dim errorDescription As String
  errorNumber = Err.Number
  errorDescription = Err.Description
  Reset
  SetError "[LoadFromLegacyPowerPoint] Error " & errorNumber & " - " & errorDescription
End Function

' Inflates the zlib-compatible payload used by compressed VbaProjectStg
' records in legacy PowerPoint files. Uses archiveint.dll under the hood.
Private Function TryZlibInflate(ByRef zlibData() As Byte, ByVal ExpectedSize As Long, ByRef result() As Byte) As Boolean
  Dim archive       As LongPtr
  Dim entry         As LongPtr
  Dim blockPtr      As LongPtr
  Dim blockSizePtr  As LongPtr
  #If Win64 Then
    Dim blockOffset As LongLong
  #Else
    Dim blockOffset As Currency
  #End If
  Dim blockOffsetValue As Long
  Dim nextOffset    As Long
  Dim status        As Long
  Dim opened        As Boolean
  Dim succeeded     As Boolean
  Dim work()        As Byte
  Dim gzipData()    As Byte

  On Error GoTo Failed
  Erase result

  Dim zlibSize As Long
  Dim lb       As Long
  zlibSize = ByteCount(zlibData)
  If zlibSize < 8 Then Exit Function
  If ExpectedSize < 8 Or ExpectedSize > MAX_EMBEDDED_PROJECT_SIZE Then Exit Function
  lb = LBound(zlibData)

  Dim cmf As Long
  Dim flg As Long
  cmf = zlibData(lb)
  flg = zlibData(lb + 1)

  If (cmf And &HF&) <> 8 Then Exit Function
  If (cmf \ &H10&) > 7 Then Exit Function
  If ((cmf * 256&) + flg) Mod 31 <> 0 Then Exit Function
  If (flg And &H20&) <> 0 Then Exit Function

  ' Preserve every byte after the zlib header. This is intentional: the
  ' working PPT variant does not expose a dependable four-byte Adler trailer.
  Dim rawDeflateSize As Long
  rawDeflateSize = zlibSize - 2
  If rawDeflateSize <= 0 Then Exit Function

  Dim bodySize As Long
  bodySize = rawDeflateSize + 5
  ReDim gzipData(0 To GZIP_HEADER_SIZE + bodySize + GZIP_TRAILER_SIZE - 1)
  gzipData(0) = &H1F
  gzipData(1) = &H8B
  gzipData(2) = &H8
  gzipData(9) = &HFF
  CopyMemory VarPtr(gzipData(GZIP_HEADER_SIZE)), VarPtr(zlibData(lb + 2)), rawDeflateSize

  Dim finalBlock As Long
  finalBlock = GZIP_HEADER_SIZE + rawDeflateSize
  gzipData(finalBlock) = &H1
  gzipData(finalBlock + 1) = &H0
  gzipData(finalBlock + 2) = &H0
  gzipData(finalBlock + 3) = &HFF
  gzipData(finalBlock + 4) = &HFF

  archive = archive_read_new()
  If archive = 0 Then GoTo CleanUp
  If archive_read_support_format_raw(archive) <> ARCHIVE_OK Then GoTo CleanUp
  If archive_read_support_filter_all(archive) <> ARCHIVE_OK Then GoTo CleanUp
  If archive_read_open_memory(archive, VarPtr(gzipData(0)), ByteCount(gzipData)) <> ARCHIVE_OK Then GoTo CleanUp
  opened = True
  If archive_read_next_header(archive, entry) <> ARCHIVE_OK Then GoTo CleanUp

  ReDim work(0 To ExpectedSize - 1)

  Do
    blockPtr = 0
    blockSizePtr = 0
    blockOffset = 0
    status = archive_read_data_block(archive, blockPtr, blockSizePtr, blockOffset)

    If status = ARCHIVE_EOF Then Exit Do
    If blockPtr = 0 Or blockSizePtr <= 0 Then Exit Do

    #If Win64 Then
      If blockOffset < 0 Or blockOffset > ExpectedSize Then GoTo CleanUp
      blockOffsetValue = CLng(blockOffset)
    #Else
      If blockOffset < 0@ Or blockOffset * 10000@ > ExpectedSize Then GoTo CleanUp
      blockOffsetValue = CLng(blockOffset * 10000@)
    #End If

    If blockOffsetValue <> nextOffset Then GoTo CleanUp
    If blockSizePtr > ExpectedSize - blockOffsetValue Then GoTo CleanUp

    CopyMemory VarPtr(work(0)) + blockOffsetValue, blockPtr, blockSizePtr
    nextOffset = blockOffsetValue + CLng(blockSizePtr)

    If nextOffset = ExpectedSize Then Exit Do
    If status <> ARCHIVE_OK Then Exit Do
  Loop

  If nextOffset <> ExpectedSize Then GoTo CleanUp
  succeeded = True

CleanUp:
  On Error Resume Next
  If opened Then archive_read_close archive
  If archive <> 0 Then archive_read_free archive
  On Error GoTo 0

  If succeeded Then
    result = work
    TryZlibInflate = True
  Else
    Erase result
  End If
  Exit Function

Failed:
  SetError "[TryZlibInflate] Error " & Err.Number & " - " & Err.Description
  Resume CleanUp
End Function

Private Sub LoadFromStorage(ByVal topStg As stdole.IUnknown, ByVal ext As String)
  
  On Error GoTo Failed
  Reset
  If topStg Is Nothing Then
    SetError "[LoadFromStorage] Invalid storage object"
    Exit Sub
  End If

  Dim topItems As Collection
  Set topItems = StorageEnumElements(topStg)
  Set m_StorageTree = BuildTreeNodes(topStg, topItems)

  Dim layout As CFBLayout
  layout = GetCFBLayout(ext)
  Dim keepAlive As New Collection
  Dim vbaStg As stdole.IUnknown
  Set vbaStg = StorageNavigate(topStg, layout.VBAPath, keepAlive)

  If vbaStg Is Nothing Then
    SetError "[LoadFromStorage] VBA storage not found at '" & layout.VBAPath & "'"
    GoTo CleanUp
  End If

  Dim vbaItems As Collection
  Set vbaItems = StorageEnumElements(vbaStg)

  If Not CollectionContains(vbaItems, "dir") Then
    SetError "[LoadFromStorage] Missing 'dir' stream in '" & layout.VBAPath & "'"
    GoTo CleanUp
  End If

  Dim dirBytes()  As Byte
  Dim decompDir() As Byte
  dirBytes = StreamReadBytes(StorageOpenStream(vbaStg, "dir"))

  If ByteCount(dirBytes) < 11 Then
    SetError "[LoadFromStorage] 'dir' stream too small"
    GoTo CleanUp
  End If

  decompDir = DecompressStream(dirBytes)
  If ByteCount(decompDir) < 11 Then
    SetError "[LoadFromStorage] Decompressed 'dir' stream is invalid"
    GoTo CleanUp
  End If

  ParseDirStream decompDir

  Dim ProjectBytes() As Byte
  ProjectBytes = ReadProjectStream(topStg, layout.projectPath, keepAlive)
  If ByteCount(ProjectBytes) > 0 Then
    ProcessProject BytesToString(ProjectBytes, LBound(ProjectBytes), ByteCount(ProjectBytes), m_Project.CodePage)
  End If

  If m_Project.ModuleCount = 0 Then
    SetError "[LoadFromStorage] No modules found in dir stream"
    GoTo CleanUp
  End If

  Dim i As Long
  For i = 1 To m_Project.ModuleCount
    ReadModuleSource vbaStg, vbaItems, i
  Next i

  LogMessage xrayInfo, "Project loaded: " & m_Project.Name & " (" & m_Project.ModuleCount & " modules)"

CleanUp:
  On Error Resume Next
  Set keepAlive = Nothing
  Set vbaStg = Nothing
  Exit Sub

Failed:
  Dim errorNumber As Long
  Dim errorDescription As String
  errorNumber = Err.Number
  errorDescription = Err.Description
  Reset
  SetError "[LoadFromStorage] Error " & errorNumber & " - " & errorDescription
  Resume CleanUp
End Sub

Private Function GetCFBLayout(ByVal ext As String) As CFBLayout
  Dim layout As CFBLayout
  Select Case LCase$(ext)
    Case "xls", "xla", "xlt"
      layout.VBAPath = "_VBA_PROJECT_CUR\VBA"
      layout.projectPath = "_VBA_PROJECT_CUR\PROJECT"
    Case "doc", "dot"
      layout.VBAPath = "Macros\VBA"
      layout.projectPath = "Macros\PROJECT"
    Case "pub"
      layout.VBAPath = "VBA\VBA"
      layout.projectPath = "VBA\PROJECT"
    Case "otm"
      layout.VBAPath = "OutlookVbaData\VBA"
      layout.projectPath = "OutlookVbaData\PROJECT"
    Case Else
      layout.VBAPath = "VBA"
      layout.projectPath = "PROJECT"
  End Select
  GetCFBLayout = layout
End Function

' keepAlive holds every storage opened along the way. Releasing an intermediate
' early leaves the child reverted, so callers pass a collection that outlives the
' returned storage. It is per-call, so the read and write modes never share one.
Private Function StorageNavigate(ByRef rootStg As stdole.IUnknown, ByVal path As String, _
                                 ByRef keepAlive As Collection, _
                                 Optional ByVal writable As Boolean = False) As stdole.IUnknown

  Dim parts()     As String
  Dim current     As stdole.IUnknown
  Dim next_stg    As stdole.IUnknown
  Dim walkedPath  As String
  Dim pathKey     As String
  Dim i           As Long

  parts = Split(path, "\")
  Set current = rootStg

  For i = 0 To UBound(parts)
    If Len(parts(i)) = 0 Then GoTo Failed

    If Len(walkedPath) = 0 Then
      walkedPath = parts(i)
    Else
      walkedPath = walkedPath & "\" & parts(i)
    End If

    pathKey = "stg:" & LCase$(walkedPath)

    Set next_stg = Nothing
    On Error Resume Next
    Set next_stg = keepAlive.item(pathKey)
    Err.Clear
    On Error GoTo 0

    If next_stg Is Nothing Then
      If writable Then
        Set next_stg = StorageOpenStorageRW(current, parts(i))
      Else
        Set next_stg = StorageOpenStorage(current, parts(i))
      End If

      If next_stg Is Nothing Then
        LogMessage xrayDebug, "[StorageNavigate] Could not open '" & parts(i) & "' in path '" & path & "'"
        GoTo Failed
      End If

      keepAlive.Add next_stg, pathKey
    End If

    Set current = next_stg
  Next i

  Set StorageNavigate = current
  Exit Function

Failed:
  Set StorageNavigate = Nothing
End Function

Private Function ReadProjectStream(ByRef rootStg As stdole.IUnknown, ByVal projectPath As String, ByRef keepAlive As Collection) As Byte()
  
  Dim lastSlash  As Long
  Dim streamName As String
  Dim parentPath As String
  Dim parentStg  As stdole.IUnknown

  lastSlash = InStrRev(projectPath, "\")
  streamName = VBA.Mid$(projectPath, lastSlash + 1)

  If lastSlash > 0 Then
    parentPath = Left$(projectPath, lastSlash - 1)
    Set parentStg = StorageNavigate(rootStg, parentPath, keepAlive)
  Else
    Set parentStg = rootStg
  End If

  If parentStg Is Nothing Then Exit Function

  Dim pStm As stdole.IUnknown
  Set pStm = StorageOpenStream(parentStg, streamName)
  If pStm Is Nothing Then
    LogMessage xrayWarning, "[ReadProjectStream] '" & projectPath & "' not found - module types may be unrefined"
    Exit Function
  End If

  ReadProjectStream = StreamReadBytes(pStm)
End Function

Private Sub ReadModuleSource(ByRef vbaStg As stdole.IUnknown, ByRef vbaItems As Collection, ByVal index As Long)
  
  If vbaStg Is Nothing Then Exit Sub
  If index < 1 Or index > m_Project.ModuleCount Then Exit Sub

  With m_Project.Modules(index)
    If Len(.streamName) = 0 Then Exit Sub
    If Not CollectionContains(vbaItems, .streamName) Then
      SetError "[ReadModuleSource] Stream '" & .streamName & "' not found for module " & .Name
      Exit Sub
    End If

    Dim pStm    As stdole.IUnknown
    Set pStm = StorageOpenStream(vbaStg, .streamName)
    If pStm Is Nothing Then
      SetError "[ReadModuleSource] Could not open stream '" & .streamName & "' for module " & .Name
      Exit Sub
    End If

    Dim modBytes() As Byte
    modBytes = StreamReadBytes(pStm)
    Set pStm = Nothing

    If ByteCount(modBytes) = 0 Then Exit Sub
    If .offset > UBound(modBytes) Then
      SetError "[ReadModuleSource] Offset (" & .offset & ") exceeds stream length for module " & .Name
      Exit Sub
    End If

    Dim compLen As Long
    compLen = UBound(modBytes) - .offset + 1
    If compLen <= 0 Then Exit Sub

    Dim compSource() As Byte
    compSource = SliceBytes(modBytes, .offset, compLen)

    Dim data() As Byte
    data = DecompressStream(compSource)
    If ByteCount(data) = 0 Then Exit Sub
    .SourceCode = BytesToString(data, LBound(data), ByteCount(data), m_Project.CodePage)
  End With
End Sub

Private Function WriteModule(ByVal index As Long, ByVal ExportPath As String) As Boolean
  
  With m_Project.Modules(index)
    Dim Extension   As String
    Dim SubFolder   As String
    Dim header      As String
    header = vbNullString

    Select Case .ModType
      Case modStandard
        Extension = ".bas"
        SubFolder = "01_Standard"
      Case modClass
        Extension = ".cls"
        SubFolder = "02_Class"
        header = "VERSION 1.0 CLASS" & vbCrLf & _
                 "BEGIN" & vbCrLf & _
                 "  MultiUse = -1  'True" & vbCrLf & _
                 "END" & vbCrLf & vbCrLf
      Case modDocument
        Extension = ".cls"
        SubFolder = "03_Document"
      Case modDesigner
        Extension = ".frm"
        SubFolder = "04_Designer"
      Case Else
        Extension = ".txt"
        SubFolder = "05_Other"
    End Select

    Dim TargetPath As String
    ExportPath = VBA.Trim$(ExportPath)
    If Right$(ExportPath, 1) <> "\" Then ExportPath = ExportPath & "\"
    TargetPath = ExportPath & SubFolder & "\"
    
    If Not FolderExists(TargetPath) Then
      If Not CreateFolderRecursive(TargetPath) Then
        SetError "[WriteModule] Could not create folder: " & TargetPath
        Exit Function
      End If
    End If

    Dim fileData() As Byte
    fileData = StringToBytes(header & .SourceCode, m_Project.CodePage)
    If ByteCount(fileData) = 0 Then
      SetError "[WriteModule] No output data produced for " & .Name
      Exit Function
    End If

    If Not WriteFile(TargetPath & .Name & Extension, fileData) Then
      SetError "[WriteModule] Failed to write " & TargetPath & .Name & Extension
      Exit Function
    End If

    LogMessage xrayInfo, "[Exported] " & SubFolder & "\" & .Name & Extension
    WriteModule = True
  End With

End Function

Private Sub ParseDirStream(ByRef data() As Byte)

  Dim tempList() As ModuleRecord
  ReDim tempList(0 To 63)
  Dim count As Long: count = 0

  m_Project.CodePage = 1252

  Dim pos        As Long: pos = 0
  Dim RecordId   As Long
  Dim RecordSize As Long

  Do While pos < UBound(data) - 5
    RecordId = CLng(data(pos)) Or (CLng(data(pos + 1)) * 256&)

    ' MS-OVBA 2.3.4.2.1.5 PROJECTCODEPAGE (0x0003)
    If RecordId = &H3 Then
      RecordSize = CLng(data(pos + 2)) Or (CLng(data(pos + 3)) * 256&) Or _
                   (CLng(data(pos + 4)) * 65536) Or (CLng(data(pos + 5)) * 16777216)
      If RecordSize = 2 Then
        m_Project.CodePage = CLng(data(pos + 6)) Or (CLng(data(pos + 7)) * 256&)
      End If
    End If

    ' Scan for MODULENAME record (0x0019) which opens each module block
    If RecordId = &H19 Then
      RecordSize = CLng(data(pos + 2)) Or (CLng(data(pos + 3)) * 256&) Or _
                   (CLng(data(pos + 4)) * 65536) Or (CLng(data(pos + 5)) * 16777216)
      If RecordSize > 0 And RecordSize < 256 Then
        Dim modRec As ModuleRecord

        Dim modPos As Long
        modPos = pos

        Do While modPos < UBound(data) - 5
          Dim m_Id   As Long
          Dim mSize As Long
          m_Id = CLng(data(modPos)) Or (CLng(data(modPos + 1)) * 256&)
          mSize = CLng(data(modPos + 2)) Or (CLng(data(modPos + 3)) * 256&) Or _
                  (CLng(data(modPos + 4)) * 65536) Or (CLng(data(modPos + 5)) * 16777216)
          modPos = modPos + 6
          If modPos + mSize > UBound(data) + 1 Then Exit Do

          Select Case m_Id
            Case &H19: modRec.Name = BytesToString(data, modPos, mSize, m_Project.CodePage)
            Case &H1A: modRec.streamName = BytesToString(data, modPos, mSize, m_Project.CodePage)
            Case &H21: modRec.ModType = modStandard
            Case &H22: modRec.ModType = modClass      ' refined later by ProcessProject
            Case &H31
              modRec.offset = CLng(data(modPos)) Or (CLng(data(modPos + 1)) * 256&) Or _
                              (CLng(data(modPos + 2)) * 65536) Or (CLng(data(modPos + 3)) * 16777216)
            Case &H2B   ' MODULETERMINATOR
              modPos = modPos + mSize
              Exit Do
          End Select
          modPos = modPos + mSize
        Loop

        If Len(modRec.streamName) > 0 Then
          If count = UBound(tempList) Then ReDim Preserve tempList(0 To count + 63)
          count = count + 1
          tempList(count) = modRec
        End If
        pos = modPos
        GoTo ContinueScan
      End If
    End If

    pos = pos + 1
ContinueScan:
  Loop

  If count > 0 Then
    ReDim m_Project.Modules(1 To count)
    Dim i As Long
    For i = 1 To count
      m_Project.Modules(i) = tempList(i)
    Next i
    m_Project.ModuleCount = count
  End If
End Sub

Private Sub ProcessProject(ByVal StreamString As String)
  StreamString = Replace(StreamString, vbCr, vbLf)
  StreamString = Replace(StreamString, vbLf & vbLf, vbLf)

  Dim lines()  As String
  Dim Key      As String
  Dim modName  As String
  Dim idx      As Long
  Dim i        As Long
  Dim eqPos    As Long

  lines = Split(StreamString, vbLf)
  Dim line As Variant
  For Each line In lines
    line = Trim$(line)
    If Len(line) = 0 Then GoTo NextLine
    If Left$(line, 1) = "'" Then GoTo NextLine

    eqPos = InStr(line, "=")
    If eqPos < 2 Then GoTo NextLine
    Key = LCase$(Trim$(Left$(line, eqPos - 1)))
    modName = Trim$(VBA.Mid$(line, eqPos + 1))

    If Left$(modName, 1) = Chr(34) Then modName = Split(modName, Chr(34))(1)
    Select Case Key
      Case "name"
        m_Project.Name = modName
        GoTo NextLine
      Case "description"
        m_Project.Description = modName
        GoTo NextLine
      Case "dpb"
        Dim passwordProtected As Boolean
        If CheckProjectLock(modName, passwordProtected) Then
          m_Project.IsPasswordProtected = passwordProtected
        Else
          ' Apparently, a valid unprotected DPB can contain 0 to 3 ignored bytes
          ' and is (at most) 22 hex characters.
          m_Project.IsPasswordProtected = (Len(modName) > 22)
        End If
        GoTo NextLine
    End Select

    If Key = "document" Then
      Dim slashPos As Long
      slashPos = InStr(modName, "/")
      If slashPos > 0 Then modName = Left$(modName, slashPos - 1)
    End If

    idx = 0
    For i = 1 To m_Project.ModuleCount
      If LCase$(m_Project.Modules(i).Name) = LCase$(modName) Or _
         LCase$(m_Project.Modules(i).streamName) = LCase$(modName) Then
        idx = i: Exit For
      End If
    Next i
    If idx = 0 Then GoTo NextLine

    With m_Project.Modules(idx)
      Select Case Key
        Case "document"
          .ModType = modDocument
          .IsForm = False
        Case "baseclass"
          .ModType = modDesigner
          .IsForm = True
        Case "class"
          If .ModType = modOther Then .ModType = modClass
        Case "module"
          If .ModType = modOther Then .ModType = modStandard
      End Select
    End With
NextLine:
  Next
End Sub

Private Function CheckProjectLock(ByVal encodedHex As String, ByRef isProtected As Boolean) As Boolean
  CheckProjectLock = False
  isProtected = False

  encodedHex = VBA.Trim$(encodedHex)
  If Len(encodedHex) < 16 Or (Len(encodedHex) And 1) <> 0 Then Exit Function
  If Len(encodedHex) > 8192 Then Exit Function

  Dim encrypted() As Byte
  Dim byteLen     As Long
  Dim i           As Long
  Dim hi          As Long
  Dim lo          As Long
  byteLen = Len(encodedHex) \ 2
  ReDim encrypted(0 To byteLen - 1)

  For i = 0 To byteLen - 1
    hi = HexNibble(AscW(VBA.Mid$(encodedHex, (i * 2) + 1, 1)))
    lo = HexNibble(AscW(VBA.Mid$(encodedHex, (i * 2) + 2, 1)))
    If hi < 0 Or lo < 0 Then Exit Function
    encrypted(i) = CByte((hi * 16) Or lo)
  Next i

  Dim seed          As Long
  Dim Version       As Long
  Dim projectKey    As Long
  Dim ignoredLength As Long
  seed = encrypted(0)
  Version = seed Xor encrypted(1)
  If Version <> 2 Then Exit Function
  projectKey = seed Xor encrypted(2)
  ignoredLength = (seed And 6) \ 2

  Dim pos              As Long
  Dim encryptedByte1   As Long
  Dim encryptedByte2   As Long
  Dim unencryptedByte1 As Long
  Dim decodedByte      As Long
  encryptedByte1 = encrypted(2)
  encryptedByte2 = encrypted(1)
  unencryptedByte1 = projectKey
  pos = 3

  If pos + ignoredLength + 4 > byteLen Then Exit Function
  For i = 1 To ignoredLength
    decodedByte = DecodeDataByte(encrypted(pos), encryptedByte1, encryptedByte2, unencryptedByte1)
    pos = pos + 1
  Next i

  Dim dataLengthValue As Double
  dataLengthValue = 0#
  For i = 0 To 3
    decodedByte = DecodeDataByte(encrypted(pos), encryptedByte1, encryptedByte2, unencryptedByte1)
    dataLengthValue = dataLengthValue + CDbl(decodedByte) * (256# ^ i)
    pos = pos + 1
  Next i

  If dataLengthValue < 1# Or dataLengthValue > CDbl(byteLen - pos) Then Exit Function
  If dataLengthValue <> CDbl(byteLen - pos) Then Exit Function

  Dim firstDataByte As Long
  For i = 1 To CLng(dataLengthValue)
    decodedByte = DecodeDataByte(encrypted(pos), encryptedByte1, _
                                    encryptedByte2, unencryptedByte1)
    If i = 1 Then firstDataByte = decodedByte
    pos = pos + 1
  Next i

  isProtected = Not (dataLengthValue = 1# And firstDataByte = 0)
  CheckProjectLock = True
End Function

Private Function DecodeDataByte(ByVal encryptedByte As Byte, ByRef encryptedByte1 As Long, ByRef encryptedByte2 As Long, ByRef unencryptedByte1 As Long) As Long
  Dim decoded As Long
  decoded = CLng(encryptedByte) Xor ((encryptedByte2 + unencryptedByte1) And &HFF&)

  encryptedByte2 = encryptedByte1
  encryptedByte1 = encryptedByte
  unencryptedByte1 = decoded
  DecodeDataByte = decoded
End Function

Private Function HexNibble(ByVal charCode As Long) As Long
  Select Case charCode
    Case 48 To 57:  HexNibble = charCode - 48
    Case 65 To 70:  HexNibble = charCode - 55
    Case 97 To 102: HexNibble = charCode - 87
    Case Else:      HexNibble = -1
  End Select
End Function

Private Function BuildTreeNodes(ByRef stg As stdole.IUnknown, ByRef items As Collection) As Collection
  Dim nodes As New Collection
  Dim entry As Variant

  For Each entry In items
    Dim node As Object
    Set node = CreateObject("Scripting.Dictionary")
    node("name") = entry("name")
    node("type") = entry("type")
    node("size") = entry("size")

    If entry("type") = STGTY_STORAGE Then
      Dim childStg   As stdole.IUnknown
      Dim childItems As Collection
      On Error Resume Next
      Set childStg = StorageOpenStorage(stg, entry("name"))
      On Error GoTo 0
      If Not childStg Is Nothing Then
        Set childItems = StorageEnumElements(childStg)
        Set node("children") = BuildTreeNodes(childStg, childItems)
        Set childStg = Nothing
      Else
        Set node("children") = New Collection
      End If
    End If

    nodes.Add node
  Next entry

  Set BuildTreeNodes = nodes
End Function

Private Sub TreeNodeRender(ByRef nodes As Collection, ByVal prefix As String, ByRef output As String)
  Dim i      As Long
  Dim node   As Object
  Dim isLast As Boolean
  Dim branch As String
  Dim childPfx As String

  i = 0
  For Each node In nodes
    i = i + 1
    isLast = (i = nodes.count)
    branch = IIf(isLast, "\-- ", "+-- ")
    childPfx = prefix & IIf(isLast, "    ", "|   ")

    Select Case node("type")
      Case STGTY_STREAM
        Dim col As String * 50
        col = prefix & branch & EscapeName(node("name"))
        output = output & col & CLng(node("size")) & " bytes" & vbCrLf
      Case STGTY_STORAGE
        output = output & prefix & branch & "[" & EscapeName(node("name")) & "]" & vbCrLf
        If Not IsEmpty(node("children")) Then
          If Not node("children") Is Nothing Then
            TreeNodeRender node("children"), childPfx, output
          End If
        End If
      Case Else
        output = output & prefix & branch & EscapeName(node("name")) & "  (type=" & node("type") & ")" & vbCrLf
    End Select
  Next node
End Sub

Private Function StorageOpenFile(ByVal FilePath As String) As stdole.IUnknown
  If Not IsCompoundFile(FilePath) Then
    SetError "[StorageOpenFile] Not a Compound Document file: " & FilePath
    Exit Function
  End If
  Dim hr As Long
  hr = StgOpenStorage(StrPtr(FilePath), 0, STGM_READ Or STGM_SHARE_DENY_WRITE, 0, 0, StorageOpenFile)
  If hr < 0 Then SetError "[StgOpenStorage] Error: &H" & Hex(hr)
End Function

Private Function StorageOpenBytes(ByRef data() As Byte) As stdole.IUnknown
  Dim lockBytes As stdole.IUnknown
  Set StorageOpenBytes = StorageOpenBytesCore(data, STGM_READ, lockBytes)
End Function

' Open an in-memory copy of a compound file.
'
' When grfMode permits writing, the caller must hold outLockBytes until
' LockBytesToArray has read the result back. The HGLOBAL behind it is
' reallocated whenever the storage grows, so the handle originally handed to
' CreateILockBytesOnHGlobal cannot be cached and reused.
Private Function StorageOpenBytesCore(ByRef data() As Byte, ByVal grfMode As Long, _
                                      ByRef outLockBytes As stdole.IUnknown) As stdole.IUnknown
  Dim hMem     As LongPtr
  Dim lpMem    As LongPtr
  Dim hr       As Long
  Dim dataSize As LongPtr

  Set outLockBytes = Nothing
  dataSize = ByteCount(data)
  If dataSize = 0 Then Exit Function

  hMem = GlobalAlloc(GMEM_MOVEABLE, dataSize)
  If hMem = 0 Then
    SetError "[StorageOpenBytes] GlobalAlloc failed"
    Exit Function
  End If

  lpMem = GlobalLock(hMem)
  If lpMem = 0 Then
    GlobalFree hMem
    SetError "[StorageOpenBytes] GlobalLock failed"
    Exit Function
  End If

  CopyMemory lpMem, VarPtr(data(LBound(data))), dataSize
  GlobalUnlock hMem

  hr = CreateILockBytesOnHGlobal(hMem, 1, outLockBytes)
  If hr < 0 Then
    GlobalFree hMem
    SetError "[StorageOpenBytes] CreateILockBytesOnHGlobal failed: &H" & Hex(hr)
    Exit Function
  End If

  hr = StgOpenStorageOnILockBytes(ObjPtr(outLockBytes), 0, grfMode Or STGM_SHARE_EXCLUSIVE, 0, 0, StorageOpenBytesCore)
  If hr < 0 Then
    Set outLockBytes = Nothing
    SetError "[StorageOpenBytes] StgOpenStorageOnILockBytes failed: &H" & Hex(hr)
  End If
End Function

Private Function StorageOpenStorageRW(ByRef pStg As stdole.IUnknown, ByVal StorageName As String) As stdole.IUnknown
  Dim hr As Long
  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_OpenStorage, StrPtr(StorageName), 0&, STGM_READWRITE Or STGM_SHARE_EXCLUSIVE, 0&, 0&, VarPtr(StorageOpenStorageRW))
  If hr < 0 Then LogMessage xrayDebug, "IStorage.OpenStorage RW(" & StorageName & ") - hResult: &H" & Hex(hr)
End Function

Private Function StorageOpenStreamRW(ByRef pStg As stdole.IUnknown, ByVal streamName As String) As stdole.IUnknown
  Dim hr As Long
  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_OpenStream, StrPtr(streamName), 0&, STGM_READWRITE Or STGM_SHARE_EXCLUSIVE, 0&, VarPtr(StorageOpenStreamRW))
  If hr < 0 Then LogMessage xrayDebug, "IStorage.OpenStream RW(" & streamName & ") - hResult: &H" & Hex(hr)
End Function

' Replace a stream's contents.
'
' SetSize comes first because the replacement is usually shorter than what it
' displaces, and ReadModuleSource reads from MODULEOFFSET to the end of the
' stream. A stale tail left behind would be decompressed as if it were source.
Private Function StreamWriteBytes(ByRef pStm As stdole.IUnknown, ByRef data() As Byte) As Boolean
  Dim hr      As Long
  Dim written As Long
  Dim total   As Long

  If pStm Is Nothing Then Exit Function
  total = ByteCount(data)

  hr = DispCallByVtbl(pStm, IStreamInterface.IStm_SetSize, SizeToCurrency(total))
  If hr < 0 Then
    SetError "[StreamWriteBytes] IStream.SetSize failed: &H" & Hex(hr)
    Exit Function
  End If

  hr = DispCallByVtbl(pStm, IStreamInterface.IStm_Seek, 0@, 0&, 0&)
  If hr < 0 Then
    SetError "[StreamWriteBytes] IStream.Seek failed: &H" & Hex(hr)
    Exit Function
  End If

  If total > 0 Then
    hr = DispCallByVtbl(pStm, IStreamInterface.IStm_Write, VarPtr(data(LBound(data))), total, VarPtr(written))
    If hr < 0 Then
      SetError "[StreamWriteBytes] IStream.Write failed: &H" & Hex(hr)
      Exit Function
    End If
    If written <> total Then
      SetError "[StreamWriteBytes] Short write: " & written & " of " & total & " bytes"
      Exit Function
    End If
  End If

  hr = DispCallByVtbl(pStm, IStreamInterface.IStm_Commit, 0&)
  If hr < 0 Then
    SetError "[StreamWriteBytes] IStream.Commit failed: &H" & Hex(hr)
    Exit Function
  End If

  StreamWriteBytes = True
End Function

Private Function StorageCommit(ByRef pStg As stdole.IUnknown) As Boolean
  Dim hr As Long
  If pStg Is Nothing Then Exit Function
  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_Commit, 0&)
  If hr < 0 Then
    SetError "[StorageCommit] IStorage.Commit failed: &H" & Hex(hr)
    Exit Function
  End If
  StorageCommit = True
End Function

' Read a modified compound file back out of the ILockBytes it was built in.
Private Function LockBytesToArray(ByRef pLock As stdole.IUnknown) As Byte()
  Dim hr       As Long
  Dim stat     As STATSTG
  Dim hMem     As LongPtr
  Dim lpMem    As LongPtr
  Dim total    As Long
  Dim result() As Byte

  If pLock Is Nothing Then Exit Function

  hr = DispCallByVtbl(pLock, ILockBytesInterface.ILkb_stat, VarPtr(stat), STATFLAG_NONAME)
  If hr < 0 Then
    SetError "[LockBytesToArray] ILockBytes.Stat failed: &H" & Hex(hr)
    Exit Function
  End If

  total = CLng(stat.cbSize * 10000#)
  If total <= 0 Then
    SetError "[LockBytesToArray] Rebuilt compound file is empty"
    Exit Function
  End If

  hr = GetHGlobalFromILockBytes(ObjPtr(pLock), hMem)
  If hr < 0 Or hMem = 0 Then
    SetError "[LockBytesToArray] GetHGlobalFromILockBytes failed: &H" & Hex(hr)
    Exit Function
  End If

  lpMem = GlobalLock(hMem)
  If lpMem = 0 Then
    SetError "[LockBytesToArray] GlobalLock failed"
    Exit Function
  End If

  ReDim result(0 To total - 1)
  CopyMemory VarPtr(result(0)), lpMem, total
  GlobalUnlock hMem

  LockBytesToArray = result
End Function

' IStream and ILockBytes take 64-bit sizes. VBA has no unsigned 64-bit type, so
' these travel as Currency, whose raw value is the real number scaled by 10000.
' The read direction is the same conversion inverted, as in StorageEnumElements.
Private Function SizeToCurrency(ByVal n As Long) As Currency
  SizeToCurrency = CCur(n) / 10000@
End Function

Private Function StorageOpenStorage(ByRef pStg As stdole.IUnknown, ByVal StorageName As String) As stdole.IUnknown
  Dim hr As Long
  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_OpenStorage, StrPtr(StorageName), 0&, STGM_READ Or STGM_SHARE_EXCLUSIVE, 0&, 0&, VarPtr(StorageOpenStorage))
  If hr < 0 Then LogMessage xrayDebug, "IStorage.OpenStorage(" & StorageName & ") - hResult: &H" & Hex(hr)
End Function

Private Function StorageOpenStream(ByRef pStg As stdole.IUnknown, ByVal streamName As String) As stdole.IUnknown
  Dim hr As Long
  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_OpenStream, StrPtr(streamName), 0&, STGM_READ Or STGM_SHARE_EXCLUSIVE, 0&, VarPtr(StorageOpenStream))
  If hr < 0 Then LogMessage xrayDebug, "IStorage.OpenStream(" & streamName & ") - hResult: &H" & Hex(hr)
End Function

Private Function StorageEnumElements(ByRef pStg As stdole.IUnknown) As Collection
  Dim pEnum   As stdole.IUnknown
  Dim hr      As Long
  Dim items   As New Collection
  Dim stat    As STATSTG
  Dim entry   As Object
  Dim fetched As Long

  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_EnumElements, 0&, 0&, 0&, VarPtr(pEnum))
  If hr < 0 Or pEnum Is Nothing Then
    LogMessage xrayDebug, "IStorage.EnumElements - hResult: &H" & Hex(hr)
    Set StorageEnumElements = items
    Exit Function
  End If

  Do
    fetched = 0
    hr = DispCallByVtbl(pEnum, IEnumSTATSTGInterface.IEnum_Next, 1&, VarPtr(stat), VarPtr(fetched))
    If hr <> 0 Or fetched = 0 Then Exit Do
    If stat.pwcsName <> 0 Then
      Set entry = CreateObject("Scripting.Dictionary")
      entry("name") = PtrToString(stat.pwcsName)
      entry("type") = stat.dwType
      entry("size") = CLng(stat.cbSize * 10000#)
      CoTaskMemFree stat.pwcsName
      stat.pwcsName = 0
      items.Add entry
    End If
  Loop

  Set StorageEnumElements = items
End Function

Private Function CollectionContains(ByRef items As Collection, ByVal Name As String) As Boolean
  Dim entry As Variant
  For Each entry In items
    If LCase$(entry("name")) = LCase$(Name) Then
      CollectionContains = True
      Exit Function
    End If
  Next
End Function

Private Function StreamReadBytes(ByRef pStm As stdole.IUnknown) As Byte()
  Dim Buffer()          As Byte
  Dim chunk(0 To 32767) As Byte
  Dim bytesRead         As Long
  Dim hr                As Long
  Dim total             As Long
  Dim capacity          As Long
  Dim required          As Long
  Dim newCapacity       As Long

  If pStm Is Nothing Then Exit Function

  hr = DispCallByVtbl(pStm, IStreamInterface.IStm_Seek, 0@, 0&, 0&)
  If hr < 0 Then
    LogMessage xrayDebug, "IStream.Seek - hResult: &H" & Hex(hr)
    Exit Function
  End If

  Do
    hr = DispCallByVtbl(pStm, IStreamInterface.IStm_Read, VarPtr(chunk(0)), 32768, VarPtr(bytesRead))
    If hr < 0 Then
      LogMessage xrayDebug, "IStream.Read - hResult: &H" & Hex(hr)
      Erase Buffer
      Exit Function
    End If
    If bytesRead <= 0 Then Exit Do

    If total > &H7FFFFFFF - bytesRead Then
      SetError "[StreamReadBytes] Stream is too large for a VBA byte array"
      Erase Buffer
      Exit Function
    End If
    required = total + bytesRead

    If required > capacity Then
      If capacity = 0 Then
        newCapacity = 65536
      Else
        newCapacity = capacity
      End If
      Do While newCapacity < required
        If newCapacity > &H3FFFFFFF Then
          newCapacity = required
          Exit Do
        End If
        newCapacity = newCapacity * 2
      Loop

      If capacity = 0 Then
        ReDim Buffer(0 To newCapacity - 1)
      Else
        ReDim Preserve Buffer(0 To newCapacity - 1)
      End If
      capacity = newCapacity
    End If

    CopyMemory VarPtr(Buffer(total)), VarPtr(chunk(0)), bytesRead
    total = required
  Loop

  If total > 0 Then
    ReDim Preserve Buffer(0 To total - 1)
    StreamReadBytes = Buffer
  End If
End Function

' LZ77 decompression per MS-OVBA 2.4.1.
Private Function DecompressStream(ByRef CompressedData() As Byte, Optional ByVal Silent As Boolean = False) As Byte()
  Dim output()    As Byte
  Dim outPos      As Long
  Dim inPos       As Long
  Dim inputLb     As Long
  Dim inputUb     As Long

  If ByteCount(CompressedData) < 3 Then Exit Function
  inputLb = LBound(CompressedData)
  inputUb = UBound(CompressedData)

  If CompressedData(inputLb) <> 1 Then
    If Not Silent Then
      SetError "[DecompressStream] Invalid signature byte: &H" & Hex(CompressedData(inputLb))
    End If
    Exit Function
  End If

  ReDim output(0 To 65535)
  inPos = inputLb + 1

  Do While inPos <= inputUb
    If inPos + 1 > inputUb Then GoTo Failed

    Dim chunkHeader  As Long
    Dim isCompressed As Boolean
    Dim payloadSize  As Long
    Dim chunkEnd     As Long
    Dim chunkOutPos  As Long

    chunkHeader = CLng(CompressedData(inPos)) Or (CLng(CompressedData(inPos + 1)) * 256&)
    inPos = inPos + 2
    If (chunkHeader And &H7000&) <> &H3000& Then GoTo Failed
    isCompressed = ((chunkHeader And &H8000&) = &H8000&)
    payloadSize = (chunkHeader And &HFFF&) + 1
    chunkEnd = inPos + payloadSize
    If chunkEnd > inputUb + 1 Then GoTo Failed
    chunkOutPos = 0

    If Not isCompressed Then
      If outPos + 4096 > UBound(output) Then ReDim Preserve output(0 To outPos + 4096 + 65535)
      Dim i As Long
      Do While inPos < chunkEnd
        output(outPos) = CompressedData(inPos)
        inPos = inPos + 1: outPos = outPos + 1: chunkOutPos = chunkOutPos + 1
      Loop
    Else
      Do While inPos < chunkEnd
        Dim flagByte As Byte
        flagByte = CompressedData(inPos)
        inPos = inPos + 1

        Dim bitIndex As Integer
        For bitIndex = 0 To 7
          If chunkOutPos >= 4096 Then Exit Do
          If inPos >= chunkEnd Then Exit Do

          If (flagByte And (2 ^ bitIndex)) = 0 Then
            If outPos > UBound(output) Then ReDim Preserve output(0 To outPos + 65535)
            output(outPos) = CompressedData(inPos)
            outPos = outPos + 1: chunkOutPos = chunkOutPos + 1: inPos = inPos + 1
          Else
            If inPos + 1 >= chunkEnd Then GoTo Failed
            Dim token As Long
            token = CLng(CompressedData(inPos)) Or (CLng(CompressedData(inPos + 1)) * 256&)
            inPos = inPos + 2

            Dim bitCount As Integer
            Select Case chunkOutPos
              Case Is <= 16:   bitCount = 4
              Case Is <= 32:   bitCount = 5
              Case Is <= 64:   bitCount = 6
              Case Is <= 128:  bitCount = 7
              Case Is <= 256:  bitCount = 8
              Case Is <= 512:  bitCount = 9
              Case Is <= 1024: bitCount = 10
              Case Is <= 2048: bitCount = 11
              Case Else:       bitCount = 12
            End Select

            Dim lengthMask As Long: lengthMask = (2 ^ (16 - bitCount)) - 1
            Dim copyLength As Long: copyLength = (token And lengthMask) + 3
            Dim copyOffset As Long: copyOffset = (token \ (2 ^ (16 - bitCount))) + 1
            If copyOffset > chunkOutPos Or copyOffset > outPos Then GoTo Failed
            If copyLength > 4096 - chunkOutPos Then GoTo Failed

            If outPos + copyLength > UBound(output) Then
              ReDim Preserve output(0 To outPos + copyLength + 65535)
            End If

            Dim j As Long
            For j = 0 To copyLength - 1
              output(outPos) = output(outPos - copyOffset)
              outPos = outPos + 1: chunkOutPos = chunkOutPos + 1
            Next j
          End If
        Next bitIndex
      Loop
    End If

    inPos = chunkEnd
  Loop

  If outPos > 0 Then
    ReDim Preserve output(0 To outPos - 1)
    DecompressStream = output
  End If
  Exit Function

Failed:
  If Not Silent Then
    SetError "[DecompressStream] Invalid or truncated compressed stream"
  End If
End Function

' LZ77 compression per MS-OVBA 2.4.1. The inverse of DecompressStream.
'
' Every chunk is emitted token-compressed. The specification also permits raw
' chunks, but the VBE refuses to load a module whose first chunk is raw - "An
' error occurred while loading <Module>" - even though such a stream decompresses
' correctly. Raw is kept only as an escape for input that cannot be token-encoded
' within 4096 bytes, which does not arise for source text.
'
' Matches come from a greedy scan over a 3-gram index. Candidates are walked in
' ascending position order and only a strictly longer match displaces the
' incumbent, so ties keep the farthest match. Those rules reproduce Office's own
' encoder, so a rewritten module carries the same bytes the VBE would have
' written for the same source. A literal-only encoder is also legal per the
' specification and far simpler, but it emits bytes no Office component ever
' produces, and pyOpenVBA found that Access rejects a module on that basis alone.
Private Function CompressStream(ByRef data() As Byte) As Byte()
  Dim output()   As Byte
  Dim encoded()  As Byte
  Dim outPos     As Long
  Dim dataLb     As Long
  Dim dataLen    As Long
  Dim cursor     As Long
  Dim chunkLen   As Long
  Dim encLen     As Long
  Dim header     As Long
  Dim k          As Long

  dataLen = ByteCount(data)

  ' Signature byte, then per chunk a 2-byte header and a payload that is at worst
  ' the input plus one flag byte per eight tokens.
  ReDim output(0 To dataLen + (dataLen \ 4096 + 1) * 520& + 16)
  output(0) = 1
  outPos = 1

  If dataLen = 0 Then
    ReDim Preserve output(0 To 0)
    CompressStream = output
    Exit Function
  End If

  dataLb = LBound(data)

  Do While cursor < dataLen
    chunkLen = dataLen - cursor
    If chunkLen > 4096 Then chunkLen = 4096

    encLen = EncodeTokenChunk(data, dataLb + cursor, chunkLen, encoded)
    If encLen <= 0 Then
      SetError "[CompressStream] Chunk encoder produced no output at offset " & cursor
      Exit Function
    End If

    If encLen <= 4096 Then
      ' CompressedChunkFlag = 1, signature 0b011, size field is length - 1.
      header = &HB000& Or (encLen - 1)
      output(outPos) = header And &HFF&
      output(outPos + 1) = (header \ 256&) And &HFF&
      outPos = outPos + 2
      For k = 0 To encLen - 1
        output(outPos) = encoded(k)
        outPos = outPos + 1
      Next k
    Else
      ' A raw chunk carries exactly 4096 bytes, so a short final chunk that
      ' overflows has no legal representation.
      If chunkLen <> 4096 Then
        SetError "[CompressStream] Final chunk of " & chunkLen & _
                 " bytes encoded to " & encLen & " bytes and cannot be stored raw"
        Exit Function
      End If
      output(outPos) = &HFF
      output(outPos + 1) = &H3F
      outPos = outPos + 2
      For k = 0 To 4095
        output(outPos) = data(dataLb + cursor + k)
        outPos = outPos + 1
      Next k
    End If

    cursor = cursor + chunkLen
  Loop

  ReDim Preserve output(0 To outPos - 1)
  CompressStream = output
End Function

' Greedy LZ encoder for a single token-compressed chunk. Fills encoded() and
' returns the number of bytes written. See CompressStream for why ties resolve
' to the farthest match.
'
' The 3-gram index is four flat arrays, which VBA indexes far faster than it can
' call into a Scripting.Dictionary: head and tail hold the first and last
' position in each hash bucket, nxt chains them in
' ascending order, and keyOf stores the full 24-bit 3-gram so a hash collision
' cannot fabricate a match. Positions are one-based so that zero means empty and
' the per-chunk reset is a plain ReDim.
Private Function EncodeTokenChunk(ByRef data() As Byte, ByVal startIdx As Long, _
                                  ByVal chunkLen As Long, ByRef encoded() As Byte) As Long
  Dim head()       As Long
  Dim tail()       As Long
  Dim nxt()        As Long
  Dim keyOf()      As Long
  Dim pow2(0 To 16) As Long

  Dim outPos       As Long
  Dim pos          As Long
  Dim indexedTo    As Long
  Dim bit          As Long
  Dim flagBits     As Long
  Dim flagPos      As Long

  Dim bitCount     As Long
  Dim bitThreshold As Long
  Dim lengthMask   As Long
  Dim offsetMask   As Long
  Dim shiftMul     As Long
  Dim maxLength    As Long

  Dim limit        As Long
  Dim bestLen      As Long
  Dim bestOff      As Long
  Dim matchLen     As Long
  Dim targetKey    As Long
  Dim h            As Long
  Dim slot         As Long
  Dim candidate    As Long
  Dim token        As Long
  Dim i            As Long

  If chunkLen <= 0 Then Exit Function

  ' 4096 literals plus one flag byte per eight tokens is the widest this can get.
  ReDim encoded(0 To 4608)
  ReDim head(0 To 65535)
  ReDim tail(0 To 65535)
  ReDim nxt(0 To 4096)
  ReDim keyOf(0 To 4096)

  pow2(0) = 1
  For i = 1 To 16
    pow2(i) = pow2(i - 1) * 2
  Next i

  bitThreshold = -1     ' force the ladder to be evaluated on the first position

  Do While pos < chunkLen
    flagBits = 0
    flagPos = outPos
    outPos = outPos + 1

    For bit = 0 To 7
      If pos >= chunkLen Then Exit For

      ' Same ladder DecompressStream uses, recomputed only when pos crosses one
      ' of its eight boundaries.
      If pos > bitThreshold Then
        Select Case pos
          Case Is <= 16:   bitCount = 4:  bitThreshold = 16
          Case Is <= 32:   bitCount = 5:  bitThreshold = 32
          Case Is <= 64:   bitCount = 6:  bitThreshold = 64
          Case Is <= 128:  bitCount = 7:  bitThreshold = 128
          Case Is <= 256:  bitCount = 8:  bitThreshold = 256
          Case Is <= 512:  bitCount = 9:  bitThreshold = 512
          Case Is <= 1024: bitCount = 10: bitThreshold = 1024
          Case Is <= 2048: bitCount = 11: bitThreshold = 2048
          Case Else:       bitCount = 12: bitThreshold = &H7FFFFFFF
        End Select
        shiftMul = pow2(16 - bitCount)
        lengthMask = shiftMul - 1
        offsetMask = &HFFFF& - lengthMask
        maxLength = lengthMask + 3
      End If

      bestLen = 0
      bestOff = 0

      If pos + 3 <= chunkLen Then
        ' Catch the index up. Positions inside a region already copied over are
        ' still indexed, which is what makes later matches reachable.
        Do While indexedTo < pos
          If indexedTo + 3 <= chunkLen Then
            slot = indexedTo + 1
            keyOf(slot) = CLng(data(startIdx + indexedTo)) _
                       Or (CLng(data(startIdx + indexedTo + 1)) * 256&) _
                       Or (CLng(data(startIdx + indexedTo + 2)) * 65536)
            h = ((CLng(data(startIdx + indexedTo)) * 256& + _
                  CLng(data(startIdx + indexedTo + 1))) Xor _
                 (CLng(data(startIdx + indexedTo + 2)) * 61&)) And &HFFFF&
            nxt(slot) = 0
            If head(h) = 0 Then
              head(h) = slot
            Else
              nxt(tail(h)) = slot
            End If
            tail(h) = slot
          End If
          indexedTo = indexedTo + 1
        Loop

        limit = maxLength
        If chunkLen - pos < limit Then limit = chunkLen - pos

        targetKey = CLng(data(startIdx + pos)) _
                 Or (CLng(data(startIdx + pos + 1)) * 256&) _
                 Or (CLng(data(startIdx + pos + 2)) * 65536)
        h = ((CLng(data(startIdx + pos)) * 256& + _
              CLng(data(startIdx + pos + 1))) Xor _
             (CLng(data(startIdx + pos + 2)) * 61&)) And &HFFFF&

        slot = head(h)
        Do While slot <> 0
          If keyOf(slot) = targetKey Then
            candidate = slot - 1
            matchLen = 3
            Do While matchLen < limit
              If data(startIdx + candidate + matchLen) <> data(startIdx + pos + matchLen) Then Exit Do
              matchLen = matchLen + 1
            Loop
            If matchLen > bestLen Then
              bestLen = matchLen
              bestOff = pos - candidate
              ' Nothing later can strictly beat the cap.
              If bestLen >= limit Then Exit Do
            End If
          End If
          slot = nxt(slot)
        Loop
      End If

      If bestLen >= 3 Then
        flagBits = flagBits Or pow2(bit)
        token = (((bestOff - 1) * shiftMul) And offsetMask) Or ((bestLen - 3) And lengthMask)
        encoded(outPos) = token And &HFF&
        encoded(outPos + 1) = (token \ 256&) And &HFF&
        outPos = outPos + 2
        pos = pos + bestLen
      Else
        encoded(outPos) = data(startIdx + pos)
        outPos = outPos + 1
        pos = pos + 1
      End If
    Next bit

    encoded(flagPos) = flagBits
  Loop

  EncodeTokenChunk = outPos
End Function

Private Function DispCallByVtbl(ByRef pUnk As stdole.IUnknown, ByVal lIndex As Long, ParamArray a() As Variant) As Variant
  Dim lIdx           As Long
  Dim vParam()       As Variant
  Dim vType(0 To 63) As Integer
  Dim vPtr(0 To 63)  As LongPtr
  Dim hResult        As Long

  vParam = a
  For lIdx = 0 To UBound(vParam)
    vType(lIdx) = VarType(vParam(lIdx))
    vPtr(lIdx) = VarPtr(vParam(lIdx))
  Next lIdx

  hResult = DispCallFunc(ObjPtr(pUnk), lIndex * PTR_SIZE, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
  If hResult < 0 Then LogMessage xrayDebug, "[DispCallByVtbl] hResult: &H" & Hex(hResult)
End Function

Private Function PtrToString(ByVal stringPtr As LongPtr) As String
  Dim buf       As String
  Dim charLen   As Long
  charLen = lstrlenW(stringPtr)
  If charLen > 0 Then
    buf = VBA.String$(charLen, vbNullChar)
    CopyMemory ByVal StrPtr(buf), ByVal stringPtr, charLen * 2
  End If
  PtrToString = buf
End Function

Private Function BytesToString(ByRef data() As Byte, ByVal pos As Long, ByVal Size As Long, ByVal CodePage As Long) As String
  If Size <= 0 Then Exit Function
  If ByteCount(data) = 0 Then Exit Function
  If pos < LBound(data) Or pos > UBound(data) Then Exit Function
  If Size > UBound(data) - pos + 1 Then Exit Function
  Dim charCount As Long
  charCount = MultiByteToWideChar(CodePage, 0, VarPtr(data(pos)), Size, 0, 0)
  If charCount = 0 Then Exit Function
  Dim result As String
  result = VBA.String$(charCount, 0)
  MultiByteToWideChar CodePage, 0, VarPtr(data(pos)), Size, StrPtr(result), charCount
  BytesToString = result
End Function

Private Function StringToBytes(ByVal s As String, ByVal CodePage As Long) As Byte()
  Dim result() As Byte
  Dim Size As Long
  Size = WideCharToMultiByte(CodePage, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
  If Size = 0 Then Exit Function
  ReDim result(0 To Size - 1)
  WideCharToMultiByte CodePage, 0, StrPtr(s), Len(s), VarPtr(result(0)), Size, 0, 0
  StringToBytes = result
End Function

Private Function IsCompoundFileBytes(ByRef data() As Byte) As Boolean
  If ByteCount(data) < 8 Then Exit Function
  IsCompoundFileBytes = HasCompoundFileSignatureAt(data, LBound(data))
End Function

Private Function HasCompoundFileSignatureAt(ByRef data() As Byte, ByVal pos As Long) As Boolean
  If ByteCount(data) < 8 Then Exit Function
  If pos < LBound(data) Or pos > UBound(data) - 7 Then Exit Function

  HasCompoundFileSignatureAt = _
      data(pos) = &HD0 And data(pos + 1) = &HCF And _
      data(pos + 2) = &H11 And data(pos + 3) = &HE0 And _
      data(pos + 4) = &HA1 And data(pos + 5) = &HB1 And _
      data(pos + 6) = &H1A And data(pos + 7) = &HE1
End Function

Private Function ByteCount(ByRef data() As Byte) As Long
  On Error Resume Next
  If UBound(data) < LBound(data) Then Exit Function
  ByteCount = UBound(data) - LBound(data) + 1
End Function

Private Function ReadUInt16LE(ByRef data() As Byte, ByVal pos As Long) As Long
  If ByteCount(data) < 2 Then Exit Function
  If pos < LBound(data) Or pos > UBound(data) - 1 Then Exit Function
  ReadUInt16LE = CLng(data(pos)) Or (CLng(data(pos + 1)) * 256&)
End Function

Private Function ReadUInt32LE(ByRef data() As Byte, ByVal pos As Long) As Double
  If ByteCount(data) < 4 Then Exit Function
  If pos < LBound(data) Or pos > UBound(data) - 3 Then Exit Function

  ReadUInt32LE = CDbl(data(pos)) + _
                 CDbl(data(pos + 1)) * 256# + _
                 CDbl(data(pos + 2)) * 65536# + _
                 CDbl(data(pos + 3)) * 16777216#
End Function

Private Function FileExists(ByVal FilePath As String) As Boolean
  Dim attrs As Long
  attrs = apiGetFileAttributesW(StrPtr(FilePath))
  If attrs = INVALID_FILE_ATTRIBUTES Then Exit Function
  FileExists = (attrs And FILE_ATTRIBUTE_DIRECTORY) = 0
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
  Dim attrs As Long
  attrs = apiGetFileAttributesW(StrPtr(folderPath))
  If attrs = INVALID_FILE_ATTRIBUTES Then Exit Function
  FolderExists = (attrs And FILE_ATTRIBUTE_DIRECTORY) <> 0
End Function

Private Function CreateFolderRecursive(ByVal folderPath As String) As Boolean
  
  folderPath = VBA.Trim$(folderPath)
  If Len(folderPath) = 0 Then Exit Function

  Do While Len(folderPath) > 3 And Right$(folderPath, 1) = "\"
    folderPath = Left$(folderPath, Len(folderPath) - 1)
  Loop

  If FolderExists(folderPath) Then CreateFolderRecursive = True: Exit Function

  Dim slashPos As Long
  Dim parent As String
  slashPos = InStrRev(folderPath, "\")
  If slashPos > 0 Then parent = Left$(folderPath, slashPos - 1)
  If Len(parent) = 2 And Right$(parent, 1) = ":" Then parent = parent & "\"

  If Len(parent) > 0 And Not FolderExists(parent) Then
    If Not CreateFolderRecursive(parent) Then Exit Function
  End If
  CreateFolderRecursive = (apiCreateDirectoryW(StrPtr(folderPath & vbNullChar), 0) <> 0)
End Function

Private Function SliceBytes(ByRef data() As Byte, ByVal pos As Long, ByVal Size As Long) As Byte()
  
  Dim result() As Byte
  If Size <= 0 Then Exit Function
  If ByteCount(data) = 0 Then Exit Function
  If pos < LBound(data) Or pos > UBound(data) Then Exit Function
  If Size > UBound(data) - pos + 1 Then Exit Function
  ReDim result(0 To Size - 1)
  CopyMemory VarPtr(result(0)), VarPtr(data(pos)), Size
  SliceBytes = result
End Function

Private Function WriteFile(ByVal FileName As String, ByRef fileData As Variant, Optional ByVal WriteBinary As Boolean = True) As Boolean
  On Error GoTo Failed
  Dim hFile As LongPtr
  Dim bytesWritten As Long
  Dim data() As Byte

  If WriteBinary Then
    data = fileData
  Else
    data = StringToBytes(CStr(fileData), 65001)
  End If

  hFile = apiCreateFile(StrPtr(FileName), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, 0, CREATE_ALWAYS, 0, 0)
  If hFile = INVALID_HANDLE_VALUE Then Exit Function

  Dim dataSize As Long
  dataSize = ByteCount(data)
  If dataSize > 0 Then
    If apiWriteFile(hFile, VarPtr(data(LBound(data))), dataSize, bytesWritten, 0) = 0 Then GoTo Failed
    If bytesWritten <> dataSize Then GoTo Failed
  End If

  apiCloseHandle hFile
  hFile = INVALID_HANDLE_VALUE
  WriteFile = True
  Exit Function

Failed:
  If hFile <> 0 And hFile <> INVALID_HANDLE_VALUE Then apiCloseHandle hFile
End Function

Private Function ReadFile(ByVal FilePath As String, _
                               Optional ByVal MaxBytes As Long = 0) As Byte()
  Dim hFile      As LongPtr
  Dim fileSize   As LongPtr
  Dim readSize   As Long
  Dim bytesRead  As Long
  Dim result()   As Byte

  hFile = apiCreateFile(StrPtr(FilePath), GENERIC_READ, _
                        FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, _
                        0, OPEN_EXISTING, 0, 0)
  If hFile = INVALID_HANDLE_VALUE Then
    SetError "[ReadFile] Could not open: " & FilePath
    Exit Function
  End If

  If apiGetFileSizeEx(hFile, fileSize) = 0 Then
    apiCloseHandle hFile
    Exit Function
  End If

  readSize = CLng(IIf(MaxBytes > 0 And MaxBytes < fileSize, MaxBytes, fileSize))
  If readSize = 0 Then apiCloseHandle hFile: Exit Function

  ReDim result(0 To readSize - 1)
  If apiReadFile(hFile, VarPtr(result(0)), readSize, bytesRead, 0) <> 0 Then
    If bytesRead = readSize Then ReadFile = result
  End If

  apiCloseHandle hFile
End Function
Private Function LoadFromAccessDB(ByVal FilePath As String) As Boolean
  Dim data() As Byte
  data = ReadFile(FilePath)

  If ByteCount(data) < 4096 Then
    SetError "[LoadFromAccessDB] File is too small to be an Access database."
    Exit Function
  End If

  Dim sig As String
  sig = StrConv(SliceBytes(data, 4, 15), vbUnicode)
  If sig <> "Standard ACE DB" And sig <> "Standard Jet DB" Then
    SetError "[LoadFromAccessDB] Invalid database signature."
    Exit Function
  End If

  ParseAccessDB data
  LoadFromAccessDB = IsLoaded
  If Not LoadFromAccessDB Then
    LogMessage xrayWarning, "[LoadFromAccessDB] No validated VBA module source was found. Access support is experimental and currently assumes 4096-byte ACE/Jet4 pages."
  End If
End Function

Private Sub ParseAccessDB(ByRef data() As Byte)
  Reset
  Dim pageCount As Long
  pageCount = (UBound(data) + 1) \ 4096

  Dim p As Long
  For p = 0 To pageCount - 1
    Dim base As Long
    base = p * 4096
    If base + 4095 > UBound(data) Then Exit For

    If data(base) = &H1 And data(base + 4) = &H4C And data(base + 5) = &H56 And data(base + 6) = &H41 And data(base + 7) = &H4C Then
      ParseLVALPage data, p
    End If
  Next p
End Sub

Private Sub ParseLVALPage(ByRef data() As Byte, ByVal p As Long)
  Dim base As Long
  base = p * 4096

  Dim slotCount As Long
  slotCount = CLng(data(base + 12)) Or (CLng(data(base + 13)) * 256&)
  If slotCount > (4096 - 14) \ 2 Then slotCount = (4096 - 14) \ 2

  Dim slot As Long
  For slot = 0 To slotCount - 1
    Dim slotOffset As Long
    slotOffset = base + 14 + (slot * 2)
    Dim raw As Long
    raw = CLng(data(slotOffset)) Or (CLng(data(slotOffset + 1)) * 256&)

    If (raw And &HF000&) <> &HD000& Then
      Dim rowStart As Long
      rowStart = raw And &HFFF&
      Dim rowEnd As Long
      rowEnd = 4096

      Dim i As Long
      For i = 0 To slotCount - 1
        Dim otherRaw As Long
        otherRaw = CLng(data(base + 14 + (i * 2))) Or (CLng(data(base + 14 + (i * 2) + 1)) * 256&)
        If (otherRaw And &HF000&) <> &HD000& Then
          Dim o As Long
          o = otherRaw And &HFFF&
          If o > rowStart And o < rowEnd Then rowEnd = o
        End If
      Next i

      Dim rowLength As Long
      rowLength = rowEnd - rowStart
      If rowLength > 0 Then
        ProcessAccessLVALRow data, p, slot, base + rowStart, rowLength
      End If
    End If
  Next slot
End Sub

Private Sub ProcessAccessLVALRow(ByRef data() As Byte, ByVal p As Long, ByVal slot As Long, ByVal rowStart As Long, ByVal rowLength As Long)
  Dim rowBytes() As Byte
  rowBytes = SliceBytes(data, rowStart, rowLength)

  Dim i As Long, hdr As Long
  For i = 0 To rowLength - 4
    If rowBytes(i) = &H1 Then
      hdr = CLng(rowBytes(i + 1)) Or (CLng(rowBytes(i + 2)) * 256&)
      If ((hdr \ 4096) And 7) = 3 Then
        TryDecompressAccessBlob SliceBytes(rowBytes, i, rowLength - i)
      End If
    End If
  Next i

  If rowLength >= 4 Then
    Dim next_p As Long
    next_p = CLng(rowBytes(1)) Or (CLng(rowBytes(2)) * 256&) Or (CLng(rowBytes(3)) * 65536)
    If next_p > 0 And next_p < ((UBound(data) + 1) \ 4096) Then
      Dim nextBase As Long
      nextBase = next_p * 4096
      If data(nextBase) = &H1 And data(nextBase + 4) = &H4C And data(nextBase + 5) = &H56 And data(nextBase + 6) = &H41 And data(nextBase + 7) = &H4C Then
        Dim chainBlob() As Byte
        If TryWalkAccessLVALChain(data, p, slot, chainBlob) Then
          If ByteCount(chainBlob) >= 4 Then
            For i = LBound(chainBlob) To UBound(chainBlob) - 3
              If chainBlob(i) = &H1 Then
                hdr = CLng(chainBlob(i + 1)) Or (CLng(chainBlob(i + 2)) * 256&)
                If ((hdr \ 4096) And 7) = 3 Then
                  TryDecompressAccessBlob SliceBytes(chainBlob, i, UBound(chainBlob) - i + 1)
                End If
              End If
            Next i
          End If
        End If
      End If
    End If
  End If
End Sub

Private Function TryWalkAccessLVALChain(ByRef data() As Byte, ByVal startPage As Long, ByVal startSlot As Long, ByRef outBlob() As Byte) As Boolean
  Dim cur_p As Long
  Dim cur_s As Long
  cur_p = startPage
  cur_s = startSlot

  Dim tempBlob() As Byte
  Dim totalLen As Long: totalLen = 0

  Dim maxChunks As Long: maxChunks = 4096
  Dim chunkCount As Long: chunkCount = 0

  Do While chunkCount < maxChunks
    chunkCount = chunkCount + 1

    Dim base As Long: base = cur_p * 4096
    If base > UBound(data) Then Exit Function

    Dim slotCount As Long
    slotCount = CLng(data(base + 12)) Or (CLng(data(base + 13)) * 256&)
    If slotCount > (4096 - 14) \ 2 Then slotCount = (4096 - 14) \ 2
    If cur_s >= slotCount Then Exit Function

    Dim slotOffset As Long: slotOffset = base + 14 + (cur_s * 2)
    Dim raw As Long: raw = CLng(data(slotOffset)) Or (CLng(data(slotOffset + 1)) * 256&)
    If (raw And &HF000&) = &HD000& Then Exit Function

    Dim rowStart As Long: rowStart = raw And &HFFF&
    Dim rowEnd As Long: rowEnd = 4096
    Dim i As Long
    For i = 0 To slotCount - 1
      Dim otherRaw As Long: otherRaw = CLng(data(base + 14 + (i * 2))) Or (CLng(data(base + 14 + (i * 2) + 1)) * 256&)
      If (otherRaw And &HF000&) <> &HD000& Then
        Dim o As Long
        o = otherRaw And &HFFF&
        If o > rowStart And o < rowEnd Then rowEnd = o
      End If
    Next i

    Dim rowLength As Long: rowLength = rowEnd - rowStart
    If rowLength < 4 Then Exit Function

    Dim next_s As Long: next_s = CLng(data(base + rowStart))
    Dim next_p As Long: next_p = CLng(data(base + rowStart + 1)) Or (CLng(data(base + rowStart + 2)) * 256&) Or (CLng(data(base + rowStart + 3)) * 65536)

    If rowLength > 4 Then
      If totalLen = 0 Then
        ReDim tempBlob(0 To rowLength - 5)
      Else
        ReDim Preserve tempBlob(0 To totalLen + rowLength - 5)
      End If
      CopyMemory VarPtr(tempBlob(totalLen)), VarPtr(data(base + rowStart + 4)), rowLength - 4
      totalLen = totalLen + rowLength - 4
    End If

    If next_p = 0 And next_s = 0 Then
      If totalLen = 0 Then Exit Function
      outBlob = tempBlob
      TryWalkAccessLVALChain = True
      Exit Function
    End If

    If next_p >= ((UBound(data) + 1) \ 4096) Then Exit Function
    cur_p = next_p
    cur_s = next_s
  Loop
End Function

Private Sub TryDecompressAccessBlob(ByRef compData() As Byte)
  
  On Error GoTo IgnoreDecompFail
  Dim decomp() As Byte
  decomp = DecompressStream(compData, True)

  If ByteCount(decomp) < 20 Then Exit Sub

  If decomp(0) = &H41 And decomp(1) = &H74 And decomp(2) = &H74 And decomp(3) = &H72 And decomp(4) = &H69 And decomp(5) = &H62 And decomp(6) = &H75 And decomp(7) = &H74 And decomp(8) = &H65 Then
    If decomp(9) = &H20 And decomp(10) = &H56 And decomp(11) = &H42 And decomp(12) = &H5F And decomp(13) = &H4E And decomp(14) = &H61 And (decomp(15) = &H6D Or decomp(15) = &H65) And decomp(16) = &H65 And decomp(17) = &H20 And decomp(18) = &H3D And decomp(19) = &H20 Then
      ExtractAccessModule decomp
    End If
  End If
IgnoreDecompFail:
End Sub

Private Sub ExtractAccessModule(ByRef decomp() As Byte)
  
  If ByteCount(decomp) = 0 Then Exit Sub
  Dim text As String
  text = BytesToString(decomp, LBound(decomp), ByteCount(decomp), m_Project.CodePage)

  Dim lines() As String
  lines = Split(text, vbCrLf)

  Dim modName As String
  Dim bodyStart As Long: bodyStart = 0
  Dim i As Long
  For i = LBound(lines) To UBound(lines)
    If Left$(lines(i), 10) = "Attribute " Then
      If Left$(lines(i), 21) = "Attribute VB_Name = """ Then
        Dim parts() As String: parts = Split(lines(i), """")
        If UBound(parts) >= 1 Then modName = parts(1)
      End If
    Else
      bodyStart = i
      Exit For
    End If
  Next i

  If Len(modName) = 0 Then Exit Sub

  If m_Project.ModuleCount > 0 Then
    For i = 1 To m_Project.ModuleCount
      If StrComp(m_Project.Modules(i).Name, modName, vbTextCompare) = 0 Then Exit Sub
    Next i
  End If

  Dim newCount As Long
  newCount = m_Project.ModuleCount + 1
  m_Project.ModuleCount = newCount

  Dim ub As Long
  ub = 0
  On Error Resume Next
  ub = UBound(m_Project.Modules)
  On Error GoTo 0

  If newCount > ub Then
    If ub = 0 Then
      ReDim m_Project.Modules(1 To 10)
    Else
      ReDim Preserve m_Project.Modules(1 To newCount + 10)
    End If
  End If

  m_Project.Modules(newCount).Name = modName
  m_Project.Modules(newCount).ModType = modStandard

  Dim isClass As Boolean: isClass = False
  For i = LBound(lines) To bodyStart - 1
    If InStr(lines(i), "VB_Exposed") > 0 Or InStr(lines(i), "VB_Creatable") > 0 Then isClass = True
  Next i
  If isClass Then m_Project.Modules(newCount).ModType = modClass

  Dim bodyLines() As String
  If bodyStart <= UBound(lines) Then
    Dim j As Long: j = 0
    ReDim bodyLines(0 To UBound(lines) - bodyStart)
    For i = bodyStart To UBound(lines)
      bodyLines(j) = lines(i)
      j = j + 1
    Next i
    m_Project.Modules(newCount).SourceCode = Join(bodyLines, vbCrLf)
  End If

End Sub

Private Sub LogMessage(ByVal level As XrayLogLevel, ByVal Msg As String)
  If level < m_LogLevel Then Exit Sub

  Dim prefix As String
  Select Case level
    Case xrayDebug:   prefix = "[DBG] "
    Case xrayInfo:    prefix = "[INF] "
    Case xrayWarning: prefix = "[WRN] "
    Case xrayError:   prefix = "[ERR] "
    Case Else:        prefix = "[   ] "
  End Select

  Debug.Print prefix & Msg
End Sub

Private Sub SetError(ByVal Msg As String)
  m_LastError = Msg
  Debug.Print "[ERR] " & Msg
End Sub

Private Sub ClearError()
  m_LastError = vbNullString
End Sub

Private Sub Class_Initialize()
  m_LogLevel = xrayInfo
  Reset
End Sub

Private Function StorageOpenZip(ByVal OfficePath As String) As stdole.IUnknown
  Dim clsid  As UUID
  Dim iid    As UUID
  Dim shExt  As stdole.IUnknown
  Dim pidl   As LongPtr
  Dim hr     As Long

  CLSIDFromString StrPtr(CLSID_ZipFolder), clsid
  CLSIDFromString StrPtr(IID_IPersistFolder), iid

  hr = CoCreateInstance(clsid, 0, CLSCTX_INPROC_SERVER, iid, shExt)
  If hr <> 0 Or shExt Is Nothing Then
    SetError "[StorageOpenZip] CoCreateInstance failed: &H" & Hex$(hr)
    Exit Function
  End If

  pidl = ILCreateFromPathW(StrPtr(OfficePath))
  If pidl = 0 Then
    SetError "[StorageOpenZip] ILCreateFromPathW failed for: " & OfficePath
    Exit Function
  End If
  hr = DispCallByVtbl(shExt, 4, pidl)
  ILFree pidl
  If hr <> 0 Then
    SetError "[StorageOpenZip] Initialize failed: &H" & Hex$(hr)
    Exit Function
  End If

  Set StorageOpenZip = COMQueryInterface(shExt, IID_IStorage)
  If StorageOpenZip Is Nothing Then
    SetError "[StorageOpenZip] QI for IStorage failed"
  End If
End Function

Private Function COMQueryInterface(ByRef pUnk As stdole.IUnknown, ByVal iid As String) As stdole.IUnknown
  If pUnk Is Nothing Then Exit Function
  Dim riid   As UUID
  Dim result As stdole.IUnknown
  CLSIDFromString StrPtr(iid), riid
  DispCallByVtbl pUnk, 0, VarPtr(riid), VarPtr(result)
  Set COMQueryInterface = result
End Function

#If False Then
  Dim xrayDebug, xrayInfo, xrayWarning, xrayError
  Dim modOther, modStandard, modClass, modDesigner, modDocument
#End If

