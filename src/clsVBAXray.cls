VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVBAXray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A tool to extract VBA code from Office files"

'  |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'  |||||||||||||||||||||||||||    vbaXray (v2.2)    ||||||||||||||||||||||||||||||
'  |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'
'  AUTHOR:   Kallun Willock
'
'  URL:      https://github.com/KallunWillock/vbaXray
'
'  LICENSE:  MIT
'
'  VERSION:  1.0    23/03/2026    Initial release. OOXML only, via Shell.Application.
'            2.0    15/08/2026    OOXML via ZipFldr IStorage, legacy compound
'                                 files, legacy PowerPoint, Access, diagnostics.
'            2.1    30/08/2026    - Added support to export valid (importable) FRX files;
'                                 - Added Project Reference extraction;
'                                 - Various bug fixes and optimisations.
'            2.2    01/09/2026    Bug fixes to improve exporting modules/classes from
'                                 Access databases.

Option Explicit

Private Const VBAXRAY_VERSION           As Single = 2.2

#If Win64 Then
  Private Const PTR_SIZE As Long = 8
#Else
  Private Const PTR_SIZE As Long = 4
#End If

Private Type LARGE_INTEGER
  LowPart  As Long
  HighPart As Long
End Type

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

  Private Declare PtrSafe Function apiCreateFile Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As LongPtr, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr
  Private Declare PtrSafe Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
  Private Declare PtrSafe Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As LongPtr, ByVal lpBuffer As LongPtr, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
  Private Declare PtrSafe Function apiGetFileSizeEx Lib "kernel32" Alias "GetFileSizeEx" (ByVal hFile As LongPtr, ByRef lpFileSize As LARGE_INTEGER) As Long
  Private Declare PtrSafe Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As LongPtr) As Long
  Private Declare PtrSafe Function apiCreateDirectoryW Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As LongPtr, ByVal lpSecurityAttributes As LongPtr) As Long
  Private Declare PtrSafe Function apiGetFileAttributesW Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As LongPtr) As Long

  Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
  Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
  Private Declare PtrSafe Function StgOpenStorage Lib "ole32" (ByVal pwcsName As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ByRef ppstgOpen As stdole.IUnknown) As Long
  Private Declare PtrSafe Function CreateILockBytesOnHGlobal Lib "ole32" (ByVal hGlobal As LongPtr, ByVal fDeleteOnRelease As Long, ByRef ppLkbyt As stdole.IUnknown) As Long
  Private Declare PtrSafe Function StgOpenStorageOnILockBytes Lib "ole32" (ByVal pLkbyt As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ByRef ppstgOpen As stdole.IUnknown) As Long
  Private Declare PtrSafe Function ILCreateFromPathW Lib "shell32" (ByVal pszPath As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

  Private Declare PtrSafe Function StgCreateDocfileOnILockBytes Lib "ole32" (ByVal pLkbyt As LongPtr, ByVal grfMode As Long, ByVal reserved As Long, ByRef ppstgOpen As stdole.IUnknown) As Long
  Private Declare PtrSafe Function GetHGlobalFromILockBytes Lib "ole32" (ByVal pLkbyt As LongPtr, ByRef phglobal As LongPtr) As Long

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
  Private Declare Function apiGetFileSizeEx Lib "kernel32" Alias "GetFileSizeEx" (ByVal hFile As LongPtr, ByRef lpFileSize As LARGE_INTEGER) As Long
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
  Private Declare Function StgOpenStorageOnILockBytes Lib "ole32" (ByVal pLkbyt As LongPtr, ByVal pstgPriority As LongPtr, ByVal grfMode As Long, ByVal snbExclude As LongPtr, ByVal reserved As Long, ByRef ppstgOpen As stdole.IUnknown) As Long

  Private Declare Function StgCreateDocfileOnILockBytes Lib "ole32" (ByVal plkbyt As Long, ByVal grfMode As Long, ByVal reserved As Long, ByRef ppstgOpen As stdole.IUnknown) As Long
  Private Declare Function GetHGlobalFromILockBytes Lib "ole32" (ByVal plkbyt As Long, ByRef phglobal As Long) As Long

  Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
  Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
  Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
  Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr

  Private Declare Function DispCallFunc Lib "oleaut32" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal lCc As Long, ByVal vtReturn As VbVarType, ByVal cActuals As Long, prgVt As Any, prgpVarg As Any, pvargResult As Variant) As Long
  Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

  Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal dest As LongPtr, ByVal src As LongPtr, ByVal cb As LongPtr)
  Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
#End If

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

Private Enum STGTY
  STGTY_STORAGE = 1
  STGTY_STREAM = 2
  STGTY_LOCKBYTES = 3
  STGTY_PROPERTY = 4
End Enum

Private Const STGM_READ                 As Long = &H0
Private Const STGM_READWRITE            As Long = &H2
Private Const STGM_SHARE_DENY_WRITE     As Long = &H20
Private Const STGM_SHARE_EXCLUSIVE      As Long = &H10
Private Const STGM_CREATE               As Long = &H1000
Private Const STATFLAG_NONAME           As Long = 1
Private Const ILOCKBYTES_STAT           As Long = 9

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
Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const CLSCTX_INPROC_SERVER      As Long = 1
Private Const ARCHIVE_OK                As Long = 0
Private Const ARCHIVE_EOF               As Long = 1
Private Const PPT_VBA_PROJECT_STG       As Long = &H1011&
Private Const MAX_EMBEDDED_PROJECT_SIZE As Long = 268435456  ' 256 Mb ceiling
Private Const GZIP_HEADER_SIZE          As Long = 10
Private Const GZIP_TRAILER_SIZE         As Long = 8

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

Private Type ModuleRecord
  Name        As String
  streamName  As String
  offset      As Long
  ModType     As ModuleType
  IsForm      As Boolean
  SourceCode  As String
End Type

Private Type ReferenceRecord
  Name          As String
  GUID          As String
  MajorVersion  As Long
  MinorVersion  As Long
  Description   As String
  FilePath      As String
  RefType       As String
End Type

Private Type ProjectRecord
  Name                As String
  Description         As String
  FilePath            As String
  CodePage            As Long
  IsPasswordProtected As Boolean
  ModuleCount         As Long
  Modules()           As ModuleRecord
  ReferenceCount      As Long
  References()        As ReferenceRecord
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
Private m_AccessPageSize    As Long
Private m_AccessRowCntOff   As Long
Private m_AccessRowArrOff   As Long
Private m_AccessRecoverCnt  As Long
Private m_AccessSeenBlobs   As Collection
Private m_vbaProjectBin()   As Byte

Private m_Pow2(0 To 16)     As Long
Private m_BitMask(0 To 7)   As Long

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

Public Property Get ReferenceCount() As Long
  ReferenceCount = m_Project.ReferenceCount
End Property

Public Property Get ReferenceName(ByVal index As Long) As String
  If index < 1 Or index > m_Project.ReferenceCount Then
    SetError "[ReferenceName] Index out of range": Exit Property
  End If
  ReferenceName = m_Project.References(index).Name
End Property

Public Property Get ReferenceGUID(ByVal index As Long) As String
  If index < 1 Or index > m_Project.ReferenceCount Then
    SetError "[ReferenceGUID] Index out of range": Exit Property
  End If
  ReferenceGUID = m_Project.References(index).GUID
End Property

Public Property Get ReferenceMajorVersion(ByVal index As Long) As Long
  If index < 1 Or index > m_Project.ReferenceCount Then
    SetError "[ReferenceMajorVersion] Index out of range": Exit Property
  End If
  ReferenceMajorVersion = m_Project.References(index).MajorVersion
End Property

Public Property Get ReferenceMinorVersion(ByVal index As Long) As Long
  If index < 1 Or index > m_Project.ReferenceCount Then
    SetError "[ReferenceMinorVersion] Index out of range": Exit Property
  End If
  ReferenceMinorVersion = m_Project.References(index).MinorVersion
End Property

Public Property Get ReferenceDescription(ByVal index As Long) As String
  If index < 1 Or index > m_Project.ReferenceCount Then
    SetError "[ReferenceDescription] Index out of range": Exit Property
  End If
  ReferenceDescription = m_Project.References(index).Description
End Property

Public Property Get ReferenceFilePath(ByVal index As Long) As String
  If index < 1 Or index > m_Project.ReferenceCount Then
    SetError "[ReferenceFilePath] Index out of range": Exit Property
  End If
  ReferenceFilePath = m_Project.References(index).FilePath
End Property

Public Property Get ReferenceType(ByVal index As Long) As String
  If index < 1 Or index > m_Project.ReferenceCount Then
    SetError "[ReferenceType] Index out of range": Exit Property
  End If
  ReferenceType = m_Project.References(index).RefType
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
    ElseIf ext = "accdb" Or ext = "accdt" Or ext = "accdr" Or ext = "mdb" Or ext = "accda" Then
      LoadFromFile = LoadFromAccessDB(FilePath)
      If LoadFromFile Then m_Project.FilePath = FilePath
    ElseIf IsCompoundFile(FilePath) Then
      Select Case ext
        Case "ppt", "pps", "pot", "ppa"
          LoadFromFile = LoadFromLegacyPowerPoint(FilePath)
          If LoadFromFile Then m_Project.FilePath = FilePath
        Case Else
          m_vbaProjectBin = ReadFile(FilePath)
          Dim fileStorage As stdole.IUnknown
          Set fileStorage = StorageOpenFile(FilePath)
          LoadFromStorage fileStorage, ext
          LoadFromFile = IsLoaded
          If LoadFromFile Then m_Project.FilePath = FilePath
      End Select
    Else
      Select Case ext
        Case "xlsm", "xlam", "xlsb", "xltm", "docm", "dotm", "pptm", "potm", "ppsm", "ppam", "sldm"
          Dim vbaBytes() As Byte
          vbaBytes = GetVBAProjectBinBytes(FilePath)
          If ByteCount(vbaBytes) > 0 Then
            LoadFromFile = LoadFromByteArray(vbaBytes)
            If LoadFromFile Then m_Project.FilePath = FilePath
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
      m_vbaProjectBin = VBAProjectBinData

      Dim byteStorage As stdole.IUnknown
      Set byteStorage = StorageOpenBytes(VBAProjectBinData)
      If Not byteStorage Is Nothing Then
        LoadFromStorage byteStorage, "bin"
        LoadFromByteArray = IsLoaded
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
  Erase m_vbaProjectBin
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
  Next
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
  Next
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
  Next
  FullSourceCode = Join(parts, vbNullString)
End Function

Public Function IsCompoundFile(ByVal FilePath As String) As Boolean
  Dim header() As Byte
  header = ReadFile(FilePath, 8)
  If ByteCount(header) < 8 Then Exit Function
  IsCompoundFile = IsCompoundFileBytes(header)
End Function

Public Sub DebugDumpStorageTree(Optional ByVal OutputFilePath As String = vbNullString, _
                                Optional ByVal ExcludeSRP As Boolean = False)
  If m_StorageTree Is Nothing Then
    SetError "[DebugDumpStorageTree] No storage loaded."
    Exit Sub
  End If
  Dim output As String
  output = "." & vbCrLf
  TreeNodeRender m_StorageTree, "", output, ExcludeSRP
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

Private Sub TreeNodeRender(ByRef nodes As Collection, ByVal prefix As String, ByRef output As String, Optional ByVal ExcludeSRP As Boolean = False)
  Dim i           As Long
  Dim Node        As Object
  Dim isLast      As Boolean
  Dim branch      As String
  Dim childPfx    As String
  Dim activeNodes As New Collection

  Dim entry As Variant
  For Each entry In nodes
    If Not (ExcludeSRP And Left$(entry("name"), 6) = "__SRP_") Then
      activeNodes.Add entry
    End If
  Next

  i = 0
  For Each Node In activeNodes
    i = i + 1
    isLast = (i = activeNodes.count)
    branch = IIf(isLast, "\-- ", "+-- ")
    childPfx = prefix & IIf(isLast, "    ", "|   ")

    Select Case Node("type")
      Case STGTY_STREAM
        Dim Col As String * 50
        Col = prefix & branch & EscapeName(Node("name"))
        output = output & Col & CLng(Node("size")) & " bytes" & vbCrLf
      Case STGTY_STORAGE
        output = output & prefix & branch & "[" & EscapeName(Node("name")) & "]" & vbCrLf
        If Not IsEmpty(Node("children")) Then
          If Not Node("children") Is Nothing Then
            TreeNodeRender Node("children"), childPfx, output, ExcludeSRP
          End If
        End If
      Case Else
        output = output & prefix & branch & EscapeName(Node("name")) & "  (type=" & Node("type") & ")" & vbCrLf
    End Select
  Next
End Sub

Private Function EscapeName(ByVal s As String) As String
  Dim i   As Long
  Dim ch  As Long
  Dim out As String
  For i = 1 To Len(s)
    ch = AscW(VBA.Mid$(s, i, 1)) And &HFFFF&
    If ch < 32 Or ch = 127 Then
      out = out & "\x" & Right$("0" & Hex$(ch), 2)
    Else
      out = out & VBA.Mid$(s, i, 1)
    End If
  Next
  EscapeName = out
End Function

Private Function GetVBAProjectBinBytes(ByVal OfficePath As String) As Byte()
  Dim target     As String
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

  Dim zipStg As stdole.IUnknown, folderStg As stdole.IUnknown, folderRaw As stdole.IUnknown, pStm As stdole.IUnknown
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

Private Function LoadFromLegacyPowerPoint(ByVal FilePath As String) As Boolean
  On Error GoTo ErrHandler
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

  Dim lb As Long, ub As Long
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
  Next

  Reset
  SetError "[LoadFromLegacyPowerPoint] No valid VBA project found"
  Exit Function

ErrHandler:
  Reset
  SetError "[LoadFromLegacyPowerPoint] Error " & Err.Number & " - " & Err.Description
End Function

Private Function TryZlibInflate(ByRef zlibData() As Byte, ByVal ExpectedSize As Long, ByRef result() As Byte) As Boolean
  ' Inflates the zlib-compatible payload used by compressed VbaProjectStg
  ' records in legacy PowerPoint files. Uses archiveint.dll under the hood.
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

  On Error GoTo ErrHandler
  Erase result

  Dim zlibSize As Long, lb As Long
  zlibSize = ByteCount(zlibData)
  If zlibSize < 8 Then Exit Function
  If ExpectedSize < 8 Or ExpectedSize > MAX_EMBEDDED_PROJECT_SIZE Then Exit Function
  lb = LBound(zlibData)

  Dim cmf As Long, flg As Long
  cmf = zlibData(lb)
  flg = zlibData(lb + 1)

  If (cmf And &HF&) <> 8 Then Exit Function
  If (cmf \ &H10&) > 7 Then Exit Function
  If ((cmf * 256&) + flg) Mod 31 <> 0 Then Exit Function
  If (flg And &H20&) <> 0 Then Exit Function
  ' Preserve every byte after the zlib header. This is very much intentional,
  ' from what I have seen, the working PPT variant does not expose a dependable
  ' four-byte Adler trailer.
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
      If blockOffset < 0@ Or (blockOffset * 10000@) > CDbl(ExpectedSize) Then GoTo CleanUp
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

ErrHandler:
  SetError "[TryZlibInflate] Error " & Err.Number & " - " & Err.Description
  Resume CleanUp
End Function

Private Sub LoadFromStorage(ByVal topStg As stdole.IUnknown, ByVal ext As String)
  On Error GoTo ErrHandler
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
  Dim vbaStg    As stdole.IUnknown
  Set vbaStg = StorageNavigate(topStg, layout.VBAPath, keepAlive)

  If vbaStg Is Nothing Then
    Dim fallbackPaths As Variant, fPath As Variant
    fallbackPaths = Array("VBA", "_VBA_PROJECT_CUR\VBA", "Macros\VBA", "VBA\VBA", "OutlookVbaData\VBA")
    For Each fPath In fallbackPaths
      Set vbaStg = StorageNavigate(topStg, CStr(fPath), keepAlive)
      If Not vbaStg Is Nothing Then
        layout.VBAPath = CStr(fPath)
        Dim lastS As Long
        lastS = InStrRev(layout.VBAPath, "\")
        If lastS > 0 Then
          layout.projectPath = Left$(layout.VBAPath, lastS - 1) & "\PROJECT"
        Else
          layout.projectPath = "PROJECT"
        End If
        Exit For
      End If
    Next
  End If

  If vbaStg Is Nothing Then
    SetError "[LoadFromStorage] VBA storage not found"
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
  Next

  LogMessage xrayInfo, "Project loaded: " & m_Project.Name & " (" & m_Project.ModuleCount & " modules)"

CleanUp:
  On Error Resume Next
  Set keepAlive = Nothing
  Set vbaStg = Nothing
  Exit Sub

ErrHandler:
  Reset
  SetError "[LoadFromStorage] Error " & Err.Number & " - " & Err.Description
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

Private Function StorageNavigate(ByRef rootStg As stdole.IUnknown, ByVal path As String, ByRef keepAlive As Collection) As stdole.IUnknown
  Dim parts()     As String
  Dim current     As stdole.IUnknown
  Dim next_stg    As stdole.IUnknown
  Dim walkedPath  As String
  Dim pathKey     As String
  Dim i           As Long

  parts = Split(path, "\")
  Set current = rootStg

  For i = 0 To UBound(parts)
    If Len(parts(i)) = 0 Then GoTo ErrHandler

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
      Set next_stg = StorageOpenStorage(current, parts(i))
      If next_stg Is Nothing Then
        LogMessage xrayDebug, "[StorageNavigate] Could not open '" & parts(i) & "' in path '" & path & "'"
        GoTo ErrHandler
      End If
      keepAlive.Add next_stg, pathKey
    End If

    Set current = next_stg
  Next

  Set StorageNavigate = current
  Exit Function

ErrHandler:
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

    Dim pStm As stdole.IUnknown
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
    If VBA.Right$(ExportPath, 1) <> "\" Then ExportPath = ExportPath & "\"
    TargetPath = ExportPath & SubFolder & "\"

    If Not FolderExists(TargetPath) Then
      If Not CreateFolderRecursive(TargetPath) Then
        SetError "[WriteModule] Could not create folder: " & TargetPath
        Exit Function
      End If
    End If

    Dim fileData() As Byte

    If .ModType = modDesigner Then
      Dim frmContent As String
      frmContent = SynthesizeFRMContent(.Name, .SourceCode, m_vbaProjectBin)
      fileData = StringToBytes(frmContent, m_Project.CodePage)
    Else
      fileData = StringToBytes(header & .SourceCode, m_Project.CodePage)
    End If

    If ByteCount(fileData) = 0 Then
      SetError "[WriteModule] No output data produced for " & .Name
      Exit Function
    End If

    If Not WriteFile(TargetPath & .Name & Extension, fileData) Then
      SetError "[WriteModule] Failed to write " & TargetPath & .Name & Extension
      Exit Function
    End If

    ' Handle FRX Auto-Extraction on Form Modules [1.4]
    If .ModType = modDesigner And ByteCount(m_vbaProjectBin) > 0 Then
      Dim layout As CFBLayout
      Dim fileExt As String
      fileExt = VBA.LCase$(VBA.Mid$(m_Project.FilePath, InStrRev(m_Project.FilePath, ".") + 1))
      layout = GetCFBLayout(fileExt)

      Dim parentPath As String
      Dim lastSlash  As Long
      lastSlash = InStrRev(layout.VBAPath, "\")
      If lastSlash > 0 Then parentPath = Left$(layout.VBAPath, lastSlash - 1)

      ExtractFormFRX m_vbaProjectBin, parentPath, .Name, TargetPath & .Name & ".frx"
    End If

    LogMessage xrayInfo, "[Exported] " & SubFolder & "\" & .Name & Extension
    WriteModule = True
  End With
End Function

Private Sub ParseDirStream(ByRef data() As Byte)
  Dim tempList() As ModuleRecord
  ReDim tempList(0 To 63)
  Dim count As Long
  count = 0

  Dim refList() As ReferenceRecord
  ReDim refList(0 To 31)
  Dim refCount As Long
  refCount = 0

  m_Project.CodePage = 1252

  Dim pos        As Long
  Dim RecordId   As Long
  Dim RecordSize As Long
  pos = 0

  Do While pos < UBound(data) - 5
    RecordId = CLng(data(pos)) Or (CLng(data(pos + 1)) * 256&)

    ' MS-OVBA 2.3.4.2.1.5 PROJECTCODEPAGE (0x0003)
    If RecordId = &H3 Then
      RecordSize = CLng(data(pos + 2)) Or (CLng(data(pos + 3)) * 256&) Or (CLng(data(pos + 4)) * 65536) Or (CLng(data(pos + 5)) * 16777216)
      If RecordSize = 2 Then
        m_Project.CodePage = CLng(data(pos + 6)) Or (CLng(data(pos + 7)) * 256&)
      End If
    End If

    ' MS-OVBA 2.3.4.2.2.2 REFERENCENAME (0x0016)
    If RecordId = &H16 Then
      RecordSize = CLng(data(pos + 2)) Or (CLng(data(pos + 3)) * 256&) Or (CLng(data(pos + 4)) * 65536) Or (CLng(data(pos + 5)) * 16777216)
      ' Validate: the two bytes after the ANSI name must be Reserved = 0x003E
      If RecordSize > 0 And RecordSize < 1024 And pos + 6 + RecordSize + 6 <= UBound(data) Then
        Dim nameEnd As Long
        nameEnd = pos + 6 + RecordSize
        Dim reserved As Long
        reserved = CLng(data(nameEnd)) Or (CLng(data(nameEnd + 1)) * 256&)

        If reserved = &H3E Then
          Dim refName As String
          refName = BytesToString(data, pos + 6, RecordSize, m_Project.CodePage)

          Dim uniSize As Long
          uniSize = CLng(data(nameEnd + 2)) Or (CLng(data(nameEnd + 3)) * 256&) Or (CLng(data(nameEnd + 4)) * 65536) Or (CLng(data(nameEnd + 5)) * 16777216)

          Dim nextRecPos As Long
          nextRecPos = nameEnd + 6 + uniSize

          If nextRecPos + 10 <= UBound(data) Then
            Dim nextId As Long
            nextId = CLng(data(nextRecPos)) Or (CLng(data(nextRecPos + 1)) * 256&)

            If nextId = &H33 Then ' REFERENCEORIGINAL
              Dim origSize As Long
              origSize = CLng(data(nextRecPos + 2)) Or (CLng(data(nextRecPos + 3)) * 256&) Or (CLng(data(nextRecPos + 4)) * 65536) Or (CLng(data(nextRecPos + 5)) * 16777216)
              nextRecPos = nextRecPos + 6 + origSize
              If nextRecPos + 10 > UBound(data) Then GoTo SkipReferenceName
              nextId = CLng(data(nextRecPos)) Or (CLng(data(nextRecPos + 1)) * 256&)
            End If

            Select Case nextId
              Case &HD  ' REFERENCEREGISTERED
                Dim regSize As Long
                regSize = CLng(data(nextRecPos + 2)) Or (CLng(data(nextRecPos + 3)) * 256&) Or (CLng(data(nextRecPos + 4)) * 65536) Or (CLng(data(nextRecPos + 5)) * 16777216)

                If regSize >= 8 And nextRecPos + 10 + regSize <= UBound(data) + 1 Then
                  Dim libidSize As Long
                  libidSize = CLng(data(nextRecPos + 6)) Or (CLng(data(nextRecPos + 7)) * 256&) Or (CLng(data(nextRecPos + 8)) * 65536) Or (CLng(data(nextRecPos + 9)) * 16777216)

                  If libidSize > 0 And libidSize <= regSize - 8 Then
                    Dim libid As String
                    libid = BytesToString(data, nextRecPos + 10, libidSize, m_Project.CodePage)

                    If refCount = UBound(refList) Then ReDim Preserve refList(0 To refCount + 31)
                    refCount = refCount + 1

                    Dim gStart As Long, gEnd As Long
                    gStart = InStr(libid, "{")
                    gEnd = InStr(gStart, libid, "}")

                    If gStart > 0 And gEnd > gStart Then
                      refList(refCount).GUID = Mid$(libid, gStart, gEnd - gStart + 1)
                      Dim afterGuid As String, hashPos As Long, verStr As String
                      afterGuid = Mid$(libid, gEnd + 2)
                      hashPos = InStr(afterGuid, "#")
                      If hashPos > 0 Then
                        verStr = Left$(afterGuid, hashPos - 1)
                        Dim dotPos As Long
                        dotPos = InStr(verStr, ".")
                        If dotPos > 0 Then
                          refList(refCount).MajorVersion = val(Left$(verStr, dotPos - 1))
                          refList(refCount).MinorVersion = val(Mid$(verStr, dotPos + 1))
                        End If
                      End If

                      Dim parts() As String: parts = Split(libid, "#")
                      If UBound(parts) >= 4 Then
                        refList(refCount).FilePath = parts(3)
                        refList(refCount).Description = parts(4)
                      End If
                    End If

                    refList(refCount).Name = refName
                    refList(refCount).RefType = "Registered"
                  End If
                End If

                pos = nextRecPos + 6 + regSize
                GoTo ContinueScan

              Case &HE  ' REFERENCEPROJECT
                Dim projSize As Long
                projSize = CLng(data(nextRecPos + 2)) Or (CLng(data(nextRecPos + 3)) * 256&) Or (CLng(data(nextRecPos + 4)) * 65536) Or (CLng(data(nextRecPos + 5)) * 16777216)

                If projSize >= 16 And nextRecPos + 6 + projSize <= UBound(data) + 1 Then
                  Dim absLibidSize As Long
                  absLibidSize = CLng(data(nextRecPos + 6)) Or (CLng(data(nextRecPos + 7)) * 256&) Or (CLng(data(nextRecPos + 8)) * 65536) Or (CLng(data(nextRecPos + 9)) * 16777216)

                  Dim absLibidPos As Long: absLibidPos = nextRecPos + 10

                  If absLibidSize > 0 And absLibidPos + absLibidSize <= UBound(data) + 1 Then
                    Dim absLibid As String
                    absLibid = BytesToString(data, absLibidPos, absLibidSize, m_Project.CodePage)

                    If refCount = UBound(refList) Then ReDim Preserve refList(0 To refCount + 31)
                    refCount = refCount + 1

                    Dim projHashPos As Long
                    projHashPos = InStr(absLibid, "#")
                    Dim projPath    As String
                    If projHashPos > 0 Then
                      projPath = Left$(absLibid, projHashPos - 1)
                      refList(refCount).Description = Mid$(absLibid, projHashPos + 1)
                    Else
                      projPath = absLibid
                    End If

                    If Left$(projPath, 2) = "*\" Then projPath = Mid$(projPath, 3)
                    refList(refCount).FilePath = projPath
                    refList(refCount).GUID = "[Project Reference]"
                    refList(refCount).Name = refName
                    refList(refCount).RefType = "Project"
                  End If
                End If
                pos = nextRecPos + 6 + projSize
                GoTo ContinueScan

              Case &H2F  ' MS-OVBA 2.3.4.2.2.3 REFERENCECONTROL
                Dim ctrlTwiddledSize As Long, ctrlWalk As Long
                ctrlTwiddledSize = CLng(data(nextRecPos + 2)) Or (CLng(data(nextRecPos + 3)) * 256&) Or (CLng(data(nextRecPos + 4)) * 65536) Or (CLng(data(nextRecPos + 5)) * 16777216)
                ctrlWalk = nextRecPos + 6 + ctrlTwiddledSize

                If ctrlWalk + 6 <= UBound(data) Then
                  Dim ctrlEmbedId As Long
                  ctrlEmbedId = CLng(data(ctrlWalk)) Or (CLng(data(ctrlWalk + 1)) * 256&)
                  If ctrlEmbedId = &H16 Then
                    Dim ctrlNameSize As Long
                    ctrlNameSize = CLng(data(ctrlWalk + 2)) Or (CLng(data(ctrlWalk + 3)) * 256&) Or (CLng(data(ctrlWalk + 4)) * 65536) Or (CLng(data(ctrlWalk + 5)) * 16777216)
                    ctrlWalk = ctrlWalk + 6 + ctrlNameSize
                    If ctrlWalk + 6 <= UBound(data) Then
                      Dim ctrlNameUniSize As Long
                      ctrlNameUniSize = CLng(data(ctrlWalk + 2)) Or (CLng(data(ctrlWalk + 3)) * 256&) Or (CLng(data(ctrlWalk + 4)) * 65536) Or (CLng(data(ctrlWalk + 5)) * 16777216)
                      ctrlWalk = ctrlWalk + 6 + ctrlNameUniSize
                    End If
                  End If
                End If

                If refCount = UBound(refList) Then ReDim Preserve refList(0 To refCount + 31)
                refCount = refCount + 1
                refList(refCount).Name = refName
                refList(refCount).RefType = "Control"

                If ctrlWalk + 12 <= UBound(data) Then
                  Dim ctrlExtId As Long
                  ctrlExtId = CLng(data(ctrlWalk)) Or (CLng(data(ctrlWalk + 1)) * 256&)
                  If ctrlExtId = &H30 Then
                    Dim ctrlExtSize As Long
                    ctrlExtSize = CLng(data(ctrlWalk + 2)) Or (CLng(data(ctrlWalk + 3)) * 256&) Or (CLng(data(ctrlWalk + 4)) * 65536) Or (CLng(data(ctrlWalk + 5)) * 16777216)
                    Dim ctrlLibidExtSize As Long
                    ctrlLibidExtSize = CLng(data(ctrlWalk + 6)) Or (CLng(data(ctrlWalk + 7)) * 256&) Or (CLng(data(ctrlWalk + 8)) * 65536) Or (CLng(data(ctrlWalk + 9)) * 16777216)

                    If ctrlLibidExtSize > 0 And ctrlWalk + 10 + ctrlLibidExtSize <= UBound(data) + 1 Then
                      Dim ctrlLibidExt As String
                      ctrlLibidExt = BytesToString(data, ctrlWalk + 10, ctrlLibidExtSize, m_Project.CodePage)
                      Dim ctrlGStart As Long, ctrlGEnd As Long
                      ctrlGStart = InStr(ctrlLibidExt, "{")
                      ctrlGEnd = InStr(ctrlGStart + 1, ctrlLibidExt, "}")
                      If ctrlGStart > 0 And ctrlGEnd > ctrlGStart Then
                        refList(refCount).GUID = Mid$(ctrlLibidExt, ctrlGStart, ctrlGEnd - ctrlGStart + 1)
                      End If
                      Dim ctrlParts() As String: ctrlParts = Split(ctrlLibidExt, "#")
                      If UBound(ctrlParts) >= 4 Then
                        refList(refCount).Description = ctrlParts(4)
                      End If
                    End If
                    ctrlWalk = ctrlWalk + 6 + ctrlExtSize
                  End If
                End If

                pos = ctrlWalk
                GoTo ContinueScan
            End Select
          End If

SkipReferenceName:
          pos = nextRecPos
          GoTo ContinueScan
        End If
      End If
    End If

    ' MS-OVBA 2.3.4.2.3.2.1 MODULENAME (0x0019)
    If RecordId = &H19 Then
      RecordSize = CLng(data(pos + 2)) Or (CLng(data(pos + 3)) * 256&) Or (CLng(data(pos + 4)) * 65536) Or (CLng(data(pos + 5)) * 16777216)
      If RecordSize > 0 And RecordSize < 256 Then
        Dim blankMod As ModuleRecord
        Dim modRec   As ModuleRecord
        modRec = blankMod

        Dim modPos As Long: modPos = pos

        Do While modPos < UBound(data) - 5
          Dim m_Id  As Long, mSize As Long
          m_Id = CLng(data(modPos)) Or (CLng(data(modPos + 1)) * 256&)
          mSize = CLng(data(modPos + 2)) Or (CLng(data(modPos + 3)) * 256&) Or (CLng(data(modPos + 4)) * 65536) Or (CLng(data(modPos + 5)) * 16777216)
          modPos = modPos + 6
          If modPos + mSize > UBound(data) + 1 Then Exit Do

          Select Case m_Id
            Case &H19: modRec.Name = BytesToString(data, modPos, mSize, m_Project.CodePage)
            Case &H1A: modRec.streamName = BytesToString(data, modPos, mSize, m_Project.CodePage)
            Case &H21: modRec.ModType = modStandard
            Case &H22: modRec.ModType = modClass
            Case &H31
              modRec.offset = CLng(data(modPos)) Or (CLng(data(modPos + 1)) * 256&) Or (CLng(data(modPos + 2)) * 65536) Or (CLng(data(modPos + 3)) * 16777216)
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
    Next
    m_Project.ModuleCount = count
  End If

  If refCount > 0 Then
    ReDim m_Project.References(1 To refCount)
    For i = 1 To refCount
      m_Project.References(i) = refList(i)
    Next
    m_Project.ReferenceCount = refCount
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
    Next
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
  Next

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
  Next

  Dim dataLengthValue As Double
  dataLengthValue = 0#
  For i = 0 To 3
    decodedByte = DecodeDataByte(encrypted(pos), encryptedByte1, encryptedByte2, unencryptedByte1)
    dataLengthValue = dataLengthValue + CDbl(decodedByte) * (256# ^ i)
    pos = pos + 1
  Next

  If dataLengthValue < 1# Or dataLengthValue > CDbl(byteLen - pos) Then Exit Function
  If dataLengthValue <> CDbl(byteLen - pos) Then Exit Function

  Dim firstDataByte As Long
  For i = 1 To CLng(dataLengthValue)
    decodedByte = DecodeDataByte(encrypted(pos), encryptedByte1, _
                                    encryptedByte2, unencryptedByte1)
    If i = 1 Then firstDataByte = decodedByte
    pos = pos + 1
  Next

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
    Dim Node As Object
    Set Node = CreateObject("Scripting.Dictionary")
    Node("name") = entry("name")
    Node("type") = entry("type")
    Node("size") = entry("size")

    If entry("type") = STGTY_STORAGE Then
      Dim childStg   As stdole.IUnknown
      Dim childItems As Collection
      On Error Resume Next
      Set childStg = StorageOpenStorage(stg, entry("name"))
      On Error GoTo 0
      If Not childStg Is Nothing Then
        Set childItems = StorageEnumElements(childStg)
        Set Node("children") = BuildTreeNodes(childStg, childItems)
        Set childStg = Nothing
      Else
        Set Node("children") = New Collection
      End If
    End If

    nodes.Add Node
  Next

  Set BuildTreeNodes = nodes
End Function

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
  Dim hMem      As LongPtr
  Dim lpMem     As LongPtr
  Dim lockBytes As stdole.IUnknown
  Dim hr        As Long
  Dim dataSize  As LongPtr

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

  hr = CreateILockBytesOnHGlobal(hMem, 1, lockBytes)
  If hr < 0 Then
    GlobalFree hMem
    SetError "[StorageOpenBytes] CreateILockBytesOnHGlobal failed: &H" & Hex(hr)
    Exit Function
  End If

  hr = StgOpenStorageOnILockBytes(ObjPtr(lockBytes), 0, STGM_READ Or STGM_SHARE_EXCLUSIVE, 0, 0, StorageOpenBytes)
  If hr < 0 Then SetError "[StorageOpenBytes] StgOpenStorageOnILockBytes failed: &H" & Hex(hr)
  Set lockBytes = Nothing
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

Private Sub CopyStreamVerbatim(ByRef srcStg As stdole.IUnknown, ByRef dstStg As stdole.IUnknown, ByVal Name As String, ByRef keepAlive As Collection)
  Dim srcStm As stdole.IUnknown, dstStm As stdole.IUnknown, b() As Byte
  Set srcStm = StorageOpenStream(srcStg, Name)
  If srcStm Is Nothing Then Exit Sub
  keepAlive.Add srcStm
  b = StreamReadBytes(srcStm)
  Set dstStm = StorageCreateStream(dstStg, Name)
  If dstStm Is Nothing Then Exit Sub
  keepAlive.Add dstStm
  StreamWriteBytes dstStm, b
End Sub

Private Sub CopyStorageVerbatim(ByRef srcParent As stdole.IUnknown, ByRef dstParent As stdole.IUnknown, ByVal Name As String, ByRef keepAlive As Collection)
  Dim srcSub As stdole.IUnknown, dstSub As stdole.IUnknown
  Set srcSub = StorageOpenStorage(srcParent, Name)
  If srcSub Is Nothing Then Exit Sub
  keepAlive.Add srcSub
  Set dstSub = StorageCreateStorage(dstParent, Name)
  If dstSub Is Nothing Then Exit Sub
  keepAlive.Add dstSub
  Dim els As Collection, el As Variant, nm As String
  Set els = StorageEnumElements(srcSub)
  For Each el In els
    nm = el("name")
    If el("type") = 1 Then
      CopyStorageVerbatim srcSub, dstSub, nm, keepAlive
    Else
      CopyStreamVerbatim srcSub, dstSub, nm, keepAlive
    End If
  Next
End Sub

Private Function StorageCreateNew(ByRef lockBytesOut As stdole.IUnknown) As stdole.IUnknown
  Dim lockBytes As stdole.IUnknown, hr As Long
  hr = CreateILockBytesOnHGlobal(0, 1, lockBytes)
  If hr < 0 Then
    SetError "[StorageCreateNew] CreateILockBytesOnHGlobal: &H" & Hex(hr)
    Exit Function
  End If
  hr = StgCreateDocfileOnILockBytes(ObjPtr(lockBytes), STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE, 0, StorageCreateNew)
  If hr < 0 Then
    SetError "[StorageCreateNew] StgCreateDocfileOnILockBytes: &H" & Hex(hr)
    Exit Function
  End If
  Set lockBytesOut = lockBytes
End Function

Private Function StorageCreateStorage(ByRef pStg As stdole.IUnknown, ByVal Name As String) As stdole.IUnknown
  Dim hr As Long
  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_CreateStorage, StrPtr(Name), STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE, 0&, 0&, VarPtr(StorageCreateStorage))
  If hr < 0 Then LogMessage xrayDebug, "IStorage.CreateStorage(" & Name & ") &H" & Hex(hr)
End Function

Private Function StorageCreateStream(ByRef pStg As stdole.IUnknown, ByVal Name As String) As stdole.IUnknown
  Dim hr As Long
  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_CreateStream, StrPtr(Name), STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_EXCLUSIVE, 0&, 0&, VarPtr(StorageCreateStream))
  If hr < 0 Then LogMessage xrayDebug, "IStorage.CreateStream(" & Name & ") &H" & Hex(hr)
End Function

Private Function StreamWriteBytes(ByRef pStm As stdole.IUnknown, ByRef data() As Byte) As Boolean
  Dim hr As Long, written As Long, n As Long
  n = ByteCount(data)
  If n = 0 Then StreamWriteBytes = True: Exit Function
  hr = DispCallByVtbl(pStm, IStreamInterface.IStm_Write, VarPtr(data(LBound(data))), n, VarPtr(written))
  If hr < 0 Then LogMessage xrayDebug, "IStream.Write &H" & Hex(hr): Exit Function
  StreamWriteBytes = (written = n)
End Function

Private Function StorageCommit(ByRef pStg As stdole.IUnknown) As Boolean
  Dim hr As Long
  hr = DispCallByVtbl(pStg, IStorageInterface.IStg_Commit, 0&)
  If hr < 0 Then LogMessage xrayDebug, "IStorage.Commit &H" & Hex(hr)
  StorageCommit = (hr >= 0)
End Function

Private Function LockBytesGetData(ByRef lockBytes As stdole.IUnknown) As Byte()
  Dim hr As Long, st As STATSTG, n As Long, hGlob As LongPtr, lp As LongPtr
  hr = DispCallByVtbl(lockBytes, ILOCKBYTES_STAT, VarPtr(st), STATFLAG_NONAME)
  If hr < 0 Then SetError "[LockBytesGetData] ILockBytes.Stat: &H" & Hex(hr): Exit Function
  n = CLng(st.cbSize * 10000#)
  If n <= 0 Then Exit Function
  hr = GetHGlobalFromILockBytes(ObjPtr(lockBytes), hGlob)
  If hr < 0 Or hGlob = 0 Then
    SetError "[LockBytesGetData] GetHGlobalFromILockBytes: &H" & Hex(hr)
    Exit Function
  End If
  lp = GlobalLock(hGlob)
  If lp = 0 Then
    SetError "[LockBytesGetData] GlobalLock failed"
    Exit Function
  End If

  Dim out() As Byte
  ReDim out(0 To n - 1)
  CopyMemory VarPtr(out(0)), lp, n
  GlobalUnlock hGlob
  LockBytesGetData = out
End Function

Private Function CollectionContains(ByRef items As Collection, ByVal Name As String) As Boolean
  Dim entry As Variant
  For Each entry In items
    If StrComp(entry("name"), Name, vbTextCompare) = 0 Then
      CollectionContains = True
      Exit Function
    End If
  Next
End Function

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
    If inPos + 1 > inputUb Then GoTo ErrHandler

    Dim chunkHeader  As Long
    Dim isCompressed As Boolean
    Dim payloadSize  As Long
    Dim chunkEnd     As Long
    Dim chunkOutPos  As Long

    chunkHeader = CLng(CompressedData(inPos)) Or (CLng(CompressedData(inPos + 1)) * 256&)
    inPos = inPos + 2
    If (chunkHeader And &H7000&) <> &H3000& Then GoTo ErrHandler
    isCompressed = ((chunkHeader And &H8000&) = &H8000&)
    payloadSize = (chunkHeader And &HFFF&) + 1
    chunkEnd = inPos + payloadSize
    If chunkEnd > inputUb + 1 Then GoTo ErrHandler
    chunkOutPos = 0

    If Not isCompressed Then
      If outPos + 4096 > UBound(output) Then ReDim Preserve output(0 To outPos + 4096 + 65535)
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

          If (flagByte And m_BitMask(bitIndex)) = 0 Then
            If outPos > UBound(output) Then ReDim Preserve output(0 To outPos + 65535)
            output(outPos) = CompressedData(inPos)
            outPos = outPos + 1: chunkOutPos = chunkOutPos + 1: inPos = inPos + 1
          Else
            If inPos + 1 >= chunkEnd Then GoTo ErrHandler
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

            Dim shiftFactor As Long: shiftFactor = m_Pow2(16 - bitCount)
            Dim lengthMask  As Long: lengthMask = shiftFactor - 1
            Dim copyLength  As Long: copyLength = (token And lengthMask) + 3
            Dim copyOffset  As Long: copyOffset = (token \ shiftFactor) + 1

            If copyOffset > chunkOutPos Or copyOffset > outPos Then GoTo ErrHandler
            If copyLength > 4096 - chunkOutPos Then GoTo ErrHandler

            If outPos + copyLength > UBound(output) Then
              ReDim Preserve output(0 To outPos + copyLength + 65535)
            End If

            Dim j As Long
            For j = 0 To copyLength - 1
              output(outPos) = output(outPos - copyOffset)
              outPos = outPos + 1: chunkOutPos = chunkOutPos + 1
            Next
          End If
        Next
      Loop
    End If

    inPos = chunkEnd
  Loop

  If outPos > 0 Then
    ReDim Preserve output(0 To outPos - 1)
    DecompressStream = output
  End If
  Exit Function

ErrHandler:
  If outPos > 0 Then
    ReDim Preserve output(0 To outPos - 1)
    DecompressStream = output
  End If
  If Not Silent Then
    SetError "[DecompressStream] Invalid or truncated compressed stream"
  End If
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
  Next

  hResult = DispCallFunc(ObjPtr(pUnk), lIndex * PTR_SIZE, CC_STDCALL, vbLong, lIdx, vType(0), vPtr(0), DispCallByVtbl)
  If hResult < 0 Then LogMessage xrayDebug, "[DispCallByVtbl] DispCallFunc invoke failed: &H" & Hex(hResult)
End Function

Private Function PtrToString(ByVal stringPtr As LongPtr) As String
  Dim buf     As String
  Dim charLen As Long
  charLen = lstrlenW(stringPtr)
  If charLen > 0 Then
    buf = VBA.String$(charLen, vbNullChar)
    CopyMemory ByVal StrPtr(buf), ByVal stringPtr, charLen * 2
  End If
  PtrToString = buf
End Function

Private Function BytesToString(ByRef data() As Byte, ByVal pos As Long, ByVal size As Long, ByVal CodePage As Long) As String
  If size <= 0 Then Exit Function
  If ByteCount(data) = 0 Then Exit Function
  If pos < LBound(data) Or pos > UBound(data) Then Exit Function
  If size > UBound(data) - pos + 1 Then Exit Function
  Dim charCount As Long
  charCount = MultiByteToWideChar(CodePage, 0, VarPtr(data(pos)), size, 0, 0)
  If charCount = 0 Then Exit Function
  Dim result As String
  result = VBA.String$(charCount, 0)
  MultiByteToWideChar CodePage, 0, VarPtr(data(pos)), size, StrPtr(result), charCount
  BytesToString = result
End Function

Private Function StringToBytes(ByVal s As String, ByVal CodePage As Long) As Byte()
  Dim result() As Byte
  Dim size     As Long
  size = WideCharToMultiByte(CodePage, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
  If size = 0 Then Exit Function
  ReDim result(0 To size - 1)
  WideCharToMultiByte CodePage, 0, StrPtr(s), Len(s), VarPtr(result(0)), size, 0, 0
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
  Dim parent   As String
  slashPos = InStrRev(folderPath, "\")
  If slashPos > 0 Then parent = Left$(folderPath, slashPos - 1)
  If Len(parent) = 2 And Right$(parent, 1) = ":" Then parent = parent & "\"

  If Len(parent) > 0 And Not FolderExists(parent) Then
    If Not CreateFolderRecursive(parent) Then Exit Function
  End If
  CreateFolderRecursive = (apiCreateDirectoryW(StrPtr(folderPath & vbNullChar), 0) <> 0)
End Function

Private Function SliceBytes(ByRef data() As Byte, ByVal pos As Long, ByVal size As Long) As Byte()
  Dim result() As Byte
  If size <= 0 Then Exit Function
  If ByteCount(data) = 0 Then Exit Function
  If pos < LBound(data) Or pos > UBound(data) Then Exit Function
  If size > UBound(data) - pos + 1 Then Exit Function
  ReDim result(0 To size - 1)
  CopyMemory VarPtr(result(0)), VarPtr(data(pos)), size
  SliceBytes = result
End Function

Private Function WriteFile(ByVal Filename As String, ByRef fileData As Variant, Optional ByVal WriteBinary As Boolean = True) As Boolean
  On Error GoTo ErrHandler
  Dim hFile        As LongPtr
  Dim bytesWritten As Long
  Dim data()       As Byte

  If WriteBinary Then
    data = fileData
  Else
    data = StringToBytes(CStr(fileData), 65001)
  End If

  hFile = apiCreateFile(StrPtr(Filename), GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, 0, CREATE_ALWAYS, 0, 0)
  If hFile = INVALID_HANDLE_VALUE Then Exit Function

  Dim dataSize As Long
  dataSize = ByteCount(data)
  If dataSize > 0 Then
    If apiWriteFile(hFile, VarPtr(data(LBound(data))), dataSize, bytesWritten, 0) = 0 Then GoTo ErrHandler
    If bytesWritten <> dataSize Then GoTo ErrHandler
  End If

  apiCloseHandle hFile
  hFile = INVALID_HANDLE_VALUE
  WriteFile = True
  Exit Function

ErrHandler:
  If hFile <> 0 And hFile <> INVALID_HANDLE_VALUE Then apiCloseHandle hFile
End Function

Private Function ReadFile(ByVal FilePath As String, Optional ByVal MaxBytes As Long = 0) As Byte()
  Dim hFile     As LongPtr
  Dim fileSize  As LARGE_INTEGER
  Dim readSize  As Long
  Dim bytesRead As Long
  Dim result()  As Byte

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

  If fileSize.HighPart <> 0 Then
    readSize = IIf(MaxBytes > 0, MaxBytes, &H7FFFFFFF)
  Else
    readSize = fileSize.LowPart
    If MaxBytes > 0 And MaxBytes < readSize Then readSize = MaxBytes
  End If

  If readSize <= 0 Then apiCloseHandle hFile: Exit Function

  ReDim result(0 To readSize - 1)
  If apiReadFile(hFile, VarPtr(result(0)), readSize, bytesRead, 0) <> 0 Then
    If bytesRead = readSize Then ReadFile = result
  End If

  apiCloseHandle hFile
End Function

Private Function LoadFromAccessDB(ByVal FilePath As String) As Boolean
  Dim data() As Byte
  data = ReadFile(FilePath)

  If ByteCount(data) < 2048 Then
    SetError "[LoadFromAccessDB] File is too small to be an Access database."
    Exit Function
  End If

  Dim sig As String
  sig = StrConv(SliceBytes(data, 4, 15), vbUnicode)
  If sig <> "Standard ACE DB" And sig <> "Standard Jet DB" Then
    SetError "[LoadFromAccessDB] Invalid database signature."
    Exit Function
  End If

  Select Case data(&H14)
    Case &H0
      m_AccessPageSize = 2048
      m_AccessRowCntOff = 8
      m_AccessRowArrOff = 10
    Case Else
      m_AccessPageSize = 4096
      m_AccessRowCntOff = 12
      m_AccessRowArrOff = 14
  End Select

  LogMessage xrayInfo, "[LoadFromAccessDB] Jet version 0x" & Right$("0" & Hex$(data(&H14)), 2) & ", page size " & m_AccessPageSize & " bytes."

  ParseAccessDB data
  LoadFromAccessDB = IsLoaded
  If Not LoadFromAccessDB Then
    LogMessage xrayWarning, "[LoadFromAccessDB] No validated VBA module source " & _
        "was found (Jet " & IIf(m_AccessPageSize = 2048, "3", "4/ACE") & ", " & _
        m_AccessPageSize & "-byte pages)."
  End If
End Function

Private Sub ParseAccessDB(ByRef data() As Byte)
  Reset
  m_AccessRecoverCnt = 0
  Set m_AccessSeenBlobs = New Collection
  Dim pageCount As Long
  pageCount = (UBound(data) + 1) \ m_AccessPageSize

  Dim p As Long
  For p = 0 To pageCount - 1
    Dim base As Long
    base = p * m_AccessPageSize
    If base + (m_AccessPageSize - 1) > UBound(data) Then Exit For

'    If data(base) = &H1 And data(base + 4) = &H4C And data(base + 5) = &H56 And data(base + 6) = &H41 And data(base + 7) = &H4C Then
    If data(base) = &H1 Then
      LogMessage xrayDebug, "[ParseAccessDB] scanning page " & p
      ParseLVALPage data, p
    End If
  Next
End Sub

Private Sub ParseLVALPage(ByRef data() As Byte, ByVal p As Long)
  Dim base As Long
  base = p * m_AccessPageSize

  Dim slotCount As Long
  slotCount = CLng(data(base + m_AccessRowCntOff)) Or (CLng(data(base + m_AccessRowCntOff + 1)) * 256&)

  If slotCount > (m_AccessPageSize - m_AccessRowArrOff) \ 2 Then slotCount = (m_AccessPageSize - m_AccessRowArrOff) \ 2

  Dim slot As Long
  For slot = 0 To slotCount - 1
    Dim slotOffset As Long
    slotOffset = base + m_AccessRowArrOff + (slot * 2)

    Dim raw As Long
    raw = CLng(data(slotOffset)) Or (CLng(data(slotOffset + 1)) * 256&)

    If (raw And &HF000&) <> &HD000& Then
      Dim rowStart As Long
      rowStart = raw And (m_AccessPageSize - 1)
      Dim rowEnd   As Long
      rowEnd = m_AccessPageSize

      Dim i As Long
      For i = 0 To slotCount - 1
        Dim otherRaw As Long
        otherRaw = CLng(data(base + m_AccessRowArrOff + (i * 2))) Or (CLng(data(base + m_AccessRowArrOff + (i * 2) + 1)) * 256&)

        If (otherRaw And &HF000&) <> &HD000& Then
          Dim o As Long
          o = otherRaw And (m_AccessPageSize - 1)
          If o > rowStart And o < rowEnd Then rowEnd = o
        End If
      Next

      Dim rowLength As Long
      rowLength = rowEnd - rowStart
      If rowLength > 0 Then
        ProcessAccessLVALRow data, p, slot, base + rowStart, rowLength
      End If
    End If
  Next
End Sub

Private Sub ProcessAccessLVALRow(ByRef data() As Byte, ByVal p As Long, ByVal slot As Long, ByVal rowStart As Long, ByVal rowLength As Long)
  Dim rowBytes() As Byte
  rowBytes = SliceBytes(data, rowStart, rowLength)

  Dim i As Long, hdr As Long
  Dim chainOK As Boolean
  For i = 0 To rowLength - 4
    If rowBytes(i) = &H1 Then
      hdr = CLng(rowBytes(i + 1)) Or (CLng(rowBytes(i + 2)) * 256&)
      If ((hdr \ m_AccessPageSize) And 7) = 3 Then
        LogMessage xrayDebug, "[ProcessAccessLVALRow] candidate at page " & p & " slot " & slot & " offset " & i
        TryDecompressAccessBlob SliceBytes(rowBytes, i, rowLength - i)
      End If
    End If
  Next

  If rowLength >= 4 Then
    Dim next_p As Long
    next_p = CLng(rowBytes(1)) Or (CLng(rowBytes(2)) * 256&) Or (CLng(rowBytes(3)) * 65536)
    If next_p > 0 And next_p < ((UBound(data) + 1) \ m_AccessPageSize) Then
      Dim nextBase As Long
      nextBase = next_p * m_AccessPageSize
      If data(nextBase) = &H1 And data(nextBase + 4) = &H4C And data(nextBase + 5) = &H56 And data(nextBase + 6) = &H41 And data(nextBase + 7) = &H4C Then
        Dim chainBlob() As Byte
        If TryWalkAccessLVALChain(data, p, slot, chainBlob) Then
          If ByteCount(chainBlob) >= 4 Then
            For i = LBound(chainBlob) To UBound(chainBlob) - 3
              If chainBlob(i) = &H1 Then
                hdr = CLng(chainBlob(i + 1)) Or (CLng(chainBlob(i + 2)) * 256&)
                If ((hdr \ m_AccessPageSize) And 7) = 3 Then
                  TryDecompressAccessBlob SliceBytes(chainBlob, i, UBound(chainBlob) - i + 1)
                End If
              End If
            Next
            chainOK = True
          End If
        End If
      End If
    End If
  End If

  If chainOK Then
    TryExtractAccessSourceTable chainBlob
  Else
    TryExtractAccessSourceTable rowBytes
  End If
End Sub

Private Function TryWalkAccessLVALChain(ByRef data() As Byte, ByVal startPage As Long, ByVal startSlot As Long, ByRef outBlob() As Byte) As Boolean
  Dim cur_p As Long, cur_s As Long
  cur_p = startPage
  cur_s = startSlot

  Dim tempBlob()   As Byte
  Dim totalLen     As Long: totalLen = 0
  Dim maxChunks    As Long: maxChunks = 4096
  Dim chunkCount   As Long: chunkCount = 0

  Do While chunkCount < maxChunks
    chunkCount = chunkCount + 1

    Dim base As Long: base = cur_p * m_AccessPageSize
    If base > UBound(data) Then Exit Function

    Dim slotCount As Long
    slotCount = CLng(data(base + m_AccessRowCntOff)) Or (CLng(data(base + m_AccessRowCntOff + 1)) * 256&)
    If slotCount > (m_AccessPageSize - m_AccessRowArrOff) \ 2 Then slotCount = (m_AccessPageSize - m_AccessRowArrOff) \ 2
    If cur_s >= slotCount Then Exit Function

    Dim slotOffset As Long
    slotOffset = base + m_AccessRowArrOff + (cur_s * 2)
    Dim raw As Long: raw = CLng(data(slotOffset)) Or (CLng(data(slotOffset + 1)) * 256&)
    If (raw And &HF000&) = &HD000& Then Exit Function

    Dim rowStart As Long: rowStart = raw And (m_AccessPageSize - 1)
    Dim rowEnd   As Long: rowEnd = m_AccessPageSize
    Dim i        As Long
    For i = 0 To slotCount - 1
      Dim otherRaw As Long
      otherRaw = CLng(data(base + m_AccessRowArrOff + (i * 2))) Or (CLng(data(base + m_AccessRowArrOff + (i * 2) + 1)) * 256&)
      If (otherRaw And &HF000&) <> &HD000& Then
        Dim o As Long
        o = otherRaw And (m_AccessPageSize - 1)
        If o > rowStart And o < rowEnd Then rowEnd = o
      End If
    Next

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

    If next_p >= ((UBound(data) + 1) \ m_AccessPageSize) Then Exit Function
    cur_p = next_p
    cur_s = next_s
  Loop
End Function

Private Function BlobHasVBASignature(ByRef blob() As Byte) As Boolean
  Dim n As Long: n = ByteCount(blob)
  If n < 10 Then Exit Function
  Dim lb As Long: lb = LBound(blob)
  Dim i As Long
  For i = lb To lb + n - 10
    If blob(i) = &H5F And blob(i + 1) = &H0 And blob(i + 2) = &H56 And _
       blob(i + 3) = &H0 And blob(i + 4) = &H42 And blob(i + 5) = &H0 And _
       blob(i + 6) = &H41 And blob(i + 7) = &H0 And blob(i + 8) = &H5F And _
       blob(i + 9) = &H0 Then
      BlobHasVBASignature = True
      Exit Function
    End If
  Next
End Function

Private Sub TryDecompressAccessBlob(ByRef compData() As Byte)
  On Error GoTo IgnoreDecompFail
  Dim decomp() As Byte
  decomp = DecompressStream(compData, True)

  If ByteCount(decomp) < 20 Then Exit Sub

  If decomp(0) = &H41 And decomp(1) = &H74 And decomp(2) = &H74 And _
     decomp(3) = &H72 And decomp(4) = &H69 And decomp(5) = &H62 And _
     decomp(6) = &H75 And decomp(7) = &H74 And decomp(8) = &H65 Then
    If decomp(9) = &H20 And decomp(10) = &H56 And decomp(11) = &H42 And _
       decomp(12) = &H5F And decomp(13) = &H4E And decomp(14) = &H61 And _
       decomp(15) = &H6D And decomp(16) = &H65 And _
       decomp(17) = &H20 And decomp(18) = &H3D And decomp(19) = &H20 Then
      ExtractAccessModule decomp
    End If
  End If
IgnoreDecompFail:
End Sub

Private Sub ExtractAccessModule(ByRef decomp() As Byte)
  If ByteCount(decomp) = 0 Then Exit Sub
  Dim text As String
  text = BytesToString(decomp, LBound(decomp), ByteCount(decomp), m_Project.CodePage)
  text = Replace(text, vbCrLf, vbLf)
  text = Replace(text, vbCr, vbLf)

  Dim lines() As String
  lines = Split(text, vbLf)

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
  Next

  If Len(modName) = 0 Then Exit Sub

  If m_Project.ModuleCount > 0 Then
    For i = 1 To m_Project.ModuleCount
      If StrComp(m_Project.Modules(i).Name, modName, vbTextCompare) = 0 Then Exit Sub
    Next
  End If

  Dim newCount As Long
  newCount = m_Project.ModuleCount + 1
  m_Project.ModuleCount = newCount

  Dim ub As Long: ub = 0
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
  Next
  If isClass Then m_Project.Modules(newCount).ModType = modClass

'  Dim bodyLines() As String
'  If bodyStart <= UBound(lines) Then
'    Dim j As Long: j = 0
'    ReDim bodyLines(0 To UBound(lines) - bodyStart)
'    For i = bodyStart To UBound(lines)
'      bodyLines(j) = lines(i)
'      j = j + 1
'    Next
'    m_Project.Modules(newCount).SourceCode = Join(bodyLines, vbCrLf)
'  End If
  m_Project.Modules(newCount).SourceCode = Join(lines, vbCrLf)
End Sub

Private Sub LogMessage(ByVal level As XrayLogLevel, ByVal msg As String)
  If level < m_LogLevel Then Exit Sub

  Dim prefix As String
  Select Case level
    Case xrayDebug:   prefix = "[DBG] "
    Case xrayInfo:    prefix = "[INF] "
    Case xrayWarning: prefix = "[WRN] "
    Case xrayError:   prefix = "[ERR] "
    Case Else:        prefix = "[   ] "
  End Select

  Debug.Print prefix & msg
End Sub

Private Sub SetError(ByVal msg As String)
  m_LastError = msg
  Debug.Print "[ERR] " & msg
End Sub

Private Sub ClearError()
  m_LastError = vbNullString
End Sub

Private Sub Class_Initialize()
  m_LogLevel = xrayInfo
  
  Dim i As Long
  For i = 0 To 16
    m_Pow2(i) = 2 ^ i
  Next i
  For i = 0 To 7
    m_BitMask(i) = 2 ^ i
  Next i
  
  Reset
End Sub

Private Function COMQueryInterface(ByRef pUnk As stdole.IUnknown, ByVal iid As String) As stdole.IUnknown
  If pUnk Is Nothing Then Exit Function
  Dim riid   As UUID
  Dim result As stdole.IUnknown
  CLSIDFromString StrPtr(iid), riid
  DispCallByVtbl pUnk, 0, VarPtr(riid), VarPtr(result)
  Set COMQueryInterface = result
End Function

Private Sub TryExtractAccessSourceTable(ByRef blob() As Byte)
  On Error GoTo Done

  Dim n As Long: n = ByteCount(blob)
  If n < 16 Then Exit Sub
  Dim lb As Long: lb = LBound(blob)
  If Not BlobHasVBASignature(blob) Then Exit Sub

  If m_AccessSeenBlobs Is Nothing Then Set m_AccessSeenBlobs = New Collection
  Dim fp As String
  fp = BlobFingerprint(blob)
  On Error Resume Next
  Dim seen As Variant
  seen = m_AccessSeenBlobs(fp)
  If Err.Number = 0 Then
    On Error GoTo Done
    Exit Sub
  End If
  Err.Clear
  On Error GoTo Done
  m_AccessSeenBlobs.Add True, fp

  Dim outBuf() As Byte, recovered As Long, i As Long, outLen As Long
  ReDim outBuf(0 To n + 2)
  outLen = 0
  i = lb

  Do While i < lb + n - 6
    If blob(i) = &HD8 And blob(i + 3) = &H0 Then
      Dim ln As Long: ln = CLng(blob(i + 4)) Or (CLng(blob(i + 5)) * 256&)
      If ln >= 1 And ln <= 200 And (i + 6 + ln) <= (lb + n) Then

        Dim startOut As Long: startOut = outLen
        Dim k As Long, b As Long, printableCnt As Long, kept As Long, hasContent As Boolean
        printableCnt = 0: kept = 0: hasContent = False

        For k = 0 To ln - 1
          b = blob(i + 6 + k)
          If b = &HA0 Then b = &H20
          If b <> &HAD Then
            If (b >= 32 And b < 127) Or b = 9 Or b = 13 Or b = 10 Then printableCnt = printableCnt + 1
            If b > 32 Then hasContent = True
            outBuf(outLen) = b: outLen = outLen + 1: kept = kept + 1
          End If
        Next

        If kept > 0 And hasContent And (printableCnt * 100 >= kept * 97) Then
          outBuf(outLen) = 13: outLen = outLen + 1
          outBuf(outLen) = 10: outLen = outLen + 1
          recovered = recovered + 1
          i = i + 6 + ln
          GoTo ContinueLoop
        Else
          outLen = startOut
        End If
      End If
    End If
    i = i + 1
ContinueLoop:
  Loop

  If recovered = 0 Then Exit Sub
  LogMessage xrayInfo, "[TryExtractAccessSourceTable] recovered " & recovered & " uncompressed source line(s)."

  Dim sb As String
  sb = BytesToString(outBuf, 0, outLen, m_Project.CodePage)
  If InStr(1, sb, "Attribute VB_Name = """, vbTextCompare) = 0 Then
    m_AccessRecoverCnt = m_AccessRecoverCnt + 1
    sb = "Attribute VB_Name = ""AccessModule_" & m_AccessRecoverCnt & """" & vbCrLf & sb
    LogMessage xrayWarning, "[TryExtractAccessSourceTable] no VB_Name in blob; assigned placeholder AccessModule_" & m_AccessRecoverCnt & "."
  End If

  ExtractAccessModule StringToBytes(sb, m_Project.CodePage)
Done:
End Sub

Private Function BlobFingerprint(ByRef blob() As Byte) As String
  Dim n As Long, lb As Long, s As String, j As Long
  n = ByteCount(blob)
  If n = 0 Then Exit Function
  lb = LBound(blob)
  s = CStr(n)
  For j = 1 To 8
    s = s & ":" & CStr(blob(lb + ((n - 1) * j) \ 8))
  Next
  BlobFingerprint = s
End Function

Public Function ExtractFormFRX(ByRef vbaProjectBin() As Byte, ByVal parentPath As String, ByVal FormName As String, ByVal FrxPath As String) As Boolean
  Dim root As stdole.IUnknown, formStg As stdole.IUnknown
  Dim lockBytes As stdole.IUnknown, dst As stdole.IUnknown
  Dim keepAlive As New Collection

  Set root = StorageOpenBytes(vbaProjectBin)
  If root Is Nothing Then
    SetError "[ExtractFormFRX] cannot open vbaProject.bin"
    Exit Function
  End If
  keepAlive.Add root

  Dim parentStg As stdole.IUnknown
  If Len(parentPath) > 0 Then
    Set parentStg = StorageNavigate(root, parentPath, keepAlive)
    If parentStg Is Nothing Then
      SetError "[ExtractFormFRX] parent path not found: " & parentPath
      Exit Function
    End If
  Else
    Set parentStg = root
  End If

  Set formStg = StorageOpenStorage(parentStg, FormName)
  If formStg Is Nothing Then
    SetError "[ExtractFormFRX] form storage not found: " & FormName
    Exit Function
  End If
  keepAlive.Add formStg

  Dim cx As Long, cy As Long
  ReadFormExtent formStg, cx, cy

  Set dst = StorageCreateNew(lockBytes)
  If dst Is Nothing Then Exit Function
  keepAlive.Add dst: keepAlive.Add lockBytes

  Dim els As Collection, el As Variant, nm As String
  Set els = StorageEnumElements(formStg)
  For Each el In els
    nm = el("name")
    If nm <> VBA.Chr$(1) & "CompObj" And nm <> VBA.Chr$(3) & "VBFrame" Then
      If el("type") = 1 Then
        CopyStorageVerbatim formStg, dst, nm, keepAlive
      Else
        CopyStreamVerbatim formStg, dst, nm, keepAlive
      End If
    End If
  Next

  Dim compObj As stdole.IUnknown
  Set compObj = StorageCreateStream(dst, VBA.Chr$(1) & "CompObj")
  If Not compObj Is Nothing Then
    keepAlive.Add compObj
    StreamWriteBytes compObj, FRXRootCompObj()
  End If

  StorageCommit dst

  Dim cfb() As Byte, outBytes() As Byte
  cfb = LockBytesGetData(lockBytes)
  If ByteCount(cfb) = 0 Then
    SetError "[ExtractFormFRX] compound file came back empty"
    Exit Function
  End If

  outBytes = PrependFRXHeader(cfb, cx, cy)
  If Not WriteFile(FrxPath, outBytes, True) Then Exit Function

  LogMessage xrayInfo, "[ExtractFormFRX] " & FormName & ".frx (" & ByteCount(outBytes) & " bytes)"
  ExtractFormFRX = True
End Function

Public Function ExportFormFRXFromFile(ByVal OfficePath As String, _
                                      ByVal FormName As String, _
                                      ByVal FrxPath As String) As Boolean
  ClearError
  Dim bin() As Byte
  bin = GetVBAProjectBinBytes(OfficePath)
  If ByteCount(bin) = 0 Then
    SetError "[ExportFormFRXFromFile] No vbaProject.bin found in " & OfficePath
    Exit Function
  End If

  Dim ext As String, layout As CFBLayout, parentPath As String, lastSlash As Long
  ext = LCase$(VBA.Mid$(OfficePath, InStrRev(OfficePath, ".") + 1))
  layout = GetCFBLayout(ext)
  lastSlash = InStrRev(layout.VBAPath, "\")
  If lastSlash > 0 Then parentPath = Left$(layout.VBAPath, lastSlash - 1)

  ExportFormFRXFromFile = ExtractFormFRX(bin, parentPath, FormName, FrxPath)
End Function

Private Function PrependFRXHeader(ByRef cfb() As Byte, ByVal cx As Long, ByVal cy As Long) As Byte()
  Dim CompoundFileSize As Long, out() As Byte
  CompoundFileSize = ByteCount(cfb)
  ReDim out(0 To CompoundFileSize + 23)
  PutDWord out, 0, &H8424C
  PutDWord out, 4, CompoundFileSize
  PutDWord out, 8, 0
  PutDWord out, 12, 0
  PutDWord out, 16, cx
  PutDWord out, 20, cy
  CopyMemory VarPtr(out(24)), VarPtr(cfb(LBound(cfb))), CompoundFileSize
  PrependFRXHeader = out
End Function

Private Sub PutDWord(ByRef b() As Byte, ByVal off As Long, ByVal val As Long)
  b(off) = val And &HFF
  b(off + 1) = (val \ &H100) And &HFF
  b(off + 2) = (val \ &H10000) And &HFF
  b(off + 3) = (val \ &H1000000) And &HFF
End Sub

Private Function FRXRootCompObj() As Byte()
  Const CompObjCode As String = "0100FEFF030A0000FFFFFFFFF0692AC6DC16CE119E9800AA00574A4F" & _
                                "190000004D6963726F736F667420466F726D7320322E3020466F726D" & _
                                "0010000000456D626564646564204F626A656374000D000000466F72" & _
                                "6D732E466F726D2E3100F439B271000000000000000000000000"
  FRXRootCompObj = HexToBytes(CompObjCode)
End Function

Private Function HexToBytes(ByVal h As String) As Byte()
  Dim n As Long, b() As Byte
  n = Len(h) \ 2
  ReDim b(0 To n - 1)
  Dim i As Long
  For i = 0 To n - 1
    b(i) = CByte("&H" & Mid$(h, i * 2 + 1, 2))
  Next
  HexToBytes = b
End Function

Private Sub ReadFormExtent(ByRef formStg As stdole.IUnknown, ByRef cx As Long, ByRef cy As Long)
  cx = 4800: cy = 3600
  Dim vbf As stdole.IUnknown, b() As Byte, t As String, w As Long, h As Long
  Set vbf = StorageOpenStream(formStg, Chr$(3) & "VBFrame")
  If vbf Is Nothing Then Exit Sub
  b = StreamReadBytes(vbf)
  If ByteCount(b) = 0 Then Exit Sub
  t = StrConv(b, vbUnicode)
  w = GrabIntAfter(t, "ClientWidth")
  h = GrabIntAfter(t, "ClientHeight")

  Dim leftVal As Long, topVal As Long
  leftVal = GrabIntAfter(t, "ClientLeft")
  topVal = GrabIntAfter(t, "ClientTop")

  If leftVal = 0 Then leftVal = 108
  If topVal = 0 Then topVal = 456

  If w > 0 Then cx = w + (2 * leftVal)
  If h > 0 Then cy = h + topVal + leftVal
End Sub

Private Function GrabIntAfter(ByVal text As String, ByVal Key As String) As Long
  Dim p As Long: p = InStr(1, text, Key, vbTextCompare)
  If p = 0 Then Exit Function
  p = InStr(p, text, "=")
  If p = 0 Then Exit Function
  Dim i As Long, ch As String, num As String
  For i = p + 1 To Len(text)
    ch = Mid$(text, i, 1)
    If ch >= "0" And ch <= "9" Then
      num = num & ch
    ElseIf Len(num) > 0 Then
      Exit For
    End If
  Next
  If Len(num) > 0 Then GrabIntAfter = CLng(num)
End Function

Private Function SynthesizeFRMContent(ByVal FormName As String, ByVal SourceCode As String, ByRef vbaProjectBin() As Byte) As String
  Dim root As stdole.IUnknown, formStg As stdole.IUnknown
  Dim keepAlive As New Collection
  Dim vbfHeader As String

  On Error Resume Next
  Set root = StorageOpenBytes(vbaProjectBin)
  If Not root Is Nothing Then
    Dim ext As String: ext = LCase$(VBA.Mid$(m_Project.FilePath, InStrRev(m_Project.FilePath, ".") + 1))
    Dim layout As CFBLayout: layout = GetCFBLayout(ext)
    Dim parentPath As String
    Dim lastSlash As Long: lastSlash = InStrRev(layout.VBAPath, "\")
    If lastSlash > 0 Then parentPath = Left$(layout.VBAPath, lastSlash - 1)

    Dim parentStg As stdole.IUnknown
    If Len(parentPath) > 0 Then
      Set parentStg = StorageNavigate(root, parentPath, keepAlive)
    Else
      Set parentStg = root
    End If

    Set formStg = StorageOpenStorage(parentStg, FormName)
    If Not formStg Is Nothing Then
      vbfHeader = GetProcessedVBFrame(FormName, formStg)
    End If
  End If
  On Error GoTo 0

  If Len(vbfHeader) = 0 Then
    vbfHeader = "VERSION 5.00" & vbCrLf & _
                "Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} " & FormName & " " & vbCrLf & _
                "   Caption         =   """ & FormName & """" & vbCrLf & _
                "   ClientHeight    =   3600" & vbCrLf & _
                "   ClientLeft      =   120" & vbCrLf & _
                "   ClientTop       =   468" & vbCrLf & _
                "   ClientWidth     =   4800" & vbCrLf & _
                "   OleObjectBlob   =   """ & FormName & ".frx"":0000" & vbCrLf & _
                "   StartUpPosition =   1  'CenterOwner" & vbCrLf & _
                "End" & vbCrLf
  End If

  SynthesizeFRMContent = vbfHeader & vbCrLf & CleanSourceAttributes(SourceCode)
End Function

Private Function GetProcessedVBFrame(ByVal FormName As String, ByRef formStg As stdole.IUnknown) As String
  Dim vbf As stdole.IUnknown, b() As Byte, t As String
  Set vbf = StorageOpenStream(formStg, Chr$(3) & "VBFrame")
  If vbf Is Nothing Then Exit Function
  b = StreamReadBytes(vbf)
  If ByteCount(b) = 0 Then Exit Function

  t = StrConv(b, vbUnicode)
  t = Replace(t, vbCrLf, vbLf)
  t = Replace(t, vbCr, vbLf)

  Dim lines() As String
  lines = Split(t, vbLf)
  Dim out() As String
  ReDim out(0 To UBound(lines) + 2)
  Dim count As Long, insertedBlob As Boolean
  count = 0: insertedBlob = False

  Dim i As Long
  For i = LBound(lines) To UBound(lines)
    Dim line As String: line = Trim$(lines(i))

    If InStr(1, line, "TypeInfoVer", vbTextCompare) = 1 Then
      ' Skip line
    ElseIf InStr(1, line, "StartUpPosition", vbTextCompare) = 1 Then
      out(count) = "   OleObjectBlob   =   """ & FormName & ".frx"":0000"
      insertedBlob = True: count = count + 1
      out(count) = lines(i): count = count + 1
    ElseIf line = "End" And Not insertedBlob Then
      out(count) = "   OleObjectBlob   =   """ & FormName & ".frx"":0000"
      insertedBlob = True: count = count + 1
      out(count) = lines(i): count = count + 1
    Else
      out(count) = lines(i): count = count + 1
    End If
  Next

  If count > 0 Then
    ReDim Preserve out(0 To count - 1)
    GetProcessedVBFrame = Join(out, vbCrLf)
  Else
    GetProcessedVBFrame = t
  End If
End Function

Private Function CleanSourceAttributes(ByVal src As String) As String
  src = Replace(src, vbCrLf, vbLf)
  src = Replace(src, vbCr, vbLf)

  Dim lines() As String: lines = Split(src, vbLf)
  Dim out()   As String: ReDim out(0 To UBound(lines))
  Dim count   As Long: count = 0
  Dim i       As Long

  For i = LBound(lines) To UBound(lines)
    Dim line As String: line = Trim$(lines(i))

    If InStr(1, line, "Attribute VB_Base", vbTextCompare) = 1 Or _
       InStr(1, line, "Attribute VB_TemplateDerived", vbTextCompare) = 1 Or _
       InStr(1, line, "Attribute VB_Customizable", vbTextCompare) = 1 Then
      ' do nothing
    Else
      out(count) = lines(i)
      count = count + 1
    End If
  Next

  If count > 0 Then
    ReDim Preserve out(0 To count - 1)
    CleanSourceAttributes = Join(out, vbCrLf)
  Else
    CleanSourceAttributes = src
  End If
End Function

#If False Then
  Dim xrayDebug, xrayInfo, xrayWarning, xrayError
  Dim modOther, modStandard, modClass, modDesigner, modDocument
#End If
