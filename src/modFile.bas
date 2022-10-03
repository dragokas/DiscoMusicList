Attribute VB_Name = "modFile"
'
' modFile module by Alex Dragokas
'

Option Explicit

Const MAX_PATH As Long = 260&
Const MAX_FILE_SIZE As Currency = 104857600@

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As Long
    pTo As Long
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type

Private Type RTL_OSVERSIONINFOEXW
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(127) As Integer
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
End Type

Public Enum VbFileAttributeExtended
    vbAll = -1&
    vbDirectory = 16& ' mean - include folders also
    vbFile = vbAll And Not vbDirectory
    vbSystem = 4&
    vbHidden = 2&
    vbReadOnly = 1
    vbNormal = 0&
    vbReparse = 1024& 'symlinks / junctions (not include hardlink to file; they reflect attributes of the target)
End Enum
#If False Then
    Dim vbAll, vbFile, vbReparse 'case sensitive protection against modification (for non-overloaded enum variables only)
#End If

Public Enum VB_FILE_ACCESS_MODE
    FOR_READ = 1
    FOR_READ_WRITE = 2
    FOR_OVERWRITE_CREATE = 4
End Enum
#If False Then
    Dim FOR_READ, FOR_READ_WRITE, FOR_OVERWRITE_CREATE
#End If

Public Enum CACHE_TYPE
    USE_CACHE
    NO_CACHE
End Enum
#If False Then
    Dim USE_CACHE, NO_CACHE
#End If

Public Enum ENUM_File_Date_Type
    Date_Created = 1
    Date_Modified = 2
    Date_Accessed = 3
End Enum
#If False Then
    Dim Date_Created, Date_Modified, Date_Accessed
#End If
 
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
 
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
 
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    lpszFileName(MAX_PATH) As Integer
    lpszAlternate(14) As Integer
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer
    dwStrucVersionh As Integer
    dwFileVersionMSl As Integer
    dwFileVersionMSh As Integer
    dwFileVersionLSl As Integer
    dwFileVersionLSh As Integer
    dwProductVersionMSl As Integer
    dwProductVersionMSh As Integer
    dwProductVersionLSl As Integer
    dwProductVersionLSh As Integer
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type

Private Type SYSTEMTIME
    wYear           As Integer
    wMonth          As Integer
    wDayOfWeek      As Integer
    wDay            As Integer
    wHour           As Integer
    wMinute         As Integer
    wSecond         As Integer
    wMilliseconds   As Integer
End Type

Private Declare Function CreateFileMapping Lib "kernel32.dll" Alias "CreateFileMappingW" (ByVal hFile As Long, ByVal lpAttributes As Long, ByVal flProtect As Long, ByVal dwMaximumSizeHigh As Long, ByVal dwMaximumSizeLow As Long, ByVal lpName As Long) As Long
Private Declare Function MapViewOfFile Lib "kernel32.dll" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32.dll" (ByVal lpBaseAddress As Long) As Long
Private Declare Function PathFileExists Lib "Shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'Private Declare Function SHFileExists Lib "shell32.dll" Alias "#45" (ByVal szPath As String) As Long
Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (OldValue As Long) As Long
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByVal OldValue As Long) As Long
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeW" (ByVal nDrive As Long) As Long
Private Declare Function GetLogicalDrives Lib "kernel32.dll" () As Long
Private Declare Function GetFileSizeEx Lib "kernel32.dll" (ByVal hFile As Long, lpFileSize As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToRead As Long, lpNumberOfByConstesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExW" (ByVal hKey As Long, ByVal lpSubKey As Long, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExW" (ByVal hKey As Long, ByVal lpValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, szData As Long, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function memcpy Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long) As Long
'Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Private Declare Function GetSystemWindowsDirectory Lib "kernel32.dll" Alias "GetSystemWindowsDirectoryW" (ByVal lpBuffer As Long, ByVal uSize As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpStringDest As Long, ByVal lpStringSrc As Long) As Long
Private Declare Function GetLongPathNameW Lib "kernel32.dll" (ByVal lpszShortPath As Long, ByVal lpszLongPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As Long, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueW" (pBlock As Any, ByVal lpSubBlock As Long, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal lSize As Long)
Private Declare Function RtlGetVersion Lib "NTDLL.DLL" (lpVersionInformation As RTL_OSVERSIONINFOEXW) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32.dll" Alias "ExpandEnvironmentStringsW" (ByVal lpSrc As Long, ByVal lpDst As Long, ByVal nSize As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function IsWow64Process Lib "kernel32.dll" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function ILCreateFromPath Lib "Shell32.dll" Alias "ILCreateFromPathW" (ByVal pszPath As Long) As Long
Private Declare Function SHOpenFolderAndSelectItems Lib "Shell32.dll" (ByVal pidlFolder As Long, ByVal cidl As Long, ByVal apidl As Long, ByVal dwFlags As Long) As Long
Private Declare Sub ILFree Lib "Shell32.dll" (ByVal pIDL As Long)
Private Declare Function OemToChar Lib "user32.dll" Alias "OemToCharA" (ByVal lpszScr As String, ByVal lpszDst As String) As Long
Private Declare Function CharToOem Lib "user32.dll" Alias "CharToOemA" (ByVal lpszScr As String, ByVal lpszDst As String) As Long
Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bDontOverwrite As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32.dll" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryW" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Long, lpSecurityAttributes As Any) As Long
Private Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationW" (lpFileOp As SHFILEOPSTRUCT) As Long

Const FILE_SHARE_READ           As Long = &H1&
Const FILE_SHARE_WRITE          As Long = &H2&
Const FILE_SHARE_DELETE         As Long = 4&
Const FILE_READ_ATTRIBUTES      As Long = &H80&
Const OPEN_EXISTING             As Long = 3&
Const CREATE_ALWAYS             As Long = 2&
Const GENERIC_READ              As Long = &H80000000
Const GENERIC_WRITE             As Long = &H40000000
Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10&
Const INVALID_HANDLE_VALUE      As Long = &HFFFFFFFF
Const ERROR_SUCCESS             As Long = 0&
Const INVALID_FILE_ATTRIBUTES   As Long = -1&
Const NO_ERROR                  As Long = 0&
Const FILE_BEGIN                As Long = 0&
Const FILE_CURRENT              As Long = 1&
Const FILE_END                  As Long = 2&
Const INVALID_SET_FILE_POINTER  As Long = &HFFFFFFFF
Const FILE_ATTRIBUTE_NORMAL     As Long = &H80
Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400&
Const ERROR_HANDLE_EOF          As Long = 38&
Const SEC_IMAGE                 As Long = &H1000000
Const PAGE_READONLY             As Long = 2&
Const FILE_MAP_READ             As Long = 4&
Const FILE_ATTRIBUTE_READONLY   As Long = 1&

Const DRIVE_FIXED               As Long = 3&
Const DRIVE_RAMDISK             As Long = 6&

Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Const KEY_QUERY_VALUE           As Long = &H1&
Const RegType_DWord             As Long = 4&

Const ch_Dot                    As String = "."
Const ch_DotDot                 As String = ".."
Const ch_Slash                  As String = "\"
Const ch_SlashAsterisk          As String = "\*"

Const S_OK                      As Long = 0&

Private lWow64Old               As Long
Private DriveTypeName           As New Collection
Private arrPathFolders()        As String
Private arrPathFiles()          As String
Private Total_Folders           As Long
Private Total_Files             As Long


Public Function FileExists(ByVal sFile As String, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
        
    AppendErrorLogCustom "FileExists - Begin", "File: " & sFile
    
    sFile = EnvironW(Trim$(sFile))
    If Len(sFile) = 0 Then Exit Function
    If Left$(sFile, 2) = "\\" Then Exit Function
    
    ' use 2 methods for reliability reason (both supported unicode pathes)
    Dim Ex(1) As Boolean
    Dim ret As Long
    Dim Redirect As Boolean, bOldStatus As Boolean
    Dim WFD     As WIN32_FIND_DATA
    Dim hFile   As Long
    
    If Not bUseWow64 Then Redirect = ToggleWow64FSRedirection(False, sFile, bOldStatus)
    
    ret = GetFileAttributes(StrPtr(sFile))
    If ret <> INVALID_HANDLE_VALUE And (0 = (ret And FILE_ATTRIBUTE_DIRECTORY)) Then
        Ex(0) = True
    ElseIf Err.LastDllError = 5 Then
        Ex(0) = True
    End If
    
    hFile = FindFirstFile(StrPtr(sFile), WFD)
    
    If hFile <> INVALID_HANDLE_VALUE Then
        If Not CBool(WFD.dwFileAttributes And vbDirectory) Then Ex(1) = True
        FindClose hFile
    ElseIf Err.LastDllError = 5 Then
        Ex(1) = True
    End If
    
    '// FS redirection reverting if need
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    FileExists = Ex(0) Or Ex(1)

    AppendErrorLogCustom "FileExists - End", "File: " & sFile, "bUseWow64: " & bUseWow64, "Exists: " & FileExists
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.FileExists", "File:", sFile
    If inIDE Then Stop: Resume Next
End Function

Public Function FolderExists(ByVal sFolder$, Optional bUseWow64 As Boolean) As Boolean
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "FolderExists - Begin", "Folder: " & sFolder, "bUseWow64: " & bUseWow64
    
    Dim ret As Long, Redirect As Boolean, bOldStatus As Boolean
    
    AppendErrorLogCustom "FolderExists - Begin", "Folder: " & sFolder
    
    sFolder = Trim$(sFolder)
    If Len(sFolder) = 0 Then Exit Function
    If Left$(sFolder, 2) = "\\" Then Exit Function 'network path
    
    If Not bUseWow64 Then Redirect = ToggleWow64FSRedirection(False, sFolder, bOldStatus)
    
    ret = GetFileAttributes(StrPtr(sFolder))
    If CBool(ret And vbDirectory) And (ret <> INVALID_FILE_ATTRIBUTES) Then
        FolderExists = True
    ElseIf Err.LastDllError = 5 Then
        FolderExists = True
    End If
    
    '// FS redirection reverting if need
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    AppendErrorLogCustom "FolderExists - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.FolderExists", "Folder:", sFolder$, "Redirection: ", bUseWow64
    If inIDE Then Stop: Resume Next
End Function


'Public Sub GetDriveTypeNames()
'    On Error GoTo ErrorHandler
'    Dim lr As Long
'    Dim i  As Long
'    Dim DT As String
'
'    For i = 65& To 90&
'
'      lr = GetDriveType(StrPtr(Chr$(i) & ":\"))
'
'      Select Case lr
'        Case 3&
'            DT = "FIXED"
'        Case 2&
'            DT = "REMOVABLE"
'        Case 5&
'            DT = "CDROM"
'        Case 4&
'            DT = "REMOTE"
'        Case 0&
'            DT = "UNKNOWN"
'        Case 1&
'            DT = "DISCONNECTED" '"NO_ROOT_DIR"
'        Case 6&
'            DT = "RAMDISK"
'        Case Else
'            DT = "UNKNOWN"
'      End Select
'
'      DriveTypeName.Add DT, Chr$(i)
'
'    Next
'
'    Exit Sub
'ErrorHandler:
'    ErrorMsg err, "modFile.GetDriveTypeNames", "Drive:", Chr$(i)
'End Sub


Function FileLenW(Optional Path As String, Optional hFileHandle As Long) As Currency ', Optional DoNotUseCache As Boolean
    On Error GoTo ErrorHandler
    
    AppendErrorLogCustom "FileLenW - Begin", "Path: " & Path, "Handle: " & hFileHandle
    
'    ' Last cached File
'    Static CachedFile As String
'    Static CachedSize As Currency
    
    Dim lr          As Long
    Dim hFile       As Long
    Dim FileSize    As Currency

'    If Not DoNotUseCache Then
'        If StrComp(Path, CachedFile, 1) = 0 Then
'            FileLenW = CachedSize
'            Exit Function
'        End If
'    End If

    If hFileHandle = 0 Then
        hFile = CreateFile(StrPtr(Path), FILE_READ_ATTRIBUTES, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    Else
        hFile = hFileHandle
    End If
    
    If hFile > 0 Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then FileLenW = FileSize * 10000&
        End If
'        If Not DoNotUseCache Then
'            CachedFile = Path
'            CachedSize = FileLenW
'        End If
        If hFileHandle = 0 Then CloseHandle hFile: hFile = 0&
    End If
    
    AppendErrorLogCustom "FileLenW - End", "Size: " & FileSize
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.FileLenW", "File:", Path, "hFile:", hFile, "FileSize:", FileLenW, "Return:", lr
End Function



Public Function OpenW(FileName As String, Access As VB_FILE_ACCESS_MODE, retHandle As Long, Optional MountToMemory As Boolean) As Boolean '// TODO: MountToMemory
    
    AppendErrorLogCustom "OpenW - Begin", "File: " & FileName, "Access: " & Access
    
    Dim FSize As Currency
    
    If Access And (FOR_READ Or FOR_READ_WRITE) Then
        If Not FileExists(FileName) Then
            retHandle = INVALID_HANDLE_VALUE
            Exit Function
        End If
    End If
        
    If Access = FOR_READ Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_OVERWRITE_CREATE Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, CREATE_ALWAYS, ByVal 0&, ByVal 0&)
    ElseIf Access = FOR_READ_WRITE Then
        retHandle = CreateFile(StrPtr(FileName), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, ByVal 0&, ByVal 0&)
    Else
        'WriteCon "Wrong access mode!", cErr
    End If

    OpenW = (INVALID_HANDLE_VALUE <> retHandle)
    
    ' ����������� �� ����������� ��������� ���� ��� �������� ( > 100 �� )
    If OpenW Then
        If Access And (FOR_READ Or FOR_READ_WRITE) Then
            FSize = LOFW(retHandle)
            If FSize > MAX_FILE_SIZE Then
                CloseHandle retHandle
                retHandle = INVALID_HANDLE_VALUE
                OpenW = False
                '"�� ���� � �� ���� ��������� ���� ����, ������ ��� ��� ������ ��������� ���������� ��������"
                Err.Clear: MsgBox "modFile.OpenW." & " Trying to open too big file" & ": (" & (FSize \ 1024 \ 1024) & " MB.) " & FileName
            End If
        End If
    Else
        MsgBox "modFile.OpenW. " & "Cannot open file: " & FileName
        'Err.Raise 75 ' Path/File Access error
    End If

    AppendErrorLogCustom "OpenW - End", "Handle: " & retHandle
End Function

                                                                  'do not change Variant type at all or you will die ^_^
Public Function GetW(hFile As Long, pos As Long, Optional vOut As Variant, Optional vOutPtr As Long, Optional cbToRead As Long) As Boolean
    
    'On Error GoTo ErrorHandler
    AppendErrorLogCustom "GetW - Being", "Handle: " & hFile, "pos: " & pos, "cbToRead: " & cbToRead
    
    Dim lBytesRead  As Long
    Dim lr          As Long
    Dim ptr         As Long
    Dim vType       As Long
    Dim UnknType    As Boolean
    
    If pos > -1 Then
        pos = pos - 1   ' VB's Get & SetFilePointer difference correction
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then
            Err.Clear: ErrorMsg Err, "Cannot set file pointer!": Err.Raise 52
        End If
        If NO_ERROR <> Err.LastDllError Then
            Err.Clear: ErrorMsg Err, "Cannot set file pointer!": Err.Raise 52
        End If
    End If
        
    vType = VarType(vOut)
            
    If 0 <> cbToRead Then   'vbError = vType
        lr = ReadFile(hFile, vOutPtr, cbToRead, lBytesRead, 0&)
                
    ElseIf vbString = vType Then
        lr = ReadFile(hFile, StrPtr(vOut), Len(vOut), lBytesRead, 0&)
        If Err.LastDllError <> 0 Or lr = 0 Then Err.Raise 52
                
        vOut = StrConv(vOut, vbUnicode)
        If Len(vOut) <> 0 Then vOut = Left$(vOut, Len(vOut) \ 2)
    Else
        'do a bit of magik :)
        memcpy ptr, ByVal VarPtr(vOut) + 8, 4& 'VT_BYREF
        Select Case vType
        Case vbByte
            lr = ReadFile(hFile, ptr, 1&, lBytesRead, 0&)
        Case vbInteger
            lr = ReadFile(hFile, ptr, 2&, lBytesRead, 0&)
        Case vbLong
            lr = ReadFile(hFile, ptr, 4&, lBytesRead, 0&)
        Case vbCurrency
            lr = ReadFile(hFile, ptr, 8&, lBytesRead, 0&)
        Case Else
            UnknType = True
            Err.Clear: ErrorMsg Err, "modFile.GetW. type #" & VarType(vOut) & " of buffer is not supported.": Err.Raise 52
        End Select
    End If
    GetW = (0 <> lr)
    If 0 = lr And Not UnknType Then Err.Clear: ErrorMsg Err, "Cannot read file!": Err.Raise 52
    
    AppendErrorLogCustom "GetW - End", "BytesRead: " & lBytesRead
'    Exit Function
'ErrorHandler:
'    AppendErrorLogFormat Now, err, "modFile.GetW"
'    Resume Next
End Function

Public Function PutW(hFile As Long, pos As Long, vInPtr As Long, cbToWrite As Long, Optional doAppend As Boolean) As Boolean
    On Error GoTo ErrorHandler
    'don't uncomment it -> recurse on DebugToFile !!!
    'AppendErrorLogCustom "PutW - Begin", "Handle: " & hFile, "pos: " & pos, "Bytes: " & cbToWrite
    
    Dim lBytesWrote  As Long
    
    pos = pos - 1   ' VB's Get & SetFilePointer difference correction
    
    If doAppend Then
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, 0&, ByVal 0&, FILE_END) Then Exit Function
    Else
        If INVALID_SET_FILE_POINTER = SetFilePointer(hFile, pos, ByVal 0&, FILE_BEGIN) Then Exit Function
    End If
    
    If NO_ERROR = Err.LastDllError Then
    
        If WriteFile(hFile, vInPtr, cbToWrite, lBytesWrote, 0&) Then PutW = True
        
    End If
    
    'AppendErrorLogCustom "PutW - End"
    Exit Function
ErrorHandler:
    'don't change/append this identifier !!! -> can cause recurse on DebugToFile !!!
    ErrorMsg Err, "modFile.PutW"
End Function

Public Function LOFW(hFile As Long) As Currency
    On Error GoTo ErrorHandler
    Dim lr          As Long
    Dim FileSize    As Currency
    
    AppendErrorLogCustom "LOFW - Begin", "Handle: " & hFile
    
    If hFile Then
        lr = GetFileSizeEx(hFile, FileSize)
        If lr Then
            If FileSize < 10000000000@ Then
                LOFW = FileSize * 10000&
            Else
                Err.Clear
                ErrorMsg Now, "File is too big. Size: " & FileSize
            End If
        End If
    End If
    
    AppendErrorLogCustom "LOFW - End", "Size: " & LOFW
ErrorHandler:
End Function

Public Function CloseW(hFile As Long) As Long
    AppendErrorLogCustom "CloseW", "Handle: " & hFile
    CloseW = CloseHandle(hFile)
End Function

Public Function LineInputW(hFile As Long, sLine As String) As Boolean
    Dim ch$, lBytesRead&, lr&
    sLine = vbNullString
    Do
        ch = vbNullChar

        lr = ReadFile(hFile, StrPtr(ch), 1, lBytesRead, 0&)
        
        If lr = 0 Or lBytesRead = 0 Or AscW(ch) = 10 Then
            If Right$(sLine, 1) = vbCr Then sLine = Left(sLine, Len(sLine) - 1)
            Exit Do
        Else
            LineInputW = True
            sLine = sLine & ch
        End If
    Loop
End Function

Public Function ToggleWow64FSRedirection(bEnable As Boolean, Optional PathNecessity As String, Optional OldStatus As Boolean) As Boolean
    'Static lWow64Old        As Long    'Warning: do not use initialized variables for this API !
                                        'Static variables is not allowed !
                                        'lWow64Old is now declared globally
    'True - enable redirector
    'False - disable redirector

    'OldStatus: current state of redirection
    'True - redirector was enabled
    'False - redirector was disabled

    'Return value is:
    'true if success

    Static IsNotRedirected  As Boolean
    Dim lr                  As Long
    
    OldStatus = Not IsNotRedirected
    
    If Not bIsWin64 Then Exit Function
    
    If Len(PathNecessity) <> 0 Then
        If StrComp(Left$(Replace(Replace(PathNecessity, "/", "\"), "\\", "\"), Len(sWinSysDir)), sWinSysDir, vbTextCompare) <> 0 Then Exit Function
    End If
    
    If bEnable Then
        If IsNotRedirected Then
            lr = Wow64RevertWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = False
        End If
    Else
        If Not IsNotRedirected Then
            lr = Wow64DisableWow64FsRedirection(lWow64Old)
            ToggleWow64FSRedirection = (lr <> 0)
            IsNotRedirected = True
        End If
    End If
    
    If ToggleWow64FSRedirection Then
        If OldStatus <> bEnable Then
            AppendErrorLogCustom "ToggleWow64FSRedirection - End", "Path: " & PathNecessity, _
                "Old State: " & OldStatus, "New State: " & bEnable
        End If
    End If
    
End Function

Public Function GetParentDir(sPath As String) As String
    Dim pos As Long
    pos = InStrRev(sPath, "\")
    If pos <> 0 Then
        GetParentDir = Left$(sPath, pos - 1)
    End If
End Function

Public Function GetExtensionName(Path As String) As String  '������ .ext
    Dim pos As Long
    pos = InStrRev(Path, ".")
    If pos <> 0 Then GetExtensionName = Mid$(Path, pos)
End Function

' �������� �� ���� �������� PE EXE
Public Function isPE_EXE(Optional FileName As String, Optional FileHandle As Long) As Boolean
    On Error GoTo ErrorHandler
    AppendErrorLogCustom "isPE_EXE - Begin", "File: " & FileName

'    #If UseHashTable Then
'        Static PE_EXE_Cache As clsTrickHashTable
'    #Else
'        Static PE_EXE_Cache As Object
'    #End If
'
'    If 0 = ObjPtr(PE_EXE_Cache) Then
'        #If UseHashTable Then
'            Set PE_EXE_Cache = New clsTrickHashTable
'        #Else
'            Set PE_EXE_Cache = CreateObject("Scripting.Dictionary")
'        #End If
'        PE_EXE_Cache.CompareMode = vbTextCompare
'    Else
'        If Len(FileName) <> 0& Then
'            If PE_EXE_Cache.Exists(FileName) Then
'                isPE_EXE = PE_EXE_Cache(FileName)
'                Exit Function
'            End If
'        End If
'    End If

    'Static PE_EXE_Cache    As New Collection ' value = true, ���� ���� �������� �������� PE EXE

    'If Len(FileName) <> 0 Then
    '    If isCollectionKeyExists(FileName, PE_EXE_Cache) Then
    '        isPE_EXE = PE_EXE_Cache(FileName)
    '        Exit Function
    '    End If
    'End If

    Dim hFile          As Long
    Dim PE_offset      As Long
    Dim MZ(1)          As Byte
    Dim pe(3)          As Byte
    Dim FSize          As Currency
  
    If FileHandle = 0& Then
        OpenW FileName, FOR_READ, hFile
    Else
        hFile = FileHandle
    End If
    If hFile <> INVALID_HANDLE_VALUE Then
        FSize = LOFW(hFile)
        If FSize >= &H3C& + 4& Then
            GetW hFile, 1&, , VarPtr(MZ(0)), ((UBound(MZ) + 1&) * CLng(LenB(MZ(0))))
            If (MZ(0) = 77& And MZ(1) = 90&) Or (MZ(1) = 77& And MZ(0) = 90&) Then  'MZ or ZM
                GetW hFile, &H3C& + 1&, PE_offset
                If PE_offset And FSize >= PE_offset + 4 Then
                    GetW hFile, PE_offset + 1&, , VarPtr(pe(0)), ((UBound(pe) + 1&) * CLng(LenB(pe(0))))
                    If pe(0) = 80& And pe(1) = 69& And pe(2) = 0& And pe(3) = 0& Then isPE_EXE = True   'PE NUL NUL
                End If
            End If
        End If
        If FileHandle = 0& Then CloseW hFile: hFile = 0&
    End If
    
    'If Len(FileName) <> 0& Then PE_EXE_Cache.Add FileName, isPE_EXE
    
    AppendErrorLogCustom "isPE_EXE - End"
    Exit Function
    
ErrorHandler:
    ErrorMsg Err, "Parser.isPE_EXE", "File:", FileName
    'On Error Resume Next
    'If Len(FileName) <> 0& Then PE_EXE_Cache.Add FileName, isPE_EXE
    If FileHandle = 0& Then
        If hFile <> 0 Then CloseW hFile: hFile = 0&
    End If
End Function

'main function to list folders

' ���������� ������ �����.
' ���� ������ �� ������� - ������������ �������������������� ������.
Public Function ListSubfolders(Path As String, Optional Recursively As Boolean = False) As String()
    On Error GoTo ErrorHandler

    AppendErrorLogCustom "ListSubfolders - Begin", "Path:", Path, "Recur:", Recursively

    Dim bRedirStateChanged As Boolean, bOldState As Boolean
    
    '������, ��� ������������ ListSubfolders_Ex, ����� ���������������� ���������� �������.
    ReDim arrPathFolders(100) As String
    '��� ������ ������ ListSubfolders_Ex ������� �������� ���������� ������� ������
    Total_Folders = 0&
    
    If bIsWin64 Then
        If StrBeginWith(Path, sWinDir) Then
            bRedirStateChanged = ToggleWow64FSRedirection(False, , bOldState)
        End If
    End If
    
    '����� �����
    Call ListSubfolders_Ex(Path, Recursively)
    If Total_Folders > 0 Then
        Total_Folders = Total_Folders - 1
        ReDim Preserve arrPathFolders(Total_Folders)      '0 to Max -1
        ListSubfolders = arrPathFolders
    End If
    
    If bRedirStateChanged Then Call ToggleWow64FSRedirection(bOldState)
    
    AppendErrorLogCustom "ListSubfolders - End"
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.ListSubfolders", "Path:", Path, "Recur:", Recursively
    If bRedirStateChanged Then Call ToggleWow64FSRedirection(bOldState)
    If inIDE Then Stop: Resume Next
End Function


Private Sub ListSubfolders_Ex(Path As String, Optional Recursively As Boolean = False)
    On Error GoTo ErrorHandler
    'On Error Resume Next
    Dim SubPathName     As String
    Dim PathName        As String
    Dim hFind           As Long
    Dim l               As Long
    Dim lpSTR           As Long
    Dim fd              As WIN32_FIND_DATA
    
    'Local module variables:
    '
    ' Total_Folders as long
    ' arrPathFolders() as string
    
    Do
        If hFind <> 0& Then
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: Exit Do
        Else
            hFind = FindFirstFile(StrPtr(Path & ch_SlashAsterisk), fd)  '"\*"
            If hFind = INVALID_HANDLE_VALUE Then Exit Do
        End If
        
        l = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT ' ���� ���������
        Do While l <> 0&
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0: Exit Do
            l = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT
        Loop
    
        If hFind <> 0& Then
            lpSTR = VarPtr(fd.dwReserved1) + 4&
            PathName = Space(lstrlen(lpSTR))
            lstrcpy StrPtr(PathName), lpSTR
        
            If fd.dwFileAttributes And vbDirectory Then
                If PathName <> ch_Dot Then  '"."
                    If PathName <> ch_DotDot Then '".."
                        SubPathName = Path & "\" & PathName
                        If UBound(arrPathFolders) < Total_Folders Then ReDim Preserve arrPathFolders(UBound(arrPathFolders) + 100&) As String
                        arrPathFolders(Total_Folders) = SubPathName
                        Total_Folders = Total_Folders + 1&
                        If Recursively Then
                            Call ListSubfolders_Ex(SubPathName, Recursively)
                        End If
                    End If
                End If
            End If
        End If
        
    Loop While hFind
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modFile.ListSubfolders", "Folder:", Path
    Resume Next
End Sub

'main function to list files

Public Function ListFiles(Path As String, Optional Extension As String = "", Optional Recursively As Boolean = False) As String()
    On Error GoTo ErrorHandler

    AppendErrorLogCustom "ListFiles - Begin", "Path: " & Path, "Ext-s: " & Extension, "Recur: " & Recursively

    Dim bRedirStateChanged As Boolean, bOldState As Boolean
    '������, ��� ������������ ListFiles_Ex, ����� ���������������� ���������� �������.
    ReDim arrPathFiles(100) As String
    '��� ������ ������ ListFiles_Ex ������� �������� ���������� ������� ������
    Total_Files = 0&
    
    If bIsWin64 Then
        If StrBeginWith(Path, sWinDir) Then
            bRedirStateChanged = ToggleWow64FSRedirection(False, , bOldState)
        End If
    End If
    
    '����� �����
    Call ListFiles_Ex(Path, Extension, Recursively)
    If Total_Files > 0 Then
        Total_Files = Total_Files - 1
        ReDim Preserve arrPathFiles(Total_Files)      '0 to Max -1
        ListFiles = arrPathFiles
    End If
    
    If bRedirStateChanged Then Call ToggleWow64FSRedirection(bOldState)
    
    AppendErrorLogCustom "ListFiles - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "modFile.ListFiles", "Path:", Path, "Ext-s:", Extension, "Recur:", Recursively
    If bRedirStateChanged Then Call ToggleWow64FSRedirection(bOldState)
    If inIDE Then Stop: Resume Next
End Function


Private Sub ListFiles_Ex(Path As String, Optional Extension As String = "", Optional Recursively As Boolean = False)
    'Example of Extension:
    '".txt" - txt files
    'empty line - all files (by default)

    On Error GoTo ErrorHandler
    'On Error Resume Next
    Dim SubPathName     As String
    Dim PathName        As String
    Dim hFind           As Long
    Dim l               As Long
    Dim lpSTR           As Long
    Dim fd              As WIN32_FIND_DATA
    
    'Local module variables:
    '
    ' Total_Files as long
    ' arrPathFiles() as string
    
    Do
        If hFind <> 0& Then
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: Exit Do
        Else
            hFind = FindFirstFile(StrPtr(Path & ch_SlashAsterisk), fd)  '"\*"
            If hFind = INVALID_HANDLE_VALUE Then Exit Do
        End If
        
        l = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT ' ���� ���������
        Do While l <> 0&
            If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0: Exit Do
            l = fd.dwFileAttributes And FILE_ATTRIBUTE_REPARSE_POINT
        Loop
    
        If hFind <> 0& Then
            lpSTR = VarPtr(fd.dwReserved1) + 4&
            PathName = Space(lstrlen(lpSTR))
            lstrcpy StrPtr(PathName), lpSTR
        
            If fd.dwFileAttributes And vbDirectory Then
                If PathName <> ch_Dot Then  '"."
                    If PathName <> ch_DotDot Then '".."
                        SubPathName = Path & "\" & PathName
                        If Recursively Then
                            Call ListFiles_Ex(SubPathName, Extension, Recursively)
                        End If
                    End If
                End If
            Else
                If inArray(GetExtensionName(PathName), SplitSafe(Extension, ";"), , , 1) Or Len(Extension) = 0 Then
                    SubPathName = Path & "\" & PathName
                    If UBound(arrPathFiles) < Total_Files Then ReDim Preserve arrPathFiles(UBound(arrPathFiles) + 100&) As String
                    arrPathFiles(Total_Files) = SubPathName
                    Total_Files = Total_Files + 1&
                End If
            End If
        End If
    Loop While hFind
    
    Exit Sub
ErrorHandler:
    ErrorMsg Err, "modFile.ListFiles_Ex", "File:", Path
    Resume Next
End Sub

Public Function GetLocalDisks$()
    Dim lDrives&, i&, sDrive$, sLocalDrives$
    lDrives = GetLogicalDrives()
    For i = 0 To 26
        If (lDrives And 2 ^ i) Then
            sDrive = Chr$(Asc("A") + i) & ":\"
            Select Case GetDriveType(StrPtr(sDrive))
                Case DRIVE_FIXED, DRIVE_RAMDISK: sLocalDrives = sLocalDrives & Chr$(Asc("A") + i) & " "
            End Select
        End If
    Next i
    GetLocalDisks = Trim$(sLocalDrives)
End Function

Public Function GetLongFilename$(sFilename$)
    Dim sLongFilename$
    If InStr(sFilename, "~") = 0 Then
        GetLongFilename = sFilename
        Exit Function
    End If
    sLongFilename = String(512, 0)
    GetLongPathNameW StrPtr(sFilename), StrPtr(sLongFilename), Len(sLongFilename)
    GetLongFilename = TrimNull(sLongFilename)
End Function

Public Function GetFilePropVersion(sFilename As String) As String
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetFilePropVersion - Begin", "File: " & sFilename
    
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$, uVFFI As VS_FIXEDFILEINFO, sVersion$, Redirect As Boolean, bOldStatus As Boolean
    
    If Not FileExists(sFilename) Then Exit Function
    
    Redirect = ToggleWow64FSRedirection(False, sFilename, bOldStatus)
    
    lDataLen = GetFileVersionInfoSize(StrPtr(sFilename), ByVal 0&)
    If lDataLen = 0 Then GoTo Finalize
    
    ReDim uBuf(0 To lDataLen - 1)
    If 0 <> GetFileVersionInfo(StrPtr(sFilename), 0&, lDataLen, uBuf(0)) Then
    
        If 0 <> VerQueryValue(uBuf(0), StrPtr("\"), hData, lDataLen) Then
        
            If hData <> 0 Then
        
                CopyMemory uVFFI, ByVal hData, Len(uVFFI)
    
                With uVFFI
                    sVersion = .dwFileVersionMSh & "." & _
                        .dwFileVersionMSl & "." & _
                        .dwFileVersionLSh & "." & _
                        .dwFileVersionLSl
                End With
            End If
        End If
    End If
    GetFilePropVersion = sVersion
    
Finalize:
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    
    AppendErrorLogCustom "GetFilePropVersion - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilePropVersion", sFilename
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Function GetVersionFromVBP(sFilename As String) As String
    On Error GoTo ErrorHandler:
    AppendErrorLogCustom "GetVersionFromVBP - Begin", "File: " & sFilename
    
    Dim hFile As Long, sLine As String, Arr() As String
    Dim MajorVer As Long, MinorVer As Long, BuildVer As Long, RevisionVer As Long
    
    OpenW sFilename, FOR_READ, hFile
    
    If hFile > 0 Then
        Do While LineInputW(hFile, sLine)
            Arr = SplitSafe(sLine, "=")
            If UBound(Arr) = 1 Then
                If StrComp(Arr(0), "MajorVer", 1) = 0 Then
                    MajorVer = Val(Arr(1))
                ElseIf StrComp(Arr(0), "MinorVer", 1) = 0 Then
                    MinorVer = Val(Arr(1))
                ElseIf StrComp(Arr(0), "BuildVer", 1) = 0 Then
                    BuildVer = Val(Arr(1))
                ElseIf StrComp(Arr(0), "RevisionVer", 1) = 0 Then
                    RevisionVer = Val(Arr(1))
                End If
            End If
        Loop
        CloseW hFile
    End If
    
    GetVersionFromVBP = MajorVer & "." & MinorVer & "." & BuildVer & "." & RevisionVer
    
    AppendErrorLogCustom "GetVersionFromVBP - End", "File: " & sFilename
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetVersionFromVBP", sFilename
    If inIDE Then Stop: Resume Next
End Function

Public Function GetFilePropCompany(sFilename As String) As String
    On Error GoTo ErrorHandler:
    Dim hData&, lDataLen&, uBuf() As Byte, uCodePage(0 To 3) As Byte
    Dim sCodePage$, sCompanyName$, Stady&, Redirect As Boolean, bOldStatus As Boolean
    
    If Not FileExists(sFilename) Then Exit Function
    
    Redirect = ToggleWow64FSRedirection(False, sFilename, bOldStatus)
    
    Stady = 1
    lDataLen = GetFileVersionInfoSize(StrPtr(sFilename), ByVal 0&)
    If lDataLen = 0 Then GoTo Finalize
    
    Stady = 2
    ReDim uBuf(0 To lDataLen - 1)
    
    Stady = 3
    If 0 <> GetFileVersionInfo(StrPtr(sFilename), 0&, lDataLen, uBuf(0)) Then
        
        Stady = 4
        VerQueryValue uBuf(0), StrPtr("\VarFileInfo\Translation"), hData, lDataLen
        If lDataLen = 0 Then GoTo Finalize
        
        Stady = 5
        CopyMemory uCodePage(0), ByVal hData, 4
        
        Stady = 6
        sCodePage = Right$("0" & Hex(uCodePage(1)), 2) & _
                Right$("0" & Hex(uCodePage(0)), 2) & _
                Right$("0" & Hex(uCodePage(3)), 2) & _
                Right$("0" & Hex(uCodePage(2)), 2)
        
        'get CompanyName string
        Stady = 7
        If VerQueryValue(uBuf(0), StrPtr("\StringFileInfo\" & sCodePage & "\CompanyName"), hData, lDataLen) = 0 Then GoTo Finalize
    
        If lDataLen > 0 And hData <> 0 Then
            Stady = 8
            sCompanyName = String$(lDataLen, 0)
            
            Stady = 9
            lstrcpy ByVal StrPtr(sCompanyName), ByVal hData
        End If
        
        Stady = 10
        GetFilePropCompany = RTrimNull(sCompanyName)
    End If
    
Finalize:
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "GetFilePropCompany", sFilename, "DataLen: ", lDataLen, "hData: ", hData, "sCodePage: ", sCodePage, _
        "Buf: ", uCodePage(0), uCodePage(1), uCodePage(2), uCodePage(3), "Stady: ", Stady
    If Redirect Then Call ToggleWow64FSRedirection(bOldStatus)
    If inIDE Then Stop: Resume Next
End Function

Public Function DirW( _
    Optional ByVal PathMaskOrFolderWithSlash As String, _
    Optional AllowedAttributes As VbFileAttributeExtended = vbNormal, _
    Optional FoldersOnly As Boolean) As String
    
    On Error GoTo ErrorHandler
    
    'WARNING note:
    'Original VB DirW$ contains bug: ReadOnly attribute incorrectly handled, so it always is in results
    'This sub properly handles 'RO' and also contains one extra flag: FILE_ATTRIBUTE_REPARSE_POINT (vbReparse)
    'Doesn't return "." and ".." folders.
    'Unicode aware
    
    Const MeaningfulBits As Long = &H417&   'D + H + R + S + Reparse
                                            '(to revert to default VB Dir behaviour, replace it by &H16 value)
    
    Dim fd      As WIN32_FIND_DATA
    Dim lpSTR   As Long
    Dim lRet    As Long
    Dim mask    As Long
    
    Static hFind        As Long
    Static lFlags       As VbFileAttributeExtended
    Static bFoldersOnly As Boolean
    
    If hFind <> 0& And Len(PathMaskOrFolderWithSlash) = 0& Then
        If FindNextFile(hFind, fd) = 0& Then FindClose hFind: hFind = 0&: Exit Function
    Else
        If hFind Then FindClose hFind: hFind = 0&
        PathMaskOrFolderWithSlash = Trim(PathMaskOrFolderWithSlash)
        lFlags = AllowedAttributes 'cache
        bFoldersOnly = FoldersOnly 'cache
        
        Select Case Right$(PathMaskOrFolderWithSlash, 1&)
        Case "", ":", "/"
            PathMaskOrFolderWithSlash = PathMaskOrFolderWithSlash & "*.*"
        End Select
        
        hFind = FindFirstFile(StrPtr(PathMaskOrFolderWithSlash), fd)
        
        If hFind = INVALID_HANDLE_VALUE Then
            If (Err.LastDllError) > 12& Then hFind = 0&: Err.Raise 52&
            Exit Function
        End If
    End If
    
    Do
        If fd.dwFileAttributes = FILE_ATTRIBUTE_NORMAL Then
            mask = 0& 'found
        Else
            mask = fd.dwFileAttributes And (Not lFlags) And MeaningfulBits
        End If
        If bFoldersOnly Then
            If Not CBool(fd.dwFileAttributes And vbDirectory) Then
                mask = 1 'continue enum
            End If
        End If
    
        If mask = 0 Then
            lpSTR = VarPtr(fd.lpszFileName(0))
            DirW = String$(lstrlen(lpSTR), 0&)
            lstrcpy StrPtr(DirW), lpSTR
            If fd.dwFileAttributes And vbDirectory Then
                If DirW <> "." And DirW <> ".." Then Exit Do 'exclude self and relative paths aliases
            Else
                Exit Do
            End If
        End If
    
        If FindNextFile(hFind, fd) = 0 Then FindClose hFind: hFind = 0: Exit Function
    Loop
    
    Exit Function
ErrorHandler:
    Debug.Print Err; Err.Description; "DirW"
End Function

Public Function GetEmptyName(ByVal sFullPath As String) As String

    Dim sExt As String
    Dim sName As String
    Dim sPath As String
    Dim i As Long

    If Not FileExists(sFullPath) Then
        GetEmptyName = sFullPath
    Else
        sExt = GetExtensionName(sFullPath)
        sPath = GetPathName(sFullPath)
        sName = GetFileName(sFullPath)
        Do
            i = i + 1
            sFullPath = BuildPath(sPath, sName & "(" & i & ")" & sExt)
        Loop While FileExists(sFullPath)
        
        GetEmptyName = sFullPath
    End If
End Function

Public Function GetFileDate(Optional file As String, Optional Date_Type As ENUM_File_Date_Type, Optional hFile As Long) As Date
    On Error GoTo ErrorHandler
    
    Dim SFCacheMode As Boolean
    Dim rval        As Long
    Dim ctime       As FILETIME
    Dim atime       As FILETIME
    Dim wtime       As FILETIME
    Dim ftime       As SYSTEMTIME
    Dim bOldRedir   As Boolean
    Dim bExternalHandle As Boolean
    
    AppendErrorLogCustom "Parser.GetFileDate - Begin: " & file
    
    If hFile <= 0 Then
        ToggleWow64FSRedirection False, file, bOldRedir
    
        hFile = CreateFile(StrPtr(file), ByVal 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, 0&, 0&)
    
        ToggleWow64FSRedirection bOldRedir
    Else
        bExternalHandle = True
    End If
    
    If hFile <> INVALID_HANDLE_VALUE Then
        rval = GetFileTime(hFile, ctime, atime, wtime)
        Select Case Date_Type
        Case Date_Modified
            rval = FileTimeToLocalFileTime(wtime, wtime)
            rval = FileTimeToSystemTime(wtime, ftime)
        Case Date_Created
            rval = FileTimeToLocalFileTime(ctime, ctime)
            rval = FileTimeToSystemTime(ctime, ftime)
        Case Date_Accessed
            rval = FileTimeToLocalFileTime(atime, atime)
            rval = FileTimeToSystemTime(atime, ftime)
        End Select
        GetFileDate = DateSerial(ftime.wYear, ftime.wMonth, ftime.wDay) + TimeSerial(ftime.wHour, ftime.wMinute, ftime.wSecond)
        If Not bExternalHandle Then
            CloseHandle hFile
        End If
    Else
        GetFileDate = CDate(0)
    End If
    
    AppendErrorLogCustom "Parser.GetFileDate - End"
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetFileDate", "File: " & file
    If inIDE Then Stop: Resume Next
End Function

Public Function IsFileOneMonthModified(sFile As String) As Boolean

    Dim bOldRedir   As Boolean
    Dim hFile       As Long

    ToggleWow64FSRedirection False, sFile, bOldRedir
    
    hFile = CreateFile(StrPtr(sFile), ByVal 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, ByVal 0&, OPEN_EXISTING, 0&, 0&)
    
    If DateDiff("d", GetFileDate(, Date_Created, hFile), Now) < 31 Then
        IsFileOneMonthModified = True
    ElseIf DateDiff("d", GetFileDate(, Date_Modified, hFile), Now) < 31 Then
        IsFileOneMonthModified = True
    End If
    
    ToggleWow64FSRedirection bOldRedir
End Function

'check file on Portable Executable
Public Function isPE(sFile As String) As Boolean
    On Error GoTo ErrorHandler:
    
    Dim bOldRedir As Boolean
    Dim hFile As Long
    Dim hMapping As Long
    Dim pBuf As Long
    
    ToggleWow64FSRedirection False, sFile, bOldRedir
    
    hFile = CreateFile(StrPtr(sFile), GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
    
    ToggleWow64FSRedirection bOldRedir
    
    If hFile <> INVALID_HANDLE_VALUE Then
    
        hMapping = CreateFileMapping(hFile, 0&, PAGE_READONLY Or SEC_IMAGE, 0&, 0&, 0&)
        
        CloseHandle hFile
        
        If hMapping <> 0 Then
            
            pBuf = MapViewOfFile(hMapping, FILE_MAP_READ, 0&, 0&, 0&)
            
            If pBuf <> 0 Then
            
                isPE = True
                UnmapViewOfFile pBuf
            End If
            
            CloseHandle hMapping
        End If
    End If
    
    Exit Function
ErrorHandler:
    ErrorMsg Err, "isPE"
    If inIDE Then Stop: Resume Next
End Function

Public Function EnvironW(ByVal SrcEnv As String, Optional UseRedir As Boolean) As String
    Dim lr As Long
    Dim buf As String
    Static LastFile As String
    Static LastResult As String
    Static isInit As Boolean
    Static PF_64 As String
    Static PF_64_Common As String
    
    If Not isInit Then
        isInit = True
        If bIsWin64 Then
            If MajorMinor >= 6.1 Then     'Win 7 and later
                PF_64 = Environ("ProgramW6432")
            Else
                PF_64 = SysDisk & "\Program Files"
            End If
        Else
            PF_64 = Environ("ProgramFiles")
        End If
        PF_64_Common = PF_64 & "\Common Files"
    End If
    
    AppendErrorLogCustom "EnvironW - Begin", "SrcEnv: " & SrcEnv
    
    If Len(SrcEnv) = 0 Then Exit Function
    If InStr(SrcEnv, "%") = 0 Then
        EnvironW = SrcEnv
    Else
        If LastFile = SrcEnv Then
            EnvironW = LastResult
            Exit Function
        End If
        'redirector correction
        If bIsWin64 Then
            If Not UseRedir Then
                If InStr(1, SrcEnv, "%PROGRAMFILES%", 1) <> 0 Then
                    SrcEnv = Replace$(SrcEnv, "%PROGRAMFILES%", PF_64, 1, 1, 1)
                End If
                If InStr(1, SrcEnv, "%COMMONPROGRAMFILES%", 1) <> 0 Then
                    SrcEnv = Replace$(SrcEnv, "%COMMONPROGRAMFILES%", PF_64_Common, 1, 1, 1)
                End If
            End If
        End If
        buf = String$(MAX_PATH, vbNullChar)
        lr = ExpandEnvironmentStrings(StrPtr(SrcEnv), StrPtr(buf), MAX_PATH + 1)
        
        If lr Then
            EnvironW = Left$(buf, lr - 1)
        Else
            EnvironW = SrcEnv
        End If
        
        If InStr(EnvironW, "%") <> 0 Then
            If MajorMinor <= 6 Then
                If InStr(1, EnvironW, "%ProgramW6432%", 1) <> 0 Then
                    EnvironW = Replace$(EnvironW, "%ProgramW6432%", SysDisk & "\Program Files", 1, -1, 1)
                End If
            End If
        End If
    End If
    LastFile = SrcEnv
    LastResult = EnvironW
    
    AppendErrorLogCustom "EnvironW - End"
End Function

Private Function inIDE() As Boolean
    inIDE = (App.LogMode = 0)
End Function

Public Function bIsWin64() As Boolean
    bIsWin64 = IsWow64
End Function

Function IsWow64() As Boolean   ' ����������� ��
    Dim hModule As Long, procAddr As Long, lIsWin64 As Long
    Static isInit As Boolean, Result As Boolean
    
    If isInit Then
        IsWow64 = Result
    Else
        isInit = True
        hModule = LoadLibrary(StrPtr("kernel32.dll"))
        If hModule Then
            procAddr = GetProcAddress(hModule, "IsWow64Process")
            If procAddr <> 0 Then
                IsWow64Process GetCurrentProcess(), lIsWin64
                Result = CBool(lIsWin64)
                IsWow64 = Result
            End If
            FreeLibrary hModule
        End If
    End If
End Function

Function sWinSysDir() As String
    sWinSysDir = Environ("SystemRoot") & "\" & "System32"
End Function

Function sWinDir() As String
    sWinDir = Environ("SystemRoot")
End Function

Public Function StrBeginWith(Text As String, BeginPart As String) As Boolean
    StrBeginWith = (StrComp(Left$(Text, Len(BeginPart)), BeginPart, 1) = 0)
End Function

Public Function StrEndWith(Text As String, EndPart As String) As Boolean
    StrEndWith = (StrComp(Right$(Text, Len(EndPart)), EndPart, 1) = 0)
End Function

Public Function SplitSafe(sComplexString As String, Optional Delimiter As String = " ") As String()
    If 0 = Len(sComplexString) Then
        ReDim Arr(0) As String
        SplitSafe = Arr
    Else
        SplitSafe = Split(sComplexString, Delimiter)
    End If
End Function

' ���������� true, ���� ������� �������� ������� � ����� �� ��������� ������� (lB, uB ������������ ��������������� �������� ��������)
Public Function inArray( _
    Stri As String, _
    MyArray() As String, _
    Optional lB As Long = -2147483647, _
    Optional uB As Long = 2147483647, _
    Optional CompareMethod As VbCompareMethod) As Boolean
    
    On Error GoTo ErrorHandler:
    If lB = -2147483647 Then lB = LBound(MyArray)   'some trick
    If uB = 2147483647 Then uB = UBound(MyArray)    'Thanks to ��������� :)
    Dim i As Long
    For i = lB To uB
        If StrComp(Stri, MyArray(i), CompareMethod) = 0 Then inArray = True: Exit For
    Next
    Exit Function
ErrorHandler:
    ErrorMsg Err, "inArray"
    If inIDE Then Stop: Resume Next
End Function

Public Function BuildPath$(sPath$, sFile$)
    BuildPath = sPath & IIf(Right$(sPath, 1) = "\", vbNullString, "\") & sFile
End Function

Public Function TrimNull(s$) As String
    TrimNull = Left$(s, lstrlen(StrPtr(s)))
End Function

Public Function RTrimNull(ByVal s$) As String
    Do While AscW(Right$(s, 1)) <> 0 And Len(s) > 0
        s = Left$(s, Len(s) - 1)
    Loop
    RTrimNull = s
End Function

Public Function GetPathName(Path As String) As String   ' �������� ������������ �������
    Dim pos As Long
    pos = InStrRev(Path, "\")
    If pos <> 0 Then GetPathName = Left$(Path, pos - 1)
End Function

' �������� ������ ��� ����� (��� ���������� �����)
Public Function GetFileName(Path As String) As String
    On Error GoTo ErrorHandler
    Dim posDot      As Long
    Dim posSl       As Long
    
    posSl = InStrRev(Path, "\")
    If posSl <> 0 Then
        posDot = InStrRev(Path, ".")
        If posDot < posSl Then posDot = 0
    Else
        posDot = InStrRev(Path, ".")
    End If
    If posDot = 0 Then posDot = Len(Path) + 1
    
    GetFileName = Mid$(Path, posSl + 1, posDot - posSl - 1)
    Exit Function
ErrorHandler:
    ErrorMsg Err, "Parser.GetFileName", "Path: ", Path
End Function

Public Function GetFileNameAndExt(Path As String) As String ' ������ ������ ��� ����� ������ � �����������
    Dim pos As Long
    pos = InStrRev(Path, "\")
    If pos <> 0 Then
        GetFileNameAndExt = Mid$(Path, pos + 1)
    Else
        GetFileNameAndExt = Path
    End If
End Function

Public Function MajorMinor() As Single
    Dim osi As RTL_OSVERSIONINFOEXW
    Dim dec             As Single
    
    osi.dwOSVersionInfoSize = Len(osi)
    RtlGetVersion osi
    
    dec = osi.dwMinorVersion
    If dec <> 0 Then Do: dec = dec / 10: Loop Until dec < 1
    MajorMinor = osi.dwMajorVersion + dec
End Function

Public Function SysDisk() As String
    SysDisk = Environ("SystemDrive")
End Function

Public Sub OpenFolderAndSelectItem(sFile As String)
    Dim hRet As Long
    Dim pIDL As Long
    
    If MajorMinor >= 5.1 Then
            
        pIDL = ILCreateFromPath(StrPtr(sFile))
    
        If pIDL <> 0 Then
            hRet = SHOpenFolderAndSelectItems(pIDL, 0, 0, 0)
    
            ILFree pIDL
        End If
    End If
            
    If pIDL = 0 Or hRet <> S_OK Then
        '�������������
        Shell "explorer.exe /select," & """" & sFile & """", vbNormalFocus
    End If
End Sub

Public Function DOS2Win(Str As String) As String
    If Len(Str) > 0 Then
        DOS2Win = String(Len(Str), 0&)
        OemToChar Str, DOS2Win
    End If
End Function

Public Function Win2Dos(Str As String) As String
    If Len(Str) > 0 Then
        Win2Dos = String(Len(Str), 0&)
        CharToOem Str, Win2Dos
    End If
End Function

Public Function URLEncode(ByVal Text As String) As String
    On Error GoTo ErrorHandler:

    Dim i As Long
    Dim acode As Long
    
    URLEncode = Text
    
    For i = Len(URLEncode) To 1 Step -1
        acode = Asc(Mid$(URLEncode, i, 1))
        Select Case acode
            Case 48 To 57, 65 To 90, 97 To 122
                ' don't touch alphanumeric chars
            Case 32
                ' replace space with "+"
                Mid$(URLEncode, i, 1) = "+"
            Case Else
                ' replace punctuation chars with "%hex"
                URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & Mid$ _
                    (URLEncode, i + 1)
        End Select
    Next
    
    Exit Function
ErrorHandler:
    Debug.Print "URLEncode" & ". Src:" & Text
End Function

'true if success
Public Function FileCopyW(FileSource As String, FileDestination As String, Optional bOverwrite As Boolean = True) As Boolean
    On Error GoTo ErrorHandler:
    Dim sFolder As String
    
    If Not FileExists(FileSource) Then Exit Function
    
    sFolder = GetParentDir(FileDestination)
    
    If Not FolderExists(sFolder) Then
        If Not MkDirW(sFolder) Then Exit Function
    End If
    
    FileCopyW = CopyFile(StrPtr(FileSource), StrPtr(FileDestination), Not bOverwrite)
    If Not FileCopyW Then
        FileCopyW = CopyFile(StrPtr(FileSource), StrPtr(FileDestination), Not bOverwrite)
        If Not FileCopyW Then
            If DeleteFileWEx(StrPtr(FileDestination)) Then
                FileCopyW = CopyFile(StrPtr(FileSource), StrPtr(FileDestination), Not bOverwrite)
            End If
        End If
    End If
       
    Exit Function
ErrorHandler:
    Debug.Print "FileCopyW. ", "Error # " & Err.Number & " - " & Err.Description
End Function

Public Function DeleteFileWEx(lpSTR As Long) As Long
    Dim iAttr As Long, lr As Long
    
    iAttr = GetFileAttributes(lpSTR)
    
    If iAttr And FILE_ATTRIBUTE_READONLY Then SetFileAttributes lpSTR, iAttr And Not FILE_ATTRIBUTE_READONLY
    lr = DeleteFileW(lpSTR)
    
    If lr <> 0 Then 'success
        DeleteFileWEx = lr
    End If
End Function

Public Function MkDirW(ByVal Path As String, Optional ByVal LastComponentIsFile As Boolean = False) As Boolean
    On Error GoTo ErrorHandler
    ' Create folders struct
    ' LastComponentIsFile - true, if you specify filename as a last part of path component
    ' Return value: true, if successfully created or if folder is already exists
    Dim FC As String, lr As Boolean, pos As Long
    Dim bRedirect As Boolean, bOldStatus As Boolean
    If LastComponentIsFile Then Path = Left(Path, InStrRev(Path, "\") - 1) ' cut off file name
    If InStr(Path, ":") = 0 Then 'if relative path
        Dim sCurDir$, nChar As Long
        sCurDir = String$(MAX_PATH, 0&)
        nChar = GetCurrentDirectory(MAX_PATH + 1, StrPtr(sCurDir))
        sCurDir = Left$(sCurDir, nChar)
        If Right$(sCurDir, 1) <> "\" Then sCurDir = sCurDir & "\"
        Path = sCurDir & Path
    End If
    If FolderExists(Path) Then
        MkDirW = True
        Exit Function
    End If
    'bRedirect = ToggleWow64FSRedirection(False, Path, bOldStatus)
    Do 'looping through each path component
        pos = pos + 1
        pos = InStr(pos, Path, "\")
        If pos Then FC = Left(Path, pos - 1) Else FC = Path
        If FolderExists(FC) Then
            lr = True 'if folder is already created
        Else
            lr = CBool(CreateDirectory(StrPtr(FC), ByVal 0&))
            If lr = 0 Then Exit Do
        End If
    Loop While (pos <> 0) And (lr <> 0)
    MkDirW = lr
    'If bRedirect Then ToggleWow64FSRedirection bOldStatus
    Exit Function
ErrorHandler:
    Debug.Print "modFile.MkDirW. Error # " & Err.Number & " - " & Err.Description
End Function

' �������� ����� � ������� (���������� 0 � ������ ������)
Public Function SendFileToRecycleBin(ByVal FileName As String, Optional Confirm As Boolean = False, Optional Silent As Boolean = True) As Long
    On Error GoTo ErrorHandler
    Const FO_DELETE             As Long = &H3
    Const FOF_ALLOWUNDO         As Long = &H40
    Const FOF_NOCONFIRMATION    As Long = &H10
    Const FOF_SILENT            As Long = &H4
    Dim FileOp As SHFILEOPSTRUCT
    
    FileName = FileName & Chr(0)
    
    With FileOp
        .wFunc = FO_DELETE
        .pFrom = StrPtr(FileName)
        .fFlags = FOF_ALLOWUNDO
        If Not Confirm Then .fFlags = .fFlags Or FOF_NOCONFIRMATION
        If Silent Then .fFlags = .fFlags Or FOF_SILENT
    End With
    SendFileToRecycleBin = SHFileOperation(FileOp)
    Exit Function
ErrorHandler:
    Debug.Print "modFile.SendFileToRecycleBin. Error # " & Err.Number & " - " & Err.Description
End Function
