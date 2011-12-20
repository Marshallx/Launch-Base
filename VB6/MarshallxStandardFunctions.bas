Attribute VB_Name = "MarshallxStandardFunctions"
Option Explicit
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function FlushPrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Integer, ByVal lpKeyName As Integer, ByVal lpString As Integer, ByVal lpFileName As String) As Integer
Private Declare Function APIPlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As LargeInt, lpTotalNumberOfBytes As LargeInt, lpTotalNumberofFreeBytes As LargeInt) As Long
Declare Function GetDiskFreeSpace Lib "kernel32.dll" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal bytes As Long)
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_FILESONLY = &H80
Private Const FOF_NOERRORUI = &H400
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = -1&
Public Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const MAX_PATH = 260
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100
Private Const STARTF_USESHOWWINDOW = &H1
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const CREATE_ALWAYS = 2
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type
Private Type LargeInt
  lngLower As Long
  lngUpper As Long
End Type
Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Public Const IE As Long = 51

Function EncryptString(ByVal Text As String, ByVal Password As String) As String
    'encrypt a string using a password
    'you must reapply the same function (and same password) on
    'the encrypted string to obtain the original, non-encrypted string
    'you get better, more secure results if you use a long password
    '(e.g. 16 chars or longer). This routine works well only with ANSI strings.
    Dim passLen As Long
    Dim i As Long
    Dim passChr As Integer
    Dim passNdx As Long
    passLen = Len(Password)
    If passLen = 0 Then Err.Raise 5 'null passwords are invalid
    'move password chars into an array of Integers to speed up code
    ReDim passChars(0 To passLen - 1) As Integer
    CopyMemory passChars(0), ByVal StrPtr(Password), passLen * 2
    'this simple algorithm XORs each character of the string
    'with a character of the password, but also modifies the
    'password while it goes, to hide obvious patterns in the
    'result string
    For i = 1 To Len(Text)
        'get the next char in the password
        passChr = passChars(passNdx)
        'encrypt one character in the string
        Mid$(Text, i, 1) = Chr$(Asc(Mid$(Text, i, 1)) Xor passChr)
        'modify the character in the password (avoid overflow)
        passChars(passNdx) = (passChr + 17) And 255
        'prepare to use next char in the password
        passNdx = (passNdx + 1) Mod passLen
    Next
    EncryptString = Text
End Function

Function GetId()
    Dim List
    Dim Msg
    Dim Object
    On Local Error Resume Next
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_BaseBoard")
    For Each Object In List
        Msg = Msg & "Motherboard Serial Number: " & Object.SerialNumber & vbCrLf
    Next
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_Processor")
    For Each Object In List
        Msg = Msg & "Processor Unique ID: " & Object.UniqueID & vbCrLf
    Next
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_BIOS")
    For Each Object In List
        Msg = Msg & "BIOS Serial Number: " & Object.SerialNumber & vbCrLf
    Next
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_LogicalDisk")
    For Each Object In List
        Msg = Msg & "Disk Serial Number: " & Object.VolumeSerialNumber & vbCrLf
    Next
    MsgBox "this is a work in progress. do not use it yet!"
    MsgBox Msg
End Function

'Old and useless - should use CmdOutput module instead
Function MxShell(ByVal Program As String, Optional ByVal WaitForProcessToClose As Boolean = False, Optional ByVal HideProgram As Boolean = False, Optional ByVal RedirectOutput As String = "") As String
    Dim pInfo As PROCESS_INFORMATION
    Dim sInfo As STARTUPINFO
    Dim sNull As String
    Dim lSuccess As Long
    Dim hFile As Long
    Dim FileHandle As Integer
    Dim TempString As String
    Dim RetVal As String
    RetVal = ""
    If Len(RedirectOutput) <> 0 Then
        sInfo.cb = Len(sInfo)
        sInfo.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
        hFile = CreateFile(RedirectOutput, GENERIC_READ Or GENERIC_WRITE, 0, ByVal 0&, CREATE_ALWAYS, 0, ByVal 0&)
        sInfo.hStdOutput = hFile
        If hFile Then
            sInfo.hStdOutput = hFile
        Else
            RedirectOutput = ""
        End If
    End If
    If HideProgram Then
        sInfo.wShowWindow = SW_HIDE
    Else
        sInfo.wShowWindow = SW_SHOWNORMAL
    End If
    lSuccess = CreateProcess(sNull, Program, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, sNull, sInfo, pInfo)
    If WaitForProcessToClose And lSuccess Then WaitForSingleObject pInfo.hProcess, INFINITE
    'lRetValue = TerminateProcess(pInfo.hProcess, 0&)
    lSuccess = CloseHandle(pInfo.hThread)
    lSuccess = CloseHandle(pInfo.hProcess)
    lSuccess = CloseHandle(hFile)
    If Len(RedirectOutput) <> 0 Then
        If FileExists(RedirectOutput) Then
            FileHandle = FreeFile
            Open RedirectOutput For Input As #FileHandle
                Do While Not EOF(FileHandle)
                    Line Input #FileHandle, TempString
                    RetVal = RetVal & TempString & vbCrLf
                Loop
            Close #FileHandle
            Call Kill(RedirectOutput)
        End If
    End If
    MxShell = RetVal
End Function

Public Function GetFileMD5(ByVal sPath As String) As String
    GetFileMD5 = Left$(GetCommandOutput(Quote(JoinPath(RESDIR, "md5deep.exe")) & " " & Quote(sPath)), 32)
End Function

Public Sub ShellAndWait(ByVal sShell As String, ByVal window_style As VbAppWinStyle)
    Dim process_id As Long
    Dim process_handle As Long
    process_id = Shell(sShell, window_style)
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
End Sub

Public Function TestFlags(FlagsProvided, FlagsRequired) As Boolean
    TestFlags = ((FlagsProvided And FlagsRequired) = FlagsRequired)
End Function

'********************************************************************************
'********************************************************************************
'********************  FILES AND FILENAME/FILEPATH STRINGS   ********************
'********************************************************************************
'********************************************************************************

'Private Function APIFunctionPresent(ByVal FunctionName As String, ByVal DllName As String) As Boolean
'   'http://www.freevbcode.com/ShowCode.Asp?ID=429
'    Dim lHandle As Long
'    Dim lAddr  As Long
'    lHandle = LoadLibrary(DllName)
'    If lHandle <> 0 Then
'        lAddr = GetProcAddress(lHandle, FunctionName)
'        FreeLibrary lHandle
'    End If
'    APIFunctionPresent = (lAddr <> 0)
'End Function

Public Function GetShortFileName(ByVal FullPath As String) As String
    Dim lAns As Long
    Dim sAns As String
    Dim ILen As Integer
    On Error Resume Next
    If Dir(FullPath) = "" Then
        'this function doesn't work if the file doesn't exist
        If Dir(FullPath, vbDirectory) = "" Then Exit Function
    End If
    sAns = Space(255)
    lAns = GetShortPathName(FullPath, sAns, 255)
    GetShortFileName = Left(sAns, lAns)
End Function

Private Function GetClusterSize(disk As String) As Currency
    Dim sectorsPerCluster As Long
    Dim bytesPerSector As Long
    Dim free As Long
    Dim total As Long
    Dim RetVal As Long
    RetVal = GetDiskFreeSpace(disk, sectorsPerCluster, bytesPerSector, free, total)
    GetClusterSize = sectorsPerCluster * bytesPerSector
End Function

Function GetDirectorySize(ByVal sPath As String, Optional ByVal GetAsDiskUsage As Boolean = True, Optional ByVal Recurse As Boolean = True) As Double
'Precondition: sPath exists and each file is less than 2GB
    'On Error Resume Next
    Dim h As Long
    Dim FD As WIN32_FIND_DATA
    Dim r As Long
    Dim dSize As Double
    Dim sName As String
    Dim ClusterSize As Long
    If GetAsDiskUsage Then ClusterSize = GetClusterSize(Left$(sPath, 3))
    'Get handle to first file or subfolder in folder.
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    h = FindFirstFile(sPath & "*", FD)
    If h <> INVALID_HANDLE_VALUE Then
        Do
            sName = Left$(FD.cFileName, InStr(FD.cFileName, vbNullChar) - 1)
            If Left$(sName, 1) <> "." Then
                'DoEvents
                If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                    'If the handle is to a folder then call the function recursively.
                    If Recurse Then dSize = dSize + GetDirectorySize(sPath & sName)
                Else
                    'Debug.Print sName & vbTab & "Low: " & FD.nFileSizeLow
                    If GetAsDiskUsage Then
                        dSize = dSize + ((Int(FD.nFileSizeLow / ClusterSize)) * ClusterSize)
                        If (FD.nFileSizeLow Mod ClusterSize) > 0 Then dSize = dSize + ClusterSize
                    Else
                        dSize = dSize + FD.nFileSizeLow
                    End If
                End If
            End If
            'DoEvents
        Loop While FindNextFile(h, FD)
        r = FindClose(h): Debug.Assert r
    End If
    GetDirectorySize = dSize
End Function

Function GetFileSize(sPath As String, Optional ByVal GetAsDiskUsage As Boolean = False) As Currency
    On Error Resume Next
    Dim fso As FileSystemObject
    Dim fsoFile As File
    Dim ClusterSize As Currency
    Set fso = New FileSystemObject
    Set fsoFile = fso.GetFile(sPath)
    GetFileSize = fsoFile.Size
    Set fsoFile = Nothing
    Set fso = Nothing
    If GetAsDiskUsage Then
        ClusterSize = GetClusterSize(Left$(sPath, 3))
        GetFileSize = CInt((GetFileSize / ClusterSize)) * ClusterSize
    End If
End Function

Function FreeDiskSpace(ByVal sDriveLetter As String) As Double
    sDriveLetter = UCase$(sDriveLetter) & ":\"
    Dim udtFreeBytesAvail As LargeInt, udtTtlBytes As LargeInt
    Dim udtTTlFree As LargeInt
    Dim dblFreeSpace As Double
    If GetDiskFreeSpaceEx(sDriveLetter, udtFreeBytesAvail, udtTtlBytes, udtTTlFree) Then
        If udtFreeBytesAvail.lngLower < 0 Then
           dblFreeSpace = udtFreeBytesAvail.lngUpper * 2 ^ 32 + udtFreeBytesAvail.lngLower + 4294967296#
        Else
           dblFreeSpace = udtFreeBytesAvail.lngUpper * 2 ^ 32 + udtFreeBytesAvail.lngLower
        End If
    End If
    FreeDiskSpace = dblFreeSpace
End Function

Function FileExists(ByVal FilePath As String) As Boolean
    On Error GoTo ErrorHandler:
    FileExists = (GetAttr(FilePath) And vbDirectory) = 0
ErrorHandler:
End Function

Function DirExists(ByVal DirPath As String) As Boolean
    On Error GoTo ErrorHandler:
    DirExists = (GetAttr(DirPath) And vbDirectory) = vbDirectory
ErrorHandler:
End Function

Public Function HDSerialNumber(ByVal sDrive As String) As Long
    If Len(sDrive) Then
        If InStr(sDrive, "\\") = 1 Then
            ' Make sure we end in backslash for UNC
            If Right$(sDrive, 1) <> "\" Then
                sDrive = sDrive & "\"
            End If
        Else
            ' If not UNC, take first letter as drive
            sDrive = Left$(sDrive, 1) & ":\"
        End If
    Else
        ' Else just use current drive
        sDrive = vbNullString
    End If
    ' Grab S/N -- Most params can be NULL
    Call GetVolumeInformation(sDrive, vbNullString, 0, HDSerialNumber, ByVal 0&, ByVal 0&, vbNullString, 0)
End Function


Function IsValidPath(ByVal PathToValidate As String) As Boolean
    Dim Folders() As String
    Dim Counter As Integer
    Dim RetVal As Boolean
    RetVal = False
    If Len(PathToValidate) <> 0 Then
        Folders() = Split(PathToValidate, "\")
        If Len(Folders(0)) = 2 Then
            If Right$(Folders(0), 1) = ":" Then
                If Len(StripInvalidChars(Left$(Folders(0), 1), InvalidFileChars)) = 1 Then
                    Counter = 1
                    Do While Counter <= UBound(Folders)
                        RetVal = False
                        If Len(Folders(Counter)) <> 0 Then
                            If Len(StripInvalidChars(Folders(Counter), InvalidFileChars)) = Len(Folders(Counter)) Then
                                RetVal = True
                            End If
                        Else
                            If Counter = UBound(Folders) Then RetVal = True
                        End If
                        If RetVal = False Then Exit Do
                        Counter = Counter + 1
                    Loop
                End If
            End If
        End If
    End If
    IsValidPath = RetVal
End Function

Public Function DirIsEmpty(ByVal DirPath As String) As Boolean
    'Requires Reference: Microsoft Scripting Runtime
    Dim bEmpty As Boolean
    Dim fso As FileSystemObject
    bEmpty = False
    Set fso = New FileSystemObject
    If fso.GetFolder(DirPath).Files.Count = 0 Then
        If fso.GetFolder(DirPath).SubFolders.Count = 0 Then
            bEmpty = True
        End If
    End If
    DirIsEmpty = bEmpty
End Function

Function MakePath(ByVal Path As String) As Boolean
    Dim i As Integer, ercode As Long
    Dim Success As Boolean
    If Mid$(Path, Len(Path), 1) <> "\" Then Path = Path & "\" 'ERM
    On Error Resume Next
    Do
        ' get the next path chunk
        i = InStr(i + 1, Path & "\", "\")
        
        ' try to create this sub-directory
        Err.Clear
        MkDir Left$(Path, i - 1)
        If Err = 0 Then
            ' the directory has been created
            ' do nothing
            Success = True
        ElseIf Err = 75 Then
            ' Path\File Access Error: the directory exists
            ' do nothing
            Success = True
            Call Err.Clear
        Else
            ' we can't continue if any other error
            ''ercode = Err
            ''On Error GoTo 0
            ''Err.Raise ercode
            Err.Raise Err
            Success = False
        End If
    Loop Until i > Len(Path)
    MakePath = Success
End Function

Public Sub KillDir(ByVal DirPath As String)
    'Requires Reference: Microsoft Scripting Runtime
    'deletes a folder and all its contents
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Call fso.DeleteFolder(DirPath, True)
End Sub

Function FileType(ByVal FileName As String) As String
    Dim DotPos As Long
    DotPos = InStrRev(FileName, ".")
    If DotPos > 0 Then
        FileType = UCase$(Mid$(FileName, DotPos + 1))
    Else
        FileType = ""
    End If
End Function

Function ChangeFileType(ByVal FileName As String, ByVal NewExtension As String) As String
    Dim DotPos As Long
    If Len(NewExtension) <> 0 Then
        If Left(NewExtension, 1) = "." Then
            If Len(NewExtension) <> 1 Then
                NewExtension = Mid$(NewExtension, 2)
            Else
                NewExtension = ""
            End If
        End If
    End If
    DotPos = InStrRev(FileName, ".")
    If DotPos > 0 Then
        ChangeFileType = Mid$(FileName, 1, DotPos) & NewExtension
    Else
        ChangeFileType = FileName & "." & NewExtension
    End If
End Function

Function StripExtension(ByVal FileName As String) As String
    Dim DotPos As Long
    DotPos = InStrRev(FileName, ".")
    If DotPos > 0 Then
        StripExtension = Mid$(FileName, 1, DotPos - 1)
    Else
        StripExtension = FileName
    End If
End Function

Sub SplitPath(ByVal FullPath As String, ByRef FileName, ByRef FilePath)
'Call SplitPath("C:\Program Files\Marshallx\Hello.txt", FileName, FilePath)
  'FileName = "Hello.txt"
  'FilePath = "C:\ProgramFiles\Marshallx"
'Call SplitPath("http://www.website.com/file.txt", FileName, FilePath)
  'FileName = "file.txt"
  'FilePath = "http://www.website.com"
    Dim SlashPos As Long
    FileName = FullPath
    FilePath = ""
    If Len(FullPath) <> 0 Then
        If Right$(FullPath, 1) = "\" Or Right$(FullPath, 1) = "/" Then FullPath = Left$(FullPath, Len(FullPath) - 1)
        SlashPos = InStrRev(FullPath, "\")
        If SlashPos = 0 Then SlashPos = InStrRev(FullPath, "/")
        If SlashPos <> 0 Then
            If SlashPos < Len(FullPath) Then FileName = Mid$(FullPath, SlashPos + 1) Else FileName = ""
            If SlashPos > 1 Then FilePath = Mid$(FullPath, 1, SlashPos - 1) Else FilePath = ""
            If Len(FilePath) = 2 Then FilePath = FilePath & "\" 'add backslash back on for root directories
        End If
    End If
End Sub

Function GetRootPath(ByVal sFullPath As String) As String
'GetRootPath("\games\pacman\pacman.exe") == "games"
    Dim SlashPos As Long
    GetRootPath = ""
    If Len(sFullPath) <> 0 Then
        If Left$(sFullPath, 1) = "\" Then sFullPath = Right$(sFullPath, Len(sFullPath) - 1)
        SlashPos = InStr(sFullPath, "\")
        If SlashPos <> 0 Then
            If SlashPos > 1 Then GetRootPath = Left$(sFullPath, SlashPos - 1)
        Else
            GetRootPath = sFullPath
        End If
    End If
End Function

Function GetRelativePath(ByVal FullPath As String, ByVal SourceDir As String, Optional ByVal UseForwardSlash As Boolean = False) As String
    Dim Counter As Long
    Dim RetVal As String
    Dim SlashPos As Long
    Dim SlashStr As String
    Select Case UseForwardSlash
    Case False: SlashStr = "\"
    Case True: SlashStr = "/"
    End Select
    If Len(FullPath) <> 0 And Len(SourceDir) <> 0 Then
        If Right$(SourceDir, 1) = SlashStr Then SourceDir = Left$(SourceDir, Len(SourceDir) - 1)
        If Right$(FullPath, 1) = SlashStr Then FullPath = Left$(FullPath, Len(FullPath) - 1)
        Counter = 1
        Do
            If Len(FullPath) < Counter Then Exit Do
            If Len(SourceDir) < Counter Then Exit Do
            If UCase$(Mid$(FullPath, Counter, 1)) <> UCase$(Mid$(SourceDir, Counter, 1)) Then Exit Do
            Counter = Counter + 1
        Loop
        If Counter = 1 Then
            'on a different drive
            RetVal = FullPath
        Else
            Counter = Counter - 1
            If Len(FullPath) > Counter Then
                If Mid$(FullPath, (Counter + 1), 1) = SlashStr Then
                    If Len(SourceDir) = Counter Then
                        Counter = Counter + 1
                        SourceDir = SourceDir & SlashStr
                    End If
                End If
            Else
                If Len(SourceDir) > Counter Then
                    If Mid$(SourceDir, (Counter + 1), 1) = SlashStr Then
                        Counter = Counter + 1
                        FullPath = FullPath & SlashStr
                    End If
                End If
            End If
            SlashPos = InStrRev(Mid$(FullPath, 1, Counter), SlashStr)
            FullPath = Mid$(FullPath, SlashPos + 1)
            SourceDir = Mid$(SourceDir, SlashPos + 1)
            SlashPos = InStrRev(SourceDir, SlashStr)
            Do While SlashPos <> 0
                RetVal = RetVal & ".." & SlashStr
                SourceDir = Mid$(SourceDir, 1, SlashPos - 1)
                SlashPos = InStrRev(SourceDir, SlashStr)
            Loop
            If Len(SourceDir) <> 0 Then RetVal = RetVal & ".." & SlashStr
            RetVal = RetVal & FullPath
        End If
    End If
    GetRelativePath = RetVal
End Function

Function GetFilePath(ByVal FullPath As String) As String
'GetFilePath("C:\Program Files\Marshallx\Hello.txt") = "C:\Program Files\Marshallx"
    Dim FileName, FilePath As String
    Call SplitPath(FullPath, FileName, FilePath)
    GetFilePath = FilePath
End Function

Function GetFileName(ByVal FullPath As String) As String
'GetFileName("C:\Program Files\Marshallx\Hello.txt") = "Hello.txt"
    Dim FileName, FilePath As String
    Call SplitPath(FullPath, FileName, FilePath)
    GetFileName = FileName
End Function

Function JoinPath(ByVal FilePath, ByVal FileName, Optional ByVal SecondJoin As String = "") As String
'JoinPath("C:\Program Files", "Marshallx") = "C:\Program Files\Marshallx"
'JoinPath("C:\Program Files\", "Marshallx") = "C:\Program Files\Marshallx"
'JoinPath(JoinPath("C:\Program Files", "Marshallx"), "Hello.txt") = "C:\Program Files\Marshallx\Hello.txt"
    If Len(FileName) <> 0 Or Len(SecondJoin) <> 0 Then
        If Len(FilePath) <> 0 Then
            If Right$(FilePath, 1) <> "\" Then FilePath = FilePath & "\"
        End If
    End If
    If Len(SecondJoin) <> 0 And Len(FileName) <> 0 Then
        If Right$(FileName, 1) <> "\" Then FileName = FileName & "\"
    End If
    JoinPath = FilePath & FileName & SecondJoin
End Function


'********************************************************************************
'********************************************************************************
'********************     INI/REGISTRY READING/WRITING      *********************
'********************************************************************************
'********************************************************************************

'Needs: Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Function ReadINIStr(ByVal SectionName As String, ByVal Flag As String, ByVal File As String, Optional DefaultVal As String = "", Optional ByVal bIncludeComment As Boolean = False) As String
'ReturnComment includes the semicolon, but strips all spaces between the end of the value and the start of the comment. Also rtrims all spaces after the comment.
    Dim RetVal As String * 255, FinalVal As String, v As Long, CommentStart As Integer, Counter As Integer
    Dim ReturnComment As String
    v = GetPrivateProfileString(SectionName, Flag, "", RetVal, 255, File)
    FinalVal = RetVal
    FinalVal = Left$(FinalVal, v)
    If Not bIncludeComment Then
        CommentStart = InStr(1, FinalVal, ";", vbTextCompare)
        Select Case CommentStart
        Case 0
            'do nothing, RetVal is Okay.
        Case 1
            ReturnComment = FinalVal
            FinalVal = ""
        Case Else
            ReturnComment = Mid$(FinalVal, CommentStart)
            FinalVal = Mid$(FinalVal, 1, CommentStart - 1)
        End Select
        ReturnComment = RTrim$(ReturnComment)
    End If
    FinalVal = RTrim$(FinalVal)
    If FinalVal = "" Then FinalVal = DefaultVal
    ReadINIStr = FinalVal
'ReturnComment is not returned in this function. You can enable this feature if you want to by altering the parameter list.
End Function

'Needs: Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Sub WriteINIStr(ByVal SectionName As String, ByVal Flag As String, ByVal IniValue As String, ByVal File As String)
    Call WritePrivateProfileString(SectionName, Flag, IniValue, File)
End Sub

Sub FlushINI(ByVal sIniFile As String)
    Call FlushPrivateProfileString(0, 0, 0, sIniFile)
End Sub

Function ReadRegStr(ByVal Value As String, Optional ByVal DefaultVal As String = "") As String
'Usage: ReadRegStr("HKLM\System\CurrentControlSet\Services\VxD\VNETSUP\Workgroup")
    Dim o As Object
    Dim s As String
    On Error Resume Next
    Set o = CreateObject("wscript.shell")
    s = o.RegRead(Value)
    If s = "" Then s = DefaultVal
    ReadRegStr = s
End Function

Sub WriteRegStr(ByVal Path As String, ByVal Value As String)
    Dim o As Object
    Dim s As String
    Set o = CreateObject("wscript.shell")
    s = o.RegWrite(Path, Value)
End Sub

'If INI string is just read in from a file, say.
Function SplitINIStr(ByVal sInput, ByRef sValue, Optional ByRef sFlag As String, Optional ByRef sComment As String) As String
    'comment includes semicolon
    'strips [only] trailing spaces from value only
    Dim iPos As Integer
    iPos = InStr(2, sInput, "=")
    sValue = ""
    sFlag = ""
    sComment = ""
    If iPos <> 0 Then
        sFlag = Left$(sInput, iPos - 1)
        If Len(sInput) - iPos <> 0 Then
            sValue = Mid$(sInput, iPos + 1)
            iPos = InStr(1, sValue, ";")
            If iPos <> 0 Then
                sComment = Mid$(sValue, iPos)
                If iPos <> 1 Then
                    sValue = Left$(sValue, iPos - 1)
                End If
            End If
            sValue = RTrim$(sValue)
        End If
    End If
    SplitINIStr = sValue
End Function

Function ConvertEOL(ByVal sBuffer As String, Optional ByVal sConvertTo As String = vbCrLf) As String
    Dim lPosCrLf As Long
    Dim lPosCr As Long
    Dim lPosLf As Long
    Dim lPos As Long
    ConvertEOL = ""
    lPos = 0
    Do While lPos <> Len(sBuffer)
        lPos = lPos + 1
        Select Case Mid$(sBuffer, lPos, 1)
        Case vbCr
            ConvertEOL = ConvertEOL & sConvertTo
            If lPos <> Len(sBuffer) Then
                If Mid$(sBuffer, lPos + 1, 1) = vbLf Then
                    lPos = lPos + 1 'extra one
                End If
            End If
        Case vbLf
            ConvertEOL = ConvertEOL & sConvertTo
        Case Else
            ConvertEOL = ConvertEOL & Mid$(sBuffer, lPos, 1)
        End Select
    Loop
End Function

'This should only be used on small INI files stored in memory
Function ReadINIStrMemory(ByVal sBuffer As String, ByVal sSection As String, ByVal sFlag As String, Optional ByRef sValue As String, Optional ByRef sComment As String, Optional ByVal sDefaultValue As String = "") As String
    'New lines must be CrLf - otherwise it won't count as a new line
    'Flags in the INI must have equals immediately after text - no spaces - otherwise it won't be recognised
    'Comment will not be trimmed and will start at the semicolon
    'Value will be trimmed both sides
    Dim lPos As Long
    sValue = ""
    sComment = ""
    'section must have square brackets round it
    If Len(sSection) <> 0 Then
        If Left$(sSection, 1) <> "[" Then sSection = "[" & sSection & "]"
    End If
    'find the section
    lPos = InStr(1, sBuffer, sSection)
    Do While lPos > 1
        Select Case Mid$(sBuffer, lPos - 1, 1)
        Case vbCr, vbLf
            Exit Do
        Case Else
            lPos = InStr(lPos + 1, sBuffer, sSection)
        End Select
    Loop
    If lPos <> 0 Then
        'found the section - trim the buffer to just this section
        sBuffer = Mid$(sBuffer, lPos)
        lPos = InStr(1, sBuffer, vbCrLf & "[")
        If lPos <> 0 Then sBuffer = Left$(sBuffer, lPos - 1)
        'flag must have equals
        If Len(sFlag) <> 0 Then
            If Right$(sFlag, 1) <> "=" Then sFlag = sFlag & "="
        End If
        'find the flag
        lPos = InStr(1, sBuffer, sFlag)
        Do While lPos > 1
            Select Case Mid$(sBuffer, lPos - 1, 1)
            Case vbCr, vbLf
                Exit Do
            Case Else
                lPos = InStr(lPos + 1, sBuffer, sFlag)
            End Select
        Loop
        If lPos <> 0 Then
            'found the flag - trim the buffer to just the rest of the line
            If Len(sBuffer) > lPos + Len(sFlag) Then
                sBuffer = Mid$(sBuffer, lPos + Len(sFlag))
                lPos = InStr(1, sBuffer, vbCrLf)
                If lPos <> 0 Then sBuffer = Left$(sBuffer, lPos - 1)
                lPos = InStr(1, sBuffer, ";")
                Select Case lPos
                Case 0
                    sValue = Trim$(sBuffer)
                Case 1
                    sComment = sBuffer
                Case Else
                    sValue = Trim$(Mid$(sBuffer, 1, lPos - 1))
                    sComment = Mid$(sBuffer, lPos)
                End Select
            End If
        End If
    End If
    ReadINIStrMemory = sValue
End Function

'********************************************************************************
'********************************************************************************
'********************                STRINGS                *********************
'********************************************************************************
'********************************************************************************

Function Quote(Optional ByVal StringToQuote As String = "") As String
    Quote = Chr(34) & StringToQuote & Chr(34)
End Function

Function DeQuote(ByVal StringToDeQuote As String) As String
    DeQuote = StringToDeQuote
    If Len(StringToDeQuote) >= 2 Then
        If Left$(StringToDeQuote, 1) = Chr(34) Then
            If Right$(StringToDeQuote, 1) = Chr(34) Then
                DeQuote = Mid$(StringToDeQuote, 2, Len(StringToDeQuote) - 2)
            End If
        End If
    End If
End Function

Function PadNum(ByVal MyNumber As Double, Optional ByVal MaskChars As Long = 1) As String
'PadNum(3, 4) = "0003"
'PadNum(12345) = "12345"
'PadNum(-9, 3) = "-009"
    Dim MyString As String
    Dim iCounter As Integer
    MyString = CStr(MyNumber)
    'Convert to integer
    iCounter = InStr(1, MyString, ".")
    If iCounter > 0 Then MyString = Left$(MyString, iCounter - 1)
    If MyString = "" Then MyString = "0"
    'Remove negative
    If Left$(MyString, 1) = "-" Then MyString = Right$(MyString, Len(MyString) - 1)
    'Padding
    Do While Len(MyString) < MaskChars
        MyString = "0" & MyString
    Loop
    'Restore negative
    If MyNumber < 0 Then MyString = "-" & MyString
    'All done
    PadNum = MyString
End Function

Function PadString(ByVal MyString As String, Optional ByVal MaskChar As String = " ", Optional ByVal MaskChars As Long = 1) As String
'PadChar("hello", "-", 7) = "--hello"
    Do While Len(MyString) < MaskChars
        MyString = MaskChar & MyString
    Loop
    PadString = MyString
End Function

Function ChangeCase(ByVal Message As String, ByVal CaseMask As String) As String
'ChangeCase("Hello", "00 1") = "helLo"
'ChangeCase("Hello", "  101101") = "HeLlO"
    Dim Counter As Integer
    Dim RetVal As String
    RetVal = ""
    Counter = 1
    Do While (Counter <= Len(Message)) And (Counter <= Len(CaseMask))
        Select Case Mid$(CaseMask, Counter, 1)
        Case " ": RetVal = RetVal & Mid$(Message, Counter, 1)
        Case "0": RetVal = RetVal & UCase$(Mid$(Message, Counter, 1))
        Case "1": RetVal = RetVal & UCase$(Mid$(Message, Counter, 1))
        End Select
        Counter = Counter + 1
    Loop
    If Counter <= Len(Message) Then
        RetVal = RetVal & Mid$(Message, Counter)
    End If
    ChangeCase = RetVal
End Function

Function BooleanStringToBoolean(ByVal BooleanString As String, Optional DefaultVal As Boolean = False) As Boolean
    Select Case UCase$(BooleanString)
    Case "0", "NO", "FALSE": BooleanStringToBoolean = False
    Case "1", "YES", "TRUE": BooleanStringToBoolean = True
    Case Else: BooleanStringToBoolean = DefaultVal
    End Select
End Function

Function BooleanStringToInteger(ByVal BooleanString As String, Optional DefaultVal As Integer = 0) As Integer
    Select Case UCase$(BooleanString)
    Case "0", "NO", "FALSE": BooleanStringToInteger = 0
    Case "1", "YES", "TRUE": BooleanStringToInteger = 1
    Case Else: BooleanStringToInteger = DefaultVal
    End Select
End Function

Function IntegerToYesNo(ByVal MyNumber As Integer) As String
    Select Case MyNumber
    Case 0: IntegerToYesNo = "no"
    Case Else: IntegerToYesNo = "yes"
    End Select
End Function

Function BooleanToYesNo(ByVal MyBoolean As Boolean) As String
    Select Case MyBoolean
    Case True: BooleanToYesNo = "yes"
    Case False: BooleanToYesNo = "no"
    End Select
End Function

Function DataSize(ByVal iBytes As Double, Optional sTargetUnit As String = "KB", Optional ByVal sSourceUnit As String = "B", Optional ByVal bIncludeTargetUnit As Boolean = True) As String
    Dim iDivisor As Double
    Dim iMultiplier As Double
    iDivisor = 1
    iMultiplier = 1
    Select Case sSourceUnit
    Case "B"
        Select Case sTargetUnit
        Case "B": iMultiplier = 1
        Case "KB": iDivisor = 1024
        Case "MB": iDivisor = 1048576
        Case "GB": iDivisor = 1073741824
        End Select
    Case "KB"
        Select Case sTargetUnit
        Case "B": iMultiplier = 1024
        Case "KB": iMultiplier = 1
        Case "MB": iDivisor = 1024
        Case "GB": iDivisor = 1048576
        End Select
    Case "MB"
        Select Case sTargetUnit
        Case "B": iMultiplier = 1048576
        Case "KB": iMultiplier = 1024
        Case "MB": iMultiplier = 1
        Case "GB": iDivisor = 1024
        End Select
    Case "GB"
        Select Case sTargetUnit
        Case "B": iMultiplier = 1073741824
        Case "KB": iMultiplier = 1048576
        Case "MB": iMultiplier = 1024
        Case "GB": iMultiplier = 1
        End Select
    End Select
    If iDivisor <> 1 Then
        iBytes = iBytes / iDivisor
    Else
        iBytes = iBytes * iMultiplier
    End If
    DataSize = Format(iBytes, "###,###,###,###,###")
    If Len(DataSize) = 0 Then DataSize = "0"
    If bIncludeTargetUnit Then DataSize = DataSize & " " & sTargetUnit
End Function

Function StripLeadingZeroes(ByVal WorkingString As String) As String
'StripLeadingZeroes("000123") = "123"
'StripLeadingZeroes("000000") = "0"
    Dim ZeroPos As Long
    ZeroPos = InStr(WorkingString, "0")
    Do While (ZeroPos = 1) And (Len(WorkingString) > 1)
        WorkingString = Mid$(WorkingString, 2)
        ZeroPos = InStr(WorkingString, "0")
    Loop
    StripLeadingZeroes = WorkingString
End Function

Function StripNonNumbers(ByVal WorkingString As String, Optional ByVal AllowFloatingPoint As Boolean = False, Optional ByVal AllowMinusSign As Boolean = False) As String
    Dim TrailerString As String
    If Len(WorkingString) = 0 Then
        StripNonNumbers = ""
    Else
        TrailerString = Mid$(WorkingString, 1, 1)
        WorkingString = Mid$(WorkingString, 2)
        If TrailerString >= Chr(48) And TrailerString <= Chr(57) Then
            StripNonNumbers = TrailerString & StripNonNumbers(WorkingString)
        Else
            If AllowFloatingPoint And TrailerString = Chr(46) Then
                StripNonNumbers = TrailerString & StripNonNumbers(WorkingString)
            Else
                If AllowMinusSign And TrailerString = Chr(45) Then
                    StripNonNumbers = TrailerString & StripNonNumbers(WorkingString)
                Else
                    StripNonNumbers = StripNonNumbers(WorkingString)
                End If
            End If
        End If
    End If
End Function

Function StripNumbers(ByVal WorkingString As String) As String
    Dim TrailerString As String
    If Len(WorkingString) = 0 Then
        StripNumbers = ""
    Else
        TrailerString = Mid$(WorkingString, 1, 1)
        WorkingString = Mid$(WorkingString, 2)
        If TrailerString >= Chr(48) And TrailerString <= Chr(57) Then
            StripNumbers = StripNonNumbers(WorkingString)
        Else
            StripNumbers = TrailerString & StripNonNumbers(WorkingString)
        End If
    End If
End Function

Function StripNonFloat(ByVal WorkingString As String) As String
    Dim TrailerString As String
    If Len(WorkingString) = 0 Then
        StripNonFloat = ""
    Else
        TrailerString = Mid$(WorkingString, 1, 1)
        WorkingString = Mid$(WorkingString, 2)
        If TrailerString <> Chr(46) And (TrailerString < Chr(48) Or TrailerString > Chr(57)) Then
            StripNonFloat = StripNonFloat(WorkingString)
        Else
            StripNonFloat = TrailerString & StripNonFloat(WorkingString)
        End If
    End If
End Function

Function InvalidFileChars() As String
    InvalidFileChars = Chr(92) & " " & Chr(47) & " " & Chr(58) & " " & Chr(42) & " " & Chr(63) & " " & Chr(34) & " " & Chr(60) & " " & Chr(62) & " " & Chr(124)
End Function

Function Min(ByVal iVal1 As Double, ByVal iVal2 As Double) As Double
    If iVal1 <= iVal2 Then
        Min = iVal1
    Else
        Min = iVal2
    End If
End Function

Function Max(ByVal iVal1 As Double, ByVal iVal2 As Double) As Double
    If iVal1 >= iVal2 Then
        Max = iVal1
    Else
        Max = iVal2
    End If
End Function

Function Restrict(ByVal iVal1 As Double, ByVal iVal2 As Double, ByVal iVal3) As Double
    If iVal2 >= iVal1 Then
        If iVal2 <= iVal3 Then
            Restrict = iVal2
        Else
            Restrict = iVal3
        End If
    Else
        Restrict = iVal1
    End If
End Function

Function Within(ByVal iVal1 As Double, ByVal iVal2 As Double, ByVal iVal3) As Boolean
    Within = False
    If iVal2 >= iVal1 Then
        If iVal2 <= iVal3 Then Within = True
    End If
End Function

Public Function RoundIt(ByVal aNumberToRound As Double, Optional ByVal aDecimalPlaces As Double = 0) As Double
'by Gary German
'uses Asymmetric Arithmetic Rounding
'RoundIt(2.5) = 3
'RoundIt(-2.5) = -2
    Dim nFactor As Double
    Dim nTemp As Double
    nFactor = 10 ^ aDecimalPlaces
    nTemp = (aNumberToRound * nFactor) + 0.5
    RoundIt = Int(CDec(nTemp)) / nFactor
'Symmetric Arithmetic Rounding:
'Function SymArith(ByVal x As Double, Optional ByVal DecimalPlaces As Double = 1) As Double
    'SymArith = Fix(x * (10 ^ DecimalPlaces) + 0.5 * Sgn(x)) / (10 ^ DecimalPlaces)
'End Function
End Function

Function StripInvalidChars(ByVal InputString As String, ByVal InvalidChars As String) As String
'InvalidChars is any string. Spaces are ignored so can never be considered invalid.
    Dim StringLen As Integer
    Dim StringPos As Integer
    Dim Counter As Integer
    StringLen = Len(InputString)
    Counter = 1
    Do While Counter <= Len(InvalidChars)
        If Not (Mid$(InvalidChars, Counter, 1) = " ") Then
            StringPos = 0
            Do While StringPos <= StringLen
                StringPos = StringPos + 1
                If Mid$(InputString, StringPos, 1) = Mid$(InvalidChars, Counter, 1) Then
                    Select Case StringPos
                    Case 1: InputString = Mid$(InputString, StringPos + 1)
                    Case Len(InputString): InputString = Mid$(InputString, 1, StringPos - 1)
                    Case Else: InputString = Mid$(InputString, 1, StringPos - 1) & Mid$(InputString, StringPos + 1)
                    End Select
                    StringPos = StringPos - 1
                End If
            Loop
        End If
        Counter = Counter + 1
    Loop
    StripInvalidChars = InputString
End Function

Function CharIsAlpha(ByVal sChar, Optional ByVal bReturnTrueIfUpper As Boolean = True, Optional ByVal bReturnTrueIfLower As Boolean = True) As Boolean
    If bReturnTrueIfUpper Then
        If Asc(sChar) >= 65 Then
            If Asc(sChar) <= 90 Then
                CharIsAlpha = True
                Exit Function
            End If
        End If
    End If
    If bReturnTrueIfLower Then
        If Asc(sChar) >= 97 Then
            If Asc(sChar) <= 122 Then
                CharIsAlpha = True
                Exit Function
            End If
        End If
    End If
    CharIsAlpha = False
End Function

Function CharIsNumber(ByVal sChar) As Boolean
    CharIsNumber = False
    If Asc(sChar) >= 48 Then
        If Asc(sChar) <= 57 Then
            CharIsNumber = True
        End If
    End If
End Function

Function ReplaceString(ByVal InputString As String, ByVal StringToReplace As String, Optional ByVal ReplacementString As String = "") As String
    'Replaces all occurrences of StringToReplace in InputString with ReplacementString
    'matching is case sensitive
    Dim StringPos As Integer
    Dim Counter As Integer
    If InStr(1, ReplacementString, StringToReplace) Then
        Call Err.Raise(vbObjectError + 1001, , "ReplaceString function encountered an infinite loop.")
    End If
    StringPos = -1
    Do While StringPos <> 0
        StringPos = InStr(1, InputString, StringToReplace)
        Select Case StringPos
        Case 0
            'do nothing, StringToReplace not found
        Case 1
            If Len(InputString) = 1 Then
                InputString = ReplacementString
            Else
                InputString = ReplacementString & Mid$(InputString, 2)
            End If
        Case Len(InputString)
            InputString = Mid$(InputString, 1, StringPos - 1) & ReplacementString
        Case Else
            InputString = Mid$(InputString, 1, StringPos - 1) & ReplacementString & Mid$(InputString, StringPos + Len(StringToReplace))
        End Select
    Loop
    ReplaceString = InputString
End Function

Function DoubleAmpersand(ByVal InputString As String) As String
    Dim StringPos As Integer
    Dim RetVal As String
    StringPos = 0
    Do While StringPos < Len(InputString)
        StringPos = StringPos + 1
        If Mid$(InputString, StringPos, 1) = "&" Then
            If StringPos < Len(InputString) Then
                If Mid$(InputString, StringPos + 1, 1) <> "&" Then
                    RetVal = RetVal & "&"
                End If
            Else
                RetVal = RetVal & "&"
            End If
        End If
        RetVal = RetVal & Mid$(InputString, StringPos, 1)
    Loop
    DoubleAmpersand = RetVal
End Function


'********************************************************************************
'********************************************************************************
'********************             MISCELLANEOUS             *********************
'********************************************************************************
'********************************************************************************

Public Function IsLeapYear(ByVal YearNum As Long) As Boolean
    Dim RetVal As Boolean
    RetVal = False
    If YearNum Mod 4 = 0 Then
        If YearNum Mod 100 = 0 Then
            If YearNum Mod 400 = 0 Then
                RetVal = True
            End If
        Else
            RetVal = True
        End If
    End If
    IsLeapYear = RetVal
End Function

Public Function Random(ByVal Low As Long, ByVal High As Long) As Long
    Random = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function IntegerToBoolean(ByVal MyNumber As Long) As Boolean
    IntegerToBoolean = (MyNumber <> 0)
End Function

Public Function BooleanToInteger(ByVal MyBool As Boolean) As Integer
    If MyBool = False Then
        BooleanToInteger = 0
    Else
        BooleanToInteger = 1
    End If
End Function

Public Function OpenLocation(ByVal WhichFilePath As String, Optional sParams As String = "", Optional sStartIn As String = vbNullString, Optional lngOpenMode As Long = 1) As Long
'Examples of usage:
'If OpenLocation("notepad.exe", "c:\myfile.txt", "", 1) < 32 Then
'    'Failed to open
'Else
'    'Opened
'End If
'OR:
'If OpenLocation("http://www.website.com", "", "", 1) < 32 Then
'    'Failed to open
'Else
'    'Opened
'End If
    OpenLocation = ShellExecute(0, "Open", WhichFilePath, sParams, sStartIn, lngOpenMode)
End Function

Function CompareVersions(ByVal VersionA As String, ByVal Comparison As String, ByVal VersionB As String) As Boolean
'Comparison can be:
'<
'=
'==
'>
'<=
'=<
'>=
'=>
'!=
'<>
'Version strings must only contain integers and floating points (example: 1.0.20.0003)
'1.02.3 = 1.02.03 = 1.02.003
'1.02.30 > 1.02.3
'Invalid versions (i.e. ones that contain non-float chars) will use string comparisons
    Dim BranchCountA As Long
    Dim BranchCountB As Long
    Dim BranchesA() As Long
    Dim BranchesB() As Long
    Dim DotPos As Long
    Dim Counter As Long
    Dim Difference As Integer
    If (StripNonFloat(VersionA) <> VersionA) Or (StripNonFloat(VersionB) <> VersionB) Then
        Select Case Comparison
        Case "<": CompareVersions = (VersionA < VersionB)
        Case ">": CompareVersions = (VersionA > VersionB)
        Case "=", "==": CompareVersions = (VersionA = VersionB)
        Case "<=", "=<": CompareVersions = (VersionA <= VersionB)
        Case ">=", "=>": CompareVersions = (VersionA >= VersionB)
        Case "!=", "<>": CompareVersions = (VersionA <> VersionB)
        End Select
    Else
        If Len(VersionA) = 0 Then VersionA = "0"
        If Len(VersionB) = 0 Then VersionB = "0"
        'VersionA
        BranchCountA = 1
        ReDim Preserve BranchesA(BranchCountA)
        Do
            DotPos = InStr(1, VersionA, ".")
            Select Case DotPos
            Case 0: BranchesA(BranchCountA) = Val(VersionA)
            Case 1: BranchesA(BranchCountA) = 0
            Case Else: BranchesA(BranchCountA) = Val(Mid$(VersionA, 1, DotPos - 1))
            End Select
            If DotPos <> 0 Then
                BranchCountA = BranchCountA + 1
                ReDim Preserve BranchesA(BranchCountA)
                VersionA = Mid$(VersionA, DotPos + 1)
            Else
                Exit Do
            End If
        Loop
        'VersionB
        BranchCountB = 1
        ReDim Preserve BranchesB(BranchCountB)
        Do
            DotPos = InStr(1, VersionB, ".")
            Select Case DotPos
            Case 0: BranchesB(BranchCountB) = Val(VersionB)
            Case 1: BranchesB(BranchCountB) = 0
            Case Else: BranchesB(BranchCountB) = Val(Mid$(VersionB, 1, DotPos - 1))
            End Select
            If DotPos <> 0 Then
                BranchCountB = BranchCountB + 1
                ReDim Preserve BranchesB(BranchCountB)
                VersionB = Mid$(VersionB, DotPos + 1)
            Else
                Exit Do
            End If
        Loop
        'balance branch numbers
        If BranchCountA > BranchCountB Then
            ReDim Preserve BranchesB(BranchCountA)
            For DotPos = (BranchCountB + 1) To BranchCountA
                BranchesB(DotPos) = 0
            Next DotPos
            BranchCountB = BranchCountA
        Else
            If BranchCountB > BranchCountA Then
                ReDim Preserve BranchesA(BranchCountB)
                For DotPos = (BranchCountA + 1) To BranchCountB
                    BranchesA(DotPos) = 0
                Next DotPos
                BranchCountA = BranchCountB
            End If
        End If
        'Compare
        For DotPos = 1 To BranchCountA
            Select Case True
            Case BranchesA(DotPos) = BranchesB(DotPos)
                Difference = 0
            Case BranchesA(DotPos) < BranchesB(DotPos)
                Difference = -1
                DotPos = BranchCountA
            Case BranchesA(DotPos) > BranchesB(DotPos)
                Difference = 1
                DotPos = BranchCountA
            End Select
        Next DotPos
        Select Case Comparison
        Case "<": CompareVersions = (Difference < 0)
        Case ">": CompareVersions = (Difference > 0)
        Case "=", "==": CompareVersions = (Difference = 0)
        Case "<=", "=<": CompareVersions = (Difference <= 0)
        Case ">=", "=>": CompareVersions = (Difference >= 0)
        Case "!=", "<>": CompareVersions = (Difference <> 0)
        End Select
    End If
End Function

Function HexToDec(HexValue As String) As String
    HexToDec = Val("&H" & HexValue)
End Function

Function HexToStr(ByRef strHex)
    Dim Length As Integer
    Dim Max As Integer
    Dim Str
    Max = Len(strHex)
    For Length = 1 To Max Step 2
        Str = Str & Chr("&h" & Mid(strHex, Length, 2))
    Next
    HexToStr = Str
End Function

'Needs: Private Declare Function GetAsyncKeyState Lib "USER32" (ByVal vKey As Long) As Integer
Function IsKeyPressed(ByVal vKey As Long) As Boolean
    'Private Const VK_LBUTTON = &H1
    Dim lKeyState As Long
    lKeyState = GetAsyncKeyState(vKey)
    IsKeyPressed = (lKeyState And &H8000)
End Function

'Needs: Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'Needs: Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Sub PlaySound(Optional ByVal FilePath As String = "", Optional ByVal ContinueExecution As Boolean = True)
'Sound played will be stopped by subsequent calls to this sub.
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_MEMORY = &H4
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10
    Const SND_PURGE = &H40
    Dim ExecVal As Long
    ExecVal = SND_NODEFAULT
    If ContinueExecution Then ExecVal = ExecVal Or SND_ASYNC
    If Len(FilePath) <> 0 Then
        Call APIPlaySound(FilePath, ByVal 0&, ExecVal)
    Else
        Call APIPlaySound(vbNullString, ByVal 0&, SND_ASYNC Or SND_PURGE)
    End If
End Sub

Function MoveToRecycleBin(ByVal FileSpec As String, Optional ByVal Confirm As Boolean = True, Optional ByVal FilesOnly As Boolean = False, Optional ShowErrorMessages As Boolean = False) As Boolean
Dim WinType_SFO As SHFILEOPSTRUCT
Dim lRet As Long
Dim lFlags As Long
    lFlags = FOF_ALLOWUNDO
    If Not Confirm Then lFlags = lFlags Or FOF_NOCONFIRMATION
    If FilesOnly Then lFlags = lFlags Or FOF_FILESONLY
    If Not ShowErrorMessages Then lFlags = lFlags Or FOF_NOERRORUI
    With WinType_SFO
        .wFunc = FO_DELETE
        .pFrom = FileSpec
        .fFlags = lFlags
    End With
    lRet = SHFileOperation(WinType_SFO)
    MoveToRecycleBin = (lRet = 0)
End Function

Function GetArgByName(ByVal ArgName As String, Optional ByVal CaseSensitive As Boolean = False) As String
'In [ c:\blah\blah.exe -arg1 one_word -arg2 -arg3 "two words" ],
'GetArgByName("arg1") = "one_word"
'GetArgByName("arg2") = "True"
'GetArgByName("arg3") = "two words"
'GetArgByName("-arg3") = "two words"
'GetArgByName("arg4") = ""
'GetArgByName("ARG3", True) = ""
    Dim ArgPos As Integer
    Dim RetVal As String
    Dim Counter As Integer
    Dim QuotedValue As Boolean
    RetVal = ""
    If Left(ArgName, 1) <> "-" Then ArgName = "-" & ArgName
    ArgPos = -1
    Do
        Select Case CaseSensitive
        Case True: ArgPos = InStrRev(UCase$(Command$), UCase$(ArgName), ArgPos)
        Case False: ArgPos = InStrRev(Command$, ArgName, ArgPos)
        End Select
        If ArgPos = 0 Then Exit Do 'because arg not found
        If (ArgPos + Len(ArgName)) > Len(Command$) Then
            'arg is present but has no value
            RetVal = "True"
        Else
            Select Case Mid$(Command$, ArgPos + Len(ArgName), 1)
            Case " "
                'get value
                Counter = 0
                'first, clear spaces between argname and argvalue
                Do
                    Counter = Counter + 1
                    If ArgPos + Len(ArgName) + Counter > Len(Command$) Then Exit Do 'because no more chars available
                    If Mid$(Command$, ArgPos + Len(ArgName) + Counter, 1) <> " " Then Exit Do
                Loop
                Counter = Counter - 1
                'now read in the value
                If ArgPos + Len(ArgName) + Counter + 1 <= Len(Command$) Then
                    'the first character of the value may alter behaviour of value fetch
                    Select Case Mid$(Command$, ArgPos + Len(ArgName) + Counter + 1, 1)
                    Case Chr(34) 'value is quoted so may contain spaces
                        QuotedValue = True
                        Counter = Counter + 1
                    Case "-" 'value is in fact another arg
                        RetVal = "True"
                        Exit Do
                    End Select
                    Do
                        Counter = Counter + 1
                        If ArgPos + Len(ArgName) + Counter > Len(Command$) Then Exit Do 'because no more chars available
                        Select Case Mid$(Command$, ArgPos + Len(ArgName) + Counter, 1)
                        Case Chr(34): If QuotedValue Then Exit Do
                        Case " ": If Not QuotedValue Then Exit Do
                        Case Else: RetVal = RetVal & Mid$(Command$, ArgPos + Len(ArgName) + Counter, 1)
                        End Select
                    Loop
                    Exit Do 'because we have a RetVal now
                Else
                    'arg is present but has no value (end of arg list reached before value found)
                    RetVal = "True"
                End If
            'Case Else
                'ArgName is only a substring of the actual arg found - we need to search again
            End Select
        End If
    Loop
    GetArgByName = RetVal
End Function

Function GetArgByNumber(ByVal ArgNumber As String, Optional ByVal OneBased As Boolean = False) As String
'In [ c:\blah\blah.exe one_word two words "two words" ],
'GetArgByNumber(0) = "one_word"
'GetArgByNumber(1) = "two"
'GetArgByNumber(2) = "words"
'GetArgByNumber(3) = "two words"
'GetArgByNumber(4) = ""
'In [ c:\blah\blah.exe -arg1 one_word hello -arg2 "two words" ],
'GetArgByNumber(0) = "-arg1"
'GetArgByNumber(1) = "one_word"
'GetArgByNumber(2) = "hello"
'GetArgByNumber(3) = "-arg2"
'GetArgByNumber(4) = "two words"
'The optional 'OneBased' boolean causes the routine to treat ArgNumber = 1 As ArgNumber = 0
    Dim RetVal As String
    Dim CharCount As Integer
    Dim QuotedValue As Boolean
    Dim ArgCount As Integer
    RetVal = ""
    ArgCount = -1
    CharCount = 0
    Do
        'first, count spaces to next arg
        Do
            CharCount = CharCount + 1
            If CharCount > Len(Command$) Then Exit Do 'because no more chars available
            If Mid$(Command$, CharCount, 1) <> " " Then Exit Do
        Loop
        If CharCount > Len(Command$) Then Exit Do 'because no more chars available
        ArgCount = ArgCount + 1
        'get the arg into RetVal
        CharCount = CharCount - 1
        If Mid$(Command$, CharCount + 1, 1) = Chr(34) Then
            QuotedValue = True
            CharCount = CharCount + 1
        Else
            QuotedValue = False
        End If
        Do
            CharCount = CharCount + 1
            If CharCount > Len(Command$) Then Exit Do
            Select Case Mid$(Command$, CharCount, 1)
            Case Chr(34): If QuotedValue Then Exit Do
            Case " ": If Not QuotedValue Then Exit Do
            End Select
            RetVal = RetVal & Mid$(Command$, CharCount, 1)
        Loop
        Select Case OneBased
        Case True: If (ArgCount - 1) = ArgNumber Then Exit Do
        Case False: If ArgCount = ArgNumber Then Exit Do
        End Select
        RetVal = ""
        QuotedValue = False
    Loop
    GetArgByNumber = RetVal
End Function

Function ResizeStringArray(ByRef vArray() As String, Optional ByVal iNewSize As Long = 0, Optional ByVal iIncrement As Long = 0, Optional ByVal iMaxSize As Long = -1) As Boolean
   'if iNewSize is negative then resizes array to abs(iNewSize)
   'if iNewSize is not negative then resizes array to max(iNewSize, ubound(array))
   'also adds iIncrement
   'caps new size at iMaxSize (if iMaxSize=-1 then no capping will take place)
   'returns false if had to be capped, true otherwise
    If iNewSize < 0 Then
        iNewSize = Abs(iNewSize)
    Else
        If iNewSize < UBound(vArray()) Then
            iNewSize = UBound(vArray())
        Else
            iNewSize = iNewSize + iIncrement
        End If
    End If
    If iNewSize = UBound(vArray()) Then
        ResizeStringArray = True
        'no resizing to be done
    Else
        If iMaxSize <> -1 Then
            If iNewSize < iMaxSize Then
                ResizeStringArray = True
            Else
                iNewSize = iMaxSize
                ResizeStringArray = False
            End If
        Else
            ResizeStringArray = True
        End If
        If UBound(vArray()) <> iNewSize Then ReDim Preserve vArray(iNewSize)
    End If
End Function

Function ResizeBooleanArray(ByRef vArray() As Boolean, Optional ByVal iNewSize As Long = 0, Optional ByVal iIncrement As Long = 0, Optional ByVal iMaxSize As Long = -1) As Boolean
   'see ResizeStringArray
    If iNewSize < 0 Then
        iNewSize = Abs(iNewSize)
    Else
        If iNewSize < UBound(vArray()) Then
            iNewSize = UBound(vArray())
        Else
            iNewSize = iNewSize + iIncrement
        End If
    End If
    If iNewSize = UBound(vArray()) Then
        ResizeBooleanArray = True
        'no resizing to be done
    Else
        If iMaxSize <> -1 Then
            If iNewSize < iMaxSize Then
                ResizeBooleanArray = True
            Else
                iNewSize = iMaxSize
                ResizeBooleanArray = False
            End If
        Else
            ResizeBooleanArray = True
        End If
        If UBound(vArray()) <> iNewSize Then ReDim Preserve vArray(iNewSize)
    End If
End Function

Sub SortListViewByDatasize(ByRef lvSort As ListView, ByVal iColumn As Integer)
    'assumes datasize is stored in Tag
    Dim i As Long
    Dim iMaxLen As Long
    'get the longest number string length
    iMaxLen = 0
    i = 1
    Do While i <= lvSort.ListItems.Count
        If Len(lvSort.ListItems(i).Tag) <> 0 Then 'don't bother with zero length strings
            If Len(lvSort.ListItems(i).Tag) > iMaxLen Then
                iMaxLen = Len(lvSort.ListItems(i).Tag)
            End If
        End If
        i = i + 1
    Loop
    lvSort.Visible = False 'improve performance
    lvSort.Sorted = False
    'pad with zeroes
    i = 1
    Do While i <= lvSort.ListItems.Count
        If iColumn <> 0 Then
            lvSort.ListItems(i).SubItems(iColumn) = PadString(lvSort.ListItems(i).Tag, "0", iMaxLen)
        Else
            lvSort.ListItems(i).Text = PadString(lvSort.ListItems(i).Tag, "0", iMaxLen)
        End If
        i = i + 1
    Loop
    lvSort.SortKey = iColumn
    lvSort.Sorted = True
    lvSort.Sorted = False
    'set to datasize string
    i = 1
    Do While i <= lvSort.ListItems.Count
        If iColumn <> 0 Then
            lvSort.ListItems(i).SubItems(iColumn) = DataSize(Val(lvSort.ListItems(i).Tag))
        Else
            lvSort.ListItems(i).Text = DataSize(Val(lvSort.ListItems(i).Tag))
        End If
        i = i + 1
    Loop
    lvSort.Visible = True
End Sub

'Function OldFileCopyByChunk(ByVal sSourceFile As String, ByVal sDestFile As String, ByVal frmProgress As Form, Optional ByVal iChunk As Currency, Optional ByVal ReturnNegativeOnError As Boolean = True) As Currency
'    'requires LargeFileReadWrite module
'    Dim iSize As Currency
'    Dim iCopied As Currency
'    Dim iRemainder As Currency
'    Dim cBuffer() As Byte
'    Dim iSourceFile As Integer
'    Dim iDestFile As Integer
'    Dim iSourceAttr As Long
'    Dim iDestAttr As Long
'    Dim iTempAttr As Long
'    On Error GoTo FileCopyByChunk_Error
'    iSourceFile = FreeFile()
'    Open sSourceFile For Binary Access Read As #iSourceFile
'    If FileExists(sDestFile) Then
'        Call SetAttr(sDestFile, vbNormal) 'make sure not hidden or read only. we'll clone the source file's attributes at the end anyway
'    End If
'    iDestFile = FreeFile()
'    Open sDestFile For Output As #iDestFile
'    Close #iDestFile 'now we've created and initialised the file
'    iSize = GetFileSize(sSourceFile)
'    If iSize > 0 Then
'        Open sDestFile For Binary Access Write As #iDestFile
'        If iChunk = 0 Then iChunk = 1048576 '1MB
'        iRemainder = iSize - (Int((iSize / iChunk)) * iChunk) 'Mod doesn't work with doubles
'        If iRemainder <> 0 Then
'            ReDim cBuffer(1 To iRemainder)
'            Get #iSourceFile, , cBuffer
'            Put #iDestFile, , cBuffer
'            iCopied = iRemainder
'            Call frmProgress.FileCopyByChunk_Progress(iCopied, iSize, iRemainder)
'        End If
'        ReDim cBuffer(1 To iChunk)
'        Do While iCopied < iSize
'            Get #iSourceFile, , cBuffer
'            Put #iDestFile, , cBuffer
'            iCopied = iCopied + iChunk
'            Call frmProgress.FileCopyByChunk_Progress(iCopied, iSize, iChunk)
'        Loop
'        Close #iDestFile
'        FileCopyByChunk = GetFileSize(sDestFile)
'        If ReturnNegativeOnError Then
'            If FileCopyByChunk <> iSize Then
'                FileCopyByChunk = -FileCopyByChunk
'            End If
'        End If
'        Call CloneFileProperties(sSourceFile, sDestFile)
'    Else
'        FileCopyByChunk = 0
'    End If
'FileCopyByChunk_Exit:
'    Close #iSourceFile
'    Exit Function
'FileCopyByChunk_Error:
'   FileCopyByChunk = -1
'   Resume FileCopyByChunk_Exit
'End Function

Function FileCopyByChunk(ByVal sSourceFile As String, ByVal sDestFile As String, ByVal frmProgress As Form, Optional ByVal iChunk As Long, Optional ByVal ReturnNegativeOnError As Boolean = True) As Currency
    'requires LargeFileReadWrite module
    Dim iSourceSize As Currency
    Dim iDestSize As Currency
    Dim iCopied As Currency
    Dim iWritten As Long
    Dim iRemainder As Long
    Dim cBuffer() As Byte
    Dim iSourceFile As Integer
    Dim iDestFile As Integer
    'On Error GoTo FileCopyByChunk_Error
    If FileExists(sDestFile) Then
        Call SetAttr(sDestFile, vbNormal) 'make sure not hidden or read only. we'll clone the source file's attributes at the end anyway
    End If
    iDestFile = FreeFile()
    Open sDestFile For Output As #iDestFile
    Close #iDestFile 'now we've created and initialised the file
    iSourceFile = API_OpenFile(sSourceFile, iSourceSize, True)
    If iSourceSize > 0 Then
        iDestFile = API_OpenFile(sDestFile, iDestSize)
        If iChunk = 0 Then iChunk = 1048576 '1MB
        iRemainder = iSourceSize - (Int((iSourceSize / iChunk)) * iChunk) 'Mod doesn't work with doubles
        iCopied = 0
        If iRemainder <> 0 Then
            ReDim cBuffer(0 To (iRemainder - 1))
            iWritten = iRemainder
            Call API_ReadFile(iSourceFile, iCopied, iWritten, cBuffer())
            If iWritten <> iRemainder Then Err.Raise 51
            Call API_WriteFile(iDestFile, iCopied, iWritten, cBuffer())
            iCopied = iWritten
            If iWritten <> iRemainder Then Err.Raise 51
            Call frmProgress.FileCopyByChunk_Progress(iCopied, iSourceSize, iRemainder)
        End If
        ReDim cBuffer(0 To (iChunk - 1))
        iWritten = iChunk
        Do While iCopied < iSourceSize
            Call API_ReadFile(iSourceFile, iCopied, iWritten, cBuffer())
            If iWritten <> iChunk Then Err.Raise 51
            Call API_WriteFile(iDestFile, iCopied, iWritten, cBuffer())
            iCopied = iCopied + iWritten
            If iWritten <> iChunk Then Err.Raise 51
            Call frmProgress.FileCopyByChunk_Progress(iCopied, iSourceSize, iChunk)
        Loop
        Call API_FileSize(iDestFile, iDestSize)
        Call API_CloseFile(iDestFile)
        FileCopyByChunk = iDestSize
        If ReturnNegativeOnError Then
            If FileCopyByChunk <> iSourceSize Then
                FileCopyByChunk = -FileCopyByChunk
            End If
        End If
    Else
        FileCopyByChunk = 0
    End If
    Call API_CloseFile(iSourceFile)
    iSourceFile = -1
    Call CloneFileProperties(sSourceFile, sDestFile)
FileCopyByChunk_Exit:
    If iSourceFile <> -1 Then Call API_CloseFile(iSourceFile)
    Exit Function
FileCopyByChunk_Error:
   FileCopyByChunk = -1
   Resume FileCopyByChunk_Exit
End Function

Public Function CloneFileProperties(ByVal sSourceFile As String, ByVal sDestFile As String) As Boolean
    Dim iFile As Long
    Dim Created As FILETIME
    Dim Accessed As FILETIME
    Dim Modified As FILETIME
    Dim iAttr As Long
    CloneFileProperties = False
    iFile = CreateFile(sSourceFile, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If iFile = 0 Then Exit Function
    If GetFileTime(iFile, Created, Accessed, Modified) = 0 Then
        Call CloseHandle(iFile)
        Exit Function
    End If
    If CloseHandle(iFile) = 0 Then Exit Function
    iAttr = GetAttr(sDestFile)
    Call SetAttr(sDestFile, vbNormal)
    iFile = CreateFile(sDestFile, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If iFile = 0 Then
        Call SetAttr(sDestFile, iAttr)
        Exit Function
    End If
    If SetFileTime(iFile, Created, Accessed, Modified) = 0 Then
        Call SetAttr(sDestFile, iAttr)
        Call CloseHandle(iFile)
        Exit Function
    End If
    If CloseHandle(iFile) = 0 Then Exit Function
    Call SetAttr(sDestFile, GetAttr(sSourceFile))
    CloneFileProperties = True
End Function
