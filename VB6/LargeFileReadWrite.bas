Attribute VB_Name = "LargeFileReadWrite"
Option Explicit
Private Const MOVEFILE_REPLACE_EXISTING = &H1
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_BEGIN = 0
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const CREATE_NEW = 1
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALLWAYS = 4
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long

Public Function API_OpenFile(ByVal FileName As String, ByRef FileSize As Currency, Optional ByVal bReadOnly As Boolean = False) As Long
    Dim FileH As Long
    Dim Ret As Long
    On Error Resume Next
    If bReadOnly Then
        FileH = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ, 0&, OPEN_ALLWAYS, 0, 0)
    Else
        FileH = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_ALLWAYS, 0, 0)
    End If
    If Err.Number > 0 Then
        Err.Clear
        API_OpenFile = -1
    Else
        API_OpenFile = FileH
        Ret = SetFilePointer(FileH, 0, 0, FILE_BEGIN)
        API_FileSize FileH, FileSize
    End If
    On Error GoTo 0
    If FileH = -1 Then Err.Raise 51
End Function

Public Sub API_FileSize(ByVal filenumber As Long, ByRef FileSize As Currency)
    Dim FileSizeL As Long
    Dim FileSizeH As Long
    FileSizeH = 0
    FileSizeL = GetFileSize(filenumber, FileSizeH)
    Long2Size FileSizeL, FileSizeH, FileSize
End Sub

Public Sub API_ReadFile(ByVal filenumber As Long, ByVal Position As Currency, ByRef BlockSize As Long, ByRef data() As Byte)
    Dim PosL As Long
    Dim PosH As Long
    Dim SizeRead As Long
    Dim Ret As Long
    Size2Long Position, PosL, PosH
    Ret = SetFilePointer(filenumber, PosL, PosH, FILE_BEGIN)
    Ret = ReadFile(filenumber, data(0), BlockSize, SizeRead, 0&)
    BlockSize = SizeRead
End Sub

Public Sub API_CloseFile(ByVal filenumber As Long)
    Dim Ret As Long
    Ret = CloseHandle(filenumber)
End Sub

Public Sub API_WriteFile(ByVal filenumber As Long, ByVal Position As Currency, ByRef BlockSize As Long, ByRef data() As Byte)
    Dim PosL As Long
    Dim PosH As Long
    Dim SizeWrit As Long
    Dim Ret As Long
    Size2Long Position, PosL, PosH
    Ret = SetFilePointer(filenumber, PosL, PosH, FILE_BEGIN)
    Ret = WriteFile(filenumber, data(0), BlockSize, SizeWrit, 0&)
    BlockSize = SizeWrit
End Sub

Private Sub Size2Long(ByVal FileSize As Currency, ByRef LongLow As Long, ByRef LongHigh As Long)
    '&HFFFFFFFF unsigned = 4294967295
    Dim Cutoff As Currency
    Cutoff = 2147483647
    Cutoff = Cutoff + 2147483647
    Cutoff = Cutoff + 1 ' now we hold the value of 4294967295 and not -1
    LongHigh = 0
    Do Until FileSize < Cutoff
        LongHigh = LongHigh + 1
        FileSize = FileSize - Cutoff
    Loop
    If FileSize > 2147483647 Then
        LongLow = -CLng(Cutoff - (FileSize - 1))
    Else
        LongLow = CLng(FileSize)
    End If
End Sub

Private Sub Long2Size(ByVal LongLow As Long, ByVal LongHigh As Long, ByRef FileSize As Currency)
    '&HFFFFFFFF unsigned = 4294967295
    Dim Cutoff As Currency
    Cutoff = 2147483647
    Cutoff = Cutoff + 2147483647
    Cutoff = Cutoff + 1 ' now we hold the value of 4294967295 and not -1
    FileSize = Cutoff * LongHigh
    If LongLow < 0 Then
        FileSize = FileSize + (Cutoff + (LongLow + 1))
    Else
        FileSize = FileSize + LongLow
    End If
End Sub
