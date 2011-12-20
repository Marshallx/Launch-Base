Attribute VB_Name = "MarshallxFileAssociation"
Option Explicit

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

'Return codes from Registration functions.
Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1&
Const ERROR_BADKEY = 2&
Const ERROR_CANTOPEN = 3&
Const ERROR_CANTREAD = 4&
Const ERROR_CANTWRITE = 5&
Const ERROR_OUTOFMEMORY = 6&
Const ERROR_INVALID_PARAMETER = 7&
Const ERROR_ACCESS_DENIED = 8&

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const MAX_PATH = 260&
Private Const REG_SZ = 1

Public Sub AssociateFileType(ByVal FileExtension As String, ByVal AppID As String, ByVal FileTypeName As String, ByVal ProgramExecutable As String, Optional ByVal DefaultIconPath As String = "", Optional ByVal DefaultIconNum As Integer = 0)
'FileExtension="aBc" registers ".abc"
'FileExtension=".aBc" registers ".aBc"
    Dim ret& 'Holds error status if any from API calls.
    Dim lphKey& 'Holds created key handle from RegCreateKey.
    If FileExtension <> "" Then
        If Left$(FileExtension, 1) <> "." Then FileExtension = "." & LCase$(FileExtension)
        'This creates a Root entry called AppID.
        ret& = RegCreateKey&(HKEY_CLASSES_ROOT, AppID, lphKey&)
        ret& = RegSetValue&(lphKey&, "", REG_SZ, FileTypeName, 0&)
        'This creates a Root entry called FileExtension associated with AppID.
        ret& = RegCreateKey&(HKEY_CLASSES_ROOT, FileExtension, lphKey&)
        ret& = RegSetValue&(lphKey&, "", REG_SZ, AppID, 0&)
        'This sets the command line for AppID.
        ret& = RegCreateKey&(HKEY_CLASSES_ROOT, AppID, lphKey&)
        ret& = RegSetValue&(lphKey&, "shell\open\command", REG_SZ, ProgramExecutable & " " & Chr(34) & "%1" & Chr(34), MAX_PATH)
        'Set default icon for this file type
        ret& = RegCreateKey&(HKEY_CLASSES_ROOT, AppID, lphKey&)
        ret& = RegSetValue&(lphKey&, "DefaultIcon", REG_SZ, DefaultIconPath & "," & CStr(DefaultIconNum), MAX_PATH)
    End If
End Sub

