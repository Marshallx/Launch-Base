VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUninstallFiles 
   Caption         =   "Launch Base Mod Creator: Uninstall Files"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvUninstallFiles 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   3625
      View            =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Menu menu_rc 
      Caption         =   "menu_rc"
      Visible         =   0   'False
      Begin VB.Menu menu_insert 
         Caption         =   "Add File To Uninstall List"
      End
   End
End
Attribute VB_Name = "frmUninstallFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_GETITEMHEIGHT = &H1A1

Private Sub Form_Load()
    Dim Temp As Long
    Temp = Val(ReadINIStr("General", "UninstallFilesTop", frmMain.ProgramINI, "0"))
    If Temp <> 0 Then Me.Top = Temp
    Temp = Val(ReadINIStr("General", "UninstallFilesLeft", frmMain.ProgramINI, "0"))
    If Temp <> 0 Then Me.Left = Temp
    Temp = Val(ReadINIStr("General", "UninstallFilesHeight", frmMain.ProgramINI, "0"))
    If Temp <> 0 Then Me.Height = Temp
    Temp = Val(ReadINIStr("General", "UninstallFilesWidth", frmMain.ProgramINI, "0"))
    If Temp <> 0 Then Me.Width = Temp
    Call Form_Resize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call WriteINIStr("General", "UninstallFilesTop", CStr(Me.Top), frmMain.ProgramINI)
    Call WriteINIStr("General", "UninstallFilesLeft", CStr(Me.Left), frmMain.ProgramINI)
    Call WriteINIStr("General", "UninstallFilesWidth", CStr(Me.Width), frmMain.ProgramINI)
    Call WriteINIStr("General", "UninstallFilesHeight", CStr(Me.Height), frmMain.ProgramINI)
    If UnloadMode <> 1 Then '0 = user clicked X, 1 = programatically, 2 = Windows
        Cancel = 1
        Call Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    lvUninstallFiles.Width = Me.ScaleWidth
    lvUninstallFiles.Height = Me.ScaleHeight
End Sub

Private Sub lvUninstallFiles_AfterLabelEdit(Cancel As Integer, NewString As String)
    If Len(NewString) <> 0 Then
        If Len(NewString) <> Len(StripInvalidChars(NewString, StripInvalidChars(InvalidFileChars, "\"))) Then
            Cancel = 1
            Call MsgBox("File names cannot contain any of the following characters: " & Mid$(StripInvalidChars(InvalidFileChars, "\"), 2), vbOKOnly + vbInformation, "Invalid Character")
        End If
    Else
        Cancel = 1
        Call MsgBox("File names cannot be blank.", vbOKOnly + vbInformation, "Invalid Character")
    End If
End Sub

Private Sub lvUninstallFiles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If Not lvUninstallFiles.SelectedItem Is Nothing Then
            Call lvUninstallFiles.ListItems.Remove(lvUninstallFiles.SelectedItem.Index)
        End If
    End If
End Sub

Private Sub lvUninstallFiles_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim X As Single
    Dim Y As Single
    Dim ListItemHeight As Long
    If KeyCode = 93 Then
        ListItemHeight = SendMessage(lvUninstallFiles.hWnd, LB_GETITEMHEIGHT, 0, ByVal 0&)
        X = lvUninstallFiles.Left + (64 * 15)
        Y = lvUninstallFiles.Top + ((ListItemHeight * 15) * lvUninstallFiles(Index).ListIndex) + ((ListItemHeight * 15) / 2)
        Call lvUninstallFiles_RightClick(X, Y)
    End If
End Sub

Private Sub lvUninstallFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call lvUninstallFiles_RightClick(X, Y)
End Sub

Private Sub lvUninstallFiles_RightClick(ByVal X As Single, ByVal Y As Single)
    Call PopupMenu(menu_rc, 2, X, Y)
End Sub

Private Sub menu_insert_Click()
    Call lvUninstallFiles.ListItems.Add(, , "example.txt")
End Sub
