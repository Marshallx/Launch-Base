VERSION 5.00
Begin VB.Form frmFileErrors 
   Caption         =   "Launch Base Mod Creator: File Errors"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileErrors 
      Height          =   4215
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
End
Attribute VB_Name = "frmFileErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Temp As Long
    Temp = Val(ReadINIStr("General", "FileErrorsTop", frmMain.ProgramINI, "0"))
    If Temp <> 0 Then Me.Top = Temp
    Temp = Val(ReadINIStr("General", "FileErrorsLeft", frmMain.ProgramINI, "0"))
    If Temp <> 0 Then Me.Left = Temp
    Temp = Val(ReadINIStr("General", "FileErrorsHeight", frmMain.ProgramINI, "0"))
    If Temp <> 0 Then Me.Height = Temp
    Temp = Val(ReadINIStr("General", "FileErrorsWidth", frmMain.ProgramINI, "0"))
    If Temp <> 0 Then Me.Width = Temp
    Call Form_Resize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call WriteINIStr("General", "FileErrorsTop", CStr(Me.Top), frmMain.ProgramINI)
    Call WriteINIStr("General", "FileErrorsLeft", CStr(Me.Left), frmMain.ProgramINI)
    Call WriteINIStr("General", "FileErrorsWidth", CStr(Me.Width), frmMain.ProgramINI)
    Call WriteINIStr("General", "FileErrorsHeight", CStr(Me.Height), frmMain.ProgramINI)
    If UnloadMode <> 1 Then '0 = user clicked X, 1 = programatically, 2 = Windows
        Cancel = 1
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    txtFileErrors.Width = Me.ScaleWidth
    txtFileErrors.Height = Me.ScaleHeight
    txtFileErrors.SelStart = Len(txtFileErrors.Text)
End Sub

Private Sub txtFileErrors_Change()
    txtFileErrors.SelStart = Len(txtFileErrors.Text)
End Sub

Private Sub txtFileErrors_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = vbCtrlMask Then
       Clipboard.Clear
       Clipboard.SetText txtFileErrors.SelText
    End If
End Sub
