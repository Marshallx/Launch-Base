VERSION 5.00
Begin VB.Form frmLiveLog 
   Caption         =   "Launch Base: LiveLog"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmLiveLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLiveLog 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   3135
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmLiveLog.frx":000C
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmLiveLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Temp As Long
    Temp = Val(ReadINIStr("LiveLog", "LiveLogTop", ProgramINI, "0"))
    If Temp <> 0 Then Me.Top = Temp
    Temp = Val(ReadINIStr("LiveLog", "LiveLogLeft", ProgramINI, "0"))
    If Temp <> 0 Then Me.Left = Temp
    Temp = Val(ReadINIStr("LiveLog", "LiveLogHeight", ProgramINI, "0"))
    If Temp <> 0 Then Me.Height = Temp
    Temp = Val(ReadINIStr("LiveLog", "LiveLogWidth", ProgramINI, "0"))
    If Temp <> 0 Then Me.Width = Temp
    Call Form_Resize
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then Cancel = 1
    If UnloadMode = 0 Then
        frmMain.menu_livelog.Checked = False
        Call Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    txtLiveLog.Width = Me.ScaleWidth
    txtLiveLog.Height = Me.ScaleHeight
    txtLiveLog.SelStart = Len(txtLiveLog.Text)
End Sub

Private Sub txtLiveLog_Change()
    txtLiveLog.SelStart = Len(txtLiveLog.Text)
End Sub

Private Sub txtLiveLog_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = vbCtrlMask Then
       Clipboard.Clear
       Clipboard.SetText txtLiveLog.SelText
    End If
End Sub
