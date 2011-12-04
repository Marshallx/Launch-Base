VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Launch Base: Help Topics"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   Begin VB.CommandButton cmdSaveHelp 
      Caption         =   "Save This Page"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton cmdHelpPage 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.TextBox txtHelpText 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   1980
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10845
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelpPage_Click(Index As Integer)
    Dim FileHandle As Integer
    Dim buffer As String
    Dim iSize As Long
    buffer = JoinPath(JoinPath(App.Path, "Help"), "help" & CStr(Index) & ".txt")
    FileHandle = FreeFile
    txtHelpText.Text = ""
    If FileExists(buffer) Then
        txtHelpText.Tag = buffer
        cmdSaveHelp.Tag = CStr(Index)
        'All at once
        Open buffer For Binary Access Read Shared As FileHandle
        iSize = LOF(FileHandle)
        buffer = Space$(iSize)
        Get #FileHandle, , buffer
        Close #FileHandle
        iSize = InStr(1, buffer, vbCrLf)
        cmdHelpPage(Index).Caption = Left$(buffer, iSize - 1)
        txtHelpText.Text = Mid$(buffer, iSize + 2)
        'Line by line
        'Open Buffer For Input As FileHandle
        'Line Input #FileHandle, Buffer
        'cmdHelpPage(Index).Caption = Buffer
        'Do While Not EOF(FileHandle)
        '    Line Input #FileHandle, Buffer
        '    txtHelpText.Text = txtHelpText.Text & Buffer & vbCrLf
        'Loop
        'Close #FileHandle
    Else
        cmdHelpPage(Index).Caption = "MISSING:" & GetFileName(buffer)
        buffer = "Missing help file: " & Quote(buffer)
        Call frmMain.WriteLogEntry(buffer)
        Call MsgBox(buffer, vbOKOnly + vbExclamation, App.Title)
    End If
    For FileHandle = 0 To cmdHelpPage.UBound
        If FileHandle <> Index Then
            cmdHelpPage(FileHandle).FontBold = False
        Else
            cmdHelpPage(FileHandle).FontBold = True
        End If
    Next FileHandle
End Sub

Private Sub cmdSaveHelp_Click()
    Dim FileHandle As Integer
    FileHandle = FreeFile
    Open txtHelpText.Tag For Output As FileHandle
        Print #FileHandle, cmdHelpPage(Val(cmdSaveHelp.Tag)).Caption
        Print #FileHandle, txtHelpText.Text;
    Close #1
End Sub

Private Sub Form_Load()
    Dim LocalString As String
    Dim Counter As Integer
    Dim FileHandle As Integer
    Counter = 0
    LocalString = JoinPath(JoinPath(App.Path, "Help"), "help" & CStr(Counter) & ".txt")
    FileHandle = FreeFile
    Do While FileExists(LocalString)
        Open LocalString For Input As FileHandle
            Line Input #FileHandle, LocalString
        Close #FileHandle
        If Counter <> 0 Then
            Load cmdHelpPage(Counter)
            cmdHelpPage(Counter).Top = cmdHelpPage(0).Top + (cmdHelpPage(0).Height * Counter)
            cmdHelpPage(Counter).Left = cmdHelpPage(0).Left
        End If
        cmdHelpPage(Counter).Caption = LocalString
        cmdHelpPage(Counter).Visible = True
        Counter = Counter + 1
        LocalString = JoinPath(JoinPath(App.Path, "Help"), "help" & CStr(Counter) & ".txt")
    Loop
    Call cmdHelpPage_Click(0)
    If Len(GetArgByName("edithelp")) <> 0 Then
        txtHelpText.Locked = False
        cmdSaveHelp.Visible = True
    End If
End Sub

Private Sub Form_Resize()
    If frmHelp.WindowState <> vbMinimized Then
        txtHelpText.Width = frmHelp.ScaleWidth - cmdHelpPage(0).Width
        txtHelpText.Height = frmHelp.ScaleHeight
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If frmMain.Visible = True Then
        If frmMain.Enabled = True Then frmMain.SetFocus
    ElseIf frmModCat.Visible = True Then
        If frmModCat.Enabled = True Then frmModCat.SetFocus
    ElseIf frmHistory.Visible = True Then
        If frmHistory.Enabled = True Then frmHistory.SetFocus
    End If
End Sub

Private Sub txtHelpText_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 67 And Shift = vbCtrlMask Then
       Clipboard.Clear
       Clipboard.SetText txtHelpText.SelText
    End If
End Sub

Private Sub txtHelpText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Screen.ActiveForm Is frmHelp Then Call txtHelpText.SetFocus
End Sub
