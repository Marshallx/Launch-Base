VERSION 5.00
Begin VB.Form frmDownloading 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Downloading file..."
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2880
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Download"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label lblDownloadKB 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   2880
   End
   Begin VB.Label lblDownloadPercent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Shape pbarDownloading 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   15
   End
End
Attribute VB_Name = "frmDownloading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function DownloadProgress(ByVal iBytesReceived As Long, ByVal iBytesExpected As Long)
    Dim DownloadPercent As Long
    If iBytesExpected = 0 Then
        DownloadPercent = 0
        lblDownloadKB.Caption = DataSize(iBytesReceived, , , False) & " / ? KB "
    Else
        DownloadPercent = (iBytesReceived * 100) / iBytesExpected
        lblDownloadKB.Caption = DataSize(iBytesReceived, , , False) & " / " & DataSize(iBytesExpected) & " "
    End If
    If DownloadPercent = 0 Then
        pbarDownloading.Width = 0
    Else
        pbarDownloading.Width = frmDownloading.ScaleWidth / (100 / DownloadPercent)
    End If
    If DownloadPercent <= 50 Then
        lblDownloadPercent.ForeColor = &H80000012
    Else
        lblDownloadPercent.ForeColor = &H8000000E
    End If
    lblDownloadPercent.Caption = CStr(DownloadPercent) & "%"
    Call Me.Refresh
End Function

Private Sub cmdCancel_Click()
    Call theInternet.CancelDownload
    cmdCancel.Enabled = False
End Sub

Private Sub Form_Load()
    If OptLBSounds Then Call PlaySound(JoinPath(RESDIR, "eva_dl.wav"), True)
End Sub
