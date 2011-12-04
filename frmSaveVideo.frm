VERSION 5.00
Begin VB.Form frmSaveVideo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base: Save Video"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   Icon            =   "frmSaveVideo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVideoDescription 
      Height          =   285
      Left            =   120
      MaxLength       =   64
      TabIndex        =   1
      Text            =   "No description"
      Top             =   720
      Width           =   3975
   End
   Begin VB.CommandButton cmdOKOptions 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelOptions 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Click OK to save the video, or click Cancel to delete it."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter a short description for the recorded video."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmSaveVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CancelVideoSave As Boolean

Private Sub cmdCancelOptions_Click()
    CancelVideoSave = True
    frmSaveVideo.Hide
End Sub

Private Sub cmdOKOptions_Click()
    CancelVideoSave = False
    frmSaveVideo.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        Cancel = 1
        Call cmdCancelOptions_Click
    End If
End Sub

Private Sub txtVideoDescription_Change()
    Dim TempString As String
    TempString = StripInvalidChars(txtVideoDescription.Text, InvalidFileChars)
    If TempString <> txtVideoDescription.Text Then
        Call MsgBox("A filename cannot contain any of the following characters: " & InvalidFileChars, vbOKOnly + vbInformation, App.Title)
        txtVideoDescription.Text = TempString
        SendKeys "{end}"
    End If
End Sub
