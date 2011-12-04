VERSION 5.00
Begin VB.Form frmSaveExcept 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base: Save except.txt"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   ControlBox      =   0   'False
   Icon            =   "frmSaveExcept.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4185
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtExceptDescription 
      Height          =   285
      Left            =   120
      MaxLength       =   64
      TabIndex        =   1
      Text            =   "No description"
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton cmdOKOptions 
      Caption         =   "OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "It will be saved in the Launch Base 'Logs' folder."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   $"frmSaveExcept.frx":000C
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3990
   End
End
Attribute VB_Name = "frmSaveExcept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CancelExceptSave As Boolean

Private Sub cmdCancelOptions_Click()
    CancelExceptSave = True
    frmSaveExcept.Hide
End Sub

Private Sub cmdOKOptions_Click()
    CancelExceptSave = False
    frmSaveExcept.Hide
End Sub

Private Sub Form_Load()
    If BooleanStringToBoolean(ReadINIStr("Restore", "ModifiedEXE", ProgramINI)) Then Call MsgBox("The Internal Error you encountered whilst running the game may have been the result of a modified executable in your Red Alert 2 directory." & vbCrLf & "It is strongly recommended that you remove any third-party patches.", vbOKOnly + vbInformation, App.Title)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        Cancel = 1
        Call cmdCancelOptions_Click
    End If
End Sub

Private Sub txtExceptDescription_Change()
    Dim TempString As String
    TempString = StripInvalidChars(txtExceptDescription.Text, InvalidFileChars)
    If TempString <> txtExceptDescription.Text Then
        Call MsgBox("A filename cannot contain any of the following characters: " & InvalidFileChars, vbOKOnly + vbInformation, App.Title)
        txtExceptDescription.Text = TempString
        SendKeys "{end}"
    End If
End Sub
