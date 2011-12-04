VERSION 5.00
Begin VB.Form frmDisclaimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base Mod Creator: Disclaimer"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4410
   Icon            =   "frmDisclaimer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      Default         =   -1  'True
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2025
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "I Agree"
      Height          =   345
      Left            =   2280
      TabIndex        =   0
      Top             =   1200
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   $"frmDisclaimer.frx":0E42
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   120
      X2              =   4320
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4320
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmDisclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    frmDisclaimer.Tag = "1"
End Sub

Private Sub cmdCancel_Click()
    frmDisclaimer.Tag = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmDisclaimer.Tag = "0"
    Cancel = 1
End Sub
