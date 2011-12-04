VERSION 5.00
Begin VB.Form frmDisclaimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base: Disclaimer"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8145
   Icon            =   "frmDisclaimer.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Quit"
      Default         =   -1  'True
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   2640
      Width           =   3825
   End
   Begin VB.Label Label2 
      Caption         =   $"frmDisclaimer.frx":0E42
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   7935
   End
   Begin VB.Label Label3 
      Caption         =   $"frmDisclaimer.frx":0F34
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   $"frmDisclaimer.frx":114D
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   7935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   120
      X2              =   8040
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   8040
      Y1              =   2520
      Y2              =   2520
   End
End
Attribute VB_Name = "frmDisclaimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Call frmMain.Shutdown
    Unload frmSplash
    Unload Me
    End
End Sub
