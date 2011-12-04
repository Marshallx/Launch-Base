VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Launch Base"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Launch Base..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4695
   End
   Begin VB.Shape pbarSplash 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   30
      Left            =   45
      Top             =   3550
      Width           =   4710
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public pbarWidth As Integer
Public pbarWidthPerUnit As Integer

Private Sub Form_Load()
    Set frmSplash.Picture = LoadPicture(JoinPath(EXEDIR, "splash.bmp"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSplash.Picture = Nothing
End Sub

Public Sub PROGRESS(Optional ByVal sMsg As String = "", Optional ByVal Units As Integer = 1)
    pbarWidth = pbarWidth + (pbarWidthPerUnit * Units)
    pbarSplash.Width = pbarWidth
    pbarSplash.Refresh
    If Len(sMsg) <> 0 Then
        lblProgress.Caption = sMsg
        lblProgress.Refresh
    End If
End Sub
