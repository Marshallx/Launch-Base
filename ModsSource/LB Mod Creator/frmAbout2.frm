VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base Mod Creator: About"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "frmAbout2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4440
      TabIndex        =   1
      Top             =   3120
      Width           =   1065
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout2.frx":0E42
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label8 
      Caption         =   "Koen van de Sande - File patching system (VPatch)"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Label Label7 
      Caption         =   "Nullsoft - Nullsoft Scriptable Install System (NSIS)"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Label Label6 
      Caption         =   "Everyone at Renegade Projects' forums - support and suggestions."
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "DCoder - DLL file for MIX/BAG/CSF file writing."
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   5175
   End
   Begin VB.Label Label4 
      Caption         =   "Marshall - Concept, design and programming."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   5175
   End
   Begin VB.Label Label2 
      Caption         =   $"frmAbout2.frx":1C84
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Credits:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   5175
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      Caption         =   "Last updated: 2010-09-11"
      Height          =   225
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   2325
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5520
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5535
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version #.##"
      Height          =   225
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label lblTitle 
      Caption         =   "Launch Base Mod Creator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   3165
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version " & CStr(App.Major) & PadNum(App.Minor, 2) & IIf(App.Revision = 0, "", "." & PadNum(App.Revision, 2))
    lblDate.Caption = "Last updated " & ReadINIStr("General", "Date", JoinPath(JoinPath(frmMain.EXEDIR, "launcher"), "liblist.gam"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub
