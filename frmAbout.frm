VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base: About"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4440
      TabIndex        =   1
      Top             =   3600
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
      Picture         =   "frmAbout.frx":0E42
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label10 
      Caption         =   "AlexB - helping users with operating system-related issues."
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   5175
   End
   Begin VB.Label lblWebsite 
      AutoSize        =   -1  'True
      Caption         =   "http://marshall.strategy-x.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   930
      MouseIcon       =   "frmAbout.frx":1A84
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   2955
      Width           =   2100
   End
   Begin VB.Label Label9 
      Caption         =   "gordon-creAtive - Launch Base splash screen."
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Label Label8 
      Caption         =   "Robert Hubley - File checksum reading code."
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1920
      Width           =   5175
   End
   Begin VB.Label Label7 
      Caption         =   "Renegade - Launch Base icon, Renegade Glass 'N' Metal skin, hosting."
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
      Top             =   2400
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "DCoder - DLL file for mix, csf and audio.bag file writing."
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
      Caption         =   $"frmAbout.frx":1D8E
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2760
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Last updated: 2011-04-24"
      Height          =   225
      Left            =   3480
      TabIndex        =   4
      Top             =   360
      Width           =   1965
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5520
      Y1              =   3465
      Y2              =   3465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   5535
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version #.##.####"
      Height          =   225
      Left            =   3480
      TabIndex        =   3
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label lblTitle 
      Caption         =   "Launch Base"
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
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   2565
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
    Select Case App.Revision
    Case 0: lblVersion.Caption = "Version " & App.Major & "." & PadNum(App.Minor, 2)
    Case Else: lblVersion.Caption = "Version " & App.Major & "." & PadNum(App.Minor, 2) & "." & PadNum(App.Revision, 4)
    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblWebsite.ForeColor = &HFF0000
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblWebsite.ForeColor = &HFF0000
End Sub

Private Sub lblWebsite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblWebsite.ForeColor = &HFF&
End Sub

Private Sub lblWebsite_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case vbLeftButton
        If OpenLocation(lblWebsite.Caption) < 32 Then
            Call MsgBox("Unable to open " & Quote(lblWebsite.Caption) & ".", vbOKOnly + vbInformation, App.Title)
        End If
    'Case vbRightButton
    '    menu_url.Tag = Index
    '    Call PopupMenu(menu_rc2, 2, X + lblWebsite.Left, Y + lblWebsite.Top)
    End Select
End Sub
