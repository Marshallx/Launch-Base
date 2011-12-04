VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base"
   ClientHeight    =   5415
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9015
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2040
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   5
   End
   Begin MSComDlg.CommonDialog dialogOpen 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "EXE"
      DialogTitle     =   "Select FinalAlert 2 YR Executable"
      Filter          =   "Final Alert 2 YR Executable|FinalAlert2YR.exe"
      Flags           =   4
   End
   Begin VB.PictureBox skinTabBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5820
      Index           =   0
      Left            =   0
      ScaleHeight     =   5820
      ScaleWidth      =   9105
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   9105
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   0
         Left            =   4680
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   840
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.CommandButton cmdLaunch 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton cmdManual 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4680
         Width           =   1935
      End
      Begin VB.ListBox lstMods 
         Appearance      =   0  'Flat
         Height          =   4320
         Index           =   0
         ItemData        =   "frmMain.frx":0E42
         Left            =   270
         List            =   "frmMain.frx":0E44
         TabIndex        =   0
         Top             =   825
         Width           =   4155
      End
      Begin VB.Label lblModUsesAres 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         Height          =   195
         Left            =   5640
         TabIndex        =   126
         Top             =   3360
         Width           =   210
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uses Ares:"
         Height          =   195
         Index           =   24
         Left            =   4680
         TabIndex        =   125
         Top             =   3360
         Width           =   765
      End
      Begin VB.Label lblModWebsite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.westwood.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   5640
         MouseIcon       =   "frmMain.frx":0E46
         MousePointer    =   99  'Custom
         TabIndex        =   87
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblTabStrip0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   86
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   85
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   84
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip0 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   83
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblModSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KB"
         Height          =   195
         Index           =   0
         Left            =   5640
         TabIndex        =   58
         Top             =   2640
         Width           =   210
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disk Usage:"
         Height          =   195
         Index           =   4
         Left            =   4680
         TabIndex        =   57
         Top             =   2640
         Width           =   870
      End
      Begin VB.Label lblModCampaigns 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Original Campaigns"
         Height          =   195
         Left            =   5640
         TabIndex        =   56
         Top             =   2880
         Width           =   1350
      End
      Begin VB.Label lblModAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Westwood Studios"
         Height          =   195
         Index           =   0
         Left            =   5640
         TabIndex        =   55
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblModVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.001"
         Height          =   195
         Index           =   0
         Left            =   5640
         TabIndex        =   54
         Top             =   1680
         Width           =   405
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         Height          =   195
         Index           =   3
         Left            =   4680
         TabIndex        =   53
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   195
         Index           =   2
         Left            =   4680
         TabIndex        =   52
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Campaigns:"
         Height          =   195
         Index           =   5
         Left            =   4680
         TabIndex        =   51
         Top             =   2880
         Width           =   825
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   50
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label lblModDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":1150
         Height          =   855
         Index           =   0
         Left            =   4680
         TabIndex        =   49
         Top             =   3720
         Width           =   3975
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   48
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label lblModDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2006-11-12"
         Height          =   195
         Index           =   0
         Left            =   5640
         TabIndex        =   47
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TX Plugin:"
         Height          =   195
         Index           =   6
         Left            =   4680
         TabIndex        =   46
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblModTX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not Required"
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   0
         Left            =   5640
         TabIndex        =   45
         Top             =   3120
         Width           =   945
      End
   End
   Begin VB.PictureBox skinTabBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5820
      Index           =   3
      Left            =   0
      ScaleHeight     =   5820
      ScaleWidth      =   9105
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   9105
      Begin VB.TextBox txtModParams 
         Height          =   285
         Left            =   5640
         MaxLength       =   255
         TabIndex        =   10
         Top             =   2880
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.ListBox lstMods 
         Appearance      =   0  'Flat
         Height          =   4320
         Index           =   3
         ItemData        =   "frmMain.frx":11E1
         Left            =   270
         List            =   "frmMain.frx":11E3
         TabIndex        =   9
         Top             =   825
         Width           =   4155
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   3
         Left            =   4680
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   60
         Top             =   840
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   61
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.CommandButton cmdLaunch 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton cmdManual 
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label lblModDescription 
         BackStyle       =   0  'Transparent
         Height          =   855
         Index           =   3
         Left            =   4680
         TabIndex        =   93
         Top             =   3720
         Width           =   3975
      End
      Begin VB.Label lblModParams 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Parameters:"
         Height          =   195
         Left            =   4680
         TabIndex        =   92
         Top             =   2910
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblModWebsite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   5640
         MouseIcon       =   "frmMain.frx":11E5
         MousePointer    =   99  'Custom
         TabIndex        =   88
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label lblTabStrip3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   74
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   73
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   72
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   71
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblModSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   70
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disk Usage:"
         Height          =   195
         Index           =   12
         Left            =   4680
         TabIndex        =   69
         Top             =   2640
         Width           =   870
      End
      Begin VB.Label lblModAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   68
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label lblModVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   67
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         Height          =   195
         Index           =   11
         Left            =   4680
         TabIndex        =   66
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   195
         Index           =   10
         Left            =   4680
         TabIndex        =   65
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         Height          =   195
         Index           =   8
         Left            =   4680
         TabIndex        =   64
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Index           =   9
         Left            =   4680
         TabIndex        =   63
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label lblModDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   5640
         TabIndex        =   62
         Top             =   1920
         Width           =   45
      End
   End
   Begin VB.PictureBox skinTabBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5820
      Index           =   2
      Left            =   0
      ScaleHeight     =   5820
      ScaleWidth      =   9105
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   9105
      Begin VB.TextBox txtFA2Folder 
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   4800
         Width           =   3735
      End
      Begin VB.ListBox lstMods 
         Appearance      =   0  'Flat
         Height          =   3540
         Index           =   2
         ItemData        =   "frmMain.frx":14EF
         Left            =   270
         List            =   "frmMain.frx":14F1
         TabIndex        =   4
         Top             =   825
         Width           =   4155
      End
      Begin VB.CheckBox cboxTX 
         Height          =   195
         Left            =   4800
         TabIndex        =   6
         Top             =   3390
         Width           =   195
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   2
         Left            =   4680
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   14
         Top             =   840
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.CommandButton cmdLaunch 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton cmdManual 
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label cmdFA2Browse 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4080
         TabIndex        =   94
         Top             =   4800
         Width           =   285
      End
      Begin VB.Label lblTX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Integrate Terrain Expansion with this FA2 mod."
         Height          =   195
         Left            =   5040
         TabIndex        =   91
         Top             =   3390
         Width           =   3285
      End
      Begin VB.Label lblModWebsite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5640
         MouseIcon       =   "frmMain.frx":14F3
         MousePointer    =   99  'Custom
         TabIndex        =   89
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label lblTabStrip2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   78
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   77
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   76
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   75
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblModSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   5640
         TabIndex        =   29
         Top             =   2640
         Width           =   45
      End
      Begin VB.Label lblModFA2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   5640
         TabIndex        =   28
         Top             =   2880
         Width           =   45
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FA2 Version:"
         Height          =   195
         Index           =   18
         Left            =   4680
         TabIndex        =   27
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TX Plugin:"
         Height          =   195
         Index           =   19
         Left            =   4680
         TabIndex        =   26
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label lblModTX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00004000&
         Height          =   195
         Index           =   2
         Left            =   5640
         TabIndex        =   25
         Top             =   3120
         Width           =   45
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disk Usage:"
         Height          =   195
         Index           =   17
         Left            =   4680
         TabIndex        =   24
         Top             =   2640
         Width           =   870
      End
      Begin VB.Label lblModAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   5640
         TabIndex        =   23
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label lblModVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   5640
         TabIndex        =   22
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         Height          =   195
         Index           =   16
         Left            =   4680
         TabIndex        =   21
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   195
         Index           =   15
         Left            =   4680
         TabIndex        =   20
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         Height          =   195
         Index           =   13
         Left            =   4680
         TabIndex        =   19
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label lblModDescription 
         BackStyle       =   0  'Transparent
         Height          =   855
         Index           =   2
         Left            =   4680
         TabIndex        =   18
         Top             =   3720
         Width           =   3975
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Index           =   14
         Left            =   4680
         TabIndex        =   17
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label lblModDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   5640
         TabIndex        =   16
         Top             =   1920
         Width           =   45
      End
   End
   Begin VB.PictureBox skinTabBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5820
      Index           =   1
      Left            =   0
      ScaleHeight     =   5820
      ScaleWidth      =   9105
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   9105
      Begin VB.ListBox lstMods 
         Appearance      =   0  'Flat
         Height          =   4320
         Index           =   1
         ItemData        =   "frmMain.frx":17FD
         Left            =   270
         List            =   "frmMain.frx":17FF
         Sorted          =   -1  'True
         TabIndex        =   98
         Top             =   825
         Width           =   4155
      End
      Begin VB.CommandButton cmdLaunch 
         Height          =   375
         Index           =   4
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton cmdLaunch 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton cmdManual 
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4680
         Width           =   1935
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   4680
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   31
         Top             =   840
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.Label lblModVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   5640
         TabIndex        =   95
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label lblModWebsite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5640
         MouseIcon       =   "frmMain.frx":1801
         MousePointer    =   99  'Custom
         TabIndex        =   90
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label lblTabStrip1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   82
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   81
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   80
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   79
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblModDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   5640
         TabIndex        =   41
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Index           =   21
         Left            =   4680
         TabIndex        =   40
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label lblModDescription 
         BackStyle       =   0  'Transparent
         Height          =   855
         Index           =   1
         Left            =   4680
         TabIndex        =   39
         Top             =   3720
         Width           =   3975
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version:"
         Height          =   195
         Index           =   20
         Left            =   4680
         TabIndex        =   38
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   195
         Index           =   22
         Left            =   4680
         TabIndex        =   37
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         Height          =   195
         Index           =   23
         Left            =   4680
         TabIndex        =   36
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label lblModAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   5640
         TabIndex        =   35
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label lblGeneral 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disk Usage:"
         Height          =   195
         Index           =   7
         Left            =   4680
         TabIndex        =   34
         Top             =   2640
         Width           =   870
      End
      Begin VB.Label lblModSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   5640
         TabIndex        =   33
         Top             =   2640
         Width           =   45
      End
   End
   Begin VB.PictureBox skinTabBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5820
      Index           =   4
      Left            =   0
      ScaleHeight     =   5820
      ScaleWidth      =   9105
      TabIndex        =   99
      Top             =   0
      Visible         =   0   'False
      Width           =   9105
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   13
         Left            =   4680
         MouseIcon       =   "frmMain.frx":1B0B
         MousePointer    =   99  'Custom
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   4080
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   124
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   12
         Left            =   360
         MouseIcon       =   "frmMain.frx":1E15
         MousePointer    =   99  'Custom
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   4080
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   122
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   11
         Left            =   4680
         MouseIcon       =   "frmMain.frx":211F
         MousePointer    =   99  'Custom
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   3240
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   120
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   10
         Left            =   360
         MouseIcon       =   "frmMain.frx":2429
         MousePointer    =   99  'Custom
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   3240
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   118
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   9
         Left            =   4680
         MouseIcon       =   "frmMain.frx":2733
         MousePointer    =   99  'Custom
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   2400
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   116
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   8
         Left            =   360
         MouseIcon       =   "frmMain.frx":2A3D
         MousePointer    =   99  'Custom
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   2400
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   114
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   7
         Left            =   4680
         MouseIcon       =   "frmMain.frx":2D47
         MousePointer    =   99  'Custom
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   1560
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   112
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   6
         Left            =   360
         MouseIcon       =   "frmMain.frx":3051
         MousePointer    =   99  'Custom
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   1560
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   110
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   5
         Left            =   4680
         MouseIcon       =   "frmMain.frx":335B
         MousePointer    =   99  'Custom
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   720
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   108
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.HScrollBar scrollbarBanners 
         Height          =   255
         Left            =   360
         Max             =   0
         TabIndex        =   106
         Top             =   4920
         Width           =   8295
      End
      Begin VB.PictureBox picMod 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   4
         Left            =   360
         MouseIcon       =   "frmMain.frx":3665
         MousePointer    =   99  'Custom
         ScaleHeight     =   735
         ScaleWidth      =   3975
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   720
         Width           =   3975
         Begin VB.Label lblNoBanner 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   101
            Top             =   120
            Width           =   3735
         End
      End
      Begin VB.Line lineBannerBottom 
         Visible         =   0   'False
         X1              =   345
         X2              =   4350
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Line lineBannerTop 
         Visible         =   0   'False
         X1              =   345
         X2              =   4350
         Y1              =   705
         Y2              =   705
      End
      Begin VB.Line lineBannerRight 
         Visible         =   0   'False
         X1              =   4335
         X2              =   4335
         Y1              =   720
         Y2              =   1455
      End
      Begin VB.Line lineBannerLeft 
         Visible         =   0   'False
         X1              =   345
         X2              =   345
         Y1              =   720
         Y2              =   1455
      End
      Begin VB.Label lblTabStrip4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   5520
         TabIndex        =   105
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   104
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   103
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTabStrip4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   102
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Menu menu_file 
      Caption         =   "&File"
      Begin VB.Menu menu_changetab 
         Caption         =   "Change &Tab"
         Begin VB.Menu menu_tab 
            Caption         =   "&Mods [List View]"
            Index           =   0
            Shortcut        =   {F5}
         End
         Begin VB.Menu menu_tab 
            Caption         =   "&Plugins"
            Index           =   1
            Shortcut        =   {F6}
         End
         Begin VB.Menu menu_tab 
            Caption         =   "&FinalAlert 2 Mods"
            Index           =   2
            Shortcut        =   {F7}
         End
         Begin VB.Menu menu_tab 
            Caption         =   "&Tools"
            Index           =   3
            Shortcut        =   {F8}
         End
         Begin VB.Menu menu_tab 
            Caption         =   "Mods [&Banner View]"
            Index           =   4
            Shortcut        =   {F9}
         End
      End
      Begin VB.Menu menu_cskin 
         Caption         =   "Change &Skin"
         Begin VB.Menu menu_skin 
            Caption         =   ""
            Index           =   0
         End
      End
      Begin VB.Menu menu_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menu_tools 
      Caption         =   "&Tools"
      Begin VB.Menu menu_modcat 
         Caption         =   "&Check For Updates/New Mods..."
      End
      Begin VB.Menu menu_fileman 
         Caption         =   "&Residual File Manager..."
         Visible         =   0   'False
      End
      Begin VB.Menu menu_history 
         Caption         =   "&Download History..."
      End
      Begin VB.Menu menu_livelog 
         Caption         =   "&LiveLog"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_tools_line0 
         Caption         =   "-"
      End
      Begin VB.Menu menu_options 
         Caption         =   "&Options..."
      End
      Begin VB.Menu menu_tools_line1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_usertool 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menu_ares 
      Caption         =   "&Ares"
      Visible         =   0   'False
      Begin VB.Menu menu_aresdoc 
         Caption         =   "Ares &Manual"
      End
      Begin VB.Menu menu_aresoptions 
         Caption         =   "Ares Update &Options..."
         Visible         =   0   'False
      End
      Begin VB.Menu menu_aresini 
         Caption         =   "View/&Edit Ares.ini Settings..."
      End
      Begin VB.Menu menu_debugfolder 
         Caption         =   "Open Ares &Debug Folder"
      End
   End
   Begin VB.Menu menu_help 
      Caption         =   "&Help"
      Begin VB.Menu menu_helptopics 
         Caption         =   "&Help Topics..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu menu_About 
         Caption         =   "&About..."
      End
      Begin VB.Menu menu_disclaimer 
         Caption         =   "&Disclaimer..."
      End
      Begin VB.Menu menu_help_line1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_website 
         Caption         =   "Website"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu menu_rc 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menu_openfolder 
         Caption         =   "Open Containing Folder"
      End
      Begin VB.Menu menu_checkforupdate 
         Caption         =   "Check For Updates"
      End
      Begin VB.Menu menu_deletemod 
         Caption         =   "Delete Mod"
      End
   End
   Begin VB.Menu menu_rc2 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menu_url 
         Caption         =   "Copy URL to Clipboard"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Misc stuff used only by frmMain
Dim FA2DIR As String
Dim TFD As Boolean
Dim TamperingDetected As Boolean
Dim SelectedTab As Integer
Dim BannerTab As Boolean
Dim BannerPage As Integer
Private Const BannerCount As Integer = 10
Dim UserToolModNum() As Integer
Dim PreventLoop As Boolean
Dim MutexValue As Long
Dim AlphabetisedMods() As Integer
'Skins
Dim SkinList() As Menu
Dim SkinPath() As String
Dim SkinCount As Integer
Dim SelectedSkinPath As String
'Colors
Dim ColorGood As Long
Dim ColorBad As Long
Dim ColorURL As Long
Dim ColorURLActive As Long
Dim ColorNeutral As Long
Dim ColorListText As Long
Dim ColorList As Long
Dim ColorComboText As Long
Dim ColorCombo As Long
Dim ColorComboTextDisabled As Long
Dim ColorComboDisabled As Long

Private Function Init_MaxMods() As Integer
    Dim fso As FileSystemObject
    Dim fso_folder As Folder
    Dim fso_file As File
    Dim iTotal As Integer
    'count how many mods we might have to load
    If DirExists(JoinPath(EXEDIR, "Mods")) Then
        Set fso = New FileSystemObject
        Set fso_folder = fso.GetFolder(JoinPath(EXEDIR, "Mods"))
        iTotal = fso_folder.SubFolders.Count * 2 'Init_Plugins reads the same directory, before we create originalyr and originalra2
        If Not DirExists(JoinPath(EXEDIR, "Mods\originalyr")) Then iTotal = iTotal + 1
        If Not DirExists(JoinPath(EXEDIR, "Mods\originalra2")) Then iTotal = iTotal + 1
        iTotal = iTotal + 5 'Init_Mods has 2 extra tasks, 'Init_Plugins has 3
    Else
        iTotal = 5 'Init_Mods has 2 extra tasks, 'Init_Plugins has 1
    End If
    Init_MaxMods = iTotal
End Function

Private Sub Init()
    Dim iCounter As Integer
    Dim sTemp As String
    Dim bOk As Boolean
    CL_noexcept = Len(GetArgByName("noexcept")) <> 0
    Call CallStackPush(Me.Name & ".Init()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    ReDim SafeFiles(0) 'must be dimmed before Shutdown runs, else will error
    Call Init_OpenMutex
    Call Init_Constants
    CL_dev = Len(GetArgByName("dev")) <> 0
    OptAdvancedMode = BooleanStringToBoolean(ReadINIStr("Options", "AdvancedMode", ProgramINI, "no"))
    'Splash screen
    frmSplash.pbarWidthPerUnit = frmSplash.pbarSplash.Width / (Init_MaxMods + 19)
    frmSplash.pbarSplash.Width = (frmSplash.pbarSplash.Width Mod frmSplash.pbarWidthPerUnit)
    If Len(GetArgByName("nosplash")) = 0 Then
        Call frmSplash.Show
        Call frmSplash.Refresh
    End If
    Call frmSplash.PROGRESS("Initializing...")
    'Start log
    Call Init_Logging
    Call WriteLogEntry
    Call WriteLogEntry("Initialization started.", LogLevel1)
    Call frmSplash.PROGRESS("Updating version/install path...")
    'Update ini
    Call WriteINIStr("General", "LaunchBaseVersion", App.Major & "." & PadNum(App.Minor, 2) & "." & PadNum(App.Revision), ProgramINI)
    'Update registry
    bOk = False
    sTemp = ReadRegStr("HKLM\SOFTWARE\Marshallx Industries\YR Launch Base\InstallPath")
    If Len(sTemp) <> 0 Then
        If DirExists(sTemp) Then
            If GetShortFileName(sTemp) = GetShortFileName(EXEDIR) Then bOk = True
        End If
    End If
    If Not bOk Then
        Call WriteRegStr("HKLM\SOFTWARE\Marshallx Industries\YR Launch Base\InstallPath", EXEDIR)
        Call WriteLogEntry("Updated InstallPath in registry.", LogLevel1)
    End If
    'Load settings from LaunchBase.ini
    Call frmSplash.PROGRESS("Loading Options...")
    Call Init_LoadOptions
    Call frmSplash.PROGRESS("Loading URLs...")
    Call Init_URLs
    Call frmSplash.PROGRESS("Showing disclaimer...")
    'Disclaimer
    If BooleanStringToBoolean(ReadINIStr("General", "ShowDisclaimer", ProgramINI, "yes")) Then
        Call frmDisclaimer.Show
        frmDisclaimer.cmdCancel.Visible = True
        frmDisclaimer.cmdOK.Caption = "I Agree"
        Do
            DoEvents
            iCounter = 0
            Do While iCounter < Forms.Count '0-based index, 1-based count
                If Forms(iCounter).Name = "frmDisclaimer" Then Exit Do
                iCounter = iCounter + 1
            Loop
            If iCounter >= Forms.Count Then Exit Do 'because we didn't find frmDisclaimer so it must have unloaded
        Loop
        Call WriteINIStr("General", "ShowDisclaimer", "no", ProgramINI)
    End If
    Call frmSplash.PROGRESS("Creating directories...")
    'Set up directories
    If Not DirExists(BACKUPDIR) Then Call LoggedMkDir(BACKUPDIR)
    If Not DirExists(JoinPath(BACKUPDIR, "Taunts")) Then Call LoggedMkDir(JoinPath(BACKUPDIR, "Taunts"))
    If Not DirExists(LOGDIR) Then Call LoggedMkDir(LOGDIR)
    Call frmSplash.PROGRESS("Determining Red Alert 2 status...")
    Call Init_RA2Check
    Call frmSplash.PROGRESS("Checking for free disk space...")
    'Check disk space
    Call WriteLogEntry("Checking for free disk space.", LogLevel1)
    Select Case True
    Case FreeDiskSpace(UCase$(Left$(EXEDIR, 1))) < OptSafetySpace
        Call WriteLogEntry("Insufficient free disk space on drive " & UCase$(Left$(EXEDIR, 1)) & ". Launch Base is configured to leave a minimum " & DataSize(OptSafetySpace, "MB") & " of free space.", LogShutdown)
    Case FreeDiskSpace(UCase$(Left$(RA2DIR, 1))) < OptSafetySpace
        Call WriteLogEntry("Insufficient free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & ". Launch Base is configured to leave a minimum " & DataSize(OptSafetySpace, "MB") & " of free space.", LogShutdown)
    Case FreeDiskSpace(UCase$(Left$(EXEDIR, 1))) < (OptSafetySpace * 2)
        Call WriteLogEntry("Free disk space on drive " & UCase$(Left$(EXEDIR, 1)) & " is very low. It is recommended that you increase free disk space to at least " & DataSize(OptSafetySpace * 2, "MB") & ".", LogMsgBox)
    Case FreeDiskSpace(UCase$(Left$(RA2DIR, 1))) < (OptSafetySpace * 2)
        Call WriteLogEntry("Free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & " is very low. It is recommended that you increase free disk space to at least " & DataSize(OptSafetySpace * 2, "MB") & ".", LogMsgBox)
    End Select
    Call frmSplash.PROGRESS("Loading safe file list...")
    'Load list of safe files
    Call SafeFiles_Load
    Call frmSplash.PROGRESS("Loading plugins...")
    'Load plugins and mods
    Call Init_LoadPlugins
    Call frmSplash.PROGRESS("Loading mods...")
    Call Init_LoadMods
    'Check FA2
    Call FA2Check(True)
    Call frmSplash.PROGRESS("Loading skins...")
    'Load skins and apply selected skin
    Call Init_LoadSkins
    Call frmSplash.PROGRESS("Performing restore process...")
    'Restore process
    If BooleanStringToBoolean(ReadINIStr("Restore", "RestorePending", ProgramINI)) Then
        Call WriteLogEntry("Restore process incomplete! Completing restore process now.")
        Call MsgBox("The last time Launch Base was run it was prematurely terminated before the restore process had a chance to run." & vbCrLf & "The restore process will be completed now.", vbOKOnly + vbInformation, App.Title)
        If menu_livelog.Checked = True Then Call frmLiveLog.SetFocus
        Call RestoreProcess
        Call MsgBox("Restore process complete.", vbOKOnly + vbInformation, App.Title)
    End If
    Call frmSplash.PROGRESS("Selecting previous tab...")
    'Set last selected tab
    SelectedTab = Val(ReadINIStr("General", "SelectedTab", ProgramINI, "0"))
    BannerTab = BooleanStringToBoolean(ReadINIStr("General", "BannerTab", ProgramINI, "no"))
    Call SelectTab(SelectedTab)
    Call frmSplash.PROGRESS("Checking DCoder DLL...")
    'Close splash form and show main form
    DCoderDLL = FileExists(JoinPath(RESDIR, "dcoder.dll"))
    If Not DCoderDLL Then Call WriteLogEntry("DCoder DLL is missing! Forcing loose file mode. Some uncompiled mods may not be activated correctly.")
    Call frmSplash.PROGRESS("Loading update catalogue...")
    Call frmModCat.LoadUpdateCat
    Call frmSplash.PROGRESS("Loading command line arguments...")
    'Command line arguments
    Call Init_CommandLine
    Call frmSplash.PROGRESS("Checking for updates to Launch Base...")
    If OptAutoUpdate Then
        Call LaunchMod_CheckForUpdate(bOk, LBModNum)
    Else
        bOk = False
    End If
    'Note: bOk is being used for "Shutdown now please, cause we just started an update"
    If Not bOk Then
        Call frmSplash.PROGRESS("Initialization complete.")
        Call WriteLogEntry("Initialization complete.", LogLevel1)
        If Len(GetArgByName("nosplash")) = 0 Then
            Call Sleep(500)
            Unload frmSplash
        End If
        Me.Show
        Me.Refresh
        'Command line arguments
        If TamperingDetected Then 'not allowed to auto-launch a game.
            If Len(CL_game) = 0 Then
                Call MsgBox("Active mod/plugin files have been tampered with outside of Launch Base!" & vbCrLf & "It is strongly recommended that you thoroughly re-read the help topics.", vbOKOnly + vbInformation, App.Title)
            Else
                Call MsgBox("Active mod/plugin files have been tampered with outside of Launch Base!" & vbCrLf & "In light of this, Launch Base will not auto-launch " & Quote(CL_game) & vbCrLf & "It is strongly recommended that you thoroughly re-read the help topics.", vbOKOnly + vbInformation, App.Title)
                CL_game = ""
            End If
            CL_modnum = -1
        End If
        If Len(CL_game) <> 0 Then
            If lstMods(Mods(CL_modnum).ModType).ListCount <> 0 Then
                For iCounter = 0 To (lstMods(Mods(CL_modnum).ModType).ListCount - 1)
                    If lstMods(Mods(CL_modnum).ModType).ItemData(iCounter) = CL_modnum Then
                        lstMods(Mods(CL_modnum).ModType).ListIndex = iCounter
                        Call PlaySound
                        iCounter = lstMods(Mods(CL_modnum).ModType).ListCount - 1
                    End If
                Next iCounter
            End If
            If CL_modnum <> -1 Then
                Call WriteLogEntry("Command line argument: Launching " & Quote(CL_game), LogLevel1)
                Call SelectTab(Mods(CL_modnum).ModType)
                Call DisplayModDetails(CL_modnum, False)
                Call LaunchMod(TypeMod, CL_modnum)
            Else
                Call WriteLogEntry("Command line argument: Unable to launch " & Quote(CL_game) & " because it doesn't exist!", LogShutdown)
            End If
        End If
    End If
    If FileExists(JoinPath(EXEDIR, "graion.txt")) Then
        Call WriteINIStr("Test", "A", GetFileMD5(JoinPath(RESDIR, "Syringe.exe")), JoinPath(EXEDIR, "graion.txt"))
        Call WriteINIStr("Test", "B", "c", JoinPath(EXEDIR, "graion.txt"))
        Call WriteINIStr("Test", "C", EncryptString("c", "password"), JoinPath(EXEDIR, "graion.txt"))
        Call WriteINIStr("Test", "D", EncryptString("c", "-6277thxtK"), JoinPath(EXEDIR, "graion.txt"))
        Call WriteINIStr("Test", "E", Base64EncodeString("c"), JoinPath(EXEDIR, "graion.txt"))
        Call WriteINIStr("Test", "F", Base64EncodeString(EncryptString("c", "password")), JoinPath(EXEDIR, "graion.txt"))
        Call WriteINIStr("Test", "G", Base64EncodeString(EncryptString("c", "-6277thxtK")), JoinPath(EXEDIR, "graion.txt"))
    End If
    Call CallStackPop
End Sub

Private Sub Init_Logging()
    If BooleanStringToBoolean(ReadINIStr("LiveLog", "LiveLogOpen", ProgramINI, "no")) Then Call menu_livelog_Click
    OptLogFile = BooleanStringToBoolean(ReadINIStr("Options", "LogFile", ProgramINI, "yes"))
    OptInitLog = BooleanStringToBoolean(ReadINIStr("Options", "InitLog", ProgramINI, "no"))
    OptLiveLog = BooleanStringToBoolean(ReadINIStr("LiveLog", "LiveLog", ProgramINI, "no"))
    OptLogLevel = Val(ReadINIStr("Options", "LogLevel", ProgramINI, "1"))
    If OptInitLog Then
        If FileExists(LOGFILE) Then Call LoggedKill(LOGFILE)
    End If
    menu_livelog.Visible = OptLiveLog
End Sub

Private Sub Init_OpenMutex()
    MutexValue = -1
    If Len(GetArgByName("nomutex")) = 0 Then
        MutexValue = CreateMutex(ByVal 0&, 1, "YRLBMUTEXERM1")
        If (Err.LastDllError = 183&) Then
            Call MsgBox("Either Launch Base itself or an installer is already running.", vbOKOnly + vbInformation, App.Title)
            Call CloseHandle(MutexValue)
            End
        End If
    End If
End Sub

Private Sub Init_CommandLine()
    Dim iCounter As Long
    Dim sTemp As String
    Dim bOk As Boolean
    Call WriteLogEntry("Loading command line arguments...", LogLevel1)
    'game
    CL_game = GetArgByName("game")
    If CL_game = "True" Or CL_game = "" Then
        CL_game = ""
        CL_modnum = -1
    Else
        iCounter = 1
        Do While iCounter <= ModCount
            If UCase$(GetFileName(Mods(iCounter).ModPath)) = UCase$(CL_game) Then
                CL_modnum = iCounter
                Exit Do
            End If
            iCounter = iCounter + 1
        Loop
    End If
    Call WriteLogEntry("CL_game = " & Quote(CL_game), LogLevel2)
    'playfile
    bOk = False
    CL_playfile = GetArgByNumber(0)
    Call WriteLogEntry("CL_playfile = " & Quote(CL_playfile), LogLevel2)
    If FileType(CL_playfile) = "IPB" Then
        sTemp = JoinPath(EXEDIR, "Mods")
        If DirExists(sTemp) Then
            sTemp = GetShortFileName(sTemp)
            If FileExists(JoinPath(sTemp, CL_playfile)) Then
                'only specified relative path to file
                bOk = True
            Else
                'specified path must be absolute
                CL_playfile = GetShortFileName(CL_playfile)
                If Len(CL_playfile) > Len(sTemp) Then
                    If FileExists(CL_playfile) Then
                        If UCase$(Left$(CL_playfile, Len(sTemp))) = UCase$(sTemp) Then
                            sTemp = Mid$(CL_playfile, Len(sTemp) + 1)
                            If Len(sTemp) > 5 Then
                                iCounter = InStr(2, sTemp, "\")
                                If iCounter <> 0 Then
                                    sTemp = Mid$(sTemp, 2, iCounter - 2)
                                    If DirExists(JoinPath(EXEDIR, "Mods\" & sTemp)) Then
                                        iCounter = 1
                                        Do While iCounter <= ModCount
                                            If UCase$(GetFileName(Mods(iCounter).ModPath)) = UCase$(sTemp) Then
                                                If CL_modnum = iCounter Or Len(CL_game) = 0 Then
                                                    CL_modnum = iCounter
                                                    CL_game = sTemp
                                                    CL_playfile = Mid$(CL_playfile, Len(JoinPath(EXEDIR, "Mods\" & sTemp & "\")) + 1)
                                                    bOk = True
                                                End If
                                                Exit Do
                                            End If
                                            iCounter = iCounter + 1
                                        Loop
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    If bOk = False Then CL_playfile = ""
    'tx
    CL_tx = Len(GetArgByName("tx")) <> 0
    Call WriteLogEntry("CL_tx = " & BooleanToYesNo(CL_tx), LogLevel2)
    'CL_dev = Len(GetArgByName("dev")) <> 0 'done at beginning of Init
    Call WriteLogEntry("CL_dev = " & BooleanToYesNo(CL_dev), LogLevel2)
    'CL_noexcept = Len(GetArgByName("noexcept")) <> 0 'done at the beginning of Init
    Call WriteLogEntry("CL_noexcept = " & BooleanToYesNo(CL_noexcept), LogLevel2)
End Sub

Private Sub Init_LoadOptions()
    'Note, some options are loaded at the beginning of Init due to needing them during Init
    Call CallStackPush(Me.Name & ".Init_LoadOptions()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Loading options...", LogLevel1)
    OptRecompile = BooleanStringToBoolean(ReadINIStr("Options", "Recompile", ProgramINI, "no"))
    OptLooseFileMode = BooleanStringToBoolean(ReadINIStr("Options", "LooseFileMode", ProgramINI, "no"))
    OptModSound1 = BooleanStringToBoolean(ReadINIStr("Options", "EnableModSound1", ProgramINI, "yes"))
    OptModSound2 = BooleanStringToBoolean(ReadINIStr("Options", "EnableModSound2", ProgramINI, "yes"))
    OptLBSounds = BooleanStringToBoolean(ReadINIStr("Options", "EnableLBSounds", ProgramINI, "yes"))
    OptLogAres = BooleanStringToBoolean(ReadINIStr("Options", "RockPatchLogging", ProgramINI, "no"))
    OptLogExcept = BooleanStringToBoolean(ReadINIStr("Options", "LogCapture", ProgramINI, "yes"))
    OptLogExceptDesc = BooleanStringToBoolean(ReadINIStr("Options", "LogIEDesc", ProgramINI, "yes"))
    OptWindowed = BooleanStringToBoolean(ReadINIStr("Options", "WindowedMode", ProgramINI, "no"))
    OptRecord = BooleanStringToBoolean(ReadINIStr("Options", "RecordVideo", ProgramINI, "no"))
    OptPlay = BooleanStringToBoolean(ReadINIStr("Options", "PlayVideo", ProgramINI, "no"))
    OptSkipLogo = BooleanStringToBoolean(ReadINIStr("Options", "SkipLogo", ProgramINI, "no"))
    OptUseCheckSums = BooleanStringToBoolean(ReadINIStr("Options", "UseCheckSums", ProgramINI, "yes"))
    OptGameChecksums = BooleanStringToBoolean(ReadINIStr("Options", "GameChecksums", ProgramINI, "yes"))
    OptVerifyPlugins = BooleanStringToBoolean(ReadINIStr("Options", "VerifyPlugins", ProgramINI, "yes"))
    OptCheckModYPLFiles = BooleanStringToBoolean(ReadINIStr("Options", "CheckModYPLFiles", ProgramINI, "no"))
    OptAutoTX = BooleanStringToBoolean(ReadINIStr("Options", "AutoTX", ProgramINI, "no"))
    OptAutoUpdate = BooleanStringToBoolean(ReadINIStr("Options", "AutoUpdate", ProgramINI, "yes"))
    OptFullDownloads = BooleanStringToBoolean(ReadINIStr("Options", "FullDownloads", ProgramINI, "no"))
    OptAutoAresUpdate = BooleanStringToBoolean(ReadINIStr("Options", "AresPrompt", ProgramINI, "yes"))
    OptShowRA2 = BooleanStringToBoolean(ReadINIStr("Options", "ShowRA2", ProgramINI, "yes"))
    OptShowYR = BooleanStringToBoolean(ReadINIStr("Options", "ShowYR", ProgramINI, "yes"))
    OptSpeedControl = BooleanStringToBoolean(ReadINIStr("Options", "SpeedControl", ProgramINI, "no"))
    OptMPDebug = BooleanStringToBoolean(ReadINIStr("Options", "MPDebug", ProgramINI, "no"))
    OptCustomSwitches = ReadINIStr("Options", "CustomSwitches", ProgramINI, "")
    OptVideoBackBuffer = BooleanStringToBoolean(ReadINIStr("Options", "VideoBackBuffer", ProgramINI, "no"))
    OptAllowVRAMSidebar = BooleanStringToBoolean(ReadINIStr("Options", "AllowVRAMSidebar", ProgramINI, "no"))
    OptMaxLogSize = Restrict(0, Val(ReadINIStr("Options", "MaxLogSize", ProgramINI, "2097152")), 16777216)
    OptSafetySpace = Restrict(33554432, Val(ReadINIStr("Options", "SafetySpace", ProgramINI, "67108864")), 1073741823)
    OptCaptureAresDebug = BooleanStringToBoolean(ReadINIStr("Options", "CaptureAresDebug", ProgramINI, "no"))
    OptAresTester = BooleanStringToBoolean(ReadINIStr("Options", "AresTester", ProgramINI, "no"))
    OptAresBranch = ReadINIStr("Options", "AresBranch", ProgramINI, "")
    OptAresRevisionDataURLURL = ReadINIStr("URL", "AresRevisionDataURLURL", ProgramINI, "http://ares.strategy-x.com/lb_data")
    OptAresRevisionDataURLHDR = ReadINIStr("URL", "AresRevisionDataURLHDR", ProgramINI, "X-Branches-File")
    OptAresRevisionDataURL = ReadINIStr("URL", "AresRevisionDataURL", ProgramINI, "http://ares.strategy-x.com/lb_data/branches")
    OptAresRevision = Val(ReadINIStr("General", "AresRevision", ProgramINI, "0"))
    OptSyringeRevision = Val(ReadINIStr("General", "SyringeRevision", ProgramINI, "0"))
    If BooleanStringToBoolean(ReadINIStr("Options", "PersistentPluginBad", ProgramINI, "no")) Then
        OptPersistentPlugin = False
        OptPersistentPluginBad = True
    Else
        OptPersistentPlugin = BooleanStringToInteger(ReadINIStr("Options", "PersistentPlugin", ProgramINI, "no"))
        OptPersistentPluginBad = False
    End If
    If BooleanStringToBoolean(ReadINIStr("Options", "PersistentModBad", ProgramINI, "no")) Then
        OptPersistentModBad = True
        OptPersistentMod = False
    Else
        OptPersistentModBad = False
        If OptPersistentPlugin Then
            OptPersistentMod = BooleanStringToInteger(ReadINIStr("Options", "PersistentMod", ProgramINI, "no"))
        Else
            OptPersistentMod = False
        End If
    End If
    OptModCatFilterModType0 = BooleanStringToBoolean(ReadINIStr("Options", "ModCatFilterModType0", ProgramINI, "yes"))
    OptModCatFilterModType1 = BooleanStringToBoolean(ReadINIStr("Options", "ModCatFilterModType1", ProgramINI, "yes"))
    OptModCatFilterModType2 = BooleanStringToBoolean(ReadINIStr("Options", "ModCatFilterModType2", ProgramINI, "yes"))
    OptModCatFilterModType3 = BooleanStringToBoolean(ReadINIStr("Options", "ModCatFilterModType3", ProgramINI, "yes"))
    OptModCatFilterModType4 = BooleanStringToBoolean(ReadINIStr("Options", "ModCatFilterModType4", ProgramINI, "yes"))
    OptModCatFilterGame0 = BooleanStringToBoolean(ReadINIStr("Options", "ModCatFilterGame0", ProgramINI, "yes"))
    OptModCatFilterGame1 = BooleanStringToBoolean(ReadINIStr("Options", "ModCatFilterGame1", ProgramINI, "yes"))
    OptModCatFilterUpdates0 = BooleanStringToBoolean(ReadINIStr("Options", "ModCatFilterUpdates0", ProgramINI, "yes"))
    OptModCatFilterUpdates1 = BooleanStringToBoolean(ReadINIStr("Options", "ModCatFilterUpdates1", ProgramINI, "yes"))
    OptModCatFilterUpdates2 = BooleanStringToBoolean(ReadINIStr("Options", "ModCatFilterUpdates2", ProgramINI, "no"))
    If OptAdvancedMode Then
        menu_ares.Visible = True
        menu_aresoptions.Visible = True
        menu_fileman.Visible = True
    End If
    If Not FileExists(JoinPath(RESDIR, "AresDocumentation\AresManual.html")) Then menu_aresdoc.Enabled = False
    ModCatUPD = ReadINIStr("URL", "ModCatUPD", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/ModCatalogue.ini")
    Call CallStackPop
End Sub

Private Sub SafeFiles_Save()
    Dim iFile As Long
    Dim sKey As String
    sKey = Decrypt(HDSerialNumber(Left$(RA2DIR, 1)), True)
    iFile = 1
    Call WriteINIStr("SafeFiles", "Count", CStr(UBound(SafeFiles())), ProgramINI)
    Do While iFile <= UBound(SafeFiles())
        Call WriteINIStr("SafeFiles", CStr(iFile), Base64EncodeString(EncryptString(StrReverse(SafeFiles(iFile)), sKey)), ProgramINI)
        iFile = iFile + 1
    Loop
End Sub

Public Sub SafeFiles_Load()
    Dim iFile As Long
    Dim sKey As String
    Dim sTemp As String
    Dim iMax As Long
    Call CallStackPush(Me.Name & ".SafeFiles_Load()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Confirming safe files...", LogLevel1)
    sKey = Decrypt(HDSerialNumber(Left$(RA2DIR, 1)), True)
    iFile = 1
    ReDim SafeFiles(0) 'also takes place on init, in case of error before we get this far
    iMax = Val(ReadINIStr("SafeFiles", "Count", ProgramINI, "0"))
    Do While iMax <> 0
        iMax = iMax - 1
        sTemp = ReadINIStr("SafeFiles", CStr(iFile), ProgramINI, "", True)
        ReDim Preserve SafeFiles(iFile)
        If Len(sTemp) Mod 4 = 0 Then 'valid length for base64 decode
            SafeFiles(iFile) = StrReverse(EncryptString(Base64DecodeString(sTemp), sKey))
            If Not FileExists(JoinPath(RA2DIR, SafeFiles(iFile))) Then
                Call SafeFiles_Find(SafeFiles(iFile), True) 'remove the file from the array
            End If
        Else
            'should never happen unless someone messes with LaunchBase.ini
            Call WriteLogEntry("Invalid SafeFiles data in LaunchBase.ini!")
        End If
        iFile = iFile + 1
    Loop
    Call CallStackPop
End Sub

Public Sub SafeFiles_Refresh()
    'in case the user has deleted any files manually since they last used the residual file manager.
    Dim iFile As Long
    Dim sKey As String
    Dim sTemp As String
    Dim iMax As Long
    Call CallStackPush(Me.Name & ".SafeFiles_Refresh()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Re-confirming safe files...", LogLevel1)
    iFile = 1
    Do While iFile <= UBound(SafeFiles())
        If Not FileExists(JoinPath(RA2DIR, SafeFiles(iFile))) Then
            Call SafeFiles_Find(SafeFiles(iFile), True) 'remove the file from the array
        Else
            iFile = iFile + 1
        End If
    Loop
    Call CallStackPop
End Sub

Public Function SafeFiles_Find(ByVal sFile As String, Optional ByVal bRemove As Boolean = False) As Long
    'if file is found, returns index
    'if bRemove, returns ubound after file removed
    Dim iCounter As Long
    SafeFiles_Find = 0
    iCounter = UBound(SafeFiles())
    Do While iCounter <> 0
        If LCase$(SafeFiles(iCounter)) = LCase$(sFile) Then
            If bRemove Then
                Do While iCounter <> UBound(SafeFiles())
                    SafeFiles(iCounter) = SafeFiles(iCounter + 1)
                    iCounter = iCounter + 1
                Loop
                iCounter = iCounter - 1
                ReDim Preserve SafeFiles(iCounter)
            End If
            SafeFiles_Find = iCounter
            Exit Do
        End If
        iCounter = iCounter - 1
    Loop
End Function

Public Function SafeFiles_Add(ByVal sFile As String) As Long
    Dim iUbound As Long
    iUbound = UBound(SafeFiles()) + 1
    ReDim Preserve SafeFiles(iUbound)
    SafeFiles(iUbound) = sFile
    SafeFiles_Add = iUbound
End Function

Public Sub Shutdown(Optional ByVal UnloadMain As Boolean = True, Optional ByVal UnloadModCat As Boolean = True)
    Dim LogSize As Long
    Dim LogBuffer() As String * 1
    Dim OldLogFile As String
    Dim TempString As String
    Dim FileHandle As Integer
    Dim FileHandle2 As Integer
    Dim iCounter As Integer
    Dim iRecord As Integer
    Call CallStackPush(Me.Name & ".Shutdown()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Program terminating.", LogLevel1)
    'Deactivate plugins
    If Not OptPersistentPlugin Then
        Call WriteLogEntry("'Persistent Plugins' is disabled. Deactivating all plugins...", LogLevel1)
        iCounter = 0
        TempString = ReadINIStr("ActivePlugins", CStr(iCounter), ProgramINI)
        Do While Len(TempString) <> 0
            Call DeactivatePlugin(TempString, True)
            iCounter = iCounter + 1
            TempString = ReadINIStr("ActivePlugins", CStr(iCounter), ProgramINI)
        Loop
    End If
    'Save settings
    Call WriteLogEntry("Saving settings.", LogLevel1)
    If Len(FA2DIR) <> 0 Then
        'LB Mod Installers need the following
        If Right$(FA2DIR, 1) = "\" Then
            FA2DIR = Left$(FA2DIR, Len(FA2DIR) - 1)
        End If
    End If
    If Len(FA2DIR) <> 0 Then Call WriteINIStr("General", "FinalAlert2Path", FA2DIR, ProgramINI)
    Call WriteINIStr("General", "SelectedTab", CStr(SelectedTab), ProgramINI)
    Call WriteINIStr("General", "BannerTab", BooleanToYesNo(BannerTab), ProgramINI)
    'remember selected mods
    For iCounter = 0 To MaxType
        If lstMods(iCounter).ListIndex <> -1 Then
            Select Case iCounter
            Case TypePlugin
                Call WriteINIStr("General", "SelectedMod" & CStr(iCounter), GetFileName(Plugins(lstMods(iCounter).ItemData(lstMods(iCounter).ListIndex)).PluginPath), ProgramINI)
            Case TypeFA2Mod
                If lstMods(iCounter).ItemData(lstMods(iCounter).ListIndex) = FA2ModNum Then
                    Call WriteINIStr("General", "SelectedMod" & CStr(iCounter), "fa2yr", ProgramINI)
                Else
                    Call WriteINIStr("General", "SelectedMod" & CStr(iCounter), GetFileName(Mods(lstMods(iCounter).ItemData(lstMods(iCounter).ListIndex)).ModPath), ProgramINI)
                End If
            Case Else
                Call WriteINIStr("General", "SelectedMod" & CStr(iCounter), GetFileName(Mods(lstMods(iCounter).ItemData(lstMods(iCounter).ListIndex)).ModPath), ProgramINI)
            End Select
        End If
    Next iCounter
    'livelog settings
    If menu_livelog.Checked = True Then
        Call WriteINIStr("LiveLog", "LiveLogOpen", "yes", ProgramINI)
    Else
        Call WriteINIStr("LiveLog", "LiveLogOpen", "no", ProgramINI)
    End If
    Call WriteINIStr("LiveLog", "LiveLogTop", CStr(frmLiveLog.Top), ProgramINI)
    Call WriteINIStr("LiveLog", "LiveLogLeft", CStr(frmLiveLog.Left), ProgramINI)
    Call WriteINIStr("LiveLog", "LiveLogWidth", CStr(frmLiveLog.Width), ProgramINI)
    Call WriteINIStr("LiveLog", "LiveLogHeight", CStr(frmLiveLog.Height), ProgramINI)
    Call SafeFiles_Save
    'Consolidate Log File
    If FileExists(LOGFILE) Then
        LogSize = GetFileSize(LOGFILE)
        If (LogSize > OptMaxLogSize) And (OptMaxLogSize <> 0) Then 'consolidate log file.
            Call WriteLogEntry("Consolidating log file.", LogLevel1)
            OldLogFile = LOGFILE & ".old"
            If FileExists(OldLogFile) Then Call Kill(OldLogFile)
            Name LOGFILE As OldLogFile
            FileHandle = FreeFile
            Open OldLogFile For Binary As #FileHandle
            Seek #FileHandle, (LogSize - OptMaxLogSize) + 1
            LogSize = OptMaxLogSize
            ReDim LogBuffer(0)
            'find start of next line
            Do While Not EOF(FileHandle)
                Get #FileHandle, , LogBuffer()
                LogSize = LogSize - 1
                Select Case LogBuffer(0)
                Case vbCr, vbLf
                    Do While Not EOF(FileHandle)
                        Get #FileHandle, , LogBuffer()
                        LogSize = LogSize - 1
                        Select Case LogBuffer(0)
                        Case vbCr, vbLf: 'do nothing
                        Case Else: Exit Do
                        End Select
                    Loop
                    Exit Do
                End Select
            Loop
            If Not EOF(FileHandle) Then
                'write new log file
                FileHandle2 = FreeFile()
                Open LOGFILE For Binary As #FileHandle2
                Put #FileHandle2, , LogBuffer()
                ReDim LogBuffer(LogSize - 1) 'not sure why -1 but last char is HEX(00)
                Get #FileHandle, , LogBuffer()
                Put #FileHandle2, , LogBuffer()
                Close #FileHandle2
                Close #FileHandle
            End If
            Call Kill(OldLogFile)
            Call WriteLogEntry("Log file consolidated.", LogLevel1)
        End If
    End If
    'Close Mutex
    Call WriteLogEntry("Closing mutex.", LogLevel2)
    If MutexValue <> -1 Then Call CloseHandle(MutexValue)
    'Terminate
    Call WriteLogEntry("Unloading forms.", LogLevel2)
    Unload frmLiveLog
    Set frmLiveLog = Nothing
    Unload frmHelp
    Set frmHelp = Nothing
    If UnloadModCat Then
        frmModCat.Busy = True
        Unload frmModCat
        Set frmModCat = Nothing
    End If
    If UnloadMain Then
        Unload frmMain
        Set frmMain = Nothing
    End If
    Call CallStackPop
End Sub

Private Sub Init_LoadSkins()
    Dim fso As FileSystemObject
    Dim fso_root As Folder
    Dim fso_folder As Folder
    Dim iCounter As Integer
    Set fso = New FileSystemObject
    Set fso_root = fso.GetFolder(JoinPath(EXEDIR, "Skins"))
    ReDim SkinPath(fso_root.SubFolders.Count)
    iCounter = 0
    For Each fso_folder In fso_root.SubFolders
        If iCounter <> 0 Then Load menu_skin(iCounter)
        SkinPath(iCounter) = JoinPath(EXEDIR, "Skins\" & fso_folder.Name)
        If FileExists(JoinPath(SkinPath(iCounter), "skin.ini")) Then
            menu_skin(iCounter).Caption = ReadINIStr("General", "Name", JoinPath(SkinPath(iCounter), "skin.ini"))
        Else
            menu_skin(iCounter).Caption = fso_folder.Name
        End If
        If UCase$(fso_folder.Name) = UCase$(DefaultSkinDir) Then menu_skin(iCounter).Caption = menu_skin(iCounter).Caption & " [default]"
        iCounter = iCounter + 1
    Next
    Call LoadSkin(-1)
End Sub

Private Sub Init_RA2Check()
    Dim R1 As String
    Dim R2 As String
    Dim bYR As Boolean
    RA2DIR = ""
    bYR = False
    'check TFD first
    R1 = ReadRegStr("HKLM\SOFTWARE\Electronic Arts\EA Games\Command and Conquer The First Decade\r2_folder")
    If Len(R1) <> 0 Then
        If FileExists(JoinPath(R1, ReadRegStr("HKLM\SOFTWARE\Electronic Arts\EA Games\Command and Conquer The First Decade\r2_executable"))) Then
            RA2DIR = R1
            TFD = True
            'check for presence of YR
            R2 = ReadRegStr("HKLM\SOFTWARE\Electronic Arts\EA Games\Command and Conquer The First Decade\yr_folder")
            If Len(R2) <> 0 Then
                If FileExists(JoinPath(R2, ReadRegStr("HKLM\SOFTWARE\Electronic Arts\EA Games\Command and Conquer The First Decade\yr_executable"))) Then
                    bYR = True
                End If
            End If
        End If
    End If
    'now check original
    R1 = ReadRegStr("HKLM\SOFTWARE\Westwood\Red Alert 2\InstallPath")
    If Len(R1) <> 0 Then
        'original reg entry exists
        If FileExists(R1) Then
            'executable exists
            R1 = GetFilePath(R1)
            If Len(RA2DIR) <> 0 Then
                'we already chose the TFD reg entry, now confirm it
                If GetShortFileName(R1) <> GetShortFileName(RA2DIR) Then
                    'original registry entry points somewhere else - use that instead
                    'unless YR is only present in TFD copy
                    If bYR Then
                        R2 = ReadRegStr("HKLM\SOFTWARE\Westwood\Yuri's Revenge\InstallPath")
                        If Len(R2) <> 0 Then
                            If FileExists(R2) Then
                                RA2DIR = R1
                                TFD = False
                            End If
                        End If
                    Else
                        RA2DIR = R1
                        TFD = False
                    End If
                End If
            Else
                RA2DIR = R1
                TFD = False
            End If
        End If
    End If
    If Len(RA2DIR) <> 0 Then
        Call WriteLogEntry("Red Alert 2 install path established: " & RA2DIR, LogLevel1)
        Call WriteLogEntry("Red Alert 2 installation is " & IIf(TFD, "", "not") & " from The First Decade DVD.")
    Else
        Call WriteLogEntry("Red Alert 2 is not installed! Please install Red Alert 2.", LogShutdown)
    End If
End Sub

Private Sub FA2Check(ByVal bInit As Boolean)
    Dim fso As FileSystemObject
    Call CallStackPush(Me.Name & ".FA2Check(" & CStr(bInit) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    If bInit Then
        FA2DIR = ReadINIStr("General", "FinalAlert2Path", ProgramINI)
        If FA2DIR = "" Then
            FA2DIR = "C:\Program Files\FinalAlert 2 Yuri's Revenge"
            If Not FileExists(JoinPath(FA2DIR, "FinalAlert2YR.exe")) Then
                FA2DIR = "C:\Program Files\FinalAlert 2 - Yuri's Revenge"
                If Not FileExists(JoinPath(FA2DIR, "FinalAlert2YR.exe")) Then
                    FA2DIR = JoinPath(RA2DIR, "fa2yr")
                    If Not FileExists(JoinPath(FA2DIR, "FinalAlert2YR.exe")) Then
                        FA2DIR = ""
                    End If
                End If
            End If
        End If
        txtFA2Folder.Text = FA2DIR
    Else
        FA2DIR = txtFA2Folder.Text
    End If
    If FileExists(JoinPath(FA2DIR, "FinalAlert2YR.exe")) Then
        Set fso = New FileSystemObject
        Mods(FA2ModNum).ModVersion = fso.GetFileVersion(JoinPath(FA2DIR, "FinalAlert2YR.exe"))
        Mods(FA2ModNum).ModSize = GetDirectorySize(FA2DIR)
        Mods(FA2ModNum).ModPath = FA2DIR
    Else
        Mods(FA2ModNum).ModSize = 0
        Mods(FA2ModNum).ModVersion = ""
        If Not bInit And Len(txtFA2Folder.Text) <> 0 Then
            If IsValidPath(FA2DIR) Then
                If DirExists(FA2DIR) Then
                    Call MsgBox("Could not detect FinalAlert 2 YR in the specified directory.", vbOKOnly + vbInformation, App.Title)
                Else
                    Call MsgBox("The specified FinalAlert 2 YR directory does not exist.", vbOKOnly + vbInformation, App.Title)
                End If
            Else
                Call MsgBox("The specified FinalAlert 2 YR directory is not a valid path.", vbOKOnly + vbInformation, App.Title)
            End If
        End If
    End If
    If Len(Mods(FA2ModNum).ModVersion) <> 0 Then
        lstMods(TypeFA2Mod).List(0) = Mods(FA2ModNum).ModName & " [" & Mods(FA2ModNum).ModVersion & "]"
    Else
        lstMods(TypeFA2Mod).List(0) = Mods(FA2ModNum).ModName
    End If
    If lstMods(TypeFA2Mod).ListIndex <> -1 Then Call DisplayModDetails(lstMods(TypeFA2Mod).ItemData(lstMods(TypeFA2Mod).ListIndex), False)
    Call CallStackPop
End Sub

Private Sub Init_LoadPlugins()
    Dim iPlugin As Integer
    Dim sPlugin As String
    Dim sFile As String
    Dim iCounter As Integer
    Dim bCounted As Boolean
    Dim sVersion As String
    Dim iTemp As Integer
    Dim bOk As Boolean
    Dim fso As FileSystemObject
    Dim fso_mods As Folder
    Dim fso_folder As Folder
    Dim fso_file As File
    Call CallStackPush(Me.Name & ".Init_LoadPlugins()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Loading plugins...", LogLevel1)
    Call frmSplash.PROGRESS("Loading plugins... ")
    If DirExists(JoinPath(EXEDIR, "Mods")) Then
        Set fso = New FileSystemObject
        Set fso_mods = fso.GetFolder(JoinPath(EXEDIR, "Mods"))
        PluginCount = 0
        For Each fso_folder In fso_mods.SubFolders
            sFile = JoinPath(fso_folder.Path, "launcher\liblist.gam")
            If FileExists(sFile) Then
                PluginCount = PluginCount + 1
                ReDim Preserve Plugins(PluginCount)
                Plugins(PluginCount).PluginPath = fso_folder.Path
                Plugins(PluginCount).PluginID = ReadINIStr("General", "PluginID", sFile)
                If Len(Plugins(PluginCount).PluginID) = 0 Then
                    'backwards compatibility for old liblist.gam
                    If BooleanStringToBoolean(ReadINIStr("General", "IsTX", sFile, "no")) Then
                        Plugins(PluginCount).PluginID = "TX"
                    End If
                End If
                Plugins(PluginCount).PluginName = ReadINIStr("General", "Name", sFile)
                Plugins(PluginCount).PluginVersion = ReadINIStr("General", "Version", sFile)
                Plugins(PluginCount).PluginDate = ReadINIStr("General", "Date", sFile)
                Plugins(PluginCount).PluginAuthor = ReadINIStr("General", "Author", sFile)
                Plugins(PluginCount).PluginWebsite = ReadINIStr("General", "Website", sFile)
                Plugins(PluginCount).PluginSize = GetDirectorySize(Plugins(PluginCount).PluginPath)
                Plugins(PluginCount).PluginDescription = ReadINIStr("General", "Description", sFile)
                Plugins(PluginCount).PluginManual = ""
                If DirExists(JoinPath(Plugins(PluginCount).PluginPath, "manual")) Then
                    Set fso_folder = fso.GetFolder(JoinPath(Plugins(PluginCount).PluginPath, "manual"))
                    For Each fso_file In fso_folder.Files
                        If Len(fso_file.Name) > 6 Then
                            If LCase$(Left$(fso_file.Name, 6)) = "index." Then
                                Plugins(PluginCount).PluginManual = fso_file.Path
                                Exit For
                            End If
                        End If
                    Next fso_file
                End If
                'Validate (except security check, which takes place on activate)
                bOk = False
                If Len(StripNumbers(Plugins(PluginCount).PluginID)) <> 0 Then 'plugin id can't be blank or wholly numeric
                    If Len(Plugins(PluginCount).PluginName) <> 0 Then 'name can't be blank
                        If Len(Plugins(PluginCount).PluginVersion) <> 0 Then 'version can't be blank for plugins
                            bOk = True
                            'add to list of plugins
                            iCounter = 0
                            bCounted = False
                            Do While iCounter < lstMods(TypePlugin).ListCount
                                If Plugins(lstMods(TypePlugin).ItemData(iCounter)).PluginID = Plugins(PluginCount).PluginID Then
                                    bCounted = True
                                    If CompareVersions(Plugins(PluginCount).PluginVersion, ">=", Plugins(lstMods(TypePlugin).ItemData(iCounter)).PluginVersion) Then lstMods(TypePlugin).ItemData(iCounter) = PluginCount
                                    Exit Do
                                End If
                                iCounter = iCounter + 1
                            Loop
                            If Not bCounted Then
                                Call lstMods(TypePlugin).AddItem(Plugins(PluginCount).PluginID)
                                lstMods(TypePlugin).ItemData(lstMods(TypePlugin).NewIndex) = PluginCount
                            End If
                        End If
                    End If
                End If
                If Not bOk Then PluginCount = PluginCount - 1
            End If
            If PluginCount <> 0 Then Call frmSplash.PROGRESS("Loading plugins... " & Plugins(PluginCount).PluginName)
        Next
        'sort lstMods alphabetically (by name of latest version) before we update the list entries with the actual versions active
        iPlugin = 0
        Do While iPlugin < lstMods(TypePlugin).ListCount
            iCounter = 0
            Do While iCounter < ((lstMods(TypePlugin).ListCount - 1) - iPlugin)
                If Plugins(lstMods(TypePlugin).ItemData(iCounter)).PluginName > Plugins(lstMods(TypePlugin).ItemData(iCounter + 1)).PluginName Then
                    iTemp = lstMods(TypePlugin).ItemData(iCounter + 1)
                    lstMods(TypePlugin).ItemData(iCounter + 1) = lstMods(TypePlugin).ItemData(iCounter)
                    lstMods(TypePlugin).ItemData(iCounter) = iTemp
                End If
                iCounter = iCounter + 1
            Loop
            iPlugin = iPlugin + 1
        Loop
        Call frmSplash.PROGRESS
        'now check for tampering and update list entry for each plugin
        iPlugin = 0
        Do While iPlugin < lstMods(TypePlugin).ListCount
            If Init_LoadPlugins_NoTampering(Plugins(lstMods(TypePlugin).ItemData(iPlugin)).PluginID) Then 'NoTampering also does the neccessary tasks if tampering is found
                sVersion = ReadINIStr("Plugin" & Plugins(lstMods(TypePlugin).ItemData(iPlugin)).PluginID, "Version", ProgramINI)
                If Len(sVersion) <> 0 Then
                    bCounted = False
                    iCounter = 0
                    Do While iCounter <= PluginCount
                        If Plugins(iCounter).PluginVersion = sVersion Then
                            bCounted = True
                            'update list entry
                            lstMods(TypePlugin).List(iPlugin) = Plugins(lstMods(TypePlugin).ItemData(iPlugin)).PluginName & " [" & Plugins(lstMods(TypePlugin).ItemData(iPlugin)).PluginVersion & "]"
                            Exit Do
                        End If
                        iCounter = iCounter + 1
                    Loop
                    If Not bCounted Then
                        'a version is active in the game that no longer exists in the Plugins folder!
                        Call DeactivatePlugin(Plugins(lstMods(TypePlugin).ItemData(iPlugin)).PluginID)
                    End If
                Else
                    'update list entry
                    iCounter = GetLatestPlugin(Plugins(lstMods(TypePlugin).ItemData(iPlugin)).PluginID)
                    lstMods(TypePlugin).List(iPlugin) = Plugins(iCounter).PluginName & " [Not Active]"
                    lstMods(TypePlugin).ItemData(iPlugin) = iCounter
                End If
            End If
            iPlugin = iPlugin + 1
        Loop
        Call frmSplash.PROGRESS
    End If
    'check any active plugins
    iPlugin = 0
    sPlugin = ReadINIStr("ActivePlugins", CStr(iPlugin), ProgramINI)
    Do While Len(sPlugin) <> 0
        iCounter = GetLatestPlugin(sPlugin)
        If iCounter = -1 Then
            'a plugin is active in the game that no longer exists in the Plugins folder!
            Call DeactivatePlugin(sPlugin)
        Else
            If GetActivePlugin(sPlugin) = -1 Then
                'plugin not active so need to reactivate (this can only happen if OptPersistentPlugins is disabled)
                Call WriteLogEntry("Reactivating plugin: " & Plugins(iCounter).PluginName, LogLevel1)
                If AuthenticatePlugin(iCounter) Then Call ActivatePlugin(iCounter)
            End If
        End If
        iPlugin = iPlugin + 1
        sPlugin = ReadINIStr("ActivePlugins", CStr(iPlugin), ProgramINI)
    Loop
    If lstMods(TypePlugin).ListCount <> 0 Then lstMods(TypePlugin).ListIndex = 0
    Call frmSplash.PROGRESS
    Call CallStackPop
End Sub

Private Function ActivatePlugin(ByVal iPlugin As Integer) As Boolean
    Dim iPluginSize As Long
    Dim iFiles As Long
    Dim iCounter As Long
    Dim sFile As String
    Dim sSourcePath As String
    Dim sDestPath As String
    Dim sBackPath As String
    Dim bOk As Boolean
    Dim fso As FileSystemObject
    Dim fso_folder As Folder
    Dim fso_file As File
    Call CallStackPush(Me.Name & ".ActivatePlugin(" & CStr(iPlugin) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    ActivatePlugin = False
    'Check for free disk space
    If ActivatePlugin_CheckDiskSpace(iPlugin) Then
        Call DeactivatePlugin(Plugins(iPlugin).PluginID) 'does nothing if PluginID not active
        'Show please wait dialog (must come after any actions that might also show the dialog, such as Authenticate)
        Call ShowPleaseWait("Activating plugin: " & Plugins(iPlugin).PluginName & " [" & Plugins(iPlugin).PluginVersion & "]")
        Call WriteLogEntry("Activating plugin: " & Plugins(iPlugin).PluginName & " [" & Plugins(iPlugin).PluginVersion & "]", LogLevel1)
        'activate the files
        Set fso = New FileSystemObject
        Set fso_folder = fso.GetFolder(JoinPath(Plugins(iPlugin).PluginPath, "video"))
        iPluginSize = 0
        iFiles = 0
        For Each fso_file In fso_folder.Files
            Call UpdatePleaseWait(, fso_file.Name)
            bOk = False
            If Not FileIsDestructive(fso_file.Name) Then
                If FileIsModfile(fso_file.Name) Then
                    If Not FileIsUserTheme(fso_file.Name) Then
                        If Not FileIsAresComponent(fso_file.Name) Then bOk = True
                    End If
                End If
            End If
            If bOk Then
                sSourcePath = fso_file.Path
                sDestPath = JoinPath(RA2DIR, fso_file.Name)
                sBackPath = JoinPath(BACKUPDIR, fso_file.Name)
                If FileExists(sDestPath) Then
                    If SafeFiles_Find(fso_file.Name) <> 0 Then
                        Call WriteLogEntry("Plugin is trying to replace a residual mod file that has been marked as safe [" & sDestPath & "]. Marking file as unsafe...")
                        Call MsgBox("A residual mod file that you have marked as safe [" & sFile & "] will be replaced by this plugin. Clearly the file is not in fact safe." & vbCrLf & "The file will now be marked as unsafe and treated as any other residual file." & vbCrLf & "It is strongly recommended that you re-read the Help Topics and consider disabling Advanced Mode.", vbOKOnly + vbExclamation, App.Title)
                        Call SafeFiles_Find(fso_file.Name, True)
                    End If
                    If FileExists(sBackPath) Then
                        Call WriteLogEntry("Unexpected backup file found! Deleting " & Quote(sBackPath) & " to make way for new file.")
                        Call Kill(sBackPath)
                    End If
                    Call LoggedMove(sDestPath, sBackPath, True)
                End If
                Call LoggedCopy(sSourcePath, sDestPath)
                Call WriteINIStr("Plugin" & Plugins(iPlugin).PluginID, CStr(iFiles), fso_file.Name, ProgramINI)
                iPluginSize = iPluginSize + GetFileSize(sSourcePath, True)
                If OptUseCheckSums Then Call WriteINIStr("Plugin" & Plugins(iPlugin).PluginID, CStr(iFiles) & "c", GetFileMD5(sSourcePath), ProgramINI)
                iFiles = iFiles + 1
            Else
                Call WriteLogEntry(Quote(sSourcePath) & " rejected for copying to " & Quote(sDestPath))
            End If
        Next
        'Update record of what version is active
        Call WriteINIStr("Plugin" & Plugins(iPlugin).PluginID, "Name", Plugins(iPlugin).PluginName, ProgramINI)
        Call WriteINIStr("Plugin" & Plugins(iPlugin).PluginID, "Version", Plugins(iPlugin).PluginVersion, ProgramINI)
        Call WriteINIStr("Plugin" & Plugins(iPlugin).PluginID, "DiskUsage", CStr(iPluginSize), ProgramINI)
        iCounter = 0
        sFile = ReadINIStr("ActivePlugins", CStr(iCounter), ProgramINI)
        Do While Len(sFile) <> 0
            If sFile = Plugins(iPlugin).PluginID Then Exit Do
            iCounter = iCounter + 1
            sFile = ReadINIStr("ActivePlugins", CStr(iCounter), ProgramINI)
        Loop
        Call WriteINIStr("ActivePlugins", CStr(iCounter), Plugins(iPlugin).PluginID, ProgramINI)
        Call WriteINIStr("ActivePlugins", CStr(iCounter) & "v", Plugins(iPlugin).PluginVersion, ProgramINI)
        iCounter = 0
        Do While iCounter < lstMods(TypePlugin).ListCount
            If Plugins(lstMods(TypePlugin).ItemData(iCounter)).PluginID = Plugins(iPlugin).PluginID Then
                lstMods(TypePlugin).List(iCounter) = Plugins(iPlugin).PluginName & " [" & Plugins(iPlugin).PluginVersion & "]"
                lstMods(TypePlugin).ItemData(iCounter) = iPlugin
                Exit Do
            End If
            iCounter = iCounter + 1
        Loop
        If Val(cmdLaunch(TypePlugin).Tag) <> -1 Then
            If Val(cmdLaunch(TypePlugin).Tag) = iPlugin Then Call DisplayPluginDetails(Val(cmdLaunch(TypePlugin).Tag))
        End If
        If lstMods(TypeMod).ListIndex <> -1 Then Call DisplayModDetails(lstMods(TypeMod).ItemData(lstMods(TypeMod).ListIndex))
        If lstMods(TypeFA2Mod).ListIndex <> -1 Then Call DisplayModDetails(lstMods(TypeFA2Mod).ItemData(lstMods(TypeFA2Mod).ListIndex))
        Call WriteLogEntry("Plugin activated.", LogLevel1)
        ActivatePlugin = True
    End If
    Call HidePleaseWait
    Call CallStackPop
End Function

Private Function ActivatePlugin_CheckDiskSpace(ByVal iPlugin As Integer) As Boolean
    Dim sPath As String
    Dim iRequired As Double
    sPath = JoinPath(Plugins(iPlugin).PluginPath, "THEME")
    If DirExists(sPath) Then iRequired = iRequired + GetDirectorySize(sPath, True, False)
    sPath = JoinPath(Plugins(iPlugin).PluginPath, "VIDEO")
    If DirExists(sPath) Then iRequired = iRequired + GetDirectorySize(sPath, True, False)
    iRequired = iRequired + OptSafetySpace
    If FreeDiskSpace(UCase$(Left$(RA2DIR, 1))) >= iRequired Then
        ActivatePlugin_CheckDiskSpace = True
    Else
        Call WriteLogEntry("Insufficient free disk space to activate plugin. This plugin requires at least " & DataSize(iRequired) & " free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & ".", LogMsgBoxExclaim)
        ActivatePlugin_CheckDiskSpace = False
    End If
End Function

Private Sub DeactivatePlugin(ByVal sPluginID As String, Optional ByVal KeepPersistent As Boolean = False, Optional ByVal CalledFromInit As Boolean = False)
    Dim iCounter As Long
    Dim sFile As String
    Dim sPath As String
    Dim sBackup As String
    Dim sName As String
    Dim iLatest As Integer
    Dim mbResult As VbMsgBoxResult
    Call CallStackPush(Me.Name & ".DeactivatePlugin(" & CStr(sPluginID) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    sName = ReadINIStr("Plugin" & sPluginID, "Name", ProgramINI)
    If Len(sName) <> 0 Then
        sName = sName & " [" & ReadINIStr("Plugin" & sPluginID, "Version", ProgramINI) & "]"
        Call WriteLogEntry("Deactivating plugin: " & sName, LogLevel1)
        Call ShowPleaseWait("Deactivating plugin: " & sName)
        iCounter = 0
        sFile = ReadINIStr("Plugin" & sPluginID, CStr(iCounter), ProgramINI)
        Do While Len(sFile) <> 0
            Call UpdatePleaseWait(, sFile)
            sPath = JoinPath(RA2DIR, sFile)
            sBackup = JoinPath(BACKUPDIR, sFile)
            If FileExists(sPath) Then
                Call LoggedKill(sPath)
            Else
                Call WriteLogEntry("Expected file " & Quote(sPath) & " not found!")
            End If
            If FileExists(sBackup) Then
CheckDiskSpace:
                If FreeDiskSpace(UCase$(Left$(RA2DIR, 1))) > GetFileSize(sBackup, True) Then
                    Call LoggedMove(sBackup, sPath, False, True)
                Else
                    mbResult = MsgBox("There is insufficient free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & " to restore residual file " & vbCrLf & Quote(sBackup) & vbCrLf & " to the Red Alert 2 directory." & vbCrLf & "This file will NOT be restored unless you free up some disk space now." & vbCrLf & vbCrLf & "Abort - skip this file and leave it in the Launch Base Backup folder." & vbCrLf & "Retry - check for disk space again and restore the file." & vbCrLf & "Ignore - permanently delete this file.", vbAbortRetryIgnore + vbDefaultButton2 + vbExclamation, App.Title)
                    If mbResult = vbIgnore Then
                        Call WriteLogEntry("Insufficient free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & " to restore residual file " & Quote(sBackup) & ". User has chosen to delete the file.")
                        Call LoggedKill(sBackup)
                    Else
                        If mbResult = vbRetry Then
                            GoTo CheckDiskSpace
                        Else
                            Call WriteLogEntry("Insufficient free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & " to restore residual file " & Quote(sBackup) & ". User has chosen to leave this file in the Launch Base Backup folder.")
                        End If
                    End If
                End If
            End If
            Call WriteINIStr("Plugin" & sPluginID, CStr(iCounter), "", ProgramINI)
            Call WriteINIStr("Plugin" & sPluginID, CStr(iCounter) & "c", "", ProgramINI)
            iCounter = iCounter + 1
            sFile = ReadINIStr("Plugin" & sPluginID, CStr(iCounter), ProgramINI)
        Loop
        'Update record of what version is active
        Call WriteINIStr("Plugin" & sPluginID, "Name", "", ProgramINI)
        Call WriteINIStr("Plugin" & sPluginID, "Version", "", ProgramINI)
        Call WriteINIStr("Plugin" & sPluginID, "DiskUsage", "", ProgramINI)
        If Not KeepPersistent Then
            iCounter = 0
            sFile = ReadINIStr("ActivePlugins", CStr(iCounter), ProgramINI)
            Do While Len(sFile) <> 0
                If sFile = sPluginID Then
                    Do While Len(sFile) <> 0
                        Call WriteINIStr("ActivePlugins", CStr(iCounter), ReadINIStr("ActivePlugins", CStr(iCounter + 1), ProgramINI), ProgramINI)
                        iCounter = iCounter + 1
                        sFile = ReadINIStr("ActivePlugins", CStr(iCounter), ProgramINI)
                    Loop
                End If
                iCounter = iCounter + 1
                sFile = ReadINIStr("ActivePlugins", CStr(iCounter), ProgramINI)
            Loop
        End If
        'Update lstMods
        iLatest = GetLatestPlugin(sPluginID)
        iCounter = 0
        Do While iCounter < lstMods(TypePlugin).ListCount
            If Plugins(lstMods(TypePlugin).ItemData(iCounter)).PluginID = sPluginID Then
                lstMods(TypePlugin).List(iCounter) = Plugins(iLatest).PluginName & " [Not Active]"
                lstMods(TypePlugin).ItemData(iCounter) = iLatest
                Exit Do
            End If
            iCounter = iCounter + 1
        Loop
        If Len(cmdLaunch(TypePlugin).Tag) <> 0 Then
            If Plugins(Val(cmdLaunch(TypePlugin).Tag)).PluginID = sPluginID Then Call DisplayPluginDetails(Val(cmdLaunch(TypePlugin).Tag))
            If lstMods(TypeMod).ListIndex <> -1 Then Call DisplayModDetails(lstMods(TypeMod).ItemData(lstMods(TypeMod).ListIndex))
            If lstMods(TypeFA2Mod).ListIndex <> -1 Then Call DisplayModDetails(lstMods(TypeFA2Mod).ItemData(lstMods(TypeFA2Mod).ListIndex))
        End If
        Call WriteLogEntry("Plugin deactivated.", LogLevel1)
        Call HidePleaseWait
    End If
    Call CallStackPop
End Sub

Public Function GetLatestPlugin(ByVal sPluginID As String) As Integer
    'Returns plugin index (or -1 if no such plugin)
    Dim iPlugin As Integer
    Dim iLatest As Integer
    iLatest = -1
    iPlugin = 1
    Do While iPlugin <= PluginCount
        If Plugins(iPlugin).PluginID = sPluginID Then
            If iLatest = -1 Then
                iLatest = iPlugin
            Else
                If CompareVersions(Plugins(iPlugin).PluginVersion, ">=", Plugins(iLatest).PluginVersion) Then iLatest = iPlugin
            End If
        End If
        iPlugin = iPlugin + 1
    Loop
    GetLatestPlugin = iLatest
End Function

Private Function GetActivePlugin(ByVal sPluginID As String) As Integer
    Dim iCounter As Integer
    Dim iPlugin As Integer
    iPlugin = -1
    iCounter = 0
    Do While iCounter < lstMods(TypePlugin).ListCount
        If Plugins(lstMods(TypePlugin).ItemData(iCounter)).PluginID = sPluginID Then
            If ReadINIStr("Plugin" & sPluginID, "Version", ProgramINI) = Plugins(lstMods(TypePlugin).ItemData(iCounter)).PluginVersion Then
                iPlugin = lstMods(TypePlugin).ItemData(iCounter)
            End If
            Exit Do
        End If
        iCounter = iCounter + 1
    Loop
    GetActivePlugin = iPlugin
End Function

Public Function GetLatestMod(ByVal sModName As String, ByVal iModType As Integer, Optional ByVal sUpdateCheckURL As String = "") As Integer
    Dim iMod As Integer
    Dim iLatest As Integer
    Dim bOk As Boolean
    iLatest = -1
    iMod = HardCodedMods
    Do While iMod <= ModCount
        bOk = False
        If Mods(iMod).ModUpdateCheckURL = sUpdateCheckURL Then
            bOk = True
        Else
            If Mods(iMod).ModName = sModName Then
                If Mods(iMod).ModType = iModType Then bOk = True
            End If
        End If
        If bOk Then
            If iLatest = -1 Then
                iLatest = iMod
            Else
                If CompareVersions(Mods(iMod).ModVersion, ">=", Mods(iLatest).ModVersion) Then iLatest = iMod
            End If
        End If
        iMod = iMod + 1
    Loop
    GetLatestMod = iLatest
End Function

Private Function Init_LoadPlugins_NoTampering(ByVal sPlugin As String) As Boolean
    Dim iFile As Integer
    Dim sFile As String
    Dim sPath As String
    Dim bOk As Boolean
    bOk = True
    iFile = 0
    sFile = ReadINIStr("Plugin" & sPlugin, CStr(iFile), ProgramINI)
    Do While Len(sFile) <> 0
        sPath = JoinPath(RA2DIR, sFile)
        If FileExists(sPath) Then
            If OptUseCheckSums Then
                sFile = ReadINIStr("Plugin" & sPlugin, CStr(iFile) & "c", ProgramINI)
                If Len(sFile) <> 0 Then
                    Call WriteLogEntry("Getting MD5 of " & Quote(sPath), LogLevel2)
                    If UCase$(GetFileMD5(sPath)) <> UCase$(sFile) Then
                        bOk = False
                        Exit Do
                    End If
                End If
            End If
        Else
            bOk = False
            Exit Do
        End If
        iFile = iFile + 1
        sFile = ReadINIStr("Plugin" & sPlugin, CStr(iFile), ProgramINI)
    Loop
    If Not bOk Then
        TamperingDetected = True
        Call WriteLogEntry(ReadINIStr("Plugin" & sPlugin, "Name", ProgramINI) & " plugin files have been tampered with outside of Launch Base!")
        Call DenyPersistentMods
        Call DeactivatePlugin(sPlugin)
    End If
    Init_LoadPlugins_NoTampering = bOk
End Function

Private Sub Init_LoadMods()
    Dim iCounter As Integer
    Dim iScrnNumFormat As Long
    Dim sVersion As String
    Dim bOk As Boolean
    Dim fso As FileSystemObject
    Dim fso_mods As Folder
    Dim fso_folder As Folder
    Dim fso_file As File
    Dim sPersistentName As String
    Dim sPersistentVersion As String
    Dim bPersistentFound As Boolean
    Dim bPersistentMatch As Boolean
    Call CallStackPush(Me.Name & ".Init_LoadMods()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Loading mods...", LogLevel1)
    ModCount = 3
    ReDim Mods(10)
    'Launch Base
    Mods(LBModNum).ModName = App.Title
    Select Case App.Revision
    Case 0: Mods(LBModNum).ModVersion = App.Major & "." & PadNum(App.Minor, 2)
    Case Else: Mods(LBModNum).ModVersion = App.Major & "." & PadNum(App.Minor, 2) & "." & App.Revision
    End Select
    Mods(LBModNum).ModType = TypeProgram
    Mods(LBModNum).ModPath = EXEDIR
    Mods(LBModNum).ModUpdateCheckURL = ReadINIStr("URL", "LBUpdateCheckURL", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/LaunchBase.upd")
    'FinalAlert 2 YR
    Mods(FA2ModNum).ModName = "FinalAlert 2 YR (unmodded)"
    Mods(FA2ModNum).ModVersion = ""
    Mods(FA2ModNum).ModDate = ""
    Mods(FA2ModNum).ModAuthor = ""
    Mods(FA2ModNum).ModWebsite = ""
    Mods(FA2ModNum).ModSize = 0
    Mods(FA2ModNum).ModDescription = "FinalAlert 2 YR, original and unmodded. Run this if you want to create maps that are compatible with any mod."
    Mods(FA2ModNum).ModBanner = JoinPath(RESDIR, "fabanner.bmp")
    Mods(FA2ModNum).ModAllowTX = True
    Mods(FA2ModNum).ModType = TypeFA2Mod
    Mods(FA2ModNum).ModUpdateCheckURL = ReadINIStr("URL", "FA2UpdateCheckURL", ProgramINI, "")
    Mods(FA2ModNum).ModPath = JoinPath(EXEDIR, "Mods\fa2yr")
    'set up RA2 folder
    Mods(RA2ModNum).ModPath = JoinPath(EXEDIR, "Mods\originalra2")
    If Not DirExists(Mods(RA2ModNum).ModPath) Then Call MakePath(Mods(RA2ModNum).ModPath)
    'set up YR folder
    Mods(YRModNum).ModPath = JoinPath(EXEDIR, "Mods\originalyr")
    If Not DirExists(Mods(YRModNum).ModPath) Then Call MakePath(Mods(YRModNum).ModPath)
    'make a note of what mod is persistent
    sPersistentName = ReadINIStr("Mod", "Name", ProgramINI)
    sPersistentVersion = ReadINIStr("Mod", "Version", ProgramINI)
    bPersistentFound = False
    bPersistentMatch = False
    'scan the Mods folder
    Set fso = New FileSystemObject
    Set fso_mods = fso.GetFolder(JoinPath(EXEDIR, "Mods"))
    For Each fso_folder In fso_mods.SubFolders
        Select Case UCase$(fso_folder.Name)
        Case "ORIGINALRA2"
            If Not TFD Then
                sVersion = ReadRegStr("HKLM\SOFTWARE\Westwood\Red Alert 2\Version")
                Select Case sVersion
                Case "65536"
                    Mods(RA2ModNum).ModVersion = "1.000"
                    Mods(RA2ModNum).ModDate = "2000-09-23"
                Case "65537"
                    Mods(RA2ModNum).ModVersion = "1.001"
                    Mods(RA2ModNum).ModDate = "2000-09-23 to ?"
                Case "65538"
                    Mods(RA2ModNum).ModVersion = "1.002"
                    Mods(RA2ModNum).ModDate = "2000-09-23 to ?"
                Case "65539"
                    Mods(RA2ModNum).ModVersion = "1.003"
                    Mods(RA2ModNum).ModDate = "2000-09-23 to ?"
                Case "65540"
                    Mods(RA2ModNum).ModVersion = "1.004"
                    Mods(RA2ModNum).ModDate = "2000-09-23 to ?"
                Case "65541"
                    Mods(RA2ModNum).ModVersion = "1.005"
                    Mods(RA2ModNum).ModDate = "2000-09-23 to ?"
                Case "65542"
                    Mods(RA2ModNum).ModVersion = "1.006"
                    Mods(RA2ModNum).ModDate = "2000-09-23 to 2001-05-25"
                Case Else: Call WriteLogEntry("Your Red Alert 2 installation is damaged." & vbCrLf & "Unrecognized Red Alert 2 version." & vbCrLf & "Please reinstall Red Alert 2.", LogShutdown)
                End Select
            Else
                Mods(RA2ModNum).ModVersion = "1.006"
                Mods(RA2ModNum).ModDate = "2000-09-23 to 2001-05-25"
            End If
            OptRA2Lang = 0
            Call WriteLogEntry("Launch Base does not yet have a language check facility. It is assumed that your Red Alert 2 language is English (US).")
            If Mods(RA2ModNum).ModVersion <> "1.006" Then
                Select Case OptRA2Lang
                Case 0: Mods(RA2ModNum).ModUpdateCheckURL = ReadINIStr("URL", "RA2UpdateCheckURLEnglish", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/ra2_english.upd")
                Case 2: Mods(RA2ModNum).ModUpdateCheckURL = ReadINIStr("URL", "RA2UpdateCheckURLGerman", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/ra2_german.upd")
                Case 3: Mods(RA2ModNum).ModUpdateCheckURL = ReadINIStr("URL", "RA2UpdateCheckURLFrench", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/ra2_french.upd")
                Case 8
                    Mods(RA2ModNum).ModUpdateCheckURL = ReadINIStr("URL", "RA2UpdateCheckURLKorean", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/ra2_korean.upd")
                    Call MsgBox("Launch Base is missing some information about the Korean version of Red Alert 2." & vbCrLf & "Please contact Marshall so that this information can be added to Launch Base.", vbOKOnly, App.Title)
                Case 9
                    Mods(RA2ModNum).ModUpdateCheckURL = ReadINIStr("URL", "RA2UpdateCheckURLChinese", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/ra2_chinese.upd")
                    Call MsgBox("Launch Base is missing some information about the Chinese version of Red Alert 2." & vbCrLf & "Please contact Marshall so that this information can be added to Launch Base.", vbOKOnly, App.Title)
                End Select
            End If
            Mods(RA2ModNum).ModName = "Red Alert 2 (unmodded)"
            Mods(RA2ModNum).ModAuthor = "Westwood Studios"
            Mods(RA2ModNum).ModWebsite = "http://www.westwood.com"
            Call WriteLogEntry("Getting size of RA2 directory.", LogLevel2)
            Mods(RA2ModNum).ModSize = GetDirectorySize(RA2DIR)
            Mods(RA2ModNum).ModDescription = "Red Alert 2, original and unmodded."
            Mods(RA2ModNum).ModCampaigns = "Original Campaigns"
            Mods(RA2ModNum).ModBanner = JoinPath(RESDIR, "ra2banner.bmp")
            Mods(RA2ModNum).ModAllowTX = True
            Mods(RA2ModNum).ModType = TypeMod
            Mods(RA2ModNum).ModScrnFormat = "SCRN%04d.pcx"
            Mods(RA2ModNum).ModSnapFormat = "Map%04d.yrm"
            Mods(RA2ModNum).ModIsForRA2 = True
            Mods(RA2ModNum).ModGameMode = "1"
            Mods(RA2ModNum).ModMapIndex = "0"
            Mods(RA2ModNum).ModUseAres = False
        Case "ORIGINALYR"
            sVersion = ""
            If Not TFD Then
                sVersion = ReadRegStr("HKLM\SOFTWARE\Westwood\Yuri's Revenge\InstallPath")
            Else
                sVersion = ReadRegStr("HKLM\SOFTWARE\Electronic Arts\EA Games\Command and Conquer The First Decade\yr_folder")
                If Len(sVersion) <> 0 Then sVersion = JoinPath(sVersion, ReadRegStr("HKLM\SOFTWARE\Electronic Arts\EA Games\Command and Conquer The First Decade\yr_executable"))
            End If
            If Len(sVersion) <> 0 Then
                If FileExists(sVersion) Then
                    If Not TFD Then
                        sVersion = ReadRegStr("HKLM\SOFTWARE\Westwood\Yuri's Revenge\Version")
                        Select Case sVersion
                        Case "65536"
                             Mods(YRModNum).ModVersion = "1.000"
                             Mods(YRModNum).ModDate = "2001-10-10"
                        Case "65537"
                             Mods(YRModNum).ModVersion = "1.001"
                             Mods(YRModNum).ModDate = "2001-10-10 to 2001-11-12"
                        End Select
                    Else
                        Mods(YRModNum).ModVersion = "1.001"
                        Mods(YRModNum).ModDate = "2001-10-10 to 2001-11-12"
                    End If
                    If Mods(YRModNum).ModVersion = "1.001" Then
                        If Not FileExists(JoinPath(RA2DIR, "expandmd01.mix")) Then Call WriteLogEntry("Your Yuri's Revenge 1.001 installation is damaged." & vbCrLf & "<expandmd01.mix> is missing." & vbCrLf & "Please reinstall Yuri's Revenge version 1.001.", LogShutdown)
                    End If
                Else
                    Call WriteLogEntry("Your Yuri's Revenge installation is damaged." & vbCrLf & "Unrecognized Yuri's Revenge version." & vbCrLf & "Please reinstall Yuri's Revenge.", LogShutdown)
                End If
            Else
                WriteLogEntry ("Yuri's Revenge is not installed!")
                Mods(YRModNum).ModVersion = ""
            End If
            OptYRLang = 0
            Call WriteLogEntry("Launch Base does not yet have a language check facility. It is assumed that your Yuri's Revenge language is English (US).")
            If Mods(YRModNum).ModVersion <> "1.001" Then
                Call MsgBox("TODO: Language of YR is unknown! Update checks are for English version only!")
                Select Case OptYRLang
                Case 0: Mods(YRModNum).ModUpdateCheckURL = ReadINIStr("URL", "YRUpdateCheckURLEnglish", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/yr_english.upd")
                Case 2: Mods(YRModNum).ModUpdateCheckURL = ReadINIStr("URL", "YRUpdateCheckURLGerman", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/yr_german.upd")
                Case 3: Mods(YRModNum).ModUpdateCheckURL = ReadINIStr("URL", "YRUpdateCheckURLFrench", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/yr_french.upd")
                Case 8
                    Mods(YRModNum).ModUpdateCheckURL = ReadINIStr("URL", "YRUpdateCheckURLKorean", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/yr_korean.upd")
                    Call MsgBox("Launch Base is missing some information about the Korean version of Yuri's Revenge." & vbCrLf & "Please contact Marshall so that this information can be added to Launch Base.", vbOKOnly, App.Title)
                Case 9
                    Mods(YRModNum).ModUpdateCheckURL = ReadINIStr("URL", "YRUpdateCheckURLChinese", ProgramINI, "http://marshall.strategy-x.com/LaunchBase/yr_chinese.upd")
                    Call MsgBox("Launch Base is missing some information about the Chinese version of Yuri's Revenge." & vbCrLf & "Please contact Marshall so that this information can be added to Launch Base.", vbOKOnly, App.Title)
                End Select
            End If
            Mods(YRModNum).ModName = "Yuri's Revenge (unmodded)"
            Mods(YRModNum).ModAuthor = "Westwood Studios"
            Mods(YRModNum).ModWebsite = "http://www.westwood.com"
            Mods(YRModNum).ModSize = Mods(RA2ModNum).ModSize
            Mods(YRModNum).ModDescription = "Yuri's Revenge, original and unmodded. Only run this if you want to play online. If you are playing single player / LAN, run the UMP instead. Find out about the UMP at http://marshall.strategy-x.com/ump"
            Mods(YRModNum).ModCampaigns = "Original Campaigns"
            Mods(YRModNum).ModBanner = JoinPath(RESDIR, "yrbanner.bmp")
            Mods(YRModNum).ModAllowTX = True
            Mods(YRModNum).ModType = TypeMod
            Mods(YRModNum).ModScrnFormat = "SCRN%04d.pcx"
            Mods(YRModNum).ModSnapFormat = "Map%04d.yrm"
            Mods(YRModNum).ModGameMode = "1"
            Mods(YRModNum).ModMapIndex = "0"
            Mods(YRModNum).ModUseAres = False
        Case Else
            If FileExists(JoinPath(fso_folder.Path, "launcher\liblist.gam")) Then
                ModCount = ModCount + 1
                If UBound(Mods()) < ModCount Then ReDim Preserve Mods(ModCount + 10)
                Mods(ModCount).ModPath = fso_folder.Path
                Mods(ModCount).ModLiblist = JoinPath(Mods(ModCount).ModPath, "launcher\liblist.gam")
                If Len(ReadINIStr("General", "PluginID", Mods(ModCount).ModLiblist)) = 0 Then
                    Mods(ModCount).ModName = ReadINIStr("General", "Name", Mods(ModCount).ModLiblist)
                    Mods(ModCount).ModVersion = ReadINIStr("General", "Version", Mods(ModCount).ModLiblist, "")
                    Mods(ModCount).ModDate = ReadINIStr("General", "Date", Mods(ModCount).ModLiblist)
                    Mods(ModCount).ModAuthor = ReadINIStr("General", "Author", Mods(ModCount).ModLiblist)
                    Mods(ModCount).ModWebsite = ReadINIStr("General", "Website", Mods(ModCount).ModLiblist)
                    Mods(ModCount).ModSize = GetDirectorySize(Mods(ModCount).ModPath)
                    Mods(ModCount).ModDescription = ReadINIStr("General", "Description", Mods(ModCount).ModLiblist)
                    Mods(ModCount).ModCampaigns = ReadINIStr("General", "Campaigns", Mods(ModCount).ModLiblist)
                    Mods(ModCount).ModAllowTX = BooleanStringToBoolean(ReadINIStr("General", "AllowTX", Mods(ModCount).ModLiblist, "yes"))
                    Mods(ModCount).ModTXVersion = ReadINIStr("General", "TXVersion", Mods(ModCount).ModLiblist, "")
                    Mods(ModCount).ModFA2Version = ReadINIStr("General", "FA2Version", Mods(ModCount).ModLiblist, "")
                    Mods(ModCount).ModUseAres = BooleanStringToBoolean(ReadINIStr("General", "UseAres", Mods(ModCount).ModLiblist, "no"))
                    Mods(ModCount).ModGameMode = ReadINIStr("General", "GameMode", Mods(ModCount).ModLiblist, "1")
                    Mods(ModCount).ModMapIndex = ReadINIStr("General", "MapIndex", Mods(ModCount).ModLiblist, "0")
                    Mods(ModCount).ModShowParams = BooleanStringToBoolean(ReadINIStr("General", "ShowParams", Mods(ModCount).ModLiblist, "no"))
                    Mods(ModCount).ModShutdownLB = BooleanStringToBoolean(ReadINIStr("General", "ShutdownLB", Mods(ModCount).ModLiblist, "no"))
                    If Mods(ModCount).ModShowParams Then
                        If FileExists(JoinPath(Mods(ModCount).ModPath, "launcher\userdata.lbu")) Then
                            Mods(ModCount).ModParams = ReadINIStr("General", "Params", JoinPath(Mods(ModCount).ModPath, "launcher\userdata.lbu"))
                        End If
                    End If
                    Select Case UCase$(ReadINIStr("General", "ModType", Mods(ModCount).ModLiblist, "MOD"))
                    Case "MOD": Mods(ModCount).ModType = TypeMod
                    Case "FA2MOD": Mods(ModCount).ModType = TypeFA2Mod
                    Case "USERTOOL", "MODTOOL", "DEVTOOL", "TOOL", "PROGRAM": Mods(ModCount).ModType = TypeProgram
                    Case Else: Mods(ModCount).ModType = -1
                    End Select
                    Mods(ModCount).ModUpdateCheckURL = ReadINIStr("General", "UpdateCheckURL", Mods(ModCount).ModLiblist, "")
                    Mods(ModCount).ModIsForRA2 = BooleanStringToBoolean(ReadINIStr("General", "IsForRA2", Mods(ModCount).ModLiblist, "no"))
                    Mods(ModCount).ModUseYuriUI = BooleanStringToBoolean(ReadINIStr("General", "UseYuriUI", Mods(ModCount).ModLiblist, "no"))
                    Mods(ModCount).ModScrnFormat = ReadINIStr("General", "ScrnFormat", Mods(ModCount).ModLiblist, "SCRN%04d.pcx")
                    Mods(ModCount).ModSnapFormat = ReadINIStr("General", "SnapFormat", Mods(ModCount).ModLiblist, "Map%04d.yrm")
                    Mods(ModCount).ModProgram = ReadINIStr("General", "Program", Mods(ModCount).ModLiblist)
                    If Len(Mods(ModCount).ModProgram) <> 0 Then Mods(ModCount).ModProgram = JoinPath(Mods(ModCount).ModPath, Mods(ModCount).ModProgram)
                    'manual
                    Mods(ModCount).ModManual = ""
                    If DirExists(JoinPath(Mods(ModCount).ModPath, "manual")) Then
                        Set fso_folder = fso.GetFolder(JoinPath(Mods(ModCount).ModPath, "manual"))
                        For Each fso_file In fso_folder.Files
                            If Len(fso_file.Name) > 6 Then
                                If LCase$(Left$(fso_file.Name, 6)) = "index." Then
                                    Mods(ModCount).ModManual = fso_file.Path
                                    Exit For
                                End If
                            End If
                        Next fso_file
                    End If
                    'sound1
                    If Not FileExists(JoinPath(Mods(ModCount).ModPath, "launcher\sound1.wav")) Then
                        If FileExists(JoinPath(Mods(ModCount).ModPath, "launcher\sound1.ogg")) Then
                            Call ConvertOggToWav(JoinPath(Mods(ModCount).ModPath, "launcher\sound1.ogg"))
                        ElseIf FileExists(JoinPath(Mods(ModCount).ModPath, "launcher\sound1.flac")) Then
                            Call ConvertFlacToWav(JoinPath(Mods(ModCount).ModPath, "launcher\sound1.flac"))
                        End If
                    End If
                    If FileExists(JoinPath(Mods(ModCount).ModPath, "launcher\sound1.wav")) Then Mods(ModCount).ModSound1 = JoinPath(Mods(ModCount).ModPath, "launcher\sound1.wav")
                    'sound2
                    If Not FileExists(JoinPath(Mods(ModCount).ModPath, "launcher\sound2.wav")) Then
                        If FileExists(JoinPath(Mods(ModCount).ModPath, "launcher\sound2.ogg")) Then
                            Call ConvertOggToWav(JoinPath(Mods(ModCount).ModPath, "launcher\sound2.ogg"))
                        ElseIf FileExists(JoinPath(Mods(ModCount).ModPath, "launcher\sound2.flac")) Then
                            Call ConvertFlacToWav(JoinPath(Mods(ModCount).ModPath, "launcher\sound2.flac"))
                        End If
                    End If
                    If FileExists(JoinPath(Mods(ModCount).ModPath, "launcher\sound2.wav")) Then Mods(ModCount).ModSound2 = JoinPath(Mods(ModCount).ModPath, "launcher\sound2.wav")
                    'Validation
                    bOk = False
                    If Len(Mods(ModCount).ModName) <> 0 Then
                        Select Case Mods(ModCount).ModType
                        Case TypeMod
                            If Not ((Mods(ModCount).ModAllowTX = False) And (Len(Mods(ModCount).ModTXVersion) <> 0)) Then
                                Call DisectScrnFormat(Mods(ModCount).ModScrnFormat, sVersion, sVersion, iScrnNumFormat)
                                If (iScrnNumFormat <> -1) Then
                                    Call DisectScrnFormat(Mods(ModCount).ModSnapFormat, sVersion, sVersion, iScrnNumFormat)
                                    If (iScrnNumFormat <> -1) Then bOk = True
                                End If
                            End If
                        Case TypeProgram
                            If Len(Mods(ModCount).ModProgram) <> 0 Then
                                If FileExists(Mods(ModCount).ModProgram) Then
                                    bOk = True
                                End If
                            Else
                                bOk = True 'tool might be documentation only
                            End If
                        Case TypeFA2Mod
                            bOk = True
                        End Select
                    End If
                    If bOk Then
                        'if this is the persistent mod then check for tampering
                        If Not bPersistentFound Then
                            If BooleanStringToBoolean(ReadINIStr("General", "ModIsActive", JoinPath(Mods(ModCount).ModPath, "launcher\userdata.lbu"))) Then
                                bPersistentFound = True
                                If Mods(ModCount).ModName = sPersistentName Then
                                    If Mods(ModCount).ModVersion = sPersistentVersion Then
                                        bPersistentMatch = True
                                        Call Init_LoadMods_NoTampering
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If Not BooleanStringToBoolean(ReadINIStr("General", "IsTX", Mods(ModCount).ModLiblist, "no")) Then
                            Call WriteLogEntry("The mod stored in " & Quote("Mods\" & GetFileName(Mods(ModCount).ModPath)) & " has been excluded because it is not valid.")
                        End If
                        ModCount = ModCount - 1
                    End If
                Else
                    'Mod is a plugin - ignore
                    ModCount = ModCount - 1
                End If
            End If
        End Select
        Call frmSplash.PROGRESS("Loading mods... " & Mods(ModCount).ModName)
    Next
    If Not bPersistentMatch Then
        If Len(sPersistentName) <> 0 Then Call DeactivateMod(True, bPersistentFound) 'persistent mod no longer available
    End If
    Call Init_LoadMods_Alphabetise
    Call frmSplash.PROGRESS
    Call Init_LoadMods_FillModLists
    Call frmSplash.PROGRESS
    Call CallStackPop
End Sub

Private Sub Init_LoadMods_Alphabetise() 'this routine is hardcoded to expect 3 stock mods (YR,RA2,FA2)
    Dim iMod As Integer
    Dim AlphaStr As String
    Dim AlphaNum As Integer
    Dim sTemp As String
    Dim ModsProcessed As Integer
    Dim ModProcessed() As Boolean
    Call CallStackPush(Me.Name & ".Init_LoadMods_Alphabetise()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Alphabetising mod list.", LogLevel2)
    ReDim ModProcessed(ModCount)
    ReDim AlphabetisedMods(ModCount)
    ModsProcessed = 0
    Do While ModsProcessed <> ModCount
        If Not ModProcessed(RA2ModNum) Then
            AlphaNum = RA2ModNum
        Else
            If Not ModProcessed(YRModNum) Then
                AlphaNum = YRModNum
            Else
                If Not ModProcessed(FA2ModNum) Then
                    AlphaNum = FA2ModNum
                Else
                    'not a stock mod
                    'first, need to know the name of the last mod in the list, for the alpha check
                    iMod = ModCount
                    Do
                        If ModProcessed(iMod) Then
                            iMod = iMod - 1
                        Else
                            AlphaNum = iMod
                            AlphaStr = Mods(AlphaNum).ModName
                            Exit Do
                        End If
                    Loop
                    'get lowest alphabetically
                    For iMod = 1 To ModCount
                        If ModProcessed(iMod) = False Then
                            If Mods(iMod).ModName = "YR Unofficial 1.002 Mini-Patch" Then
                                AlphaNum = iMod
                                Exit For
                            Else
                                If Mods(iMod).ModName <= AlphaStr Then
                                    AlphaNum = iMod
                                    AlphaStr = Mods(AlphaNum).ModName
                                End If
                            End If
                        End If
                    Next iMod
                End If
            End If
        End If
        AlphabetisedMods(ModsProcessed) = AlphaNum
        ModProcessed(AlphaNum) = True
        ModsProcessed = ModsProcessed + 1
    Loop
    Call CallStackPop
End Sub

Public Sub Init_LoadMods_FillModLists()
    Dim iMod As Integer
    Dim iCounter As Integer
    Dim iBannerMods As Integer
    Dim sMod As String
    Dim bOk As Boolean
    Call CallStackPush(Me.Name & ".Init_LoadMods_FillModLists()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Re-populating mod lists.", LogLevel2)
    For iMod = 0 To MaxType
        If iMod <> TypePlugin Then Call lstMods(iMod).Clear
    Next iMod
    iBannerMods = 0
    iCounter = 0
    Do While iCounter < ModCount
        iMod = AlphabetisedMods(iCounter)
        If Len(Mods(iMod).ModName) <> 0 Then 'mod has been deleted if name is blank
            bOk = True
            If Mods(iMod).ModType = TypeMod Then
                If Mods(iMod).ModIsForRA2 Then
                    If Not OptShowRA2 Then bOk = False
                Else
                    If Not OptShowYR Then bOk = False
                End If
            End If
            If bOk Then
                If Mods(iMod).ModVersion = "" Then
                    Call lstMods(Mods(iMod).ModType).AddItem(Mods(iMod).ModName)
                Else
                    Call lstMods(Mods(iMod).ModType).AddItem(Mods(iMod).ModName & " [" & Mods(iMod).ModVersion & "]")
                End If
                lstMods(Mods(iMod).ModType).ItemData(lstMods(Mods(iMod).ModType).ListCount - 1) = iMod
                If Mods(iMod).ModType = TypeMod Then iBannerMods = iBannerMods + 1
            End If
        End If
        iCounter = iCounter + 1
    Loop
    For iCounter = 0 To MaxType
        'can't let normal select occur otherwise it might try to play a sound
        PreventLoop = True
        sMod = ReadINIStr("General", "SelectedMod" & CStr(iCounter), ProgramINI) 'gonna try to reselect whatever mod was selected on last program shutdown
        iMod = lstMods(iCounter).ListCount - 1
        Do While iMod >= 0 'if we don't find the last selected mod then we'll just pick the first one in the list as per standard behaviour
            Select Case iCounter
            Case TypePlugin
                If GetFileName(Plugins(lstMods(iCounter).ItemData(iMod)).PluginPath) = sMod Then Exit Do
            Case TypeFA2Mod
               If Mods(lstMods(iCounter).ItemData(iMod)).ModType = iCounter Then
                    If iCounter = TypeFA2Mod Then
                        If sMod = "fa2yr" Then
                           If lstMods(iCounter).ItemData(iMod) = FA2ModNum Then Exit Do
                        End If
                    End If
                End If
            Case Else
                If Mods(lstMods(iCounter).ItemData(iMod)).ModType = iCounter Then
                   If GetFileName(Mods(lstMods(iCounter).ItemData(iMod)).ModPath) = sMod Then Exit Do
                End If
            End Select
            iMod = iMod - 1
        Loop
        If iMod = -1 Then
            If lstMods(iCounter).ListCount <> 0 Then iMod = 0 'there are some items, but none of them are the one we had selected alst time, so select the first one in the list
        End If
        lstMods(iCounter).ListIndex = iMod
        If iMod <> -1 Then iMod = lstMods(iCounter).ItemData(iMod) 'get the actual mod number from the list entry
        If iCounter = TypePlugin Then
            Call DisplayPluginDetails(iMod)
        Else
            Call DisplayModDetails(iMod, iCounter, False)
        End If
        PreventLoop = False
    Next iCounter
    'scrollbarBanners.Max = ((iBannerMods \ BannerCount) + Min(1, (iBannerMods Mod BannerCount))) - 1
    If scrollbarBanners.Value <> 0 Then
        scrollbarBanners.Value = 0
    Else
        Call scrollbarBanners_Change 'setting to zero when already zero doesn't count as a change
    End If
    scrollbarBanners.Visible = (scrollbarBanners.Max <> 0)
    Call CallStackPop
End Sub

Public Sub ScrollBanners(ByVal iPage As Integer)
    'assumes that page exists
    Dim iAlpha As Integer
    Dim iMod As Integer
    Dim iBanner As Integer
    Dim sTemp As String
    iAlpha = 0
    iBanner = 0
    Do While iAlpha < ModCount And iBanner < (BannerCount * iPage)
        iMod = AlphabetisedMods(iAlpha)
        If Mods(iMod).ModType = TypeMod Then
            If Mods(iMod).ModIsForRA2 Then
                If OptShowRA2 Then iBanner = iBanner + 1
            Else
                If OptShowYR Then iBanner = iBanner + 1
            End If
        End If
        iAlpha = iAlpha + 1
    Loop
    'iAlpha is now set to the alpha index of the first mod to show a banner for
    iBanner = MaxType + 1
    Do While iAlpha < ModCount And iBanner <= (BannerCount + MaxType)
        iMod = AlphabetisedMods(iAlpha)
        If Mods(iMod).ModType = TypeMod Then
            If (Mods(iMod).ModIsForRA2 And OptShowRA2) Or (Not Mods(iMod).ModIsForRA2 And OptShowYR) Then
                picMod(iBanner).Tag = CStr(iMod)
                lblNoBanner(iBanner).Caption = ""
                Set picMod(iBanner).Picture = Nothing
                Select Case iMod
                Case RA2ModNum
                    sTemp = JoinPath(RESDIR, "ra2banner.bmp")
                Case YRModNum
                    sTemp = JoinPath(RESDIR, "yrbanner.bmp")
                Case FA2ModNum
                    sTemp = JoinPath(RESDIR, "fabanner.bmp")
                Case Else
                    sTemp = JoinPath(Mods(iMod).ModPath, "launcher\banner.bmp")
                    If Not FileExists(sTemp) Then
                        sTemp = JoinPath(Mods(iMod).ModPath, "launcher\banner.jpg")
                        If Not FileExists(sTemp) Then
                            sTemp = JoinPath(Mods(iMod).ModPath, "launcher\banner.jpeg")
                            If Not FileExists(sTemp) Then
                                sTemp = JoinPath(Mods(iMod).ModPath, "launcher\banner.gif")
                                If Not FileExists(sTemp) Then
                                    sTemp = JoinPath(RESDIR, "nobanner.bmp")
                                    If Len(Mods(iMod).ModVersion) = 0 Then
                                        lblNoBanner(iBanner).Caption = DoubleAmpersand(Mods(iMod).ModName)
                                    Else
                                        lblNoBanner(iBanner).Caption = DoubleAmpersand(Mods(iMod).ModName) & " [" & DoubleAmpersand(Mods(iMod).ModVersion) & "]"
                                    End If
                                End If
                            End If
                        End If
                    End If
                End Select
                picMod(iBanner).Picture = LoadPicture(sTemp)
                If iBanner Mod 2 = 0 Then 'odd numbers are on the right hand side
                    picMod(iBanner).Left = 360
                Else
                    picMod(iBanner).Left = 4680
                End If
                picMod(iBanner).Visible = True
                lblNoBanner(iBanner).Visible = True
                iBanner = iBanner + 1
            End If
        End If
        iAlpha = iAlpha + 1
    Loop
    If (iBanner - 1) Mod 2 = 0 Then picMod(iBanner - 1).Left = 2520
    Do While iBanner <= (MaxType + BannerCount)
        picMod(iBanner).Visible = False
        lblNoBanner(iBanner).Visible = False
        iBanner = iBanner + 1
    Loop
End Sub

Private Sub cmdBannerScrollLeft_Click()
    Call ScrollBanners(BannerPage - 1)
End Sub

Private Sub cmdBannerScrollRight_Click()
    Call ScrollBanners(BannerPage + 1)
End Sub

Private Sub cmdFA2Browse_Click()
    dialogOpen.FileName = JoinPath(FA2DIR, "FinalAlert2YR.exe")
    dialogOpen.DialogTitle = "Select FinalAlert 2 YR Executable"
    dialogOpen.Filter = "FinalAlert 2 YR Executable|FinalAlert2YR.exe"
    dialogOpen.DefaultExt = "exe"
'RetryOpen:
    On Error GoTo CancelOpen
    dialogOpen.ShowOpen
    On Error GoTo 0
    If UCase$(GetFilePath(dialogOpen.FileName)) <> UCase$(FA2DIR) Then
        txtFA2Folder.Text = GetFilePath(dialogOpen.FileName)
        Call FA2Check(False)
    End If
CancelOpen:
End Sub

Private Sub cmdManual_Click(Index As Integer)
    If OpenLocation(cmdManual(Index).Tag) < 32 Then
        Call MsgBox("Unable to open " & Quote(cmdManual(Index).Tag) & ".", vbOKOnly + vbInformation, App.Title)
    End If
End Sub

Private Sub cmdLaunch_Click(Index As Integer)
    Call LaunchMod(Index)
End Sub

Private Function UpdateAres_LatestRevision(ByRef sBranch As String, ByRef sURL As String, ByRef sError As String) As Long
    Dim sBuffer As String
    Dim sRev As String
    Dim iBranch As Long
    Dim iBranches As Long
    Dim tempBranch As String
    Dim sBestBranch As String
    Dim sBestVersion As String
    Dim sAgent As String
    Call CallStackPush(Me.Name & ".UpdateAres_LatestRevision(" & CStr(sBranch) & ", " & CStr(sURL) & ", " & CStr(sError) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    sURL = ""
    sError = ""
    'Renegade wishes to know which header we are requesting so temporarily change the UserAgent
    sAgent = theInternet.UserAgent
    theInternet.UserAgent = sAgent & " (" & OptAresBranch & ")"
    If theInternet.CopyURLToString(OptAresRevisionDataURL, sBuffer) Then
        sBuffer = ConvertEOL(sBuffer)
        If Len(sBranch) = 0 Then
            sError = "No branch specified."
        Else
            UpdateAres_LatestRevision = Val(ReadINIStrMemory(sBuffer, sBranch, "revision", "-1"))
            If UpdateAres_LatestRevision = -1 Then
                sError = "No revision data available for branch " & sBranch
            Else
                sURL = ReadINIStrMemory(sBuffer, sBranch, "location", "")
                If Len(sURL) = 0 Then
                    sError = "No download location available for branch " & sBranch
                End If
            End If
        End If
        If Len(sError) <> 0 Then
            If sBranch <> "syringe" Then
                'check if the branch is listed
                iBranch = 0
                iBranches = 0
                sBestBranch = ""
                sBestVersion = ""
                Do
                    tempBranch = ReadINIStrMemory(sBuffer, "branches", CStr(iBranch), "")
                    If Len(tempBranch) <> 0 Then
                        iBranches = iBranches + 1
                        If tempBranch = sBranch Then
                            iBranches = -1 'we found it
                            Exit Do
                        Else
                            'we might have to revert to the latest stable branch if the selected branch has changed
                            If LCase$(ReadINIStrMemory(sBuffer, tempBranch, "stability", "?")) = "stable" Then
                                If CompareVersions(ReadINIStrMemory(sBuffer, tempBranch, "version", "?"), ">", sBestVersion) Then
                                    sBestBranch = tempBranch
                                End If
                            End If
                        End If
                    Else
                        If iBranch <> 0 Then Exit Do
                    End If
                    iBranch = iBranch + 1
                Loop
                If iBranches = 0 Then
                    sError = "Ares revision data does not list any branches."
                ElseIf iBranches <> -1 Then
                    If Len(sBestBranch) <> 0 Then
                        sBranch = sBestBranch
                        OptAresBranch = sBestBranch
                        Call WriteINIStr("Options", "AresBranch", OptAresBranch, ProgramINI)
                        Call frmMain.WriteLogEntry("Ares Update Options: 'Ares Release' changed to " & OptAresBranch & " by Launch Base.", LogLevel0)
                    Else
                        sError = "Ares revision data does not list branch " & sBranch & ". You need to select a different branch."
                    End If
                End If
            End If
        End If
    Else
        sError = "Failed to download " & IIf(sBranch = "syringe", "Syringe", "Ares") & " revision data."
        Call WriteLogEntry(sError, LogLevel0)
        UpdateAres_LatestRevision = -2
    End If
    theInternet.UserAgent = sAgent
    Call CallStackPop
End Function

Friend Sub UpdateAres(ByRef ParentForm As Form, Optional ByVal sBranch As String = "", Optional ByVal bSilent As Boolean = True, Optional ByVal bForce As Boolean = False)
    Dim sTarBall As String
    Dim sAresDir As String
    Dim mbResult As VbMsgBoxResult
    Dim sProduct As String
    Dim iOldRevision As Long
    Dim iNewRevision As Long
    Dim sError As String
    Dim sURL As String
    Dim sTemp As String
    Dim fso As FileSystemObject
    Dim fsoDir As Folder
    Dim fsoResDir As Folder
    Dim sOldFolders() As String
    Dim iOldFolder As Long
    Dim bAres As Boolean
    Call CallStackPush(Me.Name & ".UpdateAres(" & CStr(ParentForm.Name) & ", " & CStr(sBranch) & ", " & CStr(bSilent) & ", " & CStr(bForce) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    If sBranch = "" Then sBranch = OptAresBranch
    bAres = (sBranch <> "syringe")
    If bAres Then sProduct = "Ares" Else sProduct = "Syringe"
    If theInternet.Connected Then
        If Not bSilent Then Call ParentForm.ShowPleaseWait("Checking for updates to " & sProduct & "...", "Downloading revision data.")
        If bAres Then
            If Not FileExists(JoinPath(RESDIR, "Ares.dll")) Or bForce Then OptAresRevision = 0
            iOldRevision = OptAresRevision
        Else
            If Not FileExists(JoinPath(RESDIR, "Syringe.exe")) Then OptSyringeRevision = 0
            iOldRevision = OptSyringeRevision
        End If
        If bAres Then
            Call WriteLogEntry("Checking for updates to " & sProduct & ".", LogLevel1)
        Else
            Call WriteLogEntry("Checking for updates to " & sProduct & " branch " & Quote(sBranch) & ".", LogLevel1)
        End If
        iNewRevision = UpdateAres_LatestRevision(sBranch, sURL, sError)
        If Len(sError) = 0 Then
            If iNewRevision <> iOldRevision Then
                Call WriteLogEntry(sProduct & " update available." & vbCrLf & "Latest revision: " & iNewRevision, LogLevel1)
                If Not bSilent Then Call ParentForm.UpdatePleaseWait("Updating " & sProduct & "...", "Downloading update.")
                If bAres Then
                    sTarBall = JoinPath(RESDIR, "ares.tar.gz")
                Else
                    sTarBall = JoinPath(RESDIR, "syringe.tar.gz")
                End If
                If FileExists(sTarBall) Then
                    'On Error Resume Next
                    Call LoggedKill(sTarBall)
                    'If Not CL_noexcept Then On Error GoTo LocalErr
                    'If Err.Number <> 0 Then
                    '    Call WriteLogEntry("Failed to overwrite " & Quote(sTarBall) & " - " & Err.Description, IIf(bSilent, LogLevel1, LogLevel1 + LogMsgBox))
                    'End If
                End If
                Call WriteLogEntry("Downloading " & Quote(sURL) & " to " & Quote(sTarBall))
                If theInternet.CopyURLToFile(sURL, sTarBall) Then
                    Call WriteLogEntry("Download complete.")
                    If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Decompressing update.")
                    If bAres Then
                        sTemp = JoinPath(RESDIR, "ares.tar")
                    Else
                        sTemp = JoinPath(RESDIR, "syringe.tar")
                    End If
                    If FileExists(sTemp) Then Call LoggedKill(sTemp)
                    Call WriteLogEntry("Decompressing " & sTarBall, LogLevel1)
                    Call GetCommandOutput(Quote(JoinPath(RESDIR, "gunzip.exe")) & " " & Quote(sTarBall))
                    If FileExists(sTemp) Then
                        If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Extracting file(s).")
                        sTarBall = sTemp
                        Call ChDir(RESDIR)
                        'get the list of old folders so we can later find out what new folder gets created
                        Set fso = New FileSystemObject
                        Set fsoResDir = fso.GetFolder(RESDIR)
                        ReDim sOldFolders(fsoResDir.SubFolders.Count)
                        iOldFolder = 0
                        For Each fsoDir In fsoResDir.SubFolders
                            iOldFolder = iOldFolder + 1
                            sOldFolders(iOldFolder) = fsoDir.Name
                        Next

                        If bAres Then
                            'ARES
                            
                            Call WriteLogEntry("Unpacking " & sTarBall, LogLevel1)
                            Call GetCommandOutput(Quote(JoinPath(RESDIR, "tar.exe")) & " -x -f " & GetFileName(sTarBall))
                            Call LoggedKill(sTarBall)
                            'identify the new folder
                            Set fsoResDir = fso.GetFolder(RESDIR)
                            iOldFolder = 0
                            For Each fsoDir In fsoResDir.SubFolders
                                iOldFolder = 0
                                Do While iOldFolder <= UBound(sOldFolders())
                                    If (sOldFolders(iOldFolder) <> fsoDir.Name) Or (UBound(sOldFolders()) = 0) Then
                                        Call WriteLogEntry("New folder found: " & fsoDir.Name, LogLevel2)
                                        sAresDir = JoinPath(RESDIR, fsoDir.Name)
                                        Call WriteLogEntry("Ares folder accepted: " & sAresDir, LogLevel2)
                                        If Len(sAresDir) <> 0 And FileExists(JoinPath(sAresDir, "Ares.dll")) And FileExists(JoinPath(sAresDir, "Ares.dll.inj")) Then
                                            If FileExists(JoinPath(RESDIR, "Ares.dll")) Then Call LoggedKill(JoinPath(RESDIR, "Ares.dll"))
                                            Call LoggedMove(JoinPath(sAresDir, "Ares.dll"), JoinPath(RESDIR, "Ares.dll"))
                                            If FileExists(JoinPath(RESDIR, "Ares.dll.inj")) Then Call LoggedKill(JoinPath(RESDIR, "Ares.dll.inj"))
                                            Call LoggedMove(JoinPath(sAresDir, "Ares.dll.inj"), JoinPath(RESDIR, "Ares.dll.inj"))
                                            If FileExists(JoinPath(sAresDir, "ares.mix")) Then
                                                If FileExists(JoinPath(RESDIR, "ares.mix")) Then Call LoggedKill(JoinPath(RESDIR, "ares.mix"))
                                                Call LoggedMove(JoinPath(sAresDir, "ares.mix"), JoinPath(RESDIR, "ares.mix"))
                                            End If
                                            If FileExists(JoinPath(sAresDir, "Documentation\AresManual.html")) Then
                                                If DirExists(JoinPath(RESDIR, "AresDocumentation")) Then Call LoggedKillDir(JoinPath(RESDIR, "AresDocumentation"))
                                                Call LoggedMove(JoinPath(sAresDir, "Documentation"), JoinPath(RESDIR, "AresDocumentation"))
                                                menu_aresdoc.Enabled = True
                                            End If
                                            If FileExists(JoinPath(sAresDir, "readme.txt")) Then
                                                If FileExists(JoinPath(RESDIR, "AresReadme.txt")) Then Call LoggedKill(JoinPath(RESDIR, "AresReadme.txt"))
                                                Call LoggedMove(JoinPath(sAresDir, "readme.txt"), JoinPath(RESDIR, "AresReadme.txt"))
                                                'If OptAdvancedMode And OptAresUpdateNotes Then
                                                '    Call WriteLogEntry("Opening " & JoinPath(RESDIR, "AresReadme.txt"), LogLevel1)
                                                '    If OpenLocation(JoinPath(RESDIR, "AresReadme.txt")) < 32 Then
                                                '        Call WriteLogEntry("Failed to open Ares readme.")
                                                '        Call MsgBox("Failed to open Ares readme. This can be found in " & Quote(JoinPath(RESDIR, "AresReadme.txt")), vbInformation + vbOKOnly)
                                                '    End If
                                                'End If
                                            End If
                                            Call LoggedKillDir(sAresDir)
                                            OptAresRevision = iNewRevision
                                            Call WriteINIStr("General", "AresRevision", CStr(iNewRevision), ProgramINI)
                                            If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Saving checksum.")
                                            sTemp = Decrypt(HDSerialNumber(Left$(EXEDIR, 1)), True) 'key for MD5 encryption
                                            Call WriteINIStr("General", "AresDLL", Base64EncodeString(EncryptString(StrReverse(GetFileMD5(JoinPath(RESDIR, "Ares.dll"))), sTemp)), ProgramINI)
                                            Call WriteINIStr("General", "AresINJ", Base64EncodeString(EncryptString(StrReverse(GetFileMD5(JoinPath(RESDIR, "Ares.dll.inj"))), sTemp)), ProgramINI)
                                            If FileExists(JoinPath(RESDIR, "Ares.dll")) Then
                                                Call WriteINIStr("General", "AresMIX", Base64EncodeString(EncryptString(StrReverse(GetFileMD5(JoinPath(RESDIR, "ares.mix"))), sTemp)), ProgramINI)
                                            Else
                                                Call WriteINIStr("General", "AresMIX", " ", ProgramINI)
                                            End If
                                            If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Update complete.")
                                            Call WriteLogEntry("Ares update complete." & vbCrLf & "Updated to revision " & CStr(iNewRevision), IIf(bSilent, LogLevel1, LogLevel1 + LogMsgBox))
                                            iOldFolder = -1 'to say we're ok
                                            Exit For
                                        End If
                                    End If
                                    iOldFolder = iOldFolder + 1
                                Loop
                            Next
                            If iOldFolder <> -1 Then
                                If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Update failed.")
                                Call WriteLogEntry("Ares update failed - failed to unpack " & Quote(sTarBall) & vbCrLf & "Check that the Ares build you have selected is available - you may need to select a different build.", IIf(bSilent, LogLevel1, LogLevel1 + LogMsgBox))
                                If DirExists(sAresDir) Then Call LoggedKillDir(sAresDir)
                            End If
                        Else
                            'SYRINGE
                            If FileExists(JoinPath(RESDIR, "Syringe.old")) Then Call LoggedKill(JoinPath(RESDIR, "Syringe.old"))
                            If FileExists(JoinPath(RESDIR, "Syringe.exe")) Then Call LoggedMove(JoinPath(RESDIR, "Syringe.exe"), JoinPath(RESDIR, "Syringe.old"), True)
                            Call GetCommandOutput(Quote(JoinPath(RESDIR, "tar.exe")) & " -x -f " & GetFileName(sTarBall))
                            Call LoggedKill(sTarBall)
                            'identify the new folder
                            sTarBall = ""
                            Set fsoResDir = fso.GetFolder(RESDIR)
                            iOldFolder = 0
                            For Each fsoDir In fsoResDir.SubFolders
                                iOldFolder = 0
                                Do While iOldFolder <= UBound(sOldFolders())
                                    If (sOldFolders(iOldFolder) <> fsoDir.Name) Or (UBound(sOldFolders()) = 0) Then
                                        sTarBall = JoinPath(RESDIR, fsoDir.Name)
                                        If Len(sTarBall) <> 0 And FileExists(JoinPath(sTarBall, "Syringe.exe")) Then
                                            Call LoggedMove(JoinPath(sTarBall, "Syringe.exe"), JoinPath(RESDIR, "Syringe.exe"))
                                            If FileExists(JoinPath(RESDIR, "Syringe.old")) Then Call LoggedKill(JoinPath(RESDIR, "Syringe.old"))
                                            If FileExists(JoinPath(sTarBall, "license.txt")) Then
                                                If FileExists(JoinPath(RESDIR, "Syringe.txt")) Then Call LoggedKill(JoinPath(RESDIR, "Syringe.txt"))
                                                Call LoggedMove(JoinPath(sTarBall, "license.txt"), JoinPath(RESDIR, "Syringe.txt"))
                                            End If
                                            OptSyringeRevision = iNewRevision
                                            Call WriteINIStr("General", "SyringeRevision", CStr(iNewRevision), ProgramINI)
                                            If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Saving checksum.")
                                            sTemp = Decrypt(HDSerialNumber(Left$(EXEDIR, 1)), True) 'key for MD5 encryption
                                            Call WriteINIStr("General", "Syringe", Base64EncodeString(EncryptString(StrReverse(GetFileMD5(JoinPath(RESDIR, "Syringe.exe"))), sTemp)), ProgramINI)
                                            Call WriteLogEntry(JoinPath(RESDIR, "Syringe.exe") & " = " & GetFileMD5(JoinPath(RESDIR, "Syringe.exe")) & " = " & Base64EncodeString(EncryptString(StrReverse(GetFileMD5(JoinPath(RESDIR, "Syringe.exe"))), sTemp)) & " ... " & sTemp, LogLevel2)
                                            If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Update complete.")
                                            Call WriteLogEntry(sProduct & " update complete." & vbCrLf & "Updated to revision " & CStr(iNewRevision), IIf(bSilent, LogLevel1, LogLevel1 + LogMsgBox))
                                            iOldFolder = -1 'to say we found it
                                            Call LoggedKillDir(sTarBall)
                                            Exit For
                                        End If
                                    End If
                                    iOldFolder = iOldFolder + 1
                                Loop
                            Next
                            If iOldFolder <> -1 Then
                                If FileExists(JoinPath(RESDIR, "Syringe.old")) Then Call LoggedMove(JoinPath(RESDIR, "Syringe.old"), JoinPath(RESDIR, "Syringe.exe"), False, True)
                                If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Update failed.")
                                Call WriteLogEntry(sProduct & " update failed - failed to unpack " & Quote(sTarBall), IIf(bSilent, LogLevel1, LogLevel1 + LogMsgBox))
                            End If
                        End If

                        Set fsoDir = Nothing
                        Set fsoResDir = Nothing
                        Set fso = Nothing
                    Else
                        If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Update failed.")
                        Call WriteLogEntry(sProduct & " update failed - failed to decompress " & Quote(sTarBall), IIf(bSilent, LogLevel1, LogLevel1 + LogMsgBox))
                        Call LoggedKill(sTarBall)
                    End If
                Else
                    If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Update failed.")
                    Call WriteLogEntry(sProduct & " update failed - download failed.", IIf(bSilent, LogLevel1, LogLevel1 + LogMsgBox))
                End If
            Else
                If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "No update available.")
                Call WriteLogEntry("No update available - " & sProduct & " is up to date." & vbCrLf & "Latest revision: " & CStr(iNewRevision), IIf(bSilent, LogLevel1, LogLevel1 + LogMsgBox))
            End If
        Else
            If Not bSilent Then Call ParentForm.UpdatePleaseWait("", "Update check failed.")
            Call WriteLogEntry(sProduct & " update check failed. " & sError, IIf(bSilent, LogLevel1, LogLevel1 + LogMsgBox))
        End If
        If Not bSilent Then Call ParentForm.HidePleaseWait
    Else
        Call WriteLogEntry(sProduct & " update check aborted - no Internet connection available.", IIf(bSilent, LogLevel1, LogLevel1 + LogMsgBox))
    End If
    Call CallStackPop
End Sub

Private Function VerifySyringe()
    Dim sTemp As String
    Dim sEXE
    Dim bOk As Boolean
    bOk = False
    sTemp = ReadINIStr("General", "Syringe", ProgramINI)
    sEXE = JoinPath(RESDIR, "Syringe.exe")
    If FileExists(sEXE) Then
        If Len(sTemp) Mod 4 = 0 Then 'valid length for base64 decode
            'Call MsgBox(StrReverse(EncryptString(Base64DecodeString("SAAADwRHWUlDfgomcHop5BixP2t4aWxsPfO7rfdeBVk="), "-6277thxtK")) & vbCrLf & StrReverse(EncryptString(Base64DecodeString(sTemp), Decrypt(HDSerialNumber(Left$(RESDIR, 1)), True))) & vbCrLf & Base64EncodeString(EncryptString(StrReverse(GetFileMD5(JoinPath(RESDIR, "Syringe.exe"))), "-6277thxtK")), vbOKOnly)
            Call WriteLogEntry(sEXE & " = " & GetFileMD5(sEXE) & " = " & Base64EncodeString(EncryptString(StrReverse(GetFileMD5(sEXE)), sTemp)) & " != " & sTemp & " = " & StrReverse(EncryptString(Base64DecodeString(sTemp), Decrypt(HDSerialNumber(Left$(RESDIR, 1)), True))), LogLevel2)
            If StrReverse(EncryptString(Base64DecodeString(sTemp), Decrypt(HDSerialNumber(Left$(RESDIR, 1)), True))) = GetFileMD5(sEXE) Then bOk = True
        End If
    End If
    If Not bOk Then
        Call WriteLogEntry("Syringe verification failed.", LogLevel0)
        If FileExists(sEXE) Then Call LoggedKill(sEXE)
    End If
    VerifySyringe = bOk
End Function

Private Function VerifyAres()
    Dim sTemp As String
    Dim sKey As String
    Dim sDLL As String
    Dim sINJ As String
    Dim sMIX As String
    Dim bOk As Boolean
    bOk = False
    sKey = Decrypt(HDSerialNumber(Left$(RESDIR, 1)), True)
    sTemp = ReadINIStr("General", "AresDLL", ProgramINI)
    sDLL = JoinPath(RESDIR, "Ares.dll")
    sINJ = JoinPath(RESDIR, "Ares.dll.inj")
    sMIX = JoinPath(RESDIR, "ares.mix")
    If FileExists(sDLL) Then
        If FileExists(sINJ) Then
            If Len(sTemp) Mod 4 = 0 Then 'valid length for base64 decode
                If StrReverse(EncryptString(Base64DecodeString(sTemp), sKey)) = GetFileMD5(sDLL) Then
                    sTemp = ReadINIStr("General", "AresINJ", ProgramINI)
                    If Len(sTemp) Mod 4 = 0 Then 'valid length for base64 decode
                        If StrReverse(EncryptString(Base64DecodeString(sTemp), sKey)) = GetFileMD5(sINJ) Then
                            sTemp = ReadINIStr("General", "AresMIX", ProgramINI)
                            If FileExists(sMIX) Then
                                If StrReverse(EncryptString(Base64DecodeString(sTemp), sKey)) = GetFileMD5(sMIX) Then bOk = True 'mix exists and is verified
                            ElseIf Len(sTemp) = 0 Then
                                'mix does not exist and we are not expecting it
                                bOk = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    If Not bOk Then
        Call WriteLogEntry("Ares verification failed.", LogLevel0)
        If FileExists(sDLL) Then Call LoggedKill(sDLL)
        If FileExists(sINJ) Then Call LoggedKill(sINJ)
    End If
    VerifyAres = bOk
End Function


Private Sub LaunchMod(ByVal Index As Integer, Optional ByVal iMod As Integer = -1)
    Dim ShellCmd As String
    Dim bOk As Boolean
    Dim bShutdown As Boolean
    Dim iTime As Long
    Dim dDate As Date
    Dim bModified As Boolean
    Dim bSyringeOk As Boolean
    Dim bRestoreTheme As Boolean
    Dim sTemp As String
    Call CallStackPush(Me.Name & ".LaunchMod(" & CStr(Index) & "," & CStr(iMod) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    bSyringeOk = True
    bShutdown = False
    bRestoreTheme = False
    frmMain.Enabled = False
    If iMod = -1 Then
        If Index = MaxType + 1 Then
            iMod = Val(lstMods(TypePlugin).ItemData(lstMods(TypePlugin).ListIndex))
        Else
            iMod = Val(lstMods(Index).ItemData(lstMods(Index).ListIndex))
        End If
    End If
    Select Case Index
    Case TypeMod
        bModified = False
        If LaunchMod_CheckGameVersion(iMod, bModified) Then
            Call WriteLogEntry("Preparing to launch mod: " & Mods(iMod).ModName, LogLevel1)
            If PrerequisiteCheckTX(IIf(Mods(iMod).ModAllowTX, Mods(iMod).ModTXVersion, "-1"), True) Then
                If OptAutoUpdate Then
                    If iMod >= HardCodedMods Then
                        bOk = True
                        Call LaunchMod_CheckForUpdate(bShutdown, iMod, Index, bOk)
                        If (Not bOk) Or bShutdown Then GoTo AbortLaunch
                    End If
                End If
                If LaunchMod_YPLCheckOk(iMod, bRestoreTheme) Then
                    If OptModSound2 Then Call PlaySound(Mods(iMod).ModSound2, True)
                    ShellCmd = ""
                    If Len(OptCustomSwitches) <> 0 Then ShellCmd = ShellCmd & " " & OptCustomSwitches
                    If OptMPDebug Then ShellCmd = ShellCmd & " -MPDEBUG"
                    If OptLogAres Then ShellCmd = ShellCmd & " -log"
                    If OptWindowed Then ShellCmd = ShellCmd & " -WIN"
                    If OptSpeedControl Then ShellCmd = ShellCmd & " -SPEEDCONTROL"
                    If LaunchMod_PlayVideo(iMod) Then
                        Call WriteINIStr("Restore", "RestorePending", "yes", ProgramINI)
                        Call WriteINIStr("Restore", "ModifiedEXE", BooleanToYesNo(bModified), ProgramINI)
                        Call WriteINIStr("Restore", "Theme", BooleanToYesNo(bRestoreTheme), ProgramINI)
                        If Len(CL_playfile) <> 0 Then ShellCmd = ShellCmd & " -play session.ipb"
                        If LaunchMod_RecordVideo Then ShellCmd = ShellCmd & " -record " & Quote(RA2DIR)
                        iTime = -1
                        bOk = True
                        If ActivateMod(iMod) = True Then 'activateMod will also deactivate any existing mod first if neccessary
                            If Mods(iMod).ModIsForRA2 Then
                                Call WriteLogEntry("Launching Red Alert 2.", LogLevel1)
                                ShellCmd = Quote(JoinPath(RA2DIR, "ra2.exe")) & ShellCmd
                            Else
                                If BooleanStringToBoolean(ReadINIStr("Mod", "Syringe", ProgramINI, "no")) Then
                                    If OptAutoAresUpdate Then
                                        Call UpdateAres(Me, "syringe")
                                    End If
                                    bSyringeOk = VerifySyringe
                                    If Not bSyringeOk Then
                                        'Call DenyPersistentMods
                                        If OptAutoAresUpdate Then
                                            Call UpdateAres(Me, "syringe")
                                            If FileExists(JoinPath(RESDIR, "Syringe.exe")) Then
                                                bSyringeOk = True
                                            End If
                                        End If
                                    End If
                                    If bSyringeOk Then
                                        Call WriteLogEntry("Launching Yuri's Revenge via Syringe.", LogLevel1)
                                        ShellCmd = Quote(JoinPath(RESDIR, "Syringe.exe")) & " " & Quote(JoinPath(RA2DIR, "gamemd.exe")) & " --" & ShellCmd
                                    Else
                                        Call WriteLogEntry("Cannot launch Yuri's Revenge via Syringe because Syringe is not present.")
                                        Call MsgBox("Cannot launch mod!" & vbCrLf & "This mod requires Syringe, which is not present." & vbCrLf & "Syringe has most likely failed to download - check your Internet connection.", vbOKOnly + vbExclamation)
                                        bOk = False
                                    End If
                                Else
                                    Call WriteLogEntry("Launching Yuri's Revenge.", LogLevel1)
                                    ShellCmd = Quote(JoinPath(RA2DIR, "ra2md.exe")) & ShellCmd
                                End If
                            End If
                            If bOk Then
                                Call WriteLogEntry("Shelling: " & ShellCmd, LogLevel2)
                                dDate = Now()
                                Call SetCurrentDirectory(RA2DIR)
                                Call MxShell(ShellCmd, True)
                                iTime = DateDiff("s", dDate, Now())
                                If Mods(iMod).ModIsForRA2 Then
                                    If iTime > 3 Then
                                        Call WriteLogEntry("Red Alert 2 has terminated.", LogLevel1)
                                    Else
                                        Call WriteLogEntry("Red Alert 2 failed to load!")
                                        Call MsgBox("Red Alert 2 failed to load!" & vbCrLf & "Check that you can actually run Red Alert 2 outside of Launch Base.", vbOKOnly + vbExclamation, App.Title)
                                    End If
                                Else
                                    If BooleanStringToBoolean(ReadINIStr("Mod", "Syringe", ProgramINI, "no")) Then
                                        If iTime > 3 Then
                                            Call WriteLogEntry("Syringe has terminated.", LogLevel1)
                                        Else
                                            Call WriteLogEntry("Either Syringe failed to launch Yuri's Revenge or Yuri's Revenge failed to load!")
                                            Call MsgBox("Either Syringe failed to launch Yuri's Revenge or Yuri's Revenge failed to load!" & vbCrLf & "Check that you have installed the VisualC++ 2008 Redistributable Package required by Syringe" & vbCrLf & "(see Help Topics) and that you can actually run Yuri's Revenge outside of Launch Base.", vbOKOnly + vbExclamation, App.Title)
                                        End If
                                    Else
                                        If iTime > 3 Then
                                            Call WriteLogEntry("Yuri's Revenge has terminated.", LogLevel1)
                                        Else
                                            Call WriteLogEntry("Yuri's Revenge failed to load!")
                                            Call MsgBox("Yuri's Revenge failed to load!" & vbCrLf & "Check that you can actually run Yuri's Revenge outside of Launch Base.", vbOKOnly + vbExclamation, App.Title)
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            bOk = False
                        End If
                        If Not bOk Then
                            'whatever caused ActivateMod to fail should have reported it to the user
                            Call WriteINIStr("Restore", "Recording", "no", ProgramINI) 'don't want restore process thinking we should have a video
                        End If
                        Call RestoreProcess(iTime)
                    Else
                        Call WriteLogEntry("Mod launch cancelled by user.", LogLevel1)
                        If bRestoreTheme Then
                            'restore process won't run, need to restore thememd.ini now
                            If FileExists(JoinPath(RA2DIR, "thememd.ini")) Then
                                If FileExists(JoinPath(RA2DIR, "thememd.ini")) Then Call LoggedKill(JoinPath(RA2DIR, "thememd.ini"))
                                If FileExists(JoinPath(BACKUPDIR, "thememd.ini")) Then Call LoggedMove(JoinPath(BACKUPDIR, "thememd.ini"), JoinPath(RA2DIR, "thememd.ini"), False, True)
                            End If
                        End If
                    End If
                Else
                    Call WriteLogEntry("Mod launch cancelled by user.", LogLevel1)
                End If
            Else
                Call WriteLogEntry("Cannot activate " & Mods(iMod).ModName & " because Terrain Expansion prerequisites not satisfied.", LogMsgBox)
            End If
        End If
    Case TypePlugin
        Call WriteLogEntry("Preparing to activate plugin: " & Plugins(iMod).PluginName, LogLevel1)
        If AuthenticatePlugin(iMod) Then
            If OptAutoUpdate Then
                bOk = True
                Call LaunchMod_CheckForUpdate(bShutdown, iMod, Index, bOk)
                If (Not bOk) Or bShutdown Then GoTo AbortLaunch
            End If
            Call ActivatePlugin(iMod)
        End If
    Case (MaxType + 1)
        Call DeactivatePlugin(Plugins(iMod).PluginID)
    Case TypeFA2Mod
        Call WriteLogEntry("Preparing to launch FA2 mod: " & Mods(iMod).ModName, LogLevel1)
        If PrerequisiteCheckTX(IIf(Mods(iMod).ModAllowTX, Mods(iMod).ModTXVersion, "-1"), True) Then
            If CL_tx And cboxTX.Enabled = True Then cboxTX.Value = 1
            If cboxTX.Value = 1 Or Not CL_tx Then
                If OptAutoUpdate Then
                    If iMod >= HardCodedMods Then
                        bOk = True
                        Call LaunchMod_CheckForUpdate(bShutdown, iMod, Index, bOk)
                        If (Not bOk) Or bShutdown Then GoTo AbortLaunch
                    End If
                End If
                If OptModSound2 Then Call PlaySound(Mods(iMod).ModSound2, True)
                Call WriteINIStr("Restore", "RestorePending", "yes", ProgramINI)
                ShellCmd = Quote(JoinPath(FA2DIR, "FinalAlert2YR.exe"))
                If ActivateMod(iMod) = True Then
                    Call WriteLogEntry("Launching FinalAlert 2 YR.", LogLevel1)
                    Call MxShell(ShellCmd, True)
                    Call WriteLogEntry("FinalAlert 2 YR has terminated.", LogLevel1)
                End If
                Call RestoreProcess
            Else
                If cboxTX.Value <> 1 Then Call WriteLogEntry("Launch Base was run with the -tx switch however the Terrain Expansion cannot be incorporated with the specified FinalAlert 2 mod!", LogMsgBox)
            End If
        Else
            Call WriteLogEntry("Cannot activate " & Mods(iMod).ModName & " because Terrain Expansion prerequisites not satisfied.", LogMsgBox)
        End If
    Case TypeProgram
        If CL_modnum = -1 Then
            iMod = Val(lstMods(Index).ItemData(lstMods(Index).ListIndex))
        Else
            iMod = CL_modnum
            Call SelectTab(Mods(iMod).ModType)
            Call DisplayModDetails(iMod, Index, False)
        End If
        Call WriteLogEntry("Preparing to launch tool: " & Mods(iMod).ModName, LogLevel1)
        If OptAutoUpdate Then
            bOk = True
            Call LaunchMod_CheckForUpdate(bShutdown, iMod, Index, bOk)
            If (Not bOk) Or bShutdown Then GoTo AbortLaunch
        End If
        If OptModSound2 Then Call PlaySound(Mods(iMod).ModSound2, True)
        ShellCmd = Quote(Mods(iMod).ModProgram)
        If Len(txtModParams.Text) <> 0 Then ShellCmd = ShellCmd & " " & txtModParams.Text
        Call WriteLogEntry("Launching " & Mods(iMod).ModName, LogLevel1)
        If Mods(iMod).ModShutdownLB Then
            Call Shutdown(False)
            bShutdown = True
        End If
        Call Shell(ShellCmd, vbNormal)
    End Select
AbortLaunch:
    Select Case True
    Case bShutdown
        Unload Me
    Case CL_modnum <> -1
        Call Shutdown
    Case Else
        frmMain.Enabled = True
        If Mods(iMod).ModType <> TypeProgram Then Call frmMain.SetFocus
        Call CallStackPop
    End Select
End Sub

Private Function LaunchMod_CheckGameVersion(ByVal iMod As Long, ByRef bModified As Boolean) As Boolean
    Dim bOk As Boolean
    Dim mbResult As VbMsgBoxResult
    bOk = True
    If Mods(iMod).ModIsForRA2 Then
        If Mods(RA2ModNum).ModVersion <> "1.006" Then
            If iMod <> RA2ModNum Then
                Call WriteLogEntry("Cannot launch a Red Alert 2 mod without version 1.006 of Red Alert 2.")
                Call MsgBox("You cannot launch a Red Alert 2 mod if you don't have version 1.006 of Red Alert 2." & vbCrLf & "Please upgrade your installation of Red Alert 2.", vbOKOnly + vbInformation, App.Title)
                bOk = False
            End If
        Else
            'checksum - only check if latest version 'cause we don't know the other checksums
            If OptGameChecksums Then
                If GetFileMD5(JoinPath(RA2DIR, "ra2.exe")) <> GameChecksum("RA2", TFD) Then
                    bOk = False
                Else
                    If GetFileMD5(JoinPath(RA2DIR, "game.exe")) <> GameChecksum("GAME", TFD) Then bOk = False
                End If
                If bOk = False Then
                    bModified = True
                    Call WriteLogEntry("One or more game executable files have been modified by a third party!")
                    bOk = (MsgBox("One or more game executable files have been modified by a third party!" & vbCrLf & "Press OK to launch the game anyway or press Cancel to abort.", vbOKCancel + vbInformation, App.Title) = vbOK)
                End If
            Else
                Call WriteLogEntry("Verify Executables is disabled - Red Alert 2 verification skipped.", LogLevel1)
            End If
        End If
    Else
        If Mods(YRModNum).ModVersion = "" Then
            bOk = False
            Call WriteLogEntry("Cannot launch Yuri's Revenge because it is not installed.")
            Call MsgBox("Yuri's Revenge is not installed." & vbCrLf & "Please install Yuri's Revenge.", vbOKOnly + vbInformation, App.Title)
        Else
            If Mods(YRModNum).ModVersion <> "1.001" Then
               If iMod <> YRModNum Then
                    Call WriteLogEntry("Cannot launch a Yuri's Revenge mod without version 1.001 of Yuri's Revenge.")
                    Call MsgBox("You cannot launch a Yuri's Revenge mod if you don't have version 1.001 of Yuri's Revenge." & vbCrLf & "Please upgrade your installation of Yuri's Revenge.", vbOKOnly + vbInformation, App.Title)
                    bOk = False
                End If
            Else
                If OptGameChecksums Then
                    If GetFileMD5(JoinPath(RA2DIR, "ra2md.exe")) <> GameChecksum("RA2MD", TFD) Then
                        bOk = False
                    Else
                        If GetFileMD5(JoinPath(RA2DIR, "gamemd.exe")) <> GameChecksum("GAMEMD", TFD) Then
                            bOk = False
                        'Else
                        '    If GetFileMD5(JoinPath(RA2DIR, "yuri.exe")) <> GameChecksum("YURI", TFD) Then
                        '        bOk = False
                        '    End If
                        End If
                    End If
                    If bOk = False Then
                        bModified = True
                        Call WriteLogEntry("One or more game executable files have been modified by a third party!")
                        bOk = (MsgBox("One or more game executable files have been modified by a third party!" & vbCrLf & "Press OK to launch the game anyway or press Cancel to abort.", vbOKCancel + vbInformation, App.Title) = vbOK)
                    End If
                Else
                    Call WriteLogEntry("Verify Executables is disabled - Yuri's Revenge verification skipped.", LogLevel1)
                End If
            End If
        End If
    End If
    LaunchMod_CheckGameVersion = bOk
End Function

Private Function LaunchMod_PlayVideo(ByVal iMod As Long) As Boolean
    Dim CancelLaunch As Boolean
    CancelLaunch = False
    If Len(CL_playfile) <> 0 Then
        If FileExists(JoinPath(Mods(iMod).ModPath, CL_playfile)) Then
            Call WriteLogEntry("Command line argument: Automatically selecting video.", LogLevel1)
        Else
            Call WriteLogEntry("Command line argument: Unable to automatically select video because specified video does not exist!")
            CL_playfile = ""
        End If
    Else
        If OptPlay Then
            Call frmPlayVideo.RefreshList(Mods(iMod).ModPath, Mods(iMod).ModVersion)
            If frmPlayVideo.lstVideos.ListCount > 1 Then
                Call frmPlayVideo.Show(vbModal)
                CancelLaunch = frmPlayVideo.CancelLaunch
            End If
            Unload frmPlayVideo
        End If
    End If
    If Not CancelLaunch Then
        If Len(CL_playfile) <> 0 Then
            Call WriteLogEntry("Video file selected: " & Quote(JoinPath(Mods(iMod).ModPath, CL_playfile)), LogLevel1)
            If FileExists(JoinPath(RA2DIR, "session.ipb")) Then
                If FileExists(JoinPath(BACKUPDIR, "session.ipb")) Then
                    Call Kill(JoinPath(BACKUPDIR, "session.ipb"))
                    Call WriteLogEntry("Unexpected backup file found! " & Quote(JoinPath(BACKUPDIR, "session.ipb")) & " deleted to make way for new backup file.")
                End If
                Call LoggedMove(JoinPath(RA2DIR, "session.ipb"), JoinPath(BACKUPDIR, "session.ipb"))
            End If
            Call LoggedCopy(JoinPath(Mods(iMod).ModPath, CL_playfile), JoinPath(RA2DIR, "session.ipb"))
        End If
    End If
    LaunchMod_PlayVideo = Not CancelLaunch
End Function

Private Function LaunchMod_RecordVideo() As Boolean
    Dim RetVal As Boolean
    If (OptRecord And Len(CL_playfile) = 0) Then
        RetVal = True
        Call WriteLogEntry("Preparing to record video.", LogLevel1)
        Call WriteINIStr("Restore", "Recording", "yes", ProgramINI)
        If FileExists(JoinPath(RA2DIR, "session.ipb")) Then
            If FileExists(JoinPath(BACKUPDIR, "session.ipb")) Then
                Call WriteLogEntry("Unexpected backup file found! Deleting " & Quote(JoinPath(BACKUPDIR, "session.ipb")) & " to make way for new backup file.")
                Call Kill(JoinPath(BACKUPDIR, "session.ipb"))
            End If
            Call LoggedMove(JoinPath(RA2DIR, "session.ipb"), JoinPath(BACKUPDIR, "session.ipb"), True)
        End If
    Else
        RetVal = False
    End If
    LaunchMod_RecordVideo = RetVal
End Function

Private Function LaunchMod_YPLCheckOk(ByVal iMod As Long, ByRef bRestoreTheme As Boolean) As Boolean
    Dim CancelLaunch As Boolean
    Dim sPlaylist As String
    Dim bRemove As Boolean
    Dim sBackup As String
    Dim sTheme As String
    CancelLaunch = False
    If Not Mods(iMod).ModIsForRA2 And OptCheckModYPLFiles Then
        Call frmPlaylist.RefreshList(Mods(iMod).ModPath)
        If frmPlaylist.lstPlaylists.ListCount > 1 Then
            Call frmPlaylist.Show(vbModal)
            CancelLaunch = frmPlaylist.CancelLaunch
            sPlaylist = frmPlaylist.Playlist
        End If
        Unload frmPlaylist
    End If
    If Not CancelLaunch Then
        bRemove = False
        If Len(sPlaylist) = 0 Then
            'use original thememd.ini
            bRemove = True
            Call WriteLogEntry("Playlist file selected: " & "Original Yuri's Revenge playlist.", LogLevel1)
        ElseIf FileType(sPlaylist) = "YPL" Then
            bRemove = True
            Call WriteLogEntry("Playlist file selected: " & Quote(JoinPath(Mods(iMod).ModPath, sPlaylist)), LogLevel1)
        Else
            sPlaylist = ""
            Call WriteLogEntry("Playlist file selected: " & "Active YR Playlist Manager playlist.", LogLevel1)
        End If
        If bRemove Then
            sTheme = JoinPath(RA2DIR, "thememd.ini")
            If FileExists(sTheme) Then
                sBackup = JoinPath(BACKUPDIR, "thememd.ini")
                If FileExists(sBackup) Then
                    Call Kill(sBackup)
                    Call WriteLogEntry("Unexpected backup file found! " & Quote(sBackup) & " deleted to make way for new backup file.")
                End If
                Call LoggedMove(sTheme, sBackup, True, False)
                bRestoreTheme = True 'remember to restore it after mod launch is done
            End If
            If Len(sPlaylist) <> 0 Then Call LoggedCopy(JoinPath(Mods(iMod).ModPath, sPlaylist), sTheme)
        End If
    End If
    LaunchMod_YPLCheckOk = Not CancelLaunch
End Function

Private Sub LaunchMod_CheckForUpdate(ByRef bShutdown As Boolean, ByVal iMod As Integer, Optional ByVal iType As Integer = TypeMod, Optional ByRef bLaunching As Boolean = False, Optional ByRef bSingle As Boolean = False)
'returns true if we tried to download (no indication of whether or not this was successful)
    Dim sLocal As String
    Dim sRemote As String
    Dim mbResult As VbMsgBoxResult
    Dim sName As String
    Dim sVersion As String
    Dim iRecord As Integer
    Dim iRecordCount As Integer
    Call CallStackPush(Me.Name & ".LaunchMod_CheckForUpdate(" & CStr(bShutdown) & ", " & CStr(iMod) & ", " & CStr(iType) & ", " & CStr(bLaunching) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    bShutdown = False
    Select Case iType
    Case TypePlugin
        sName = Plugins(iMod).PluginName
        sVersion = Plugins(iMod).PluginVersion
        iRecord = 0
        iRecordCount = PluginCount
    Case Else
        sName = Mods(iMod).ModName
        sVersion = Mods(iMod).ModVersion
        iRecord = 1
        iRecordCount = ModCount
    End Select
    'please wait dialog
    Call WriteLogEntry("Checking for updates to " & sName & "...", LogLevel1)
    If iType = TypePlugin Or iMod <> LBModNum Then Call ShowPleaseWait("Checking for updates to " & sName & "...", "Checking for installed legacy versions.")
    'check for legacy versions
    If iMod <> LBModNum Then
        Do While iRecord <= iRecordCount
            If iRecord <> iMod Then
                If iType = TypePlugin Then
                    'check by plugin id
                    If Plugins(iRecord).PluginID = Plugins(iMod).PluginID Then
                        If CompareVersions(Plugins(iRecord).PluginVersion, ">", sVersion) Then Exit Do
                    End If
                Else
                    'check by update check url
                    If Len(Mods(iMod).ModUpdateCheckURL) <> 0 And Mods(iRecord).ModUpdateCheckURL = Mods(iMod).ModUpdateCheckURL Then
                        If CompareVersions(Mods(iRecord).ModVersion, ">", sVersion) Then Exit Do
                    Else
                        'check by name and type
                        If Mods(iRecord).ModType = iType Then
                            If Mods(iRecord).ModName = sName Then
                                If CompareVersions(Mods(iRecord).ModVersion, ">", sVersion) Then Exit Do
                            End If
                        End If
                    End If
                End If
            End If
            iRecord = iRecord + 1
        Loop
    Else
        iRecord = iRecordCount + 1
    End If
    If iRecord = (iRecordCount + 1) Then
        'not a legacy version - find the update record and update it
        If iType = TypePlugin Or iMod <> LBModNum Then Call UpdatePleaseWait(, "Identifying update record.")
        If iType = TypePlugin Then
            Call frmModCat.UpdateUpdateRecord(0)
            iRecord = 1
            Do While iRecord <= UpdateRecordCount
                If UpdateRecords(iRecord).ModPluginID = Plugins(iMod).PluginID Then
                    If iType = TypePlugin Or iMod <> LBModNum Then Call UpdatePleaseWait(, "Updating update record.")
                    Call frmModCat.UpdateUpdateRecord(iRecord)
                    Exit Do
                End If
                iRecord = iRecord + 1
            Loop
        Else
            If Len(Mods(iMod).ModUpdateCheckURL) <> 0 Then
                'search by modnum because it might have been replaced by something in the online catalogue
                iRecord = 0
                Do While iRecord <= UpdateRecordCount
                    If UpdateRecords(iRecord).CheckModNum = iMod Then
                        If iType = TypePlugin Or iMod <> LBModNum Then Call UpdatePleaseWait(, "Updating update record.")
                        Call frmModCat.UpdateUpdateRecord(iRecord)
                        Exit Do
                    End If
                    iRecord = iRecord + 1
                Loop
            Else
                iRecord = UpdateRecordCount + 1
            End If
        End If
        'hide the dialog
        If iMod <> LBModNum Or iType = TypePlugin Then
            'we aren't in the splash screen so there is a dialog to hide
            Call HidePleaseWait
            Call Me.SetFocus
        End If
        If iRecord <> UpdateRecordCount + 1 Then
            'offer update if there is one
            If Len(UpdateRecords(iRecord).ModLatestVersion) <> 0 Then
                If CompareVersions(UpdateRecords(iRecord).ModLatestVersion, ">", sVersion) Then
                    sRemote = UpdateRecords(iRecord).CheckDownloadURL
                    If Len(sRemote) <> 0 Then
                        'Offer
                        Call WriteLogEntry("Update available. Latest version of " & sName & " is " & UpdateRecords(iRecord).ModLatestVersion & ".", LogLevel1)
                        mbResult = MsgBox("An update is available for " & sName & "." & vbCrLf & "Your version: " & vbTab & sVersion & vbCrLf & "Latest version: " & vbTab & UpdateRecords(iRecord).ModLatestVersion & vbCrLf & "Download size: " & vbTab & DataSize(UpdateRecords(iRecord).CheckDownloadSize) & vbCrLf & vbCrLf & "Would you like to download this update now?", vbYesNo + vbQuestion, App.Title)
                        If mbResult = vbYes Then
                            bLaunching = False
                            Call WriteLogEntry("Checking for free disk space on drive " & UCase$(Left(EXEDIR, 1)), LogLevel1)
                            If FreeDiskSpace(Left(EXEDIR, 1)) > (OptSafetySpace + UpdateRecords(iRecord).CheckDownloadSize) Then
                                Call frmMain.WriteLogEntry("Downloading " & Quote(sRemote) & " to " & Quote(sLocal), LogLevel1)
                                Call frmDownloading.Show
                                Call frmDownloading.DownloadProgress(0, UpdateRecords(iRecord).CheckDownloadSize)
                                sLocal = JoinPath(SETUPDIR, GetFileName(sRemote))
                                If theInternet.CopyURLToFile(sRemote, sLocal, frmDownloading) Then
                                    If Not theInternet.DownloadCancelled Then
                                        Call frmMain.WriteLogEntry("Download complete.", LogLevel1)
                                        Call CheckForUpdate_DownloadHistory(sLocal, iType, sName, sVersion, IIf(UpdateRecords(iRecord).CheckUpdateOnly, UpdateRecords(iRecord).ModUserVersion, ""))
                                        Unload frmDownloading
                                        mbResult = MsgBox("Launch Base will now close so that " & UpdateRecords(iRecord).ModName & " can be installed.", vbOKCancel + vbInformation, App.Title)
                                        If mbResult = vbOK Then
                                            Call WriteLogEntry("Executing " & sLocal & "...", LogLevel1)
                                            Call Shutdown(False, True)
                                            Call Shell(Quote(sLocal), vbNormal)
                                            Unload Me
                                            Exit Sub
                                        Else
                                            Call WriteLogEntry("User chose not to execute " & sLocal, LogLevel1)
                                        End If
                                    Else
                                        Unload frmDownloading
                                        Call WriteLogEntry("Download cancelled by user.", LogLevel1)
                                    End If
                                Else
                                    Unload frmDownloading
                                    Call WriteLogEntry("Failed to download " & Quote(sRemote) & " to " & Quote(sLocal), LogLevel1)
                                    Call MsgBox("Failed to download " & Quote(sRemote), vbOKOnly + vbExclamation, App.Title)
                                End If
                            Else
                                Call WriteLogEntry("Insufficient free disk space on drive " & UCase$(Left(EXEDIR, 1)) & " to download the update.", LogMsgBox)
                            End If
                        Else
                            Call WriteLogEntry("User chose not to download the update.", LogLevel1)
                        End If
                    Else
                        Call WriteLogEntry("A newer version of " & sName & " is available however no download locations could be found!", LogLevel1)
                        If bLaunching Then
                            bLaunching = (MsgBox("A newer version of " & sName & " is available however no download locations could be found!" & vbCrLf & "Visit the " & sName & " website to download the latest version." & vbCrLf & vbCrLf & "Do you wish to abort launching " & sName & "?", vbYesNo + vbQuestion, App.Title) = vbNo)
                        Else
                            If iMod <> LBModNum Then Call MsgBox("A newer version of " & sName & " is available however no download locations could be found!" & vbCrLf & "Visit the " & sName & " website to download the latest version.", vbOKOnly + vbInformation, App.Title)
                        End If
                    End If
                Else
                    Call WriteLogEntry("No updates available - " & sName & " is up to date.", IIf(bLaunching Or (iType <> TypePlugin And iMod = LBModNum), LogLevel1, LogLevel1 + LogMsgBox))
                End If
            Else
                'failed to update record. this has already been logged but no messagebox
                If Not bLaunching And iMod <> LBModNum Then
                    Call MsgBox("There was a problem downloading the update check file for " & sName & "." & vbCrLf & "Visit the " & sName & " website to download the latest version.", vbOKOnly + vbInformation, App.Title)
                End If
            End If
        Else
            'no update record or no url provided
            If iType = TypePlugin Then
                Call WriteLogEntry("Update check aborted - plugin " & sName & " was not found in the catalogue.", LogLevel1)
                If bSingle Then Call MsgBox(sName & " was not found in the update catalogue." & vbCrLf & sName & " may not be an officially recognised plugin.", vbOKOnly + vbInformation, App.Title)
            Else
                If Len(Mods(iMod).ModUpdateCheckURL) <> 0 Then
                    Call Panic("Missing update record for " & sName & ".")
                Else
                    Call WriteLogEntry("Update check aborted - " & sName & " has not specified an update check URL.", IIf(bLaunching, LogLevel1, LogMsgBox))
                    If bSingle Then Call MsgBox(sName & " has no update check URL so an update check cannot be performed." & vbCrLf & "Check the " & sName & " website to see if there are any updates.", vbOKOnly + vbInformation, App.Title)
                End If
            End If
        End If
    Else
        Call WriteLogEntry("Update check aborted - user is activating a legacy version of " & sName & ".", LogLevel1)
        'we can't be in the splash (no legacy version for LB itself) screen so there is a dialog to hide
        Call HidePleaseWait
        If bSingle Then Call MsgBox("You already have another, more up-to-date version of " & sName & "." & vbCrLf & "Update checks are not performed for legacy versions of mods.", vbOKOnly + vbInformation, App.Title)
        Call Me.SetFocus
    End If
    Call CallStackPop
End Sub

Public Sub CheckForUpdate_DownloadHistory(ByVal sFilePath As String, ByVal iModType As Integer, ByVal sName As String, ByVal sNewVersion As String, Optional ByVal sOldVersion As String = "")
    Dim sHistory As String
    Dim sFileName As String
    Call WriteLogEntry("Updating download history...", LogLevel1)
    sFileName = UCase$(GetFileName(sFilePath))
    Call WriteINIStr(sFileName, "ModType", CStr(iModType), SetupsINI)
    Call WriteINIStr(sFileName, "Name", sName, SetupsINI)
    Call WriteINIStr(sFileName, "NewVersion", sNewVersion, SetupsINI)
    Call WriteINIStr(sFileName, "OldVersion", sOldVersion, SetupsINI)
    Call WriteINIStr(sFileName, "Date", PadNum(Year(Now), 4) & "-" & PadNum(Month(Now), 2) & "-" & PadNum(Day(Now), 2) & " " & PadNum(Hour(Now), 2) & ":" & PadNum(Minute(Now), 2), SetupsINI)
End Sub

Private Sub RestoreProcess(Optional ByVal iTime As Long = -1)
    Dim sModPath As String
    Dim sModName As String
    Dim sModVersion As String
    Dim sModNewVersion As String
    Dim iModType As Integer
    Dim sSnapFormat As String
    Dim sScrnFormat As String
    Dim bOfficial As Boolean
    Dim bGameIsRA2 As Boolean
    Dim iPos As Integer
    Call CallStackPush(Me.Name & ".RestoreProcess(" & CStr(iTime) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Restore process started.", LogLevel1)
    sModName = ReadINIStr("Mod", "Name", ProgramINI)
    sModVersion = ReadINIStr("Mod", "Version", ProgramINI)
    iModType = Val(ReadINIStr("Mod", "ModType", ProgramINI))
    sSnapFormat = ReadINIStr("Mod", "SnapFormat", ProgramINI)
    sScrnFormat = ReadINIStr("Mod", "ScrnFormat", ProgramINI)
    bOfficial = BooleanStringToBoolean(ReadINIStr("Mod", "Official", ProgramINI, "no"))
    bGameIsRA2 = BooleanStringToBoolean(ReadINIStr("Mod", "GameIsRA2", ProgramINI))
    'user-generated content
    If iModType = TypeMod Then
        If FileExists(JoinPath(RA2DIR, "syringe.log")) Then
            If FileExists(JoinPath(RESDIR, "syringe.log")) Then
                Call LoggedKill(JoinPath(RESDIR, "syringe.log"))
            End If
            Call LoggedMove(JoinPath(RA2DIR, "syringe.log"), JoinPath(RESDIR, "syringe.log"))
        End If
        sModPath = DeactivateMod_CheckModGone(sModName, sModVersion, sModNewVersion, sScrnFormat, sSnapFormat, bOfficial, bGameIsRA2, True)
        Call DeactivateMod_GameConfig(sModPath, bGameIsRA2)
        Call DeactivateMod_SaveGames(sModPath, sModVersion, sModNewVersion, sSnapFormat, True)
        Call DeactivateMod_Screenshots(sModPath, sScrnFormat, True)
        Call RestoreProcess_IPBVideo(sModPath, iTime)
        Call RestoreProcess_Logs
        Call RestoreProcess_AresDebug(iTime)
    End If
    'uninstall mod if persistent mods turned off
    If Not OptPersistentMod Or CL_dev Then
        Call DeactivateMod(True, False) 'this will restore residual files too
    End If
    'thememd.ini
    If BooleanStringToBoolean(ReadINIStr("Restore", "Theme", ProgramINI)) Then
        If FileExists(JoinPath(RA2DIR, "thememd.ini")) Then Call LoggedKill(JoinPath(RA2DIR, "thememd.ini"))
        If FileExists(JoinPath(BACKUPDIR, "thememd.ini")) Then Call LoggedMove(JoinPath(BACKUPDIR, "thememd.ini"), JoinPath(RA2DIR, "thememd.ini"), False, True)
    End If
    'Now that the restore process is complete, we can blank out the 'restore pending' record.
    Call WriteINIStr("Restore", "RestorePending", "no", ProgramINI)
    Call WriteINIStr("Restore", "Recording", "", ProgramINI)
    Call WriteLogEntry("Restore process complete.", LogLevel1)
    Call CallStackPop
End Sub

Private Function DeactivateMod_CheckModGone(ByVal sModName As String, ByVal sModVersion As String, ByRef sModNewVersion As String, ByVal sScrnFormat As String, ByVal sSnapFormat As String, ByVal bOfficial As Boolean, ByVal bGameIsRA2 As Boolean, Optional ByVal RestoreMode As Boolean = False) As String 'returns path of a mod folder to put the files
    Dim iCounter As Integer
    Dim bUserStuff As Boolean
    Dim sModPath As String
    Dim sModUserdata As String
    Dim sModLiblist As String
    Dim sTargetVersion As String
    Dim MsgBoxResult As VbMsgBoxResult
    Dim fso As FileSystemObject
    Dim fso_folder As Folder
    Dim fso_file As File
    Call CallStackPush(Me.Name & ".DeactivateMod_CheckModGone(" & CStr(sModName) & ", " & CStr(sModVersion) & ", " & CStr(sModNewVersion) & ", " & CStr(sScrnFormat) & ", " & CStr(sSnapFormat) & ", " & CStr(bOfficial) & ", " & CStr(bGameIsRA2) & ", " & CStr(RestoreMode) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    If bOfficial Then
        If bGameIsRA2 Then
            sModPath = Mods(RA2ModNum).ModPath
            sModNewVersion = Mods(RA2ModNum).ModVersion
        Else
            sModPath = Mods(YRModNum).ModPath
            sModNewVersion = Mods(YRModNum).ModVersion
        End If
    Else
        sModPath = ""
        For iCounter = 1 To ModCount
            If ModIsInstalled(iCounter) Then
                sModUserdata = JoinPath(Mods(iCounter).ModPath, "launcher\userdata.lbu")
                sModLiblist = JoinPath(Mods(iCounter).ModPath, "launcher\liblist.gam")
                If ReadINIStr("General", "Name", sModLiblist) = sModName Then
                    sModPath = Mods(iCounter).ModPath
                    sModNewVersion = ReadINIStr("General", "Version", sModLiblist)
                    Exit For
                Else
                    Call WriteINIStr("General", "ModIsActive", "no", sModUserdata)
                End If
            End If
        Next iCounter
    End If
    If Len(sModPath) = 0 And Not RestoreMode Then
        'mod has been deleted - check if there are any files to copy back
        Set fso_folder = fso.GetFolder(RA2DIR)
        bUserStuff = False
        For Each fso_file In fso_folder.Files
            If FileIsSaveGame(fso_file.Name) Then
                bUserStuff = True
                Exit For
            ElseIf ConfirmScrnFormat(fso_file.Name, sScrnFormat) Then
                bUserStuff = True
                Exit For
            ElseIf ConfirmScrnFormat(fso_file.Name, sSnapFormat) Then
                bUserStuff = True
                Exit For
            End If
        Next
        If bUserStuff Then
            MsgBoxResult = MsgBox(Quote(sModName) & " has either been updated or removed from Launch Base since it was installed." & vbCrLf & "There are one or more user-generated files (such as saved games or screenshots) for this mod in the Red Alert 2 game directory that were waiting to be copied back to the mod's folder." & vbCrLf & "Do you want these files to be saved?", vbYesNo + vbQuestion, App.Title)
            If MsgBoxResult = vbYes Then
                dialogOpen.FileName = JoinPath(EXEDIR, "Mods\Save Files Here")
                dialogOpen.DialogTitle = "Select a directory to save files..."
                dialogOpen.Filter = "Save Files Here"
                dialogOpen.DefaultExt = ""
                Do
                    sModPath = ""
                    On Error GoTo CancelUGDir
                    dialogOpen.ShowSave
                    On Error GoTo 0
                    sModPath = GetFilePath(dialogOpen.FileName)
                    If DirExists(sModPath) Then
                        sModLiblist = JoinPath(sModPath, "launcher\liblist.gam")
                        If FileExists(sModLiblist) Then
                            If sModName = ReadINIStr("General", "Name", sModLiblist) Then
                                sTargetVersion = ReadINIStr("General", "Version", sModLiblist)
                                If CompareVersions(sTargetVersion, "=", sModVersion) Then
                                    MsgBoxResult = MsgBox("The directory you selected appears to contain the same version of the mod that the user-generated files were created with." & vbCrLf & vbCrLf & "Would you like to choose another directory?" & vbCrLf & "If you click 'No' then ALL existing saved games in the folder will be deleted before the user-generated files are saved here." & vbCrLf & "If you click 'Cancel' then the user-generated files will be deleted.", vbYesNoCancel + vbQuestion, App.Title)
                                    If MsgBoxResult = vbCancel Then sModPath = ""
                                    Select Case MsgBoxResult
                                    Case vbNo, vbCancel: Exit Do
                                    End Select
                                Else
                                    MsgBoxResult = MsgBox("The directory you selected appears to contain a different version of the mod that the user-generated files were created with." & vbCrLf & vbCrLf & "Would you like to choose another directory?" & vbCrLf & "If you click 'No' then any saved games will be stored in the 'oldsaves' directory, and may overwrite existing 'old saved games' without warning." & vbCrLf & "If you click 'Cancel' then the user-generated files will be deleted.", vbYesNoCancel + vbQuestion, App.Title)
                                    If MsgBoxResult = vbCancel Then sModPath = ""
                                    Select Case MsgBoxResult
                                    Case vbNo, vbCancel: Exit Do
                                    End Select
                                End If
                            Else
                                MsgBoxResult = MsgBox("The directory you selected contains another mod." & vbCrLf & "You cannot save the user-generated files here." & vbCrLf & vbCrLf & "Would you like to choose another directory?" & vbCrLf & "If you click 'No' then the user-generated files will be deleted.", vbYesNo + vbQuestion, App.Title)
                                If MsgBoxResult = vbNo Then
                                    sModPath = ""
                                    Exit Do
                                End If
                            End If
                        Else
                            If DirIsEmpty(sModPath) Then
                                Exit Do
                            Else
                                MsgBoxResult = MsgBox("The directory you selected is not empty." & vbCrLf & "It is not recommended to save the files here." & vbCrLf & vbCrLf & "Would you like to choose another directory?" & vbCrLf & "If you click 'No' then the user-generated files will be saved in the selected directory, and may overwrite existing files without warning." & vbCrLf & "If you click 'Cancel' then the user-generated files will be deleted.", vbYesNoCancel + vbQuestion, App.Title)
                                If MsgBoxResult = vbCancel Then sModPath = ""
                                Select Case MsgBoxResult
                                Case vbNo, vbCancel: Exit Do
                                End Select
                            End If
                        End If
                    Else
                        Call MakePath(sModPath)
                    End If
                    GoTo LoopUGDir
CancelUGDir:
                    MsgBoxResult = MsgBox("Are you sure that you want the user-generated files to be deleted?", vbYesNo, App.Title)
                    If MsgBoxResult = vbYes Then Exit Do
LoopUGDir:
                Loop
            End If
        End If
    End If
    DeactivateMod_CheckModGone = sModPath
    Call CallStackPop
End Function

Private Sub DeactivateMod_SaveGames(ByVal sModPath As String, ByVal sModVersion As String, ByVal sModNewVersion As String, ByVal sSnapFormat As String, Optional ByVal RestoreMode As Boolean = False)
    Dim sDestFile As String
    Dim sModSaveDir As String
    Dim fso As FileSystemObject
    Dim fso_folder As Folder
    Dim fso_file As File
    Call CallStackPush(Me.Name & ".DeactivateMod_SaveGames(" & CStr(sModPath) & ", " & CStr(sModVersion) & ", " & CStr(sModNewVersion) & ", " & CStr(sSnapFormat) & ", " & CStr(RestoreMode) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Processing generated save games.", LogLevel1)
    Set fso = New FileSystemObject
    If Len(sModPath) <> 0 Then
        sModSaveDir = JoinPath(sModPath, "saves")
        'check we are saving to the correct version
        If CompareVersions(sModVersion, "=", sModNewVersion) Then
            'delete existing saves to make way for new files
            If DirExists(sModSaveDir) Then
                Set fso_folder = fso.GetFolder(sModSaveDir)
                For Each fso_file In fso_folder.Files
                    Call LoggedKill(fso_file.Path)
                Next
            End If
        Else
            'not the correct version so will put files in oldsaves instead
            sModSaveDir = JoinPath(sModSaveDir, "oldsaves")
        End If
    End If
    'add the new/replacement saves and map snapshots
    Set fso_folder = fso.GetFolder(RA2DIR)
    For Each fso_file In fso_folder.Files
        If FileIsSaveGame(fso_file.Name) Or ConfirmScrnFormat(fso_file.Name, sSnapFormat) Then
            If Len(sModPath) <> 0 Then
                If Not DirExists(sModSaveDir) Then Call LoggedMkDir(sModSaveDir)
                sDestFile = JoinPath(sModSaveDir, fso_file.Name)
                If FileExists(sDestFile) Then Call LoggedKill(sDestFile)
                Call LoggedMove(fso_file.Path, sDestFile)
            Else
                'mod has been deleted
                If Not RestoreMode Then Call LoggedKill(fso_file.Path) 'user has authorised file deletion
            End If
        End If
    Next
    Call CallStackPop
End Sub

Private Sub RestoreProcess_IPBVideo(ByVal sModPath As String, Optional ByVal iTime As Long = -1)
    Dim sFile As String
    Dim sDestFile As String
    Dim sModLiblist As String
    Dim MsgBoxResult As VbMsgBoxResult
    Call CallStackPush(Me.Name & ".RestoreProcess_IPBVideo(" & CStr(sModPath) & ", " & CStr(iTime) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    sModLiblist = JoinPath(sModPath, "liblist.gam")
    sFile = JoinPath(RA2DIR, "session.ipb")
    If BooleanStringToBoolean(ReadINIStr("Restore", "Recording", ProgramINI, "no")) Then
        If FileExists(sFile) Then
            If iTime > 30 Or iTime = -1 Then
                'we probably want this video
                If Len(sModPath) <> 0 Then
                    Call frmSaveVideo.Show(vbModal)
                    If Not frmSaveVideo.CancelVideoSave Then
                        sDestFile = PadNum(Year(Now()), 4) & "-" & PadNum(Month(Now()), 2) & "-" & PadNum(Day(Now()), 2) & " " & PadNum(Hour(Now()), 2) & "." & PadNum(Minute(Now()), 2) & "." & PadNum(Second(Now()), 2) & " [" & ReadINIStr("Mod", "Version", ProgramINI) & "] - " & frmSaveVideo.txtVideoDescription.Text & ".ipb"
                        sDestFile = JoinPath(sModPath, sDestFile)
                        If FileExists(sDestFile) Then Call Kill(sDestFile)
                        Call LoggedMove(sFile, sDestFile)
                    Else
                        Call WriteLogEntry("User chose not to save recorded video.", LogLevel1)
                        Call LoggedKill(sFile)
                    End If
                    Unload frmSaveVideo
                Else
                    dialogOpen.FileName = JoinPath(EXEDIR, "Mods\No description.ipb")
                    dialogOpen.DialogTitle = App.Title & ": Save Video"
                    dialogOpen.Filter = "IPB Videos (*.ipb)|*.ipb"
                    dialogOpen.DefaultExt = "ipb"
                    Do
                        sDestFile = ""
                        On Error GoTo CancelIPBSave
                        dialogOpen.ShowSave
                        On Error GoTo 0
                        sDestFile = GetFileName(dialogOpen.FileName)
                        If DirExists(GetFilePath(dialogOpen.FileName)) Then
                            If FileExists(dialogOpen.FileName) Then
                                MsgBoxResult = MsgBox(Quote(sDestFile) & " already exists." & vbCrLf & "Are you sure you want to overwrite this file?" & vbCrLf & vbCrLf & "Click 'No' to choose a different filename." & vbCrLf & "If you click 'Cancel' then the video will not be saved.", vbYesNoCancel, App.Title)
                                If MsgBoxResult = vbCancel Then sDestFile = ""
                                Select Case MsgBoxResult
                                Case vbYes
                                    Call Kill(dialogOpen.FileName)
                                    Exit Do
                                Case vbCancel
                                    sDestFile = ""
                                    Exit Do
                                End Select
                            Else
                                Exit Do
                            End If
                        Else
                            Call MakePath(GetFilePath(dialogOpen.FileName))
                            Exit Do
                        End If
                        GoTo LoopIPBSave
CancelIPBSave:
                        MsgBoxResult = MsgBox("Are you sure that you do not want to save the recorded video?", vbYesNo, App.Title)
                        If MsgBoxResult = vbYes Then
                            sDestFile = ""
                            Exit Do
                        End If
LoopIPBSave:
                    Loop
                    If Len(sDestFile) <> 0 Then
                        sDestFile = PadNum(Year(Now()), 4) & "-" & PadNum(Month(Now()), 2) & "-" & PadNum(Day(Now()), 2) & " " & PadNum(Hour(Now()), 2) & "." & PadNum(Minute(Now()), 2) & "." & PadNum(Second(Now()), 2) & " [" & ReadINIStr("Mod", "Version", ProgramINI) & "] - " & sDestFile
                        sDestFile = JoinPath(GetFilePath(dialogOpen.FileName), sDestFile)
                        If FileExists(sDestFile) Then sDestFile = dialogOpen.FileName
                        Call LoggedMove(sFile, sDestFile)
                    Else
                        Call WriteLogEntry("User chose not to save recorded video.", LogLevel1)
                        Call LoggedKill(sFile)
                    End If
                End If
            Else
                'we don't want this video because it is too short to be interesting
                Call WriteLogEntry("Game ran for less than 30 seconds.", LogLevel1)
                Call LoggedKill(sFile)
            End If
        Else
            'we were expecting a video but didn't get one
            If iTime > 30 Or iTime = -1 Then
                'video was long enough to be interesting
                If BooleanStringToBoolean(ReadINIStr("Mod", "GameIsRA2", ProgramINI)) Then
                    Call WriteLogEntry("Red Alert 2 failed to record a session.ipb file.", LogMsgBox)
                Else
                    Call WriteLogEntry("Yuri's Revenge failed to record a session.ipb file.", LogMsgBox)
                End If
            Else
                'we don't really care
                Call WriteLogEntry("Yuri's Revenge failed to record a session.ipb file, however the game ran for less than 30 seconds so we don't really care.")
            End If
        End If
    End If
    If FileExists(sFile) Then Call LoggedKill(sFile)
    sDestFile = sFile
    sFile = JoinPath(BACKUPDIR, "session.ipb")
    If FileExists(sFile) Then Call LoggedMove(sFile, sDestFile, False, True)
    Call CallStackPop
End Sub

Private Sub RestoreProcess_Logs()
    Dim sFile As String
    Dim sDestFile As String
    Dim mbResult As VbMsgBoxResult
    Dim sMod As String
    Dim h As Long
    Dim r As Long
    Dim FD As WIN32_FIND_DATA
    Call CallStackPush(Me.Name & ".RestoreProcess_Logs()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    sMod = StripInvalidChars(ReadINIStr("Mod", "Name", ProgramINI), InvalidFileChars) & " [" & StripInvalidChars(ReadINIStr("Mod", "Version", ProgramINI), InvalidFileChars) & "] - " & PadNum(Year(Now()), 4) & "-" & PadNum(Month(Now()), 2) & "-" & PadNum(Day(Now()), 2) & " " & PadNum(Hour(Now()), 2) & "." & PadNum(Minute(Now()), 2) & "." & PadNum(Second(Now()), 2) & " - "
    'Except.txt
    h = FindFirstFile(JoinPath(RA2DIR, "except*.txt"), FD)
    If h <> INVALID_HANDLE_VALUE Then
        Do
            sFile = Left$(FD.cFileName, InStr(FD.cFileName, vbNullChar) - 1)
            If Not ((FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY) Then
                'remove residual files should have gotten rid of any except.txt files
                'so we can assume that this is a new one
                If OptLogExcept Then
                    If OptLogExceptDesc Then
                        Call frmSaveExcept.Show(vbModal)
                        Call LoggedMove(sFile, JoinPath(LOGDIR, sMod & "Except - " & frmSaveExcept.txtExceptDescription.Text & ".txt"))
                        Unload frmSaveExcept
                    Else
                        Call LoggedMove(sFile, JoinPath(LOGDIR, sMod & "Except.txt"))
                    End If
                Else
                    Call LoggedKill(sFile)
                End If
                If BooleanStringToBoolean(ReadINIStr("Restore", "ModifiedEXE", ProgramINI)) Then Call MsgBox("The Internal Error you encountered whilst running the game may have been the result of a modified executable in your Red Alert 2 directory." & vbCrLf & "It is strongly recommended that you remove any third-party patches.", vbOKOnly + vbInformation, App.Title)
            End If
        Loop While FindNextFile(h, FD)
        r = FindClose(h): Debug.Assert r
    End If
    Call CallStackPop
End Sub

Private Sub RestoreProcess_AresDebug(ByRef iTime As Long)
    Dim sDebugDir As String
    Dim mbResult As VbMsgBoxResult
    Dim sMod As String
    Dim fso As FileSystemObject
    Dim fsoDir As Folder
    Dim fsoFile As File
    Dim dateNow As Date
    Dim sFile As String
    Call CallStackPush(Me.Name & ".RestoreProcess_Logs()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    dateNow = Now()
    If BooleanStringToBoolean(ReadINIStr("Mod", "Syringe", ProgramINI, "no")) Then
        If OptCaptureAresDebug And iTime <> -1 Then
            sDebugDir = JoinPath(RA2DIR, "Debug")
            If DirExists(sDebugDir) Then
                sMod = StripInvalidChars(ReadINIStr("Mod", "Name", ProgramINI), InvalidFileChars) & " [" & StripInvalidChars(ReadINIStr("Mod", "Version", ProgramINI), InvalidFileChars) & "] - " & PadNum(Year(Now()), 4) & "-" & PadNum(Month(Now()), 2) & "-" & PadNum(Day(Now()), 2) & " " & PadNum(Hour(Now()), 2) & "." & PadNum(Minute(Now()), 2) & "." & PadNum(Second(Now()), 2) & " - "
                Set fso = New FileSystemObject
                Set fsoDir = fso.GetFolder(sDebugDir)
                For Each fsoFile In fsoDir.Files
                    If DateDiff("s", fsoFile.DateCreated, dateNow) <= iTime Then
                        sFile = fsoFile.Name
                        Call LoggedMove(fsoFile.Path, JoinPath(LOGDIR, sMod & sFile))
                        If BooleanStringToBoolean(ReadINIStr("Restore", "ModifiedEXE", ProgramINI)) Then
                            If Len(sFile) >= 6 Then
                                If LCase$(Left$(sFile, 6)) = "except" Then
                                    Call MsgBox("The Internal Error you encountered whilst running the game may have been the result of a modified executable in your Red Alert 2 directory." & vbCrLf & "It is strongly recommended that you remove any third-party patches.", vbOKOnly + vbInformation, App.Title)
                                End If
                            End If
                        End If
                    End If
                Next
                Set fsoFile = Nothing
                Set fsoDir = Nothing
                Set fso = Nothing
            End If
        End If
    End If
    Call CallStackPop
End Sub

Private Sub DeactivateMod_Screenshots(ByVal sModPath As String, ByVal sScrnFormat As String, Optional ByVal RestoreMode As Boolean = False)
    Dim iCounter As Long
    Dim sPre As String
    Dim sPost As String
    Dim sDestFile As String
    Dim sFile As String
    Dim iScreenNum As Long
    Dim iRetNum As Long
    Dim bOk As Boolean
    Dim fso As FileSystemObject
    Dim fso_folder As Folder
    Dim fso_file As File
    Call CallStackPush(Me.Name & ".DeactivateMod_Screenshots(" & CStr(sModPath) & ", " & CStr(sScrnFormat) & ", " & CStr(RestoreMode) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Processing generated screenshots.", LogLevel1)
    Set fso = New FileSystemObject
    If Len(sModPath) <> 0 Then
        iScreenNum = 0
        Set fso_folder = fso.GetFolder(sModPath)
        For Each fso_file In fso_folder.Files
            If ConfirmScrnFormatReturnNum(fso_file.Name, sScrnFormat, iRetNum) Then
                If iRetNum > iScreenNum Then iScreenNum = iRetNum
            End If
        Next
    End If
    Set fso_folder = fso.GetFolder(RA2DIR)
    Call DisectScrnFormat(sScrnFormat, sPre, sPost, iRetNum)
    For Each fso_file In fso_folder.Files
        If ConfirmScrnFormat(fso_file.Name, sScrnFormat) Then
            If Len(sModPath) <> 0 Then
                iScreenNum = iScreenNum + 1
                sDestFile = JoinPath(sModPath, sPre & PadNum(iScreenNum, iRetNum) & sPost)
                sFile = fso_file.Path
                Call LoggedMove(sFile, sDestFile)
            Else
                'mod has been deleted
                If Not RestoreMode Then Call LoggedKill(fso_file.Path) 'user has authorised file deletion
            End If
        End If
    Next
    Call CallStackPop
End Sub

Private Sub DeactivateMod_ResidualFiles()
    Dim iCounter As Long
    Dim iMissingFileCount As Long
    Dim sFile As String
    Dim sDestFile As String
    Dim sBackupFile As String
    Dim mbResult As VbMsgBoxResult
    Call CallStackPush(Me.Name & ".DeactivateMod_ResidualFiles()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call WriteLogEntry("Restoring residual files...", LogLevel1)
    iCounter = 0
    iMissingFileCount = 0
    sFile = ReadINIStr("Restore", CStr(iCounter), ProgramINI)
    Do While Len(sFile) <> 0
        sDestFile = JoinPath(RA2DIR, sFile)
        sBackupFile = JoinPath(BACKUPDIR, sFile)
        If FileExists(sDestFile) Then
            'do nothing - this should only be an LB plugin file that has been activated since the non-LB plugin file was removed
            If FileExists(sBackupFile) Then
                Call WriteLogEntry(Quote(sBackupFile) & " could not be restored because a file with the same name is already active.")
            Else
                iMissingFileCount = iMissingFileCount + 1
                Call WriteLogEntry(Quote(sBackupFile) & " could not be restored because a file with the same name is already active. In addition, " & Quote(sBackupFile) & " is missing!")
            End If
        Else
            If FileExists(sBackupFile) Then
CheckDiskSpace:
                If FreeDiskSpace(UCase$(Left$(RA2DIR, 1))) > GetFileSize(sBackupFile, True) Then
                    Call LoggedMove(sBackupFile, sDestFile, False, True)
                Else
                    mbResult = MsgBox("There is insufficient free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & " to restore residual file " & vbCrLf & Quote(sBackupFile) & vbCrLf & " to the Red Alert 2 directory." & vbCrLf & "This file will NOT be restored unless you free up some disk space now." & vbCrLf & vbCrLf & "Abort - skip this file and leave it in the Launch Base Backup folder." & vbCrLf & "Retry - check for disk space again and restore the file." & vbCrLf & "Ignore - permanently delete this file.", vbAbortRetryIgnore + vbDefaultButton2 + vbExclamation, App.Title)
                    If mbResult = vbIgnore Then
                        Call WriteLogEntry("Insufficient free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & " to restore residual file " & Quote(sBackupFile) & ". User has chosen to delete the file.")
                        Call LoggedKill(sBackupFile)
                    Else
                        If mbResult = vbRetry Then
                            GoTo CheckDiskSpace
                        Else
                            Call WriteLogEntry("Insufficient free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & " to restore residual file " & Quote(sBackupFile) & ". User has chosen to leave this file in the Launch Base Backup folder.")
                        End If
                    End If
                End If
            Else
                iMissingFileCount = iMissingFileCount + 1
                Call WriteLogEntry("Missing backup file! Expecting to restore " & Quote(sBackupFile) & " but it does not exist.")
            End If
        End If
        Call WriteINIStr("Restore", CStr(iCounter), "", ProgramINI)
        iCounter = iCounter + 1
        sFile = ReadINIStr("Restore", CStr(iCounter), ProgramINI)
    Loop
    If iMissingFileCount <> 0 Then
        Call WriteLogEntry(CStr(iMissingFileCount) & " files that Launch Base was expecting to restore could not be found.")
        Call DenyPersistentMods
        Call DenyPersistentPlugins
    End If
    iCounter = iCounter - iMissingFileCount 'don't need to subtract 1 from iCounter because iCounter starts at 0
    Call WriteLogEntry(CStr(iCounter) & " residual files restored.", LogLevel1)
    Call CallStackPop
End Sub

Private Sub Form_Load()
    Call Init
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then '0 = user clicked X, 1 = programatically, 2 = Windows
        Call Shutdown
    End If
End Sub

Private Sub lblModWebsite_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblModWebsite(Index).ForeColor = ColorURLActive
End Sub

Private Sub lblNoBanner_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picMod_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub menu_ares_Click()
    menu_debugfolder.Enabled = DirExists(JoinPath(RA2DIR, "Debug"))
End Sub

Private Sub menu_aresdoc_Click()
    If OpenLocation(JoinPath(RESDIR, "AresDocumentation\AresManual.html")) < 32 Then Call MsgBox("Unable to open " & Quote(JoinPath(RESDIR, "AresDocumentation\AresManual.html")) & ".", vbOKOnly + vbInformation, App.Title)
End Sub

Private Sub menu_loadcsf_Click()
    Dim csffile As MarshallxCSFClass
    Set csffile = New MarshallxCSFClass
    Call csffile.Initialise
    Call csffile.LoadCSF(JoinPath(EXEDIR, "temp1.csf"))
    Call csffile.UpdateWith(JoinPath(EXEDIR, "unicode.txt"))
    Call csffile.SaveCSF(JoinPath(EXEDIR, "temp2.csf"))
    Set csffile = Nothing
End Sub

Private Sub menu_debugfolder_Click()
    If DirExists(JoinPath(RA2DIR, "Debug")) Then
        If OpenLocation(JoinPath(RA2DIR, "Debug")) < 32 Then Call MsgBox("Failed to open folder: " & JoinPath(RA2DIR, "Debug"), vbOKOnly + vbInformation, App.Title)
    Else
        Call MsgBox("Folder """ & JoinPath(RA2DIR, "Debug") & """ has disappeared!", vbOKOnly + vbInformation, App.Title)
        menu_debugfolder.Enabled = False
    End If
End Sub

Private Sub picMod_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index > MaxType Then
        Select Case Button
        Case vbLeftButton: Call LaunchMod(TypeMod, Val(picMod(Index).Tag))
        Case vbRightButton: Call lstMods_RightClick(Index, X, Y)
        End Select
    End If
End Sub

Private Sub picMod_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index > MaxType Then
        'only for banner page
        lineBannerTop.X1 = picMod(Index).Left - 15
        lineBannerTop.X2 = picMod(Index).Left + picMod(Index).Width + 15
        lineBannerBottom.X1 = picMod(Index).Left - 15
        lineBannerBottom.X2 = picMod(Index).Left + picMod(Index).Width + 15
        lineBannerTop.Y1 = picMod(Index).Top - 15
        lineBannerTop.Y2 = picMod(Index).Top - 15
        lineBannerBottom.Y1 = picMod(Index).Top + picMod(Index).Height
        lineBannerBottom.Y2 = picMod(Index).Top + picMod(Index).Height
        lineBannerLeft.X1 = picMod(Index).Left - 15
        lineBannerLeft.X2 = picMod(Index).Left - 15
        lineBannerRight.X1 = picMod(Index).Left + picMod(Index).Width
        lineBannerRight.X2 = picMod(Index).Left + picMod(Index).Width
        lineBannerLeft.Y1 = picMod(Index).Top
        lineBannerLeft.Y2 = picMod(Index).Top + picMod(Index).Height
        lineBannerRight.Y1 = picMod(Index).Top
        lineBannerRight.Y2 = picMod(Index).Top + picMod(Index).Height
        lineBannerTop.Visible = True
        lineBannerBottom.Visible = True
        lineBannerLeft.Visible = True
        lineBannerRight.Visible = True
        If OptModSound1 Then Call PlaySound(Mods(Val(picMod(Index).Tag)).ModSound1)
    End If
End Sub

Private Sub lblModWebsite_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case vbLeftButton
        If OpenLocation(ReplaceString(lblModWebsite(Index).Caption, "&&", "&")) < 32 Then
            Call MsgBox("Unable to open " & Quote(ReplaceString(lblModWebsite(Index).Caption, "&&", "&")) & ".", vbOKOnly + vbInformation, App.Title)
        End If
    Case vbRightButton
        menu_url.Tag = Index
        Call PopupMenu(menu_rc2, 2, X + lblModWebsite(Index).Left, Y + lblModWebsite(Index).Top)
    End Select
End Sub

Private Sub menu_url_Click()
    Clipboard.Clear
    Clipboard.SetText lblModWebsite(menu_url.Tag).Caption
End Sub

Private Sub lblTabStrip0_Click(Index As Integer)
    Call lblTabStripX_Click(Index)
End Sub

Private Sub lblTabStrip1_Click(Index As Integer)
    Call lblTabStripX_Click(Index)
End Sub

Private Sub lblTabStrip2_Click(Index As Integer)
    Call lblTabStripX_Click(Index)
End Sub

Private Sub lblTabStrip3_Click(Index As Integer)
    Call lblTabStripX_Click(Index)
End Sub

Private Sub lblTabStrip4_Click(Index As Integer)
    Call lblTabStripX_Click(Index)
End Sub

Private Sub lblTabStripX_Click(Index As Integer)
    If Index = 0 Then
        Select Case SelectedTab
        Case 0
            Call SelectTab(4)
            BannerTab = True
        Case 4
            Call SelectTab(0)
            BannerTab = False
        Case Else
            If BannerTab Then
                Call SelectTab(4)
            Else
                Call SelectTab(0)
            End If
        End Select
    Else
        Call SelectTab(Index)
    End If
End Sub


Private Sub lblTX_Click()
    Select Case cboxTX.Value
    Case 0: cboxTX.Value = 1
    Case 1: cboxTX.Value = 0
    End Select
End Sub

Public Sub DisplayPluginDetails(ByVal iPlugin As Integer)
    Dim sFile As String
    Dim iCounter As Integer
    Dim iListIndex As Integer
    Call CallStackPush(Me.Name & ".DisplayPluginDetails(" & CStr(iPlugin) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    lblNoBanner(TypePlugin).Caption = ""
    If iPlugin = -1 Then
        'clear details
        picMod(TypePlugin).ToolTipText = ""
        Set picMod(TypePlugin).Picture = Nothing
        picMod(TypePlugin).Picture = LoadPicture(JoinPath(RESDIR, "nobanner.bmp"))
        lblModVersion(TypePlugin).Caption = ""
        lblModDate(TypePlugin).Caption = ""
        lblModAuthor(TypePlugin).Caption = ""
        lblModWebsite(TypePlugin).Caption = ""
        lblModSize(TypePlugin).Caption = ""
        lblModDescription(TypePlugin).Caption = ""
        cmdManual(TypePlugin).Tag = ""
        cmdLaunch(TypePlugin).Enabled = False
        cmdLaunch(MaxType + 1).Enabled = False 'otherwise we get a deactivate button that will crash
        cmdLaunch(MaxType + 1).Visible = False
    Else
        'display banner
        sFile = JoinPath(Plugins(iPlugin).PluginPath, "launcher\banner.bmp")
        If Not FileExists(sFile) Then
            sFile = JoinPath(Plugins(iPlugin).PluginPath, "launcher\banner.jpg")
            If Not FileExists(sFile) Then
                sFile = JoinPath(Plugins(iPlugin).PluginPath, "launcher\banner.jpeg")
                If Not FileExists(sFile) Then
                    sFile = JoinPath(Plugins(iPlugin).PluginPath, "launcher\banner.gif")
                    If Not FileExists(sFile) Then
                       sFile = JoinPath(RESDIR, "nobanner.bmp")
                       lblNoBanner(TypePlugin).Caption = Plugins(iPlugin).PluginName
                    End If
                End If
            End If
        End If
        picMod(TypePlugin).Picture = LoadPicture(sFile)
        'display details
        picMod(TypePlugin).ToolTipText = DoubleAmpersand(Plugins(iPlugin).PluginName)
        lblModVersion(TypePlugin).Caption = DoubleAmpersand(Plugins(iPlugin).PluginVersion)
        lblModDate(TypePlugin).Caption = DoubleAmpersand(Plugins(iPlugin).PluginDate)
        lblModAuthor(TypePlugin).Caption = DoubleAmpersand(Plugins(iPlugin).PluginAuthor)
        lblModWebsite(TypePlugin).Caption = DoubleAmpersand(Plugins(iPlugin).PluginWebsite)
        If ReadINIStr("Plugin" & Plugins(iPlugin).PluginID, "Version", ProgramINI) = Plugins(iPlugin).PluginVersion Then
            lblModSize(TypePlugin).Caption = DataSize(Plugins(iPlugin).PluginSize) & " (+" & DataSize(ReadINIStr("Plugin" & Plugins(iPlugin).PluginID, "DiskUsage", ProgramINI)) & ")"
        Else
            lblModSize(TypePlugin).Caption = DataSize(Plugins(iPlugin).PluginSize)
        End If
        lblModDescription(TypePlugin).Caption = DoubleAmpersand(Plugins(iPlugin).PluginDescription)
        cmdManual(TypePlugin).Tag = Plugins(iPlugin).PluginManual
        'buttons
        cmdLaunch(TypePlugin).Enabled = True
        'show deactivate button instead?
        cmdLaunch(MaxType + 1).Visible = (GetActivePlugin(Plugins(iPlugin).PluginID) = iPlugin)
    End If
    cmdLaunch(TypePlugin).Visible = Not cmdLaunch(MaxType + 1).Visible
    lblModWebsite(TypePlugin).Visible = Len(lblModWebsite(TypePlugin).Caption) <> 0
    cmdManual(TypePlugin).Enabled = Len(cmdManual(TypePlugin).Tag) <> 0
    cmdLaunch(TypePlugin).Tag = CStr(iPlugin)
    Call CallStackPop
End Sub

Public Sub DisplayModDetails(ByVal iMod As Integer, Optional Index As Integer, Optional ByVal PlayModSound1 As Boolean = True)
    Dim sFile As String
    Dim bOk As Boolean
    Dim iForLoopA As Integer
    Call CallStackPush(Me.Name & ".DisplayModDetails(" & CStr(iMod) & ", " & CStr(Index) & ", " & CStr(PlayModSound1) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    If iMod = -1 Then
        lblNoBanner(Index).Caption = ""
        Set picMod(Index).Picture = Nothing
        picMod(Index).ToolTipText = ""
        picMod(Index).Picture = LoadPicture(JoinPath(RESDIR, "nobanner.bmp"))
        lblModVersion(Index).Caption = ""
        lblModDate(Index).Caption = ""
        lblModAuthor(Index).Caption = ""
        lblModWebsite(Index).Caption = ""
        lblModSize(Index).Caption = ""
        lblModDescription(Index).Caption = ""
        cmdManual(Index).Tag = ""
        cmdLaunch(Index).Enabled = False
        Select Case Index
        Case TypeMod
            lblModCampaigns.Caption = ""
            lblModTX(Index).Caption = ""
            lblModUsesAres.Caption = ""
        Case TypeFA2Mod
            lblModTX(Index).Caption = ""
            lblModFA2.Caption = ""
            cboxTX.Value = 0
            cboxTX.Enabled = False
            cboxTX.Visible = False
            lblTX.Enabled = False
            lblTX.Visible = False
        Case TypeProgram
            txtModParams.Text = ""
            txtModParams.Visible = False
        End Select
    Else
        Index = Mods(iMod).ModType
        lblNoBanner(Index).Caption = ""
        Set picMod(Index).Picture = Nothing
        'display banner
        Select Case iMod
        Case RA2ModNum
            sFile = JoinPath(RESDIR, "ra2banner.bmp")
        Case YRModNum
            sFile = JoinPath(RESDIR, "yrbanner.bmp")
        Case FA2ModNum
            sFile = JoinPath(RESDIR, "fabanner.bmp")
        Case Else
            sFile = JoinPath(Mods(iMod).ModPath, "launcher\banner.bmp")
            If Not FileExists(sFile) Then
                sFile = JoinPath(Mods(iMod).ModPath, "launcher\banner.jpg")
                If Not FileExists(sFile) Then
                    sFile = JoinPath(Mods(iMod).ModPath, "launcher\banner.jpeg")
                    If Not FileExists(sFile) Then
                        sFile = JoinPath(Mods(iMod).ModPath, "launcher\banner.gif")
                        If Not FileExists(sFile) Then
                            sFile = JoinPath(RESDIR, "nobanner.bmp")
                            lblNoBanner(Index).Caption = DoubleAmpersand(Mods(iMod).ModName)
                        End If
                    End If
                End If
            End If
        End Select
        picMod(Index).Picture = LoadPicture(sFile)
        'display details
        picMod(Index).ToolTipText = DoubleAmpersand(Mods(iMod).ModName)
        lblModVersion(Index).Caption = DoubleAmpersand(Mods(iMod).ModVersion)
        lblModDate(Index).Caption = DoubleAmpersand(Mods(iMod).ModDate)
        lblModAuthor(Index).Caption = DoubleAmpersand(Mods(iMod).ModAuthor)
        lblModWebsite(Index).Caption = DoubleAmpersand(Mods(iMod).ModWebsite)
        lblModSize(Index).Caption = DataSize(Mods(iMod).ModSize)
        If BooleanStringToBoolean(ReadINIStr("General", "ModIsActive", JoinPath(Mods(iMod).ModPath, "launcher\userdata.lbu"))) Then
            If ReadINIStr("Mod", "Name", ProgramINI) = Mods(iMod).ModName Then
                If ReadINIStr("Mod", "Version", ProgramINI) = Mods(iMod).ModVersion Then
                    lblModSize(Index).Caption = lblModSize(Index).Caption & " (+" & DataSize(Val(ReadINIStr("Mod", "DiskUsage", ProgramINI, "0")))
                End If
            End If
        End If
        lblModDescription(Index).Caption = DoubleAmpersand(Mods(iMod).ModDescription)
        cmdManual(Index).Tag = Mods(iMod).ModManual
        cmdLaunch(Index).Enabled = True
        'Index specific controls:
        Select Case Index
        Case TypeMod
            lblModCampaigns.Caption = DoubleAmpersand(Mods(iMod).ModCampaigns)
            'display corect launch button image
            If Mods(iMod).ModIsForRA2 Then
                If cmdLaunch(Index).Tag <> "RA2" Then
                    Set cmdLaunch(Index).Picture = Nothing
                    Select Case FileExists(JoinPath(SelectedSkinPath, "btnb" & CStr(Index) & "r.bmp"))
                    Case True: cmdLaunch(Index).Picture = LoadPicture(JoinPath(SelectedSkinPath, "btnb" & CStr(Index) & "r.bmp"))
                    Case False: cmdLaunch(Index).Picture = LoadPicture(JoinPath(RESDIR, "btnb" & CStr(Index) & "r.bmp"))
                    End Select
                    cmdLaunch(Index).Tag = "RA2"
                End If
            Else
                If cmdLaunch(Index).Tag <> "YR" Then
                    Set cmdLaunch(Index).Picture = Nothing
                    Select Case FileExists(JoinPath(SelectedSkinPath, "btnb" & CStr(Index) & "y.bmp"))
                    Case True: cmdLaunch(Index).Picture = LoadPicture(JoinPath(SelectedSkinPath, "btnb" & CStr(Index) & "y.bmp"))
                    Case False: cmdLaunch(Index).Picture = LoadPicture(JoinPath(RESDIR, "btnb" & CStr(Index) & "y.bmp"))
                    End Select
                    cmdLaunch(Index).Tag = "YR"
                End If
            End If
            'Terrain Expansion prerequisite check
            If Len(Mods(iMod).ModTXVersion) <> 0 Then
                lblModTX(Index).Caption = "Version " & DoubleAmpersand(Mods(iMod).ModTXVersion) & " Required"
            End If
            If Mods(iMod).ModAllowTX Then
                sFile = Mods(iMod).ModTXVersion
            Else
                sFile = "-1"
            End If
            If PrerequisiteCheckTX(sFile) Then
                lblModTX(Index).ForeColor = ColorGood
                lblModTX(Index).Tag = "GOOD"
                If Len(Mods(iMod).ModTXVersion) = 0 Then lblModTX(Index).Caption = "Not Required"
            Else
                lblModTX(Index).ForeColor = ColorBad
                lblModTX(Index).Tag = "BAD"
                cmdLaunch(Index).Enabled = False
                If Not Mods(iMod).ModAllowTX Then lblModTX(Index).Caption = "Not Allowed"
            End If
            'Uses Ares
            If Mods(iMod).ModUseAres Then
                lblModUsesAres.Caption = "Yes"
            ElseIf FileExists(JoinPath(Mods(iMod).ModPath, "syringe\ares.dll")) Or FileExists(JoinPath(Mods(iMod).ModPath, "syringe\ares.dll.inj")) Then
                lblModUsesAres.Caption = "Custom Ares DLL"
            ElseIf DirExists(JoinPath(Mods(iMod).ModPath, "syringe")) Then
                lblModUsesAres.Caption = "Non-Ares DLL(s)"
            Else
                lblModUsesAres.Caption = "No"
            End If
        Case TypeFA2Mod
            'FinalAlert 2 version check
            lblModFA2.ForeColor = ColorGood
            lblModFA2.Tag = "GOOD"
            Select Case Mods(iMod).ModFA2Version
            Case ""
                lblModFA2.Caption = "Any Version Required"
                If Len(Mods(FA2ModNum).ModVersion) = 0 Then
                    lblModFA2.ForeColor = ColorBad
                    lblModFA2.Tag = "BAD"
                    cmdLaunch(Index).Enabled = False
                End If
            Case Else
                lblModFA2.Caption = "Version " & Mods(iMod).ModFA2Version & " Required"
                If Mods(FA2ModNum).ModVersion <> Mods(iMod).ModFA2Version Then
                    lblModFA2.ForeColor = ColorBad
                    lblModFA2.Tag = "BAD"
                    cmdLaunch(Index).Enabled = False
                End If
            End Select
            'Terrain Expansion prerequisite check
            If Len(Mods(iMod).ModTXVersion) <> 0 Then lblModTX(Index).Caption = "Version " & DoubleAmpersand(Mods(iMod).ModTXVersion) & " Required"
            If PrerequisiteCheckTX(IIf(Mods(iMod).ModAllowTX, Mods(iMod).ModTXVersion, "-1")) Then
                lblModTX(Index).ForeColor = ColorGood
                lblModTX(Index).Tag = "GOOD"
                If Len(Mods(iMod).ModTXVersion) = 0 Then lblModTX(Index).Caption = "Not Required"
            Else
                lblModTX(Index).ForeColor = ColorBad
                lblModTX(Index).Tag = "BAD"
                cmdLaunch(Index).Enabled = False
                If Not Mods(iMod).ModAllowTX Then lblModTX(Index).Caption = "Not Allowed"
            End If
            'now set up the TX integration checkbox
            cboxTX.Value = 0
            cboxTX.Visible = False
            lblTX.Visible = False
            If Mods(iMod).ModAllowTX Then
                If Len(Mods(iMod).ModTXVersion) = 0 Then
                    If GetActivePlugin("TX") <> -1 Then
                        cboxTX.Visible = True
                        lblTX.Visible = True
                    'Else
                        'do we need to indicate that TX integration is available?
                    End If
                Else
                    cboxTX.Value = 1
                End If
            End If
        Case TypeProgram
            lblModParams.Visible = Mods(iMod).ModShowParams
            txtModParams.Visible = Mods(iMod).ModShowParams
            txtModParams.Text = Mods(iMod).ModParams
            If Len(Mods(iMod).ModProgram) = 0 Then cmdLaunch(Index).Enabled = False
        End Select
        If (PlayModSound1 = True) And (OptModSound1) And (SelectedTab = Index) Then Call PlaySound(Mods(iMod).ModSound1, True)
    End If
    lblModWebsite(Index).Visible = Len(lblModWebsite(Index).Caption) <> 0
    cmdManual(Index).Enabled = Len(cmdManual(Index).Tag) <> 0
    Call CallStackPop
End Sub

Public Function PrerequisiteCheckTX(ByVal sRequired As String, Optional ByVal bAutoInstall As Boolean = False) As Boolean
    Dim bOk As Boolean
    Dim bPoss As Boolean
    Dim sActive As String
    Dim iPlugin As Integer
    Dim iCounter As Integer
    Call CallStackPush(Me.Name & ".PrerequisiteCheckTX()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    bOk = False
    sActive = ReadINIStr("PluginTX", "Version", ProgramINI)
    Select Case sRequired
    Case "-1"
        'Mod does not allow the TX
        If Len(sActive) = 0 Then
            bOk = True
        Else
            bPoss = OptAutoTX
            iPlugin = -1 'remember that we want to deactivate the plugin
        End If
    Case ""
        'Mod doesn't care
        bOk = True
        If bAutoInstall And OptAutoTX Then
            'we should activate anyway
            iPlugin = GetLatestPlugin("TX")
            If iPlugin <> -1 Then
                sRequired = Plugins(iPlugin).PluginVersion
                If CompareVersions(sActive, "<>", sRequired) Then
                    bOk = False
                    bPoss = True
                End If
            End If
        End If
    Case Else
        If Len(sActive) <> 0 Then
            If CompareVersions(sActive, "=", sRequired) Then bOk = True
        End If
        'now check autoTX
        If Not bOk Then
            'active TX version is not correct
            If OptAutoTX Then
                'check that we have the needed version
                iPlugin = GetLatestPlugin("TX")
                If iPlugin <> -1 Then
                    bPoss = CompareVersions(Plugins(iPlugin).PluginVersion, ">=", sRequired)
                Else
                    bPoss = False
                End If
            End If
        End If
    End Select
    If Not bOk Then
        If bPoss Then
            bOk = True
            If bAutoInstall Then
                If iPlugin <> -1 Then
                    bOk = ActivatePlugin(iPlugin) 'set bOk because plugin could fail the security check
                Else
                    Call DeactivatePlugin("TX")
                End If
            End If
        End If
    End If
    PrerequisiteCheckTX = bOk
    Call CallStackPop
End Function

Public Function AuthenticatePlugin(ByVal iPlugin As Integer) As Boolean
    Dim bOk As Boolean
    Dim AuthKeyFile As String
    Dim AuthKeyExt As String
    Dim AuthKeyInt As String
    Dim AuthKeyLen As Long
    Dim Counter As Long
    Dim FileHandle As Integer
    bOk = False
    If CL_dev Or Not OptVerifyPlugins Then
        Call WriteLogEntry("Skipping plugin authentication.", LogMsgBox)
        bOk = True
    Else
        Call WriteLogEntry("Authenticating plugin: " & Plugins(iPlugin).PluginName & " [" & Plugins(iPlugin).PluginVersion & "]", LogLevel1)
        Call ShowPleaseWait("Authenticating plugin: " & Plugins(iPlugin).PluginName & " [" & Plugins(iPlugin).PluginVersion & "]")
        AuthKeyFile = JoinPath(Plugins(iPlugin).PluginPath, "launcher\security.key")
        Call UpdatePleaseWait(, GetFileName(AuthKeyFile))
        If FileExists(AuthKeyFile) Then
            FileHandle = FreeFile
            Open AuthKeyFile For Input As #FileHandle
                Line Input #FileHandle, AuthKeyInt
            Close #FileHandle
            AuthKeyExt = ""
            AuthKeyLen = Len(AuthKeyInt)
            If AuthKeyLen Mod 2 = 0 Then
                For Counter = 1 To (AuthKeyLen \ 2)
                    AuthKeyExt = AuthKeyExt & Mid$(AuthKeyInt, Counter * 2, 1)
                Next Counter
                AuthKeyExt = StrReverse(Decrypt(AuthKeyExt))
                AuthKeyInt = AuthenticatePlugin_GetArchiveMD5(Plugins(iPlugin).PluginPath)
                AuthKeyInt = UCase$(AuthKeyInt)
                AuthKeyExt = UCase$(AuthKeyExt)
                If AuthKeyInt = AuthKeyExt Then bOk = True
                Call WriteLogEntry(UCase$(AuthKeyInt))
                Call WriteLogEntry(UCase$(AuthKeyExt))
            End If
        End If
        Call HidePleaseWait
        If bOk Then
            Call WriteLogEntry("Plugin authenticated.", LogLevel1)
        Else
            Call WriteLogEntry("Plugin failed to authenticate. Activation aborted.")
            Call MsgBox("Plugin could not be activated because it failed to authenticate!" & vbCrLf & "You should verify that the plugin is from a trusted source.", vbOKOnly + vbExclamation, App.Title)
        End If
    End If
    AuthenticatePlugin = bOk
End Function

Private Function AuthenticatePlugin_GetArchiveMD5(ByVal ArchiveRootPath As String) As String
    Dim RestoreDirPath As String
    Dim RunningMD5 As String
    Dim TempMD5 As String
    Dim bOk As Boolean
    Dim fso As FileSystemObject
    Dim fso_root As Folder
    Dim fso_folder As Folder
    Dim fso_file As File
    bOk = True
    RunningMD5 = ""
    Set fso = New FileSystemObject
    Set fso_root = fso.GetFolder(ArchiveRootPath)
    For Each fso_folder In fso_root.SubFolders 'we don't care about the root
        Select Case UCase$(fso_folder.Name)
        Case "CAMEO", "FA2FILES", "HVA", "INI", "INTERFACE", "MANUAL", "MAP", "MIX", "SCREEN", "SHP", "SIDE 1", "SIDE 2", "SIDE 3", "SIDE 4", "SOUND", "SPEECH", "STRING TABLE", "SYRINGE", "TAUNTS", "THEME", "VIDEO", "VXL"
            For Each fso_file In fso_folder.Files
                Call UpdatePleaseWait(, fso_file.Name)
                TempMD5 = GetFileMD5(fso_file.Path)
                If Len(TempMD5) = 0 Then
                    Call WriteLogEntry("Failed to get checksum of file: " & fso_file.Path)
                    bOk = False
                    Exit For
                Else
                    RunningMD5 = RunningMD5 & TempMD5
                    Call WriteLogEntry(TempMD5 & " " & fso_file.Path, LogLevel2)
                End If
            Next
            If Not bOk Then Exit For
        Case "LAUNCHER"
            For Each fso_file In fso_folder.Files
                Select Case UCase$(fso_file.Name)
                Case "LIBLIST.GAM", "BANNER.BMP", "BANNER.GIF", "BANNER.JPEG", "BANNER.JPG"
                    Call UpdatePleaseWait(, fso_file.Name)
                    TempMD5 = GetFileMD5(fso_file.Path)
                    If Len(TempMD5) = 0 Then
                        Call WriteLogEntry("Failed to get checksum of file: " & fso_file.Path)
                        bOk = False
                        Exit For
                    Else
                        RunningMD5 = RunningMD5 & TempMD5
                        Call WriteLogEntry(TempMD5 & " " & fso_file.Path, LogLevel2)
                    End If
                End Select
            Next
        End Select
    Next
    If Not bOk Then RunningMD5 = ""
    AuthenticatePlugin_GetArchiveMD5 = RunningMD5
End Function

Private Function Decrypt(ByVal MessageE As String, Optional ByVal Encrypt As Boolean = False) As String
    Dim AlphabetD As String
    Dim MessageD As String
    Dim AlphabetE As String
    Dim Counter As Long
    Dim Counter2 As Integer
    Dim Ok As Boolean
    Select Case Encrypt
    Case False
        AlphabetD = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890"
        AlphabetE = "VamBO0jIzfU9W1GvXqNeYkCgTJb4PriSlyHp5MRwcEA8soQ3dFZnhK2tx7Lu6D"
    Case True
        AlphabetE = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890"
        AlphabetD = "VamBO0jIzfU9W1GvXqNeYkCgTJb4PriSlyHp5MRwcEA8soQ3dFZnhK2tx7Lu6D"
    End Select
    For Counter = 1 To Len(MessageE)
        Ok = False
        For Counter2 = 1 To Len(AlphabetE)
            If Mid$(MessageE, Counter, 1) = Mid$(AlphabetE, Counter2, 1) Then
                MessageD = MessageD & Mid$(AlphabetD, Counter2, 1)
                Ok = True
                Counter2 = Len(AlphabetE)
            End If
        Next Counter2
        If Not Ok Then MessageD = MessageD & Mid$(MessageE, Counter, 1)
    Next Counter
    Decrypt = MessageD
End Function

Public Sub WriteLogEntry(Optional ByVal Entry As String = "", Optional ByVal Flags As MxLogTypeConstant = 0)
'Special flags:
'LogIE - adds "Internal Error!" to the message and talks about what to do. Also forces LogShutdown.
'LogShutdown - Tells user we need to close. Displays a message box using Critical icon.
'LogMsgBox - will display log entry in a message box. Uses Information icon.
'LogMsgBoxExclaim - will display log entry in a message box. Uses Exclamation icon.
'other flags determine whether or not to display it at all - LogIE ignores such flags
    Dim ThisLogIE As Boolean
    Dim ThisLogShutdown As Boolean
    Dim bLog As Boolean
    Dim hFile As Integer
    Dim TimeStamp As String
    Dim OneLineEntry As String
    Dim EntryArray() As String
    'establish mode
    bLog = False
    ThisLogIE = (Flags And LogIE) = LogIE
    If Not ThisLogIE Then
        ThisLogShutdown = TestFlags(Flags, LogShutdown)
        If OptLogLevel >= 1 Or Not TestFlags(Flags, LogLevel1) Then
            If OptLogLevel = 2 Or Not TestFlags(Flags, LogLevel2) Then
                bLog = True
            End If
        End If
    Else
        bLog = True
        ThisLogShutdown = True
    End If
    'the actual message
    If ThisLogIE Then Entry = "Internal Error! " & Entry
    If bLog Then
        If Len(Entry) <> 0 Or ThisLogIE Then TimeStamp = CStr(Year(Now())) & "-" & PadNum(Month(Now()), 2) & "-" & PadNum(Day(Now()), 2) & " " & PadNum(Hour(Now()), 2) & ":" & PadNum(Minute(Now()), 2) & ":" & PadNum(Second(Now()), 2) & "  " Else TimeStamp = ""
        If ThisLogIE Then
            OneLineEntry = Replace(Entry, vbCrLf, vbCrLf & Space$(Len(TimeStamp)))
        Else
            OneLineEntry = Replace(Entry, vbCrLf, " ")
        End If
        If OptLogFile Then
            hFile = FreeFile()
            Open LOGFILE For Append As #hFile
            Print #hFile, TimeStamp & OneLineEntry
            Close hFile
        End If
        If Len(frmLiveLog.txtLiveLog.Text) + Len(TimeStamp) + Len(OneLineEntry) >= 65533 Then
            frmLiveLog.txtLiveLog.Text = Right$(frmLiveLog.txtLiveLog.Text, 65533 - (Len(TimeStamp) + Len(OneLineEntry))) & TimeStamp & OneLineEntry & vbCrLf
        Else
            frmLiveLog.txtLiveLog.Text = frmLiveLog.txtLiveLog.Text & TimeStamp & OneLineEntry & vbCrLf
        End If
        If menu_livelog.Checked = True Then
            Call frmLiveLog.Refresh
        End If
    End If
    If ThisLogShutdown Then
        If ThisLogIE Then
            EntryArray = Split(Entry, vbCrLf)
            Call MsgBox(App.Title & " has encountered a problem and needs to close." & vbCrLf & vbCrLf & EntryArray(0) & vbCrLf & vbCrLf & "Please contact Marshall immediately with the following:" & vbCrLf & "Detailed information about what you were doing at the time." & vbCrLf & "Instructions on how to replicate the problem if you can." & vbCrLf & "The <LaunchBase.log> file if you have one.", vbOKOnly + vbCritical, "Internal Error")
        Else
            Call MsgBox(App.Title & " has encountered a problem and needs to close." & vbCrLf & vbCrLf & Entry, vbOKOnly + vbCritical, App.Title)
        End If
        Call Shutdown
        End
    Else
        If TestFlags(Flags, LogMsgBoxExclaim) Then
            Call MsgBox(Entry, vbOKOnly + vbExclamation, App.Title)
        ElseIf TestFlags(Flags, LogMsgBox) Then
            Call MsgBox(Entry, vbOKOnly + vbInformation, App.Title)
        End If
    End If
End Sub

Private Sub ActivateMod_ActivateFile(ByVal sFile As String, ByVal sSource As String, ByVal iMod As Integer, ByRef iFileCount As Long, ByRef iModSize As Double, Optional ByVal bMoveNotCopy As Boolean = False, Optional ByVal bFA2Mode As Boolean = False, Optional ByVal iRestore As Long = -1)
    Dim sDest As String
    Call CallStackPush(Me.Name & ".ActivateMod_ActivateFile(" & CStr(sFile) & ", " & CStr(sSource) & ", " & CStr(iMod) & ", " & CStr(iFileCount) & ", " & CStr(iModSize) & ", " & CStr(bMoveNotCopy) & ", " & CStr(bFA2Mode) & ", " & CStr(iRestore) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    sDest = JoinPath(IIf(Not bFA2Mode, RA2DIR, FA2DIR), sFile)
    If FileExists(sDest) Then
        If Not bFA2Mode Then
            If SafeFiles_Find(sFile) = 0 Then
                Call WriteLogEntry("Unexpected file found: " & Quote(sDest))
                Call MsgBox("An unexpected file has been found in your Red Alert 2 folder." & vbCrLf & Quote(sDest) & " will be deleted to make way for a replacement file." & vbCrLf & "If you believe that Launch Base should be anticipating this file then please report it." & vbCrLf & vbCrLf & "Please make a backup of this file now if it is important." & vbCrLf & "When you click OK the file will be deleted.", vbOKOnly + vbExclamation, App.Title)
                If FileExists(sDest) Then
                    Call LoggedKill(sDest)
                Else
                    Call WriteLogEntry(Quote(sDest) & " has been removed by the user.")
                End If
            Else
                Call WriteLogEntry("Mod is trying to replace a residual mod file that has been marked as safe [" & sDest & "]. Marking file as unsafe...")
                Call MsgBox("A residual mod file that you have marked as safe [" & sFile & "] will be replaced by this mod. Clearly the file is not in fact safe." & vbCrLf & "The file will now be marked as unsafe and treated as any other residual file." & vbCrLf & "It is strongly recommended that you re-read the Help Topics and consider disabling Advanced Mode.", vbOKOnly + vbExclamation, App.Title)
                Call SafeFiles_Find(sFile, True)
                Call ActivateMod_RemoveResidualFile(sFile, iRestore, True)
            End If
        Else
            'special case residual files - not handled by main routine
            Call ActivateMod_RemoveResidualFile(sFile, iRestore, True)
        End If
    End If
    If bMoveNotCopy Then
        Call LoggedMove(sSource, sDest)
    Else
        Call LoggedCopy(sSource, sDest)
    End If
    iModSize = iModSize + GetFileSize(sDest, True)
    Call WriteINIStr("Mod", CStr(iFileCount), sFile, ProgramINI)
    If OptUseCheckSums Then Call WriteINIStr("Mod", CStr(iFileCount) & "c", GetFileMD5(sDest), ProgramINI)
    iFileCount = iFileCount + 1
    Call CallStackPop
End Sub

Private Sub ShowPleaseWait(ByVal Message As String, Optional ByVal Message2 As String = "")
    Me.Enabled = False
    Call frmPleaseWait.Show
    Call UpdatePleaseWait(Message, Message2)
End Sub

Private Sub UpdatePleaseWait(Optional ByVal Message As String, Optional ByVal Message2 As String)
    If Message <> "" Then frmPleaseWait.Label1.Caption = Message
    frmPleaseWait.Label2.Caption = Message2
    Call frmPleaseWait.Refresh
End Sub

Private Sub HidePleaseWait()
    Unload frmPleaseWait
    Me.Enabled = True
End Sub

Public Sub LoggedMove(ByVal sSource As String, ByVal sDest As String, Optional ByVal IsBackup As Boolean = False, Optional ByVal IsRestore As Boolean = False)
    If Not IsBackup Then
        If Not IsRestore Then
            Call WriteLogEntry("Moving " & Quote(sSource) & " to " & Quote(sDest))
        Else
            Call WriteLogEntry("Restoring " & Quote(sSource) & " to " & Quote(sDest))
        End If
    Else
        Call WriteLogEntry("Backing up (moving) " & Quote(sSource) & " to " & Quote(sDest))
    End If
    Name sSource As sDest
End Sub

Private Sub LoggedCopy(ByVal sSource As String, ByVal sDest As String)
    Call WriteLogEntry("Copying " & Quote(sSource) & " to " & Quote(sDest))
    Call FileCopy(sSource, sDest)
End Sub

Private Sub LoggedKill(ByVal sPath As String)
    Call WriteLogEntry("Deleting " & Quote(sPath))
    Call Kill(sPath)
End Sub

Private Sub LoggedKillDir(ByVal sPath As String)
    Call WriteLogEntry("Deleting " & Quote(sPath))
    Call KillDir(sPath)
End Sub

Private Sub LoggedMkDir(ByVal sPath As String)
    Call WriteLogEntry("Creating " & Quote(sPath))
    Call MkDir(sPath)
End Sub

Private Sub LoggedMakePath(ByVal sPath As String)
    Call WriteLogEntry("Creating " & Quote(sPath))
    Call MakePath(sPath)
End Sub

Private Function GetNameVersion(ByVal sName As String, ByVal sVersion As String) As String
    If Len(sVersion) <> 0 Then
        GetNameVersion = sName & " [" & sVersion & "]"
    Else
        GetNameVersion = sName
    End If
End Function

Private Function ActivateMod(ByVal iMod As Integer) As Boolean
    Dim sKiln As String
    Dim iInstCount As Long
    Dim sFolder As String
    Dim sFile As String
    Dim sSource As String
    Dim sTemp As String
    Dim iModSize As Double
    Dim iCounter As Long
    Dim iRestore As Long
    Dim fso As FileSystemObject
    Dim fso_root As Folder
    Dim fso_folder As Folder
    Dim fso_file As File
    Dim hExpand As Integer
    Dim sExpand As String
    Dim hSide1 As Integer
    Dim sSide1 As String
    Dim hSide2 As Integer
    Dim sSide2 As String
    Dim hSide3 As Integer
    Dim sSide3 As String
    Dim hSide4 As Integer
    Dim sSide4 As String
    Dim hBAG As Integer
    Dim sCSF As String
    Dim hTXT As Integer
    Dim sLang As String
    Dim sAudio As String
    Dim bReject As Boolean
    Dim bRecompileExpand As Boolean
    Dim bRecompileBAG As Boolean
    Dim bRecompileCSF As Boolean
    Dim bAresOk As Boolean
    Dim csfcsf As MarshallxCSFClass
    Call CallStackPush(Me.Name & ".ActivateMod(" & CStr(iMod) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    ActivateMod = False
    bAresOk = True
    'remove residual files
    If ActivateMod_RemoveResidualFiles(iRestore) Then
        'Check if mod is already active
        If ActivateMod_ReactivateNeeded(iMod) Then
            'Show please wait dialog
            Call ShowPleaseWait("Activating" & IIf(Mods(iMod).ModType = TypeMod, " mod: ", " FA2 mod: ") & GetNameVersion(Mods(iMod).ModName, Mods(iMod).ModVersion), "")
            'Setup paths/filenames
            sKiln = JoinPath(Mods(iMod).ModPath, "Kiln")
            If Mods(iMod).ModIsForRA2 Or Mods(iMod).ModType = TypeFA2Mod Then
                sExpand = "expand98.mix"
                sCSF = "ra2.csf"
                sLang = "language.mix"
                sAudio = "audio.mix"
            Else
                sExpand = "expandmd98.mix"
                sCSF = "ra2md.csf"
                sLang = "langmd.mix"
                sAudio = "audiomd.mix"
            End If
            sSide1 = "sidec01.mix"
            sSide2 = "sidec02.mix"
            If Mods(iMod).ModUseYuriUI Then
                sSide3 = "sidec03.mix"
            Else
                sSide3 = "sidec02md.mix"
            End If
            sSide4 = "sidec04.mix"
            'Clear compiled files
            bRecompileExpand = True
            bRecompileCSF = True
            bRecompileBAG = True
            If OptRecompile Or CL_dev Then
                If DirExists(sKiln) Then Call LoggedKillDir(sKiln)
                Call LoggedMakePath(sKiln)
            Else
                If Not DirExists(sKiln) Then
                    Call LoggedMakePath(sKiln)
                Else
                    bRecompileExpand = Not FileExists(JoinPath(sKiln, sExpand))
                    bRecompileCSF = Not FileExists(JoinPath(sKiln, sCSF))
                    If FileExists(JoinPath(sKiln, "audio.bag")) Then
                        If FileExists(JoinPath(sKiln, "audio.idx")) Then
                            bRecompileBAG = False
                        Else
                            bRecompileBAG = True
                            Call LoggedKill(JoinPath(sKiln, "audio.bag"))
                        End If
                    Else
                        bRecompileBAG = True
                        If FileExists(JoinPath(sKiln, "audio.idx")) Then Call LoggedKill(JoinPath(sKiln, "audio.idx"))
                    End If
                End If
            End If
            'deactivate any active mod
            Call DeactivateMod(False, True)
            'Check for free disk space
            If ActivateMod_CheckDiskSpace(iMod, bRecompileBAG, bRecompileCSF, bRecompileExpand, sKiln) Then
                'log start of activation
                Call WriteLogEntry("Activating " & IIf(Mods(iMod).ModType = TypeMod, "mod: ", "FA2 mod: ") & GetNameVersion(Mods(iMod).ModName, Mods(iMod).ModVersion), LogLevel1)
                'create record of activation
                Call WriteINIStr("Mod", "Name", Mods(iMod).ModName, ProgramINI)
                Call WriteINIStr("Mod", "ModType", CStr(Mods(iMod).ModType), ProgramINI)
                Call WriteINIStr("Mod", "Version", Mods(iMod).ModVersion, ProgramINI)
                Call WriteINIStr("Mod", "ScrnFormat", Mods(iMod).ModScrnFormat, ProgramINI)
                Call WriteINIStr("Mod", "SnapFormat", Mods(iMod).ModSnapFormat, ProgramINI)
                Call WriteINIStr("Mod", "GameIsRA2", BooleanToYesNo(Mods(iMod).ModIsForRA2), ProgramINI)
                Select Case iMod
                Case YRModNum, RA2ModNum, FA2ModNum: Call WriteINIStr("Mod", "Official", "yes", ProgramINI)
                Case Else: Call WriteINIStr("Mod", "Official", "no", ProgramINI)
                End Select
                Call WriteINIStr("General", "ModIsActive", "yes", JoinPath(Mods(iMod).ModPath, "launcher\userdata.lbu"))
                'Now install the files
                iModSize = 0
                iInstCount = 0
                Set fso = New FileSystemObject
                Set fso_root = fso.GetFolder(Mods(iMod).ModPath)
                For Each fso_folder In fso_root.SubFolders
                    sFolder = UCase$(fso_folder.Name)
                    Select Case sFolder
                    Case "SCREEN", "THEME", "VIDEO", "CAMEO", "SHP", "SPEECH", "TMP", "HVA", "INI", "INTERFACE", "MAP", "MIX", "VXL", "SIDE1", "SIDE 1", "SIDE_1", "SIDE2", "SIDE 2", "SIDE_2", "SIDE3", "SIDE 3", "SIDE_3", "SIDE4", "SIDE 4", "SIDE_4", "TAUNTS", "STRING TABLE", "STRINGTABLE", "STRING_TABLE", "SOUND", "SOUNDS", "SYRINGE"
                        For Each fso_file In fso_folder.Files
                            'prepare the file
                            sFile = fso_file.Name
                            sSource = fso_file.Path
                            Call UpdatePleaseWait(, sFile)
                            Select Case FileType(sFile)
                            Case "OGG", "FLAC"
                                Select Case sFolder
                                Case "THEME", "VIDEO", "SPEECH", "SOUND", "SOUNDS", "TAUNTS"
                                    sFile = ChangeFileType(sFile, "WAV")
                                    sTemp = JoinPath(sKiln, sFile)
                                    If Not FileExists(sTemp) Then
                                        Select Case FileType(sFile)
                                        Case "OGG": Call ConvertOggToWav(sSource, sTemp)
                                        Case "FLAC": Call ConvertFlacToWav(sSource, sTemp)
                                        End Select
                                        If Not FileExists(sTemp) Then GoTo NextModFile
                                    End If
                                    sSource = sTemp
                                End Select
                            End Select
                            'validate and process file
                            bReject = True 'assume the file will fail validation
                            If Not FileIsAresComponent(sFile) Then
                            Select Case sFolder
                            Case "SCREEN", "THEME", "VIDEO" 'LOOSE FILES
                                If Not FileIsDestructive(sFile) Then
                                    If FileIsModfile(sFile) Then
                                        If Not FileIsUserTheme(sFile) Then
                                            If Not FileIsSoundtrack(sFile) Then
                                                If Not FileIsOfficialMapPackMap(sFile) Then
                                                    If Not FileIsReservedMix(sFile) Then
                                                        Call ActivateMod_ActivateFile(sFile, sSource, iMod, iInstCount, iModSize)
                                                        bReject = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Case "CAMEO", "HVA", "INI", "INTERFACE", "MAP", "MIX", "SHP", "SPEECH", "TMP", "VXL" 'EXPANDMD98.MIX
                                If bRecompileExpand Then
                                    If DCoderDLL And Not OptLooseFileMode Then
                                        If hExpand = 0 Then hExpand = DCoderMIXOpen(JoinPath(sKiln, hExpand))
                                        Call DCoderMIXInsert(hExpand, sSource, JoinPath(sKiln, sExpand))
                                    Else
                                        If LCase$(sFile) <> "thememd.ini" Then 'thememd.ini is not allowed to be loose because of YR Playlist Manager.
                                            Call ActivateMod_ActivateFile(sFile, sSource, iMod, iInstCount, iModSize)
                                        End If
                                    End If
                                End If
                                bReject = False
                            Case "SIDE1", "SIDE 1", "SIDE_1"
                                If bRecompileExpand Then
                                    If hSide1 = 0 Then
                                        'haven't set up the base file yet
                                        If FileExists(JoinPath(Mods(iMod).ModPath, sFolder, sSide1)) Then
                                            'use that as the base
                                            If DCoderDLL Then
                                                Call FileCopy(JoinPath(Mods(iMod).ModPath, sFolder, sSide1), JoinPath(sKiln, sSide1))
                                                Call WriteLogEntry(Quote(JoinPath(Mods(iMod).ModPath, sFolder, sSide1)) & " copied to " & Quote(JoinPath(sKiln, sSide1)) & ".")
                                                hSide1 = DCoderMIXOpen(JoinPath(sKiln, sSide1))
                                            Else
                                                Call WriteLogEntry("DCoder DLL is missing! Unable to open " & sSide1 & " for appending.")
                                            End If
                                        Else
                                            'extract it from ra2.mix
                                            If DCoderDLL Then
                                                hSide1 = DCoderMIXOpen(JoinPath(RA2DIR, "ra2.mix"))
                                                Call DCoderMIXExtract(hSide1, sSide1, JoinPath(RA2DIR, "ra2.mix"), JoinPath(sKiln, sSide1))
                                                Call DCoderMIXClose(hSide1)
                                                hSide1 = DCoderMIXOpen(JoinPath(sKiln, sSide1))
                                            Else
                                                Call WriteLogEntry("DCoder DLL is missing! Unable to extract " & sSide1 & " from " & JoinPath(RA2DIR, "ra2.mix"))
                                            End If
                                        End If
                                    End If
                                    If hSide1 <> 0 Then
                                        bReject = False
                                        If UCase$(sSide1) <> UCase$(sFile) Then
                                            Call DCoderMIXInsert(hSide1, sSource, JoinPath(sKiln, sSide1))
                                        End If
                                    Else
                                        If Not DCoderDLL Then
                                            If UCase$(sSide1) = UCase$(sFile) Then Call ActivateMod_ActivateFile(sFile, sSource, iMod, iInstCount, iModSize)
                                        Else
                                            'why have we failed to open the file?
                                            Call Panic
                                        End If
                                    End If
                                Else
                                    bReject = False
                                End If
                            Case "SIDE2", "SIDE 2", "SIDE_2"
                                If bRecompileExpand Then
                                    If hSide2 = 0 Then
                                        'haven't set up the base file yet
                                        If FileExists(JoinPath(Mods(iMod).ModPath, sFolder, sSide2)) Then
                                            'use that as the base
                                            If DCoderDLL Then
                                                Call FileCopy(JoinPath(Mods(iMod).ModPath, sFolder, sSide2), JoinPath(sKiln, sSide2))
                                                Call WriteLogEntry(Quote(JoinPath(Mods(iMod).ModPath, sFolder, sSide2)) & " copied to " & Quote(JoinPath(sKiln, sSide2)) & ".")
                                                hSide2 = DCoderMIXOpen(JoinPath(sKiln, sSide2))
                                            Else
                                                Call WriteLogEntry("DCoder DLL is missing! Unable to open " & sSide2 & " for appending.")
                                            End If
                                        Else
                                            'extract it from ra2.mix
                                            If DCoderDLL Then
                                                hSide2 = DCoderMIXOpen(JoinPath(RA2DIR, "ra2.mix"))
                                                Call DCoderMIXExtract(hSide2, sSide2, JoinPath(RA2DIR, "ra2.mix"), JoinPath(sKiln, sSide2))
                                                Call DCoderMIXClose(hSide2)
                                                hSide2 = DCoderMIXOpen(JoinPath(sKiln, sSide2))
                                            Else
                                                Call WriteLogEntry("DCoder DLL is missing! Unable to extract " & sSide2 & " from " & JoinPath(RA2DIR, "ra2.mix"))
                                            End If
                                        End If
                                    End If
                                    If hSide2 <> 0 Then
                                        bReject = False
                                        If UCase$(sSide2) <> UCase$(sFile) Then
                                            Call DCoderMIXInsert(hSide2, sSource, JoinPath(sKiln, sSide2))
                                        End If
                                    Else
                                        If Not DCoderDLL Then
                                            If UCase$(sSide2) = UCase$(sFile) Then Call ActivateMod_ActivateFile(sFile, sSource, iMod, iInstCount, iModSize)
                                        Else
                                            'why have we failed to open the file?
                                            Call Panic
                                        End If
                                    End If
                                Else
                                    bReject = False
                                End If
                            Case "SIDE3", "SIDE 3", "SIDE_3"
                                If bRecompileExpand Then
                                    If hSide3 = 0 Then
                                        'haven't set up the base file yet
                                        If FileExists(JoinPath(Mods(iMod).ModPath, sFolder, sSide3)) Then
                                            'use that as the base
                                            If DCoderDLL Then
                                                Call FileCopy(JoinPath(Mods(iMod).ModPath, sFolder, sSide3), JoinPath(sKiln, sSide3))
                                                Call WriteLogEntry(Quote(JoinPath(Mods(iMod).ModPath, sFolder, sSide3)) & " copied to " & Quote(JoinPath(sKiln, sSide3)) & ".")
                                                hSide3 = DCoderMIXOpen(JoinPath(sKiln, sSide3))
                                            Else
                                                Call WriteLogEntry("DCoder DLL is missing! Unable to open " & sSide3 & " for appending.")
                                            End If
                                        Else
                                            'extract it from ra2.mix
                                            If Mods(iMod).ModUseYuriUI Then
                                                If DCoderDLL Then
                                                    hSide3 = DCoderMIXOpen(JoinPath(RA2DIR, "ra2md.mix"))
                                                    Call DCoderMIXExtract(hSide3, sSide3, JoinPath(RA2DIR, "ra2md.mix"), JoinPath(sKiln, sSide3))
                                                    Call DCoderMIXClose(hSide3)
                                                    hSide3 = DCoderMIXOpen(JoinPath(sKiln, sSide3))
                                                Else
                                                    Call WriteLogEntry("DCoder DLL is missing! Unable to extract " & sSide3 & " from " & JoinPath(RA2DIR, "ra2md.mix"))
                                                End If
                                            Else
                                                If DCoderDLL Then
                                                    hSide3 = DCoderMIXOpen(JoinPath(RA2DIR, "ra2.mix"))
                                                    Call DCoderMIXExtract(hSide3, sSide3, JoinPath(RA2DIR, "ra2.mix"), JoinPath(sKiln, sSide3))
                                                    Call DCoderMIXClose(hSide3)
                                                    hSide3 = DCoderMIXOpen(JoinPath(sKiln, sSide3))
                                                Else
                                                    Call WriteLogEntry("DCoder DLL is missing! Unable to extract " & sSide3 & " from " & JoinPath(RA2DIR, "ra2.mix"))
                                                End If
                                            End If
                                        End If
                                    End If
                                    If hSide3 <> 0 Then
                                        bReject = False
                                        If UCase$(sSide3) <> UCase$(sFile) Then
                                            Call DCoderMIXInsert(hSide3, sSource, JoinPath(sKiln, sSide3))
                                        End If
                                    Else
                                        If Not DCoderDLL Then
                                            If UCase$(sSide3) = UCase$(sFile) Then Call ActivateMod_ActivateFile(sFile, sSource, iMod, iInstCount, iModSize)
                                        Else
                                            'why have we failed to open the file?
                                            Call Panic
                                        End If
                                    End If
                                Else
                                    bReject = False
                                End If
                            Case "SIDE4", "SIDE 4", "SIDE_4"
                                If bRecompileExpand Then
                                    If hSide4 = 0 Then
                                        'haven't set up the base file yet
                                        If FileExists(JoinPath(Mods(iMod).ModPath, sFolder, sSide4)) Then
                                            'use that as the base
                                            If DCoderDLL Then
                                                Call FileCopy(JoinPath(Mods(iMod).ModPath, sFolder, sSide4), JoinPath(sKiln, sSide4))
                                                Call WriteLogEntry(Quote(JoinPath(Mods(iMod).ModPath, sFolder, sSide4)) & " copied to " & Quote(JoinPath(sKiln, sSide4)) & ".")
                                                hSide4 = DCoderMIXOpen(JoinPath(sKiln, sSide4))
                                            Else
                                                Call WriteLogEntry("DCoder DLL is missing! Unable to open " & sSide4 & " for appending.")
                                            End If
                                        End If
                                    End If
                                    If hSide4 <> 0 Then
                                        bReject = False
                                        If UCase$(sSide4) <> UCase$(sFile) Then
                                            Call DCoderMIXInsert(hSide4, sSource, JoinPath(sKiln, sSide4))
                                        End If
                                    Else
                                        If Not DCoderDLL Then
                                            If UCase$(sSide4) = UCase$(sFile) Then Call ActivateMod_ActivateFile(sFile, sSource, iMod, iInstCount, iModSize)
                                        Else
                                            'why have we failed to open the file?
                                            Call Panic
                                        End If
                                    End If
                                Else
                                    bReject = False
                                End If
                            Case "TAUNTS" 'TAUNTS
                                If FileIsOfficialTaunt(sFile) Or FileIsRPTaunt(sFile) Then
                                    Call ActivateMod_ActivateFile("Taunts\" & sFile, sSource, iMod, iInstCount, iModSize)
                                    bReject = False
                                End If
                            Case "STRING TABLE", "STRINGTABLE", "STRING_TABLE"
                                If bRecompileCSF Then
                                    Select Case FileType(sFile)
                                    Case "CSF", "TXT"
                                        If csfcsf Is Nothing Then
                                            Call ActivateMod_PrepareCSF(iMod, sCSF, sKiln, sLang, csfcsf)
                                        End If
                                        If Not (csfcsf Is Nothing) Then
                                            If UCase$(sFile) <> UCase$(sCSF) Then
                                                Call csfcsf.UpdateWith(sSource)
                                                Call WriteLogEntry(Quote(sSource) & " merged into " & Quote(JoinPath(sKiln, sCSF)), LogLevel2)
                                            End If
                                            bReject = False
                                        End If
                                    End Select
                                Else
                                    bReject = False
                                End If
                            Case "SOUND", "SOUNDS" 'AUDIO.BAG/IDX
                                If bRecompileBAG Then
                                    If hBAG = 0 Then
                                        'haven't set up the base file yet
                                        If FileExists(JoinPath(Mods(iMod).ModPath, sFolder, "audio.bag")) And FileExists(FileExists(JoinPath(Mods(iMod).ModPath, sFolder, "audio.idx"))) Then
                                            'use that as the base
                                            If DCoderDLL Then
                                                Call FileCopy(JoinPath(Mods(iMod).ModPath, sFolder, "audio.bag"), JoinPath(sKiln, "audio.bag"))
                                                Call WriteLogEntry(Quote(JoinPath(Mods(iMod).ModPath, sFolder, "audio.idx")) & " copied to " & Quote(JoinPath(sKiln, "audio.idx")) & ".")
                                                Call FileCopy(JoinPath(Mods(iMod).ModPath, sFolder, "audio.idx"), JoinPath(sKiln, "audio.idx"))
                                                hBAG = DCoderBAGOpen(JoinPath(sKiln, "audio.bag"))
                                            Else
                                                Call WriteLogEntry("DCoder DLL is missing! Unable to open audio.bag/idx for appending.")
                                            End If
                                        Else
                                            'extract it from language.mix
                                            If DCoderDLL Then
                                                hBAG = DCoderMIXOpen(JoinPath(RA2DIR, sLang))
                                                Call DCoderMIXExtract(hBAG, sAudio & "/" & "audio.bag", JoinPath(sKiln, "audio.bag"), JoinPath(RA2DIR, sLang) & "/" & sAudio & "/" & "audio.bag")
                                                Call DCoderMIXExtract(hBAG, sAudio & "/" & "audio.idx", JoinPath(sKiln, "audio.idx"), JoinPath(RA2DIR, sLang) & "/" & sAudio & "/" & "audio.idx")
                                                Call DCoderMIXClose(hBAG)
                                                hBAG = DCoderBAGOpen(JoinPath(sKiln, "audio.bag"))
                                            Else
                                                Call WriteLogEntry("DCoder DLL is missing! Unable to extract audio.bag/idx from " & JoinPath(RA2DIR, sLang) & "/" & sAudio)
                                            End If
                                        End If
                                    End If
                                    If hBAG <> 0 Then
                                        Select Case UCase$(sFile)
                                        Case "AUDIO.BAG", "AUDIO.IDX"
                                            bReject = False
                                        Case Else
                                            Select Case FileType(sFile)
                                            Case "BAG"
                                                Call DCoderBAGMerge(hBAG, sSource, JoinPath(sKiln, "audio.bag"))
                                                bReject = False
                                            Case "WAV"
                                                Call DCoderWAVInsert(hBAG, sSource, JoinPath(sKiln, "audio.bag"))
                                                bReject = False
                                            End Select
                                        End Select
                                    Else
                                        If Not DCoderDLL Then
                                            Select Case UCase$(sFile)
                                            Case "AUDIO.BAG", "AUDIO.IDX"
                                                bReject = False
                                                Call ActivateMod_ActivateFile(sFile, sSource, iMod, iInstCount, iModSize)
                                            End Select
                                        Else
                                            'why have we failed to open the file?
                                            Call Panic
                                        End If
                                    End If
                                Else
                                    bReject = False
                                End If
                            Case "FA2FILES"
                                If Mods(iMod).ModType = TypeFA2Mod Then
                                    If FileIsFA2File(sFile) Then
                                        Call ActivateMod_ActivateFile(sFile, sSource, iMod, iInstCount, iModSize, , True, iRestore)
                                        bReject = False
                                    End If
                                End If
                            Case "SYRINGE"
                                If Not FileIsDestructive(sFile) Then
                                    Select Case FileType(sFile)
                                    Case "DLL"
                                        'activate both the dll and the inj
                                        If UCase$(sFile) <> "ARES.DLL" Or Not Mods(iMod).ModUseAres Then
                                            sTemp = sFile & ".inj"
                                            If Not FileIsDestructive(sTemp) Then
                                                If FileExists(JoinPath(GetFilePath(sSource), sTemp)) Then
                                                    Call ActivateMod_ActivateFile(sFile, sSource, iMod, iInstCount, iModSize)
                                                    Call ActivateMod_ActivateFile(sTemp, JoinPath(GetFilePath(sSource), sTemp), iMod, iInstCount, iModSize)
                                                    bReject = False
                                                    'remember that we are using Syringe
                                                    Call WriteINIStr("Mod", "Syringe", "yes", ProgramINI)
                                                End If
                                            End If
                                        End If
                                    Case "INJ"
                                        If UCase$(sFile) <> "ARES.DLL.INJ" Or Not Mods(iMod).ModUseAres Then
                                            'just test that the dll is there too
                                            If Len(sFile) >= 8 Then
                                                sTemp = Left$(sFile, Len(sFile) - 4)
                                                If FileType(sTemp) = "DLL" Then
                                                    If Not FileIsDestructive(sTemp) Then
                                                        If FileExists(JoinPath(GetFilePath(sSource), sTemp)) Then bReject = False
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End Select
                                End If
                            End Select
                            End If
                            If bReject Then Call WriteLogEntry(Quote(sSource) & " rejected for copying to " & Quote(JoinPath(RA2DIR, sFile)))
NextModFile:
                        Next
                    End Select
                Next
                'kiln - activate compiled files
                If hSide1 <> 0 Then
                    Call UpdatePleaseWait(, sSide1)
                    Call DCoderMIXWrite(hSide1, JoinPath(sKiln, sSide1))
                    Call DCoderMIXClose(hSide1)
                    If hExpand = 0 Then hExpand = DCoderMIXOpen(JoinPath(sKiln, sExpand))
                    Call DCoderMIXInsert(hExpand, JoinPath(sKiln, sSide1), JoinPath(sKiln, sExpand))
                    If OptRecompile Then Call LoggedKill(JoinPath(sKiln, sSide1))
                End If
                If hSide2 <> 0 Then
                    Call UpdatePleaseWait(, sSide2)
                    Call DCoderMIXWrite(hSide2, JoinPath(sKiln, sSide2))
                    Call DCoderMIXClose(hSide2)
                    If hExpand = 0 Then hExpand = DCoderMIXOpen(JoinPath(sKiln, sExpand))
                    Call DCoderMIXInsert(hExpand, JoinPath(sKiln, sSide2), JoinPath(sKiln, sExpand))
                    If OptRecompile Then Call LoggedKill(JoinPath(sKiln, sSide2))
                End If
                If hSide3 <> 0 Then
                    Call UpdatePleaseWait(, sSide3)
                    Call DCoderMIXWrite(hSide3, JoinPath(sKiln, sSide3))
                    Call DCoderMIXClose(hSide3)
                    If hExpand = 0 Then hExpand = DCoderMIXOpen(JoinPath(sKiln, sExpand))
                    Call DCoderMIXInsert(hExpand, JoinPath(sKiln, sSide3), JoinPath(sKiln, sExpand))
                    If OptRecompile Then Call LoggedKill(JoinPath(sKiln, sSide3))
                End If
                If hSide4 <> 0 Then
                    Call UpdatePleaseWait(, sSide4)
                    Call DCoderMIXWrite(hSide4, JoinPath(sKiln, sSide4))
                    Call DCoderMIXClose(hSide4)
                    If hExpand = 0 Then hExpand = DCoderMIXOpen(JoinPath(sKiln, sExpand))
                    Call DCoderMIXInsert(hExpand, JoinPath(sKiln, sSide4), JoinPath(sKiln, sExpand))
                    If OptRecompile Then Call LoggedKill(JoinPath(sKiln, sSide4))
                End If
                If hExpand <> 0 Then
                    Call UpdatePleaseWait(, sExpand)
                    Call DCoderMIXWrite(hExpand, JoinPath(sKiln, sExpand))
                    Call DCoderMIXClose(hExpand)
                End If
                If FileExists(JoinPath(sKiln, sExpand)) Then Call ActivateMod_ActivateFile(sExpand, JoinPath(sKiln, sExpand), iMod, iInstCount, iModSize, OptRecompile)
                If hBAG <> 0 Then
                    Call UpdatePleaseWait(, "audio.bag")
                    Call DCoderBAGWrite(hBAG, JoinPath(sKiln, "audio.bag"))
                    Call DCoderBAGClose(hBAG)
                    Call ActivateMod_ActivateFile("audio.bag", JoinPath(sKiln, "audio.bag"), iMod, iInstCount, iModSize, OptRecompile)
                    Call ActivateMod_ActivateFile("audio.idx", JoinPath(sKiln, "audio.idx"), iMod, iInstCount, iModSize, OptRecompile)
                End If
                If FileExists(JoinPath(sKiln, "audio.bag")) And FileExists(JoinPath(sKiln, "audio.idx")) Then
                    Call ActivateMod_ActivateFile("audio.bag", JoinPath(sKiln, "audio.bag"), iMod, iInstCount, iModSize, OptRecompile)
                    Call ActivateMod_ActivateFile("audio.idx", JoinPath(sKiln, "audio.idx"), iMod, iInstCount, iModSize, OptRecompile)
                End If
                If bRecompileCSF Then
                    If csfcsf Is Nothing Then
                        If Mods(iMod).ModType = TypeMod And Not Mods(iMod).ModIsForRA2 Then Call ActivateMod_PrepareCSF(iMod, sCSF, sKiln, sLang, csfcsf)
                    End If
                    If Not (csfcsf Is Nothing) Then
                        Call UpdatePleaseWait(, sCSF)
                        Call csfcsf.SaveCSF
                        Set csfcsf = Nothing
                    End If
                End If
                If FileExists(JoinPath(sKiln, sCSF)) Then Call ActivateMod_ActivateFile(sCSF, JoinPath(sKiln, sCSF), iMod, iInstCount, iModSize, OptRecompile)
                'Trash any WAV files converted from OGG
                Set fso_folder = fso.GetFolder(sKiln)
                For Each fso_file In fso_folder.Files
                    If FileType(fso_file.Name) = "WAV" Then Call LoggedKill(JoinPath(sKiln, fso_file.Name))
                Next
                'Now the TX FA2 mod integration
                If Mods(iMod).ModType = TypeFA2Mod And cboxTX.Value = 1 Then
                    Call WriteLogEntry("Integrating Terrain Expansion FA2 mod.", LogLevel1)
                    sSource = Plugins(GetLatestPlugin("TX")).PluginPath
                    Set fso_folder = fso.GetFolder(JoinPath(sSource, "fa2files"))
                    For Each fso_file In fso_folder.Files
                        sFile = fso_file.Name
                        Call UpdatePleaseWait(, sFile)
                        If FileIsFA2File(sFile) Then
                            Call ActivateMod_ActivateFile(sFile, JoinPath(sSource, "fa2files", sFile), iMod, iInstCount, iModSize, , True, iRestore)
                        Else
                            Call ActivateMod_ActivateFile(sFile, JoinPath(sSource, "fa2files", sFile), iMod, iInstCount, iModSize)
                        End If
                    Next
                    Call UpdatePleaseWait(, "expand06.mix")
                    Call ActivateMod_ActivateFile("expand06.mix", JoinPath(sSource, "video\expandmd06.mix"), iMod, iInstCount, iModSize)
                End If
                If Mods(iMod).ModType = TypeMod Then
                    'add save games and map snapshots
                    sSource = JoinPath(Mods(iMod).ModPath, "saves")
                    If DirExists(sSource) Then
                        Set fso_folder = fso.GetFolder(sSource)
                        For Each fso_file In fso_folder.Files
                            sFile = fso_file.Name
                            Call UpdatePleaseWait(, sFile)
                            If FileIsSaveGame(sFile) Or ConfirmScrnFormat(sFile, Mods(iMod).ModSnapFormat) Then
                                Call LoggedCopy(JoinPath(sSource, sFile), JoinPath(RA2DIR, sFile))
                            End If
                        Next
                    End If
                    'load game config
                    Call UpdatePleaseWait(, "Loading game configuration.")
                    Call ActivateMod_GameConfig(iMod, Mods(iMod).ModIsForRA2)
                    'add blank logo video if appropriate
                    If OptSkipLogo Then
                        If Not FileExists(JoinPath(RA2DIR, "ea_wwlogo.bik")) Then
                            Call WriteLogEntry("Adding fake logo video.", LogLevel2)
                            sFile = "ea_wwlogo.bik"
                            Call UpdatePleaseWait(, sFile)
                            sSource = JoinPath(RESDIR, sFile)
                            Call ActivateMod_ActivateFile(sFile, sSource, iMod, iInstCount, iModSize)
                        End If
                    End If
                    'add Ares.dll
                    If Mods(iMod).ModUseAres Then
                        bAresOk = False
                        If OptAutoAresUpdate Then
                            Call UpdatePleaseWait(, "Checking for updates to Ares.")
                            Call UpdateAres(Me)
                        End If
                        If FileExists(JoinPath(RESDIR, "Ares.dll")) And FileExists(JoinPath(RESDIR, "Ares.dll.inj")) Then
                            bAresOk = VerifyAres
                            If Not bAresOk Then
                                'Call DenyPersistentMods
                                If OptAutoAresUpdate Then
                                    Call UpdateAres(Me)
                                    If FileExists(JoinPath(RESDIR, "Ares.dll")) And FileExists(JoinPath(RESDIR, "Ares.dll.inj")) Then
                                        bAresOk = True
                                    End If
                                End If
                            End If
                        End If
                        If bAresOk Then
                            Call UpdatePleaseWait(, "Activating Ares.")
                            Call ActivateMod_ActivateFile("Ares.dll", JoinPath(RESDIR, "Ares.dll"), iMod, iInstCount, iModSize)
                            Call ActivateMod_ActivateFile("Ares.dll.inj", JoinPath(RESDIR, "Ares.dll.inj"), iMod, iInstCount, iModSize)
                            If FileExists(JoinPath(RESDIR, "ares.mix")) Then Call ActivateMod_ActivateFile("ares.mix", JoinPath(RESDIR, "ares.mix"), iMod, iInstCount, iModSize)
                            Call WriteINIStr("Mod", "Syringe", "yes", ProgramINI) 'remember we are using Syringe
                        Else
                            Call WriteLogEntry("This mod requires the Ares DLL, which is not present.")
                            Call MsgBox("Failed to activate mod!" & vbCrLf & "This mod requires the Ares DLL, which is not present." & vbCrLf & "The DLL has most likely failed to download - check your Internet connection.", vbOKOnly + vbExclamation)
                        End If
                    End If
                End If
            End If
            If bAresOk Then
                Call UpdatePleaseWait(, "Recalculating mod disk usage.")
                Call WriteLogEntry("Mod activated.", LogLevel1)
                Call WriteLogEntry("Recalculating mod disk usage.", LogLevel1)
                Call WriteINIStr("Mod", "DiskUsage", CStr(iModSize), ProgramINI)
                Mods(iMod).ModSize = GetDirectorySize(Mods(iMod).ModPath)
                'could just call DisplayModDetails(iMod) but for now this is the only thing that is updated
                lblModSize(Mods(iMod).ModType).Caption = DataSize(Mods(iMod).ModSize) & " (+" & DataSize(iModSize) & ")"
                ActivateMod = True
                Call UpdatePleaseWait(, "Mod activated and ready to launch.")
            Else
                ActivateMod = False
            End If
            Call HidePleaseWait
        Else
            Call WriteLogEntry(Mods(iMod).ModName & " already active. Skipping activation procedure.", LogLevel1)
            ActivateMod = True
        End If
    End If
    Call CallStackPop
End Function

Private Function ActivateMod_PrepareCSF(ByVal iMod As Integer, ByVal sCSF As String, ByVal sKiln As String, ByVal sLang As String, ByRef csfcsf As MarshallxCSFClass)
    Dim sSource As String
    Dim hMIX As Integer
    sSource = JoinPath(Mods(iMod).ModPath, "string table\" & sCSF)
    If Not FileExists(sSource) Then
        sSource = JoinPath(Mods(iMod).ModPath, "stringtable\" & sCSF)
        If Not FileExists(sSource) Then
            sSource = JoinPath(Mods(iMod).ModPath, "string_table\" & sCSF)
        End If
    End If
    If FileExists(sSource) Then
        'use that as the base
        Call FileCopy(sSource, JoinPath(sKiln, sCSF))
        Call WriteLogEntry(Quote(sSource) & " copied to " & Quote(JoinPath(sKiln, sCSF)) & ".")
        Set csfcsf = New MarshallxCSFClass
        Call csfcsf.LoadCSF(JoinPath(sKiln, sCSF))
        If Mods(iMod).ModType = TypeMod And Not Mods(iMod).ModIsForRA2 Then
            Call csfcsf.UpdateWith(JoinPath(RESDIR, "yrpm.csf"), False)
            Call WriteLogEntry(Quote(JoinPath(RESDIR, "yrpm.csf")) & " merged into " & Quote(JoinPath(sKiln, sCSF)) & ".")
        End If
    Else
        'extract it from language.mix
        If DCoderDLL Then
            hMIX = DCoderMIXOpen(JoinPath(RA2DIR, sLang))
            Call DCoderMIXExtract(hMIX, sCSF, JoinPath(sKiln, sCSF), JoinPath(RA2DIR, sLang))
            Call DCoderMIXClose(hMIX)
            Set csfcsf = New MarshallxCSFClass
            Call csfcsf.LoadCSF(JoinPath(sKiln, sCSF))
            If Mods(iMod).ModType = TypeMod And Not Mods(iMod).ModIsForRA2 Then
                Call csfcsf.UpdateWith(JoinPath(RESDIR, "yrpm.csf"), False)
                Call WriteLogEntry(Quote(JoinPath(RESDIR, "yrpm.csf")) & " merged into " & Quote(JoinPath(sKiln, sCSF)) & ".")
            End If
        Else
            Call WriteLogEntry("DCoder DLL is missing! Unable to extract " & sCSF & " from " & JoinPath(RA2DIR, sLang))
            Call WriteLogEntry("Cannot merge " & Quote(JoinPath(RESDIR, "yrpm.csf")) & " into " & sCSF & ".")
        End If
    End If
End Function

Private Function DeleteMod_DeleteFolder(ByRef bNoErrors As Boolean, ByVal sPath As String, ByVal sRelative As String, ByRef sUninstData() As String, ByVal iModType As Integer) As Boolean
    Dim h As Long
    Dim FD As WIN32_FIND_DATA
    Dim r As Long
    Dim sName As String
    Dim sSubPath As String
    Dim sSubPathRel As String
    Dim bOk As Boolean
    Dim bSubOk As Boolean
    Dim iPass As Integer
    Dim iUninFile As Long
    Dim mbResult As VbMsgBoxResult
    Dim tSaves As Integer
    Dim tIPB As Integer
    Dim tPCX As Integer
    bOk = True
    If Not DirIsEmpty(sPath) Then
        tSaves = 0
        tIPB = 0
        tPCX = 0
        For iPass = 1 To 2
            If iPass = 1 Or Not bOk Then
                h = FindFirstFile(JoinPath(sPath, "*"), FD)
                If h <> INVALID_HANDLE_VALUE Then
                    Do
                        sName = Left$(FD.cFileName, InStr(FD.cFileName, vbNullChar) - 1)
                        Select Case UCase$(sName)
                        Case ".", ".."
                            'do nothing
                        Case Else
                            sSubPath = JoinPath(sPath, sName)
                            If Len(sRelative) <> 0 Then
                                sSubPathRel = sRelative & "\" & sName
                            Else
                                sSubPathRel = sName
                            End If
                            If (FD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
                                'It's a subfolder. Is it an auto-remove one?
                                bSubOk = False
                                Select Case UCase$(sSubPathRel)
                                Case "LAUNCHER"
                                    bSubOk = True
                                Case "SAVES"
                                    If iModType = TypeMod Then
                                        If tSaves = 0 Then
                                            mbResult = MsgBox("Do you wish to delete any saved games for this mod?", vbYesNo + vbQuestion, App.Title)
                                            Select Case mbResult
                                            Case vbYes: tSaves = 1
                                            Case vbNo: tSaves = -1
                                            End Select
                                        End If
                                        If tSaves = 1 Then bSubOk = True
                                    End If
                                Case "KILN", "MANUAL", "SCREEN", "THEME", "VIDEO", "CAMEO", "SHP", "SPEECH", "TMP", "HVA", "INI", "INTERFACE", "MAP", "MIX", "VXL", "SIDE1", "SIDE 1", "SIDE_1", "SIDE2", "SIDE 2", "SIDE_2", "SIDE3", "SIDE 3", "SIDE_3", "SIDE4", "SIDE 4", "SIDE_4", "TAUNTS", "STRING TABLE", "STRINGTABLE", "STRING_TABLE", "SOUND", "SOUNDS", "FA2FILES", "SYRINGE"
                                    If iModType <> TypeProgram Then
                                        bSubOk = True
                                    End If
                                End Select
                                If Not bSubOk Then
                                    'Not an auto folder so check the uninst data
                                    iUninFile = 1
                                    Do While iUninFile <= UBound(sUninstData())
                                        If StrComp(UCase$(sUninstData(iUninFile)), UCase$(sSubPathRel), vbBinaryCompare) = 0 Then
                                            bSubOk = True
                                            Exit Do
                                        End If
                                        iUninFile = iUninFile + 1
                                    Loop
                                    If Not bSubOk Then
                                        'Not in uninst data so descend into the directory
                                        bSubOk = DeleteMod_DeleteFolder(bNoErrors, sSubPath, sSubPathRel, sUninstData(), iModType)
                                    End If
                                End If
                                If bSubOk Then
                                    'This subfolder is eligible for removal...
                                    If Not bOk Then
                                        '...and we want to remove it now
                                        Call UpdatePleaseWait("", sSubPathRel)
                                        Call WriteLogEntry("Moving " & Quote(sSubPath) & " to the recycle bin.")
                                        If Not MoveToRecycleBin(sSubPath, False) Then
                                            bNoErrors = False
                                            Call WriteLogEntry("Failed to move " & Quote(sSubPath) & " to the recycle bin.")
                                        End If
                                    End If
                                Else
                                    bOk = False 'can't remove this subfolder so start removing its siblings
                                End If
                            Else
                                'It's a file. Check if it's eligible for auto-remove
                                bSubOk = False
                                If iModType = TypeMod Then
                                    If Len(sRelative) = 0 Then
                                        'files in the root of the mod's folder
                                        Select Case FileType(sSubPath)
                                        Case "YPL"
                                            bSubOk = True
                                        Case "IPB"
                                            If tIPB = 0 Then
                                                mbResult = MsgBox("Do you wish to delete any scripted videos (*.IPB) for this mod?", vbYesNo + vbQuestion, App.Title)
                                                Select Case mbResult
                                                Case vbYes: tIPB = 1
                                                Case vbNo: tIPB = -1
                                                End Select
                                            End If
                                            If tIPB = 1 Then bSubOk = True
                                        Case "PCX"
                                            If tPCX = 0 Then
                                                mbResult = MsgBox("Do you wish to delete any screenshots (*.PCX) for this mod?", vbYesNo + vbQuestion, App.Title)
                                                Select Case mbResult
                                                Case vbYes: tPCX = 1
                                                Case vbNo: tPCX = -1
                                                End Select
                                            End If
                                            If tPCX = 1 Then bSubOk = True
                                        End Select
                                    End If
                                End If
                                If Not bSubOk Then
                                    'check if the file is in the uninstall data
                                    bSubOk = False
                                    iUninFile = 1
                                    Do While iUninFile <= UBound(sUninstData())
                                        If StrComp(UCase$(sUninstData(iUninFile)), UCase$(sSubPathRel), vbBinaryCompare) = 0 Then
                                            bSubOk = True
                                            Exit Do
                                        End If
                                        iUninFile = iUninFile + 1
                                    Loop
                                End If
                                If bSubOk Then
                                    'This file is eligible for removal...
                                    If Not bOk Then
                                        '...and we want to remove it now
                                        Call UpdatePleaseWait("", sSubPathRel)
                                        Call WriteLogEntry("Moving " & Quote(sSubPath) & " to the recycle bin.")
                                        If Not MoveToRecycleBin(sSubPath, False) Then
                                            bNoErrors = False
                                            Call WriteLogEntry("Failed to move " & Quote(sSubPath) & " to the recycle bin.")
                                        End If
                                    End If
                                Else
                                    bOk = False 'can't remove this subfolder so start removing its siblings
                                End If
                            End If
                        End Select
                    Loop While FindNextFile(h, FD)
                    r = FindClose(h): Debug.Assert r
                Else
                    bOk = False
                End If
            End If
        Next iPass
    End If
    DeleteMod_DeleteFolder = bOk
End Function

Private Function ActivateMod_CheckDiskSpace(ByVal iMod As Integer, ByVal RecompileAudio As Boolean, ByVal RecompileStrings As Boolean, ByVal RecompileExpand As Boolean, ByVal sKiln As String) As Boolean
    Dim lSpaceRequiredMod As Long
    Dim lSpaceRequiredKiln As Long
    Dim lSpaceUsedDir As Long
    Dim bRA2Ok As Boolean
    Dim bKilnOk As Boolean
    Dim bOk As Boolean
    Dim sMessage As String
    Dim fso As FileSystemObject
    Dim fso_root As Folder
    Dim fso_folder As Folder
    Dim fso_file As File
    Set fso = New FileSystemObject
    Set fso_root = fso.GetFolder(Mods(iMod).ModPath)
    For Each fso_folder In fso_root.SubFolders
        lSpaceUsedDir = GetDirectorySize(JoinPath(Mods(iMod).ModPath, fso_folder.Name), False, False)
        Select Case UCase$(fso_folder.Name)
        Case "SCREEN"
            lSpaceRequiredMod = lSpaceRequiredMod + lSpaceUsedDir
        Case "THEME", "VIDEO"
            For Each fso_file In fso_folder.Files
                Select Case FileType(fso_file.Name)
                Case "OGG", "FLAC"
                    sMessage = ChangeFileType(fso_file.Path, "WAV")
                    If FileExists(sMessage) Then
                        lSpaceRequiredMod = lSpaceRequiredMod + GetFileSize(sMessage)
                    Else
                        lSpaceRequiredMod = lSpaceRequiredMod + (fso_file.Size * 6)
                    End If
                Case Else
                    lSpaceRequiredMod = lSpaceRequiredMod + fso_file.Size
                End Select
            Next
        Case "SPEECH"
            For Each fso_file In fso_folder.Files
                Select Case FileType(fso_file.Name)
                Case "OGG", "FLAC"
                    sMessage = ChangeFileType(fso_file.Path, "WAV")
                    If FileExists(sMessage) Then
                        lSpaceRequiredMod = lSpaceRequiredMod + GetFileSize(sMessage)
                    Else
                        lSpaceRequiredMod = lSpaceRequiredMod + (fso_file.Size * 6)
                        If RecompileExpand Then lSpaceRequiredKiln = lSpaceRequiredKiln + (fso_file.Size * 6)
                    End If
                Case Else
                    lSpaceRequiredMod = lSpaceRequiredMod + fso_file.Size
                    If RecompileExpand Then lSpaceRequiredKiln = lSpaceRequiredKiln + fso_file.Size
                End Select
            Next
        Case "CAMEO", "HVA", "INI", "INTERFACE", "MAP", "MIX", "SHP", "TMP", "VXL" 'EXPAND
            If RecompileExpand Then lSpaceRequiredKiln = lSpaceRequiredKiln + lSpaceUsedDir
            lSpaceRequiredMod = lSpaceRequiredMod + lSpaceUsedDir
        Case "SIDE1", "SIDE 1", "SIDE_1"
            If RecompileExpand Then lSpaceRequiredKiln = lSpaceRequiredKiln + lSpaceUsedDir
            lSpaceRequiredMod = lSpaceRequiredMod + lSpaceUsedDir
            If Not FileExists(JoinPath(fso_folder.Name, "sidec01.mix")) Then
                If RecompileExpand Then lSpaceRequiredKiln = lSpaceRequiredKiln + 2099412
                lSpaceRequiredMod = lSpaceRequiredMod + 2099412
            End If
        Case "SIDE2", "SIDE 2", "SIDE_2"
            If RecompileExpand Then lSpaceRequiredKiln = lSpaceRequiredKiln + lSpaceUsedDir
            lSpaceRequiredMod = lSpaceRequiredMod + lSpaceUsedDir
            If Not FileExists(JoinPath(fso_folder.Name, "sidec02.mix")) Then
                If RecompileExpand Then lSpaceRequiredKiln = lSpaceRequiredKiln + 2102564
                lSpaceRequiredMod = lSpaceRequiredMod + 2102564
            End If
        Case "SIDE3", "SIDE 3", "SIDE_3"
            If RecompileExpand Then lSpaceRequiredKiln = lSpaceRequiredKiln + lSpaceUsedDir
            lSpaceRequiredMod = lSpaceRequiredMod + lSpaceUsedDir
            If Not Mods(iMod).ModUseYuriUI Then
                If Not FileExists(JoinPath(fso_folder.Name, "sidec02md.mix")) Then
                    If RecompileExpand Then lSpaceRequiredKiln = lSpaceRequiredKiln + 1823972
                    lSpaceRequiredMod = lSpaceRequiredMod + 1823972
                End If
            End If
        Case "SIDE4", "SIDE 4", "SIDE_4"
            If RecompileExpand Then lSpaceRequiredKiln = lSpaceRequiredKiln + lSpaceUsedDir
            lSpaceRequiredMod = lSpaceRequiredMod + lSpaceUsedDir
        Case "TAUNTS"
            For Each fso_file In fso_folder.Files
                Select Case FileType(fso_file.Name)
                Case "OGG", "FLAC"
                    sMessage = ChangeFileType(fso_file.Path, "WAV")
                    If FileExists(sMessage) Then
                        lSpaceRequiredMod = lSpaceRequiredMod + GetFileSize(sMessage)
                    Else
                        lSpaceRequiredMod = lSpaceRequiredMod + (fso_file.Size * 6)
                    End If
                Case Else
                    lSpaceRequiredMod = lSpaceRequiredMod + fso_file.Size
                End Select
            Next
        Case "STRING TABLE", "STRINGTABLE", "STRING_TABLE"
            If RecompileStrings Then
                lSpaceRequiredKiln = lSpaceRequiredKiln + lSpaceUsedDir
                lSpaceRequiredMod = lSpaceRequiredMod + lSpaceUsedDir
                If Mods(iMod).ModIsForRA2 Then
                    If Not FileExists(JoinPath(fso_folder.Name, "ra2.csf")) Then
                        Select Case OptRA2Lang
                        Case 0 'US
                            lSpaceRequiredMod = lSpaceRequiredMod + 485776
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 485776
                        Case 2 'German
                            lSpaceRequiredMod = lSpaceRequiredMod + 518358
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 518358
                        Case 3 'French
                            lSpaceRequiredMod = lSpaceRequiredMod + 525116
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 525116
                        Case 8 'Korean
                            Call WriteLogEntry("File size of Korean ra2.csf is unknown. Overestimating...")
                            lSpaceRequiredMod = lSpaceRequiredMod + 1024000
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 1024000
                        Case 9 'Chinese
                            Call WriteLogEntry("File size of Chinese ra2.csf is unknown. Overestimating...")
                            lSpaceRequiredMod = lSpaceRequiredMod + 1024000
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 1024000
                        End Select
                    End If
                Else
                     If Not FileExists(JoinPath(fso_folder.Name, "ra2md.csf")) Then
                        Select Case OptRA2Lang
                        Case 0
                            lSpaceRequiredMod = lSpaceRequiredMod + 573269 'US
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 573269 'US
                        Case 2 'German
                            lSpaceRequiredMod = lSpaceRequiredMod + 615723
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 615723
                        Case 3 'French
                            lSpaceRequiredMod = lSpaceRequiredMod + 621061
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 621061
                        Case 8 'Korean
                            Call WriteLogEntry("File size of Korean ra2md.csf is unknown. Overestimating...")
                            lSpaceRequiredMod = lSpaceRequiredMod + 1024000
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 1024000
                        Case 9 'Chinese
                            Call WriteLogEntry("File size of Chinese ra2md.csf is unknown. Overestimating...")
                            lSpaceRequiredMod = lSpaceRequiredMod + 1024000
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 1024000
                        End Select
                    End If
                End If
            Else
                If Mods(iMod).ModIsForRA2 Then
                    lSpaceRequiredMod = lSpaceRequiredMod + GetFileSize(JoinPath(sKiln, "ra2.csf"))
                Else
                    lSpaceRequiredMod = lSpaceRequiredMod + GetFileSize(JoinPath(sKiln, "ra2md.csf"))
                End If
            End If
        Case "SOUND", "SOUNDS" 'AUDIO.BAG/IDX
            'the bag itself
            If RecompileAudio Then
                If Not FileExists(JoinPath(fso_folder.Name, "audio.bag")) Or Not FileExists(JoinPath(fso_folder.Name, "audio.idx")) Then
                    If Mods(iMod).ModIsForRA2 Then
                        Select Case OptRA2Lang
                        Case 0 'US
                            lSpaceRequiredMod = lSpaceRequiredMod + 17684696 'audio.bag=, audio.idx=
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 17684696 'audio.bag=, audio.idx=
                        Case 2 'German
                            Call WriteLogEntry("File size of German audio.bag is unknown. Overestimating...")
                            lSpaceRequiredMod = lSpaceRequiredMod + 18109886 'audio.bag=18068636, audio.idx=41250
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 18109886
                        Case 3 'French
                            lSpaceRequiredMod = lSpaceRequiredMod + 17983701 'audio.bag=17942181, audio.idx=41520
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 17983701
                        Case 8 'Korean
                            Call WriteLogEntry("File size of Korean audio.bag is unknown. Overestimating...")
                            lSpaceRequiredMod = lSpaceRequiredMod + 32000000 'audio.bag=, audio.idx=
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 32000000
                        Case 9 'Chinese
                            Call WriteLogEntry("File size of Chinese audio.bag is unknown. Overestimating...")
                            lSpaceRequiredMod = lSpaceRequiredMod + 32000000 'audio.bag=, audio.idx=
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 32000000
                        End Select
                    Else
                        Select Case OptYRLang
                        Case 0 'US
                            lSpaceRequiredMod = lSpaceRequiredMod + 37707904 'audio.bag=, audio.idx=
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 37707904 'audio.bag=, audio.idx=
                        Case 2 'German
                            lSpaceRequiredMod = lSpaceRequiredMod + 42103816 'audio.bag=42021544, audio.idx=82272
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 42103816
                        Case 3 'French
                            lSpaceRequiredMod = lSpaceRequiredMod + 39172408 'audio.bag=39090136, audio.idx=82272
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 39172408
                        Case 8 'Korean
                            Call WriteLogEntry("File size of Korean audio.bag is unknown. Overestimating...")
                            lSpaceRequiredMod = lSpaceRequiredMod + 64000000 'audio.bag=, audio.idx=
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 64000000
                        Case 9 'Chinese
                            Call WriteLogEntry("File size of Chinese audio.bag is unknown. Overestimating...")
                            lSpaceRequiredMod = lSpaceRequiredMod + 64000000 'audio.bag=, audio.idx=
                            lSpaceRequiredKiln = lSpaceRequiredKiln + 64000000
                        End Select
                    End If
                End If
                'individual sounds
                For Each fso_file In fso_folder.Files
                    Select Case FileType(fso_file.Name)
                    Case "OGG", "FLAC"
                        sMessage = ChangeFileType(fso_file.Path, "WAV")
                        If FileExists(sMessage) Then
                            lSpaceRequiredMod = lSpaceRequiredMod + GetFileSize(sMessage)
                            lSpaceRequiredKiln = lSpaceRequiredKiln + GetFileSize(sMessage)
                        Else
                            lSpaceRequiredMod = lSpaceRequiredMod + (fso_file.Size * 6)
                            lSpaceRequiredKiln = lSpaceRequiredKiln + (fso_file.Size * 6)
                        End If
                    Case Else
                        lSpaceRequiredMod = lSpaceRequiredMod + fso_file.Size
                        lSpaceRequiredKiln = lSpaceRequiredKiln + fso_file.Size
                    End Select
                Next
            Else
                lSpaceRequiredMod = lSpaceRequiredMod + GetFileSize(JoinPath(sKiln, "audio.bag")) + GetFileSize(JoinPath(sKiln, "audio.idx"))
            End If
        Case "FA2FILES"
            lSpaceRequiredMod = lSpaceRequiredMod + lSpaceUsedDir
        Case "SAVES"
            lSpaceRequiredMod = lSpaceRequiredMod + lSpaceUsedDir
        End Select
    Next
    bOk = True
    bRA2Ok = True
    bKilnOk = True
    If UCase$(Left$(RA2DIR, 1)) = UCase$(Left$(Mods(iMod).ModPath, 1)) Then
        lSpaceRequiredMod = lSpaceRequiredMod + lSpaceRequiredKiln
        lSpaceRequiredKiln = 0
    End If
    If lSpaceRequiredMod <> 0 Then
        lSpaceRequiredMod = lSpaceRequiredMod + OptSafetySpace
        If FreeDiskSpace(UCase$(Left$(RA2DIR, 1))) < lSpaceRequiredMod Then bRA2Ok = False
    End If
    If lSpaceRequiredKiln <> 0 Then
        lSpaceRequiredKiln = lSpaceRequiredKiln + OptSafetySpace
        If FreeDiskSpace(UCase$(Left$(Mods(iMod).ModPath, 1))) < lSpaceRequiredKiln Then bKilnOk = False
    End If
    If Not bRA2Ok Then
        bOk = False
        If Not bKilnOk Then
            sMessage = "Insufficient free disk space to activate mod. This mod requires at least " & DataSize(lSpaceRequiredMod) & " free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & " and " & DataSize(lSpaceRequiredKiln) & " free disk space on drive " & UCase$(Left$(Mods(iMod).ModPath, 1)) & "."
        Else
            sMessage = "Insufficient free disk space to activate mod. This mod requires at least " & DataSize(lSpaceRequiredMod) & " free disk space on drive " & UCase$(Left$(RA2DIR, 1)) & "."
        End If
    ElseIf Not bKilnOk Then
        bOk = False
        sMessage = "Insufficient free disk space to activate mod. This mod requires at least " & DataSize(lSpaceRequiredKiln) & " free disk space on drive " & UCase$(Left$(Mods(iMod).ModPath, 1)) & "."
    End If
    If Not bOk Then Call WriteLogEntry(sMessage, LogMsgBoxExclaim)
    ActivateMod_CheckDiskSpace = bOk
End Function

Private Function ActivateMod_ReactivateNeeded(ByVal iMod As Integer) As Boolean
    Dim sLBU As String
    Dim bOk As Boolean
    bOk = True
    If Not OptRecompile Then 'always reinstall if RecompileMods is on
        If Not CL_dev Then 'CL_dev forces RecompileMod
            sLBU = JoinPath(Mods(iMod).ModPath, "launcher\userdata.lbu")
            If FileExists(sLBU) Then
                If BooleanStringToBoolean(ReadINIStr("General", "ModIsActive", sLBU, "no")) Then
                    If Mods(iMod).ModName = ReadINIStr("Mod", "Name", ProgramINI) Then
                        If Mods(iMod).ModVersion = ReadINIStr("Mod", "Version", ProgramINI) Then
                            bOk = False 'mod is already installed
                        End If
                    End If
                End If
            End If
        End If
    End If
    ActivateMod_ReactivateNeeded = bOk
End Function

Private Sub DCoderMIXExtract(ByVal hMIX As Integer, ByVal sWhatToExtract As String, ByVal sSource, ByVal sDest As String)
    Call WriteLogEntry("DCoder DLL is missing! Cannot extract " & Quote(sDest) & " from " & Quote(sSource))
End Sub

Private Sub DCoderBAGWrite(ByVal hBAG As Integer, ByVal sDestFile As String)
    Call WriteLogEntry("DCoder DLL is missing! Cannot assemble " & Quote(sDestFile))
End Sub

Private Sub DCoderCSFWrite(ByVal hCSF As Integer, ByVal sDestFile As String)
    Call WriteLogEntry("DCoder DLL is missing! Cannot assemble " & Quote(sDestFile))
End Sub

Private Sub DCoderMIXWrite(ByVal hMIX As Integer, ByVal sDestFile As String)
    Call WriteLogEntry("DCoder DLL is missing! Cannot assemble " & Quote(sDestFile))
End Sub

Private Function DCoderBAGOpen(ByVal sBAGFile As String) As Integer
    Call WriteLogEntry("DCoder DLL is missing! Cannot open " & Quote(sBAGFile) & ".", LogLevel2)
    DCoderBAGOpen = 0
End Function

Private Function DCoderCSFOpen(ByVal sCSFFile As String) As Integer
    Call WriteLogEntry("DCoder DLL is missing! Cannot open " & Quote(sCSFFile) & ".", LogLevel2)
    DCoderCSFOpen = 0
End Function

Private Function DCoderMIXOpen(ByVal sMIXFile As String) As Integer
    Call WriteLogEntry("DCoder DLL is missing! Cannot open " & Quote(sMIXFile) & ".", LogLevel2)
    DCoderMIXOpen = 0
End Function

Private Function DCoderMIXCreate(ByVal sMIXFile As String) As Integer
    Call WriteLogEntry("DCoder DLL is missing! Cannot create " & Quote(sMIXFile) & ".", LogLevel2)
    DCoderMIXCreate = 0
End Function

Private Sub DCoderCsfClose(ByVal hCSF As Integer)
    Call WriteLogEntry("DCoder DLL is missing! Cannot close CSF file #" & CStr(hCSF) & ".", LogLevel2)
End Sub

Private Sub DCoderBAGClose(ByVal hBAG As Integer)
    Call WriteLogEntry("DCoder DLL is missing! Cannot close BAG file #" & CStr(hBAG) & ".", LogLevel2)
End Sub

Private Sub DCoderMIXClose(ByVal hMIX As Integer)
    Call WriteLogEntry("DCoder DLL is missing! Cannot close MIX file #" & CStr(hMIX) & ".", LogLevel2)
End Sub

Private Sub DCoderCSFMerge(ByVal hCSF As Integer, ByVal sCSFFile As String, ByVal sDestFile As String)
    Call WriteLogEntry("DCoder DLL is missing! Cannot merge " & Quote(sCSFFile) & " into " & Quote(sDestFile) & ".", LogLevel2)
End Sub

Private Sub DCoderCSFInsert(ByVal hCSF As Integer, ByVal sID As String, ByVal sValue, ByVal sDestFile As String)
    Call WriteLogEntry("DCoder DLL is missing! Cannot insert " & Quote(sID) & " into " & Quote(sDestFile) & ".", LogLevel2)
End Sub

Private Sub DCoderBAGMerge(ByVal hBAG As Integer, ByVal sBAGFile As String, ByVal sDestFile As String)
    Call WriteLogEntry("DCoder DLL is missing! Cannot merge " & Quote(sBAGFile) & " into " & Quote(sDestFile) & ".", LogLevel2)
End Sub

Private Sub DCoderWAVInsert(ByVal hBAG As Integer, ByVal sWavFile As String, ByVal sDestFile As String)
    Call WriteLogEntry("DCoder DLL is missing! Cannot insert " & Quote(sWavFile) & " into " & Quote(sDestFile) & ".", LogLevel2)
End Sub

Private Sub DCoderMIXInsert(ByVal hMIX As Integer, ByVal sSourceFile As String, ByVal sDestFile As String)
    Call WriteLogEntry("DCoder DLL is missing! Cannot insert " & Quote(sSourceFile) & " into " & Quote(sDestFile) & ".", LogLevel2)
End Sub

Private Function ActivateMod_RemoveResidualFiles(ByRef iRestore As Long) As Boolean
    'This will only remove non-LB mod files
    'Should only be run by ActivateMod
    Dim bOk As Boolean
    Dim fso As FileSystemObject
    Dim fso_folder As Folder
    Dim fso_file As File
    Call CallStackPush(Me.Name & ".ActivateMod_RemoveResidualFiles()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    bOk = True
    Call WriteLogEntry("Removing residual files...", LogLevel1)
    iRestore = 0
    Do While Len(ReadINIStr("Restore", CStr(iRestore), ProgramINI)) <> 0
        iRestore = iRestore + 1
    Loop
    Set fso = New FileSystemObject
    Set fso_folder = fso.GetFolder(RA2DIR)
    For Each fso_file In fso_folder.Files
        If FileIsDirty(fso_file.Name) Then
            If SafeFiles_Find(fso_file.Name) = 0 Then
                If Not ActivateMod_RemoveResidualFile(fso_file.Name, iRestore) Then
                    bOk = False
                    Exit For
                End If
            Else
                Call WriteLogEntry("Residual file " & Quote(JoinPath(RA2DIR, fso_file.Name)) & " is marked as safe by user... ignoring.")
            End If
        End If
    Next
    If iRestore = 0 Then
        Call WriteLogEntry("No [unsafe] residual files detected.", LogLevel1)
    Else
        Call WriteLogEntry(CStr(iRestore) & " residual files temporarily removed.")
        Call DenyPersistentMods
    End If
    ActivateMod_RemoveResidualFiles = bOk
    Call CallStackPop
End Function

Private Function ActivateMod_RemoveResidualFile(ByVal sFile As String, ByRef iRestore As Long, Optional ByVal bFA2Mode As Boolean = False) As Boolean
    Dim sBackup As String
    Dim sSource As String
    Dim bMoved As Boolean
    Call CallStackPush(Me.Name & ".ActivateMod_RemoveResidualFile()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    bMoved = False
    sBackup = JoinPath(BACKUPDIR, sFile)
    sSource = JoinPath(IIf(Not bFA2Mode, RA2DIR, FA2DIR), sFile)
    If FileExists(sBackup) Then
        Call WriteLogEntry("Unexpected backup file found! Deleting " & Quote(sBackup) & " to make way for new backup file.")
        Call Kill(sBackup)
    End If
    '---move the file---
    If FreeDiskSpace(UCase$(Left$(sBackup, 1))) >= (GetFileSize(sSource) + OptSafetySpace) Then
        Call WriteLogEntry("Backing up (moving) " & Quote(sSource) & " to " & Quote(sBackup))
        On Error GoTo CannotMove
        Name sSource As sBackup
        On Error GoTo LocalErr
        bMoved = True
        Call WriteINIStr("Restore", CStr(iRestore), sFile, ProgramINI)
        iRestore = iRestore + 1
CannotMove:
        On Error GoTo LocalErr
        If Not bMoved Then
            Call WriteLogEntry("Failed to backup (move) " & Quote(sSource) & " to " & Quote(sBackup))
            Call MsgBox("Failed to activate mod!" & vbCrLf & "Failed to backup (move) " & Quote(sSource) & " to " & Quote(sBackup), vbOKOnly + vbExclamation)
        End If
    Else
        Call WriteLogEntry("Insufficient disk space on drive " & Left$(sBackup, 1) & " to backup (move) " & Quote(sSource) & " to " & Quote(sBackup))
        Call MsgBox("Failed to activate mod!" & vbCrLf & "Insufficient disk space on drive " & Left$(sBackup, 1) & " to backup (move) " & Quote(sSource) & " to " & Quote(sBackup), vbOKOnly + vbExclamation)
    End If
    '---file moved (or not)---
    ActivateMod_RemoveResidualFile = bMoved
    Call CallStackPop
End Function

Friend Function FileIsDirty(ByVal FileName As String) As Boolean
    Dim iCounter As Integer
    Dim iFile As Integer
    Dim sFile As String
    Dim sPluginID As String
    Dim Ok As Boolean
    Ok = True
    FileName = UCase$(FileName)
    Select Case FileName
    Case "DEBUG.TXT"
        Ok = False
    Case "THEMEMD.INI"
        If Len(ReadINIStr("YRPMOPTS", "Music", JoinPath(RA2DIR, FileName))) <> 0 Then Ok = False
    Case "SESSION.IPB" 'because cmdLaunch might create this file before ActivateMod_RemoveResidualFiles runs. The LaunchMod_PlayVideo routine will take care of this file.
        Ok = False
    Case "EXCEPT.TXT" 'this is destructive to prevent mods including it, but is residual because user-generated content
        Ok = True
    Case Else
        If FileIsDestructive(FileName) Then
            Ok = False
        Else
            If FileIsSoundtrack(FileName) Then
                Ok = False
            Else
                If FileIsUserTheme(FileName) Then
                    Ok = False
                Else
                    If FileIsCustomMap(FileName) Then
                        Ok = False
                    Else
                        If FileIsSeed(FileName) Then
                            Ok = False
                        Else
                            If FileIsAresComponent(FileName) Then
                                Ok = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Select
    If Ok Then
        iFile = 0
        sFile = ReadINIStr("Mod", CStr(iFile), ProgramINI)
        Do While Len(sFile) <> 0
            If UCase$(sFile) = FileName Then
                Ok = False
                Exit Do
            End If
            iFile = iFile + 1
            sFile = ReadINIStr("Mod", CStr(iFile), ProgramINI)
        Loop
        If Ok Then
            iCounter = 0
            sPluginID = ReadINIStr("ActivePlugins", CStr(iCounter), ProgramINI)
            Do While Len(sPluginID) <> 0
                iFile = 0
                sFile = ReadINIStr("Plugin" & sPluginID, CStr(iFile), ProgramINI)
                Do While Len(sFile) <> 0
                    If UCase$(sFile) = FileName Then
                        Ok = False
                        Exit Do
                    End If
                    iFile = iFile + 1
                    sFile = ReadINIStr("Plugin" & sPluginID, CStr(iFile), ProgramINI)
                Loop
                If Not Ok Then Exit Do
                iCounter = iCounter + 1
                sPluginID = ReadINIStr("ActivePlugins", CStr(iCounter), ProgramINI)
            Loop
        End If
    End If
    FileIsDirty = Ok
End Function

Private Sub ConvertOggToWav(ByVal SourcePath As String, Optional ByVal DestPath As String = "")
    Dim process_id
    Dim process_handle
    If DestPath = "" Then DestPath = ChangeFileType(SourcePath, "wav")
    If FileExists(DestPath) Then
        Call WriteLogEntry("Deleting " & Quote(DestPath) & " to make way for new file.")
        Call Kill(DestPath)
    End If
    process_id = Shell(Quote(JoinPath(EXEDIR, "Resource", "oggdec.exe")) & " " & Quote(SourcePath), vbHide)
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    If FileExists(ChangeFileType(SourcePath, "wav")) Then
        Call WriteLogEntry(Quote(SourcePath) & " converted to " & Quote(ChangeFileType(SourcePath, "wav")))
        SourcePath = ChangeFileType(SourcePath, "wav")
        If UCase$(SourcePath) <> UCase$(DestPath) Then
            If Not DirExists(GetFilePath(DestPath)) Then Call LoggedMakePath(GetFilePath(DestPath))
            Call LoggedMove(SourcePath, DestPath)
        End If
    Else
        Call WriteLogEntry("Failed to convert " & Quote(SourcePath) & " to " & Quote(DestPath), LogLevel1)
    End If
End Sub

Private Sub ConvertFlacToWav(ByVal SourcePath As String, Optional ByVal DestPath As String = "")
    Dim process_id
    Dim process_handle
    If DestPath = "" Then DestPath = ChangeFileType(SourcePath, "wav")
    If FileExists(DestPath) Then
        Call WriteLogEntry("Deleting " & Quote(DestPath) & " to make way for new file.")
        Call Kill(DestPath)
    End If
    process_id = Shell(Quote(JoinPath(EXEDIR, "Resource", "flac.exe")) & " -d " & Quote(SourcePath), vbHide)
    'not using " --keep-foreign-metadata" because operation will fail if there is no foreign metadata
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    If FileExists(ChangeFileType(SourcePath, "wav")) Then
        Call WriteLogEntry(Quote(SourcePath) & " converted to " & Quote(ChangeFileType(SourcePath, "wav")))
        SourcePath = ChangeFileType(SourcePath, "wav")
        If UCase$(SourcePath) <> UCase$(DestPath) Then
            If Not DirExists(GetFilePath(DestPath)) Then Call LoggedMakePath(GetFilePath(DestPath))
            Call LoggedMove(SourcePath, DestPath)
        End If
    Else
        Call WriteLogEntry("Failed to convert " & Quote(SourcePath) & " to " & Quote(DestPath), LogLevel1)
    End If
End Sub

Public Sub DeactivateMod(Optional ByVal RestoreResidualFiles As Boolean = True, Optional ByVal HandleUserGeneratedContent As Boolean = True)
    Dim FileCount As Long
    Dim DestFile As String
    Dim BackPath As String
    Dim DestPath As String
    Dim MissingFileCount As Long
    Dim sSnapFormat As String
    Dim sScrnFormat As String
    Dim iModType As Integer
    Dim sModName As String
    Dim sModVersion As String
    Dim sModNewVersion As String
    Dim sModPath As String
    Dim bGameIsRA2 As Boolean
    Dim bOfficial As Boolean
    Dim iPos As Integer
    Dim ErrVars(1) As Variant
    Call CallStackPush(Me.Name & ".DeactivateMod(" & CStr(RestoreResidualFiles) & ", " & CStr(HandleUserGeneratedContent) & ")")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    sModName = ReadINIStr("Mod", "Name", ProgramINI)
    sModVersion = ReadINIStr("Mod", "Version", ProgramINI)
    iModType = Val(ReadINIStr("Mod", "ModType", ProgramINI))
    sSnapFormat = ReadINIStr("Mod", "SnapFormat", ProgramINI)
    sScrnFormat = ReadINIStr("Mod", "ScrnFormat", ProgramINI)
    bOfficial = BooleanStringToBoolean(ReadINIStr("Mod", "Official", ProgramINI, "no"))
    bGameIsRA2 = BooleanStringToBoolean(ReadINIStr("Mod", "GameIsRA2", ProgramINI))
    MissingFileCount = 0
    If Len(sModName) <> 0 Then
        If Len(sModVersion) = 0 Then
            Call WriteLogEntry("Deactivating " & IIf(iModType = TypeMod, "mod", "FA2 mod") & ": " & sModName, LogLevel1)
            Call ShowPleaseWait("Deactivating " & IIf(iModType = TypeMod, "mod", "FA2 mod") & ": " & sModName)
        Else
            Call WriteLogEntry("Deactivating " & IIf(iModType = TypeMod, "mod", "FA2 mod") & ": " & sModName & " [" & sModVersion & "]", LogLevel1)
            Call ShowPleaseWait("Deactivating " & IIf(iModType = TypeMod, "mod", "FA2 mod") & ": " & sModName & " [" & sModVersion & "]")
        End If
        'remove the core files
        FileCount = 0
        DestFile = ReadINIStr("Mod", CStr(FileCount), ProgramINI)
        Call UpdatePleaseWait(, DestFile)
        Do While Len(DestFile) <> 0
            If (iModType = TypeFA2Mod) And FileIsFA2File(DestFile) Then
                DestPath = JoinPath(FA2DIR, DestFile)
            Else
                DestPath = JoinPath(RA2DIR, DestFile)
            End If
            BackPath = JoinPath(BACKUPDIR, DestFile)
            If FileExists(DestPath) Then
                Call LoggedKill(DestPath)
            Else
                MissingFileCount = MissingFileCount + 1
                Call WriteLogEntry("Expected file not found! " & Quote(DestPath) & " could not be deleted.")
            End If
            Call WriteINIStr("Mod", CStr(FileCount), "", ProgramINI)
            Call WriteINIStr("Mod", CStr(FileCount) & "c", "", ProgramINI)
            FileCount = FileCount + 1
            DestFile = ReadINIStr("Mod", CStr(FileCount), ProgramINI)
        Loop
        If MissingFileCount <> 0 Then
            Call WriteLogEntry(CStr(MissingFileCount) & " files that Launch Base was expecting to delete could not be found.")
            Call DenyPersistentMods
        End If
        'user-generated content
        If (iModType = TypeMod) Then
            If HandleUserGeneratedContent Then
                sModPath = DeactivateMod_CheckModGone(sModName, sModVersion, sModNewVersion, sScrnFormat, sSnapFormat, bOfficial, bGameIsRA2)
                Call DeactivateMod_GameConfig(sModPath, bGameIsRA2)
                Call DeactivateMod_SaveGames(sModPath, sModVersion, sModNewVersion, sSnapFormat)
                Call DeactivateMod_Screenshots(sModPath, sScrnFormat)
            End If
            'mod has been deactivated so game config needs to be reset (unless about to activate another mod)
            If RestoreResidualFiles Then Call ActivateMod_GameConfig(-1, bGameIsRA2)
        End If
        'erase mod record - this could go at the beginning but leaving here in case of crash mid-deactivate
        Call WriteINIStr("Mod", "Name", "", ProgramINI)
        Call WriteINIStr("Mod", "Version", "", ProgramINI)
        Call WriteINIStr("Mod", "ModType", "", ProgramINI)
        Call WriteINIStr("Mod", "SnapFormat", "", ProgramINI)
        Call WriteINIStr("Mod", "ScrnFormat", "", ProgramINI)
        Call WriteINIStr("Mod", "GameIsRA2", "", ProgramINI)
        Call WriteINIStr("Mod", "Official", "", ProgramINI)
        Call WriteINIStr("Mod", "DiskUsage", "", ProgramINI)
        Call WriteINIStr("Mod", "Syringe", "", ProgramINI)
        If FileExists(JoinPath(sModPath, "launcher\userdata.lbu")) Then Call WriteINIStr("General", "ModIsActive", "no", JoinPath(sModPath, "launcher\userdata.lbu"))
        Call WriteLogEntry("Mod deactivated.", LogLevel1)
        'Need to update disk usage on details display if selected
        iPos = InStr(1, lblModSize(iModType).Caption, " (+")
        If iPos <> 0 Then lblModSize(iModType).Caption = Left$(lblModSize(iModType).Caption, iPos - 1)
        Call HidePleaseWait
    End If
    'Finally, restore any residual files
    If RestoreResidualFiles Then Call DeactivateMod_ResidualFiles
    Call CallStackPop
End Sub

Private Sub DeactivateMod_GameConfig(ByVal sModPath As String, ByVal bGameIsRA2 As Boolean)
    Dim sUserData As String
    Dim sConfig As String
    Dim Sections(5) As String
    Dim Flags(5, 17) As String
    Dim FlagCount(5) As Integer
    Dim FlagCounter As Integer
    Dim SectionCounter As Integer
    Dim sValue As String
    Call WriteLogEntry("Saving mod/game configuration.", LogLevel1)
    'Multiplayer
    Sections(1) = "Multiplayer"
    FlagCount(1) = 2
    Flags(1, 1) = "GameMode"
    Flags(1, 2) = "ScenIndex"
    'MultiPlayer
    Sections(2) = "MultiPlayer"
    FlagCount(2) = 5
    Flags(2, 1) = "Color"
    Flags(2, 2) = "ColorEx"
    Flags(2, 3) = "Side"
    Flags(2, 4) = "SideEx"
    Flags(2, 5) = "GameMode"
    'Skirmish
    Sections(3) = "Skirmish"
    FlagCount(3) = 17
    Flags(3, 1) = "GameMode"
    Flags(3, 2) = "ScenIndex"
    Flags(3, 3) = "GameSpeed"
    Flags(3, 4) = "Credits"
    Flags(3, 5) = "UnitCount"
    Flags(3, 6) = "ShortGame"
    Flags(3, 7) = "SuperWeaponsAllowed"
    Flags(3, 8) = "BuildOffAlly"
    Flags(3, 9) = "MCVRepacks"
    Flags(3, 10) = "CratesAppear"
    Flags(3, 11) = "Slot01"
    Flags(3, 12) = "Slot02"
    Flags(3, 13) = "Slot03"
    Flags(3, 14) = "Slot04"
    Flags(3, 15) = "Slot05"
    Flags(3, 16) = "Slot06"
    Flags(3, 17) = "Slot07"
    'LAN
    Sections(4) = "LAN"
    FlagCount(4) = 17
    For FlagCounter = 1 To 17
        Flags(4, FlagCounter) = Flags(3, FlagCounter)
    Next FlagCounter
    'WonlinePref
    Sections(5) = "WonlinePref"
    FlagCount(5) = 17
    For FlagCounter = 1 To 17
        Flags(5, FlagCounter) = Flags(3, FlagCounter)
    Next FlagCounter
    If bGameIsRA2 Then
        sConfig = JoinPath(RA2DIR, "ra2.ini")
    Else
        sConfig = JoinPath(RA2DIR, "ra2md.ini")
    End If
    sUserData = JoinPath(sModPath, "launcher\userdata.lbu")
    If FileExists(sUserData) Then
        For SectionCounter = 1 To UBound(Sections)
            For FlagCounter = 1 To FlagCount(SectionCounter)
                sValue = ReadINIStr(Sections(SectionCounter), Flags(SectionCounter, FlagCounter), sConfig)
                Call WriteINIStr(Sections(SectionCounter), Flags(SectionCounter, FlagCounter), sValue, sUserData)
            Next FlagCounter
        Next SectionCounter
    End If
End Sub

Private Sub ActivateMod_GameConfig(ByVal iMod As Integer, ByVal bGameIsRA2 As Boolean) 'if iMod is -1 then reset
    Dim sUserData As String
    Dim sConfig As String
    Dim Sections(5) As String
    Dim Flags(5, 17) As String
    Dim FlagCount(5) As Integer
    Dim Defaults(5, 17) As String
    Dim FlagCounter As Integer
    Dim SectionCounter As Integer
    Dim sValue As String
    If iMod = -1 Then
        Call WriteLogEntry("Resetting game configuration.", LogLevel1)
    Else
        Call WriteLogEntry("Restoring mod/game configuration.", LogLevel1)
    End If
    'Multiplayer
    Sections(1) = "Multiplayer"
    FlagCount(1) = 2
    Flags(1, 1) = "GameMode"
    Flags(1, 2) = "ScenIndex"
    Defaults(1, 1) = "1"
    Defaults(1, 2) = "0"
    'MultiPlayer
    Sections(2) = "MultiPlayer"
    FlagCount(2) = 5
    Flags(2, 1) = "Color"
    Flags(2, 2) = "ColorEx"
    Flags(2, 3) = "Side"
    Flags(2, 4) = "SideEx"
    Flags(2, 5) = "GameMode"
    Defaults(2, 1) = "1"
    Defaults(2, 2) = "-2"
    Defaults(2, 3) = "Americans"
    Defaults(2, 4) = "-2"
    Defaults(2, 5) = "1"
    'Skirmish
    Sections(3) = "Skirmish"
    FlagCount(3) = 17
    Flags(3, 1) = "GameMode"
    Flags(3, 2) = "ScenIndex"
    Flags(3, 3) = "GameSpeed"
    Flags(3, 4) = "Credits"
    Flags(3, 5) = "UnitCount"
    Flags(3, 6) = "ShortGame"
    Flags(3, 7) = "SuperWeaponsAllowed"
    Flags(3, 8) = "BuildOffAlly"
    Flags(3, 9) = "MCVRepacks"
    Flags(3, 10) = "CratesAppear"
    Flags(3, 11) = "Slot01"
    Flags(3, 12) = "Slot02"
    Flags(3, 13) = "Slot03"
    Flags(3, 14) = "Slot04"
    Flags(3, 15) = "Slot05"
    Flags(3, 16) = "Slot06"
    Flags(3, 17) = "Slot07"
    Defaults(3, 1) = "1"
    Defaults(3, 2) = "0"
    Defaults(3, 3) = "0"
    Defaults(3, 4) = "10000"
    Defaults(3, 5) = "10"
    Defaults(3, 6) = "yes"
    Defaults(3, 7) = "yes"
    Defaults(3, 8) = "yes"
    Defaults(3, 9) = "yes"
    Defaults(3, 10) = "yes"
    Defaults(3, 11) = "1,-2,-2"
    Defaults(3, 12) = "1,-2,-2"
    Defaults(3, 13) = "1,-2,-2"
    Defaults(3, 14) = "1,-2,-2"
    Defaults(3, 15) = "1,-2,-2"
    Defaults(3, 16) = "1,-2,-2"
    Defaults(3, 17) = "1,-2,-2"
    'LAN
    Sections(4) = "LAN"
    FlagCount(4) = 17
    For FlagCounter = 1 To 17
        Flags(4, FlagCounter) = Flags(3, FlagCounter)
        Select Case FlagCounter
        Case 11, 12, 13, 14, 15, 16
            Defaults(4, FlagCounter) = "2,-2,-2"
        Case 17
            Defaults(4, FlagCounter) = "3,-2,-2"
        Case Else
            Defaults(4, FlagCounter) = Defaults(3, FlagCounter)
        End Select
    Next FlagCounter
    'WonlinePref
    Sections(5) = "WonlinePref"
    FlagCount(5) = 17
    For FlagCounter = 1 To 17
        Flags(5, FlagCounter) = Flags(3, FlagCounter)
        Select Case FlagCounter
        Case 11, 12, 13, 14, 15, 16, 17
            Defaults(5, FlagCounter) = "2,-2,-2"
        Case Else
            Defaults(5, FlagCounter) = Defaults(3, FlagCounter)
        End Select
    Next FlagCounter
    If bGameIsRA2 Then
        sConfig = JoinPath(RA2DIR, "ra2.ini")
    Else
        sConfig = JoinPath(RA2DIR, "ra2md.ini")
    End If
    If iMod <> -1 Then sUserData = JoinPath(Mods(iMod).ModPath, "launcher\userdata.lbu")
    For SectionCounter = 1 To UBound(Sections)
        For FlagCounter = 1 To FlagCount(SectionCounter)
            If iMod = -1 Then
                'force default
                Call WriteINIStr(Sections(SectionCounter), Flags(SectionCounter, FlagCounter), Defaults(SectionCounter, FlagCounter), sConfig)
            Else
                'get value from userdata.lbu
                sValue = ReadINIStr(Sections(SectionCounter), Flags(SectionCounter, FlagCounter), sUserData)
                If Len(sValue) <> 0 Then
                    'game config has been saved so use saved value
                    Call WriteINIStr(Sections(SectionCounter), Flags(SectionCounter, FlagCounter), sValue, sConfig)
                Else
                    'game config not saved so use a default value
                    Select Case Flags(SectionCounter, FlagCounter)
                    Case "GameMode"
                        If Len(Mods(iMod).ModGameMode) <> 0 Then
                            Call WriteINIStr(Sections(SectionCounter), Flags(SectionCounter, FlagCounter), Mods(iMod).ModGameMode, sConfig)
                        Else
                            Call WriteINIStr(Sections(SectionCounter), Flags(SectionCounter, FlagCounter), Defaults(SectionCounter, FlagCounter), sConfig)
                        End If
                    Case "MapIndex"
                        If Len(Mods(iMod).ModMapIndex) <> 0 Then
                            Call WriteINIStr(Sections(SectionCounter), Flags(SectionCounter, FlagCounter), Mods(iMod).ModMapIndex, sConfig)
                        Else
                            Call WriteINIStr(Sections(SectionCounter), Flags(SectionCounter, FlagCounter), Defaults(SectionCounter, FlagCounter), sConfig)
                        End If
                    Case Else
                        Call WriteINIStr(Sections(SectionCounter), Flags(SectionCounter, FlagCounter), Defaults(SectionCounter, FlagCounter), sConfig)
                    End Select
                End If
            End If
        Next FlagCounter
    Next SectionCounter
    If OptVideoBackBuffer Then
        Call WriteINIStr("Video", "VideoBackBuffer", "yes", sConfig)
    Else
        Call WriteINIStr("Video", "VideoBackBuffer", "no", sConfig)
    End If
    If OptAllowVRAMSidebar Then
        Call WriteINIStr("Video", "AllowVRAMSidebar", "yes", sConfig)
    Else
        Call WriteINIStr("Video", "AllowVRAMSidebar", "no", sConfig)
    End If
End Sub

Private Function Init_LoadMods_NoTampering() As Boolean
    Dim iFile As Integer
    Dim sFile As String
    Dim sPath As String
    Dim bOk As Boolean
    bOk = True
    iFile = 0
    sFile = ReadINIStr("Mod", CStr(iFile), ProgramINI)
    Do While Len(sFile) <> 0
        sPath = JoinPath(RA2DIR, sFile)
        If FileExists(sPath) Then
            If OptUseCheckSums Then
                sFile = ReadINIStr("Mod", CStr(iFile) & "c", ProgramINI)
                If Len(sFile) <> 0 Then
                    Call WriteLogEntry("Getting MD5 of " & Quote(sPath), LogLevel2)
                    If UCase$(GetFileMD5(sPath)) <> UCase$(sFile) Then
                        bOk = False
                        Exit Do
                    End If
                End If
            End If
        Else
            bOk = False
            Exit Do
        End If
        iFile = iFile + 1
        sFile = ReadINIStr("Mod", CStr(iFile), ProgramINI)
    Loop
    If Not bOk Then
        TamperingDetected = True
        Call WriteLogEntry("Mod files have been tampered with outside of Launch Base!")
        Call DenyPersistentMods
        Call DeactivateMod(True, False)
    End If
    Init_LoadMods_NoTampering = bOk
End Function

Private Sub DenyPersistentMods()
    If OptPersistentMod Then
        OptPersistentMod = False
        OptPersistentModBad = True
        Call WriteINIStr("Options", "PersistentMod", "no", ProgramINI)
        Call WriteINIStr("Options", "PersistentModBad", "yes", ProgramINI)
        Call WriteLogEntry("Program Options: 'Persistent Mods' disabled by Launch Base.")
    End If
End Sub

Private Sub DenyPersistentPlugins()
    If OptPersistentPlugin Then
        OptPersistentPlugin = False
        OptPersistentPluginBad = True
        Call WriteINIStr("Options", "PersistentPlugin", "no", ProgramINI)
        Call WriteINIStr("Options", "PersistentPluginBad", "yes", ProgramINI)
        Call WriteLogEntry("Program Options: 'Persistent Plugins' disabled by Launch Base.")
    End If
End Sub

Private Sub lstMods_Click(Index As Integer)
    If Not PreventLoop Then
        If Index <> TypePlugin Then
            Call DisplayModDetails(lstMods(Index).ItemData(lstMods(Index).ListIndex), Index)
        Else
            Call DisplayPluginDetails(lstMods(Index).ItemData(lstMods(Index).ListIndex))
        End If
    End If
End Sub

Private Sub lstMods_DblClick(Index As Integer)
    If Index <> TypePlugin Then
        If cmdLaunch(Index).Enabled = True Then Call cmdLaunch_Click(Index)
    End If
End Sub

Private Sub lstMods_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim X As Single
    Dim Y As Single
    Call ListItemCoords(Index, X, Y)
    If KeyCode = 93 Then Call lstMods_RightClick(Index, X, Y)
    'Case 13, 32: Call cmdLaunch_Click(Index) 'this fires automatically if you press enter to switch tabs
End Sub
       
Function ListItemClicked(ByVal X As Single, ByVal Y As Single) As Integer
    Dim curritem As Long
    Dim pt As POINTAPI
    pt.X = X \ Screen.TwipsPerPixelX
    pt.Y = Y \ Screen.TwipsPerPixelY
    Call ClientToScreen(lstMods(menu_rc.Tag).hWnd, pt)
    curritem = LBItemFromPt(lstMods(menu_rc.Tag).hWnd, pt.X, pt.Y, False)
    ListItemClicked = curritem
End Function

Private Sub ListItemCoords(ByVal Index As Integer, ByRef X, ByRef Y)
    Dim ListItemHeight As Long
    ListItemHeight = SendMessage(lstMods(Index).hWnd, LB_GETITEMHEIGHT, 0, ByVal 0&)
    X = lstMods(Index).Left + (64 * 15)
    Y = lstMods(Index).Top + ((ListItemHeight * 15) * lstMods(Index).ListIndex) + ((ListItemHeight * 15) / 2)
End Sub

Private Sub lstMods_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then Call lstMods_RightClick(Index, X, Y)
End Sub

Private Sub lstMods_RightClick(ByVal Index As Integer, ByVal X As Single, ByVal Y As Single)
    Dim iListItem As Integer
    Dim iMod As Integer
    Dim iPlugin As Integer
    Dim iVisible As Integer
    iMod = -1
    If Index > MaxType Then
        menu_rc.Tag = CStr(TypeMod)
        'right-clicked on a banner
        iMod = Val(picMod(Index).Tag)
        X = X + picMod(Index).Left
        Y = Y + picMod(Index).Top
        'reset the index to the correct type
        Index = TypeMod
        'select the item on the other view mode
        iListItem = 0
        Do While iListItem < lstMods(Index).ListCount
            If lstMods(TypeMod).ItemData(iListItem) = iMod Then
                lstMods(Index).ListIndex = iListItem
                Exit Do
            End If
            iListItem = iListItem + 1
        Loop
    Else
        menu_rc.Tag = CStr(Index)
        'right-clicked on a list - figure out which listitem
        iListItem = ListItemClicked(X, Y)
        If iListItem = -1 Then
            'iListItem = lstMods(Index).ListIndex 'if we want the popup menu to appear anyway
            Exit Sub 'if we don't want to display the popup menu
        End If
        X = X + lstMods(Index).Left
        Y = Y + lstMods(Index).Top
        If lstMods(Index).ListCount <> 0 Then
            lstMods(Index).ListIndex = iListItem
            iMod = lstMods(Index).ItemData(lstMods(Index).ListIndex)
        End If
    End If
    If iMod <> -1 Then
        menu_checkforupdate.Visible = True 'at least one menu item must always be visible - this will get hidden later if appropriate
        iVisible = 0
        'delete mod
        If iMod >= HardCodedMods Or Index = TypePlugin Then
            menu_deletemod.Tag = CStr(iMod)
            menu_deletemod.Visible = True
            iVisible = iVisible + 1
            Select Case Index
            Case TypeMod: menu_deletemod.Caption = "Delete Mod"
            Case TypeFA2Mod: menu_deletemod.Caption = "Delete FA2 Mod"
            Case TypeProgram: menu_deletemod.Caption = "Delete Tool"
            Case TypePlugin: menu_deletemod.Caption = "Delete Plugin"
            End Select
        Else
            menu_deletemod.Visible = False
        End If
        'open containing folder
        Select Case Index
        Case TypeMod, TypeProgram
            If iMod <> FA2ModNum Then
                menu_openfolder.Visible = True
                menu_openfolder.Tag = CStr(iMod)
                iVisible = iVisible + 1
            Else
                menu_openfolder.Visible = False
            End If
        Case Else
            'no point allowing opening of TypeFA2 or TypePlugin because no user-generated content
            If CL_dev Then
                menu_openfolder.Visible = True
                menu_openfolder.Tag = CStr(iMod)
                iVisible = iVisible + 1
            Else
                menu_openfolder.Visible = False
            End If
        End Select
        'check for update
        If Index <> TypePlugin Then
            Select Case iMod
            Case RA2ModNum: If Mods(RA2ModNum).ModVersion = "1.006" Then menu_checkforupdate.Visible = False
            Case YRModNum: If Mods(YRModNum).ModVersion = "1.001" Then menu_checkforupdate.Visible = False
            End Select
        End If
        If menu_checkforupdate.Visible Then
            menu_checkforupdate.Tag = CStr(iMod)
            iVisible = iVisible + 1
        Else
            If iVisible <> 0 Then menu_checkforupdate.Visible = False 'must always have at least one visible - doesn't matter which because iVisible keeps track of whether we will show the menu at all
        End If
        'show the menu
        If iVisible <> 0 Then Call PopupMenu(menu_rc, 2, X, Y)
    End If
End Sub

Private Function ModIsInstalled(ByVal iMod As Integer) As Boolean
    Dim bInstalled As Boolean
    Dim sLBU As String
    bInstalled = False
    sLBU = JoinPath(Mods(iMod).ModPath, "launcher\userdata.lbu")
    If FileExists(sLBU) Then
        bInstalled = BooleanStringToBoolean(ReadINIStr("General", "ModIsActive", sLBU, "no"))
    End If
    ModIsInstalled = bInstalled
End Function

Private Sub menu_openfolder_Click()
    Dim iMod As Integer
    iMod = Val(menu_openfolder.Tag)
    Select Case Val(menu_rc.Tag)
    Case TypePlugin: If OpenLocation(Plugins(iMod).PluginPath) < 32 Then Call MsgBox("Failed to open folder: " & Plugins(iMod).PluginPath, vbOKOnly + vbInformation, App.Title)
    Case Else: If OpenLocation(Mods(iMod).ModPath) < 32 Then Call MsgBox("Failed to open folder: " & Mods(iMod).ModPath, vbOKOnly + vbInformation, App.Title)
    End Select
End Sub

Private Sub menu_checkforupdate_Click()
    Dim bShutdown As Boolean
    Call LaunchMod_CheckForUpdate(bShutdown, Val(menu_checkforupdate.Tag), Val(menu_rc.Tag), False, True)
    If bShutdown Then Unload Me
End Sub

Private Sub menu_deletemod_Click()
    Dim mbResult As VbMsgBoxResult
    Dim iMod As Integer
    Dim iModType As Integer
    Dim sMessage As String
    Dim sName As String
    Dim sPath As String
    Dim iUninFile As Long
    Dim sUninFile As String
    Dim sUninstData() As String
    Dim sVersion As String
    Dim sPluginID As String
    Dim bOk As Boolean
    bOk = True
    iMod = Val(menu_deletemod.Tag)
    iModType = Val(menu_rc.Tag)
    'display the message
    Select Case iModType
    Case TypeMod, TypeFA2Mod, TypeProgram
        sVersion = Mods(iMod).ModVersion
        sName = Mods(iMod).ModName
        sPath = Mods(iMod).ModPath
    Case TypePlugin
        sVersion = Plugins(iMod).PluginVersion
        sName = Plugins(iMod).PluginName
        sPath = Plugins(iMod).PluginPath
    End Select
    If Len(sVersion) <> 0 Then sVersion = " [" & sVersion & "]"
    sMessage = "Are you sure that you want to remove " & sName & sVersion & "?" & vbCrLf & "This will completely remove "
    Select Case iModType
    Case TypeMod: sMessage = sMessage & "the mod"
    Case TypeFA2Mod: sMessage = sMessage & "the FA2 mod"
    Case TypeProgram: sMessage = sMessage & "the program"
    Case TypePlugin: sMessage = sMessage & "version " & Plugins(iMod).PluginVersion & " of the plugin"
    End Select
    sMessage = sMessage & " from Launch Base and your computer."
    mbResult = MsgBox(sMessage, vbYesNo + vbQuestion, App.Title)
    'remove the mod
    If mbResult = vbYes Then
        'deactivate first
        Select Case iModType
        Case TypeMod, TypeFA2Mod: If ReadINIStr("Mod", "Name", ProgramINI) = Mods(iMod).ModName Then Call DeactivateMod(True, True)
        Case TypePlugin: If ReadINIStr("Plugin" & Plugins(iMod).PluginID, "Version", ProgramINI) = Plugins(iMod).PluginVersion Then Call DeactivatePlugin(Plugins(iMod).PluginID)
        End Select
        'get uninstall data
        ReDim sUninstData(0)
        If iModType <> TypePlugin Then
            iUninFile = 0
            sUninFile = UCase$(ReadINIStr("Uninstall", CStr(iUninFile), Mods(iMod).ModLiblist))
            Do While Len(sUninFile) <> 0
                iUninFile = iUninFile + 1
                ReDim Preserve sUninstData(iUninFile)
                sUninstData(iUninFile) = sUninFile
                sUninFile = ReadINIStr("Uninstall", CStr(iUninFile), Mods(iMod).ModLiblist)
            Loop
        End If
        'recycle files
        Call ShowPleaseWait("Removing " & Quote(sName & sVersion) & "...", "")
        Call WriteLogEntry("Removing " & Quote(sName & sVersion) & "...")
        If DeleteMod_DeleteFolder(bOk, sPath, "", sUninstData(), iModType) Then
            Call ChDir(App.Path)
            Call UpdatePleaseWait("", GetFileName(sPath))
            Call WriteLogEntry("Moving " & Quote(sPath) & " to the recycle bin.")
            If Not MoveToRecycleBin(sPath, False, , True) Then
                bOk = False
                Call WriteLogEntry("Failed to move " & Quote(sPath) & " to the recycle bin.")
            Else
                Call MsgBox(sName & sVersion & " successfully removed.", vbOKOnly + vbInformation, App.Title)
            End If
        Else
            Call ChDir(App.Path)
            If bOk Then
                Call WriteLogEntry("Failed to remove " & Quote(sPath) & ". One or more residual files still exist.")
                Call MsgBox("One or more residual files were not removed from" & vbCrLf & Quote(sPath) & vbCrLf & "You should review the contents of this folder and remove it yourself.", vbOKOnly + vbInformation, App.Title)
                Call OpenLocation(sPath)
            End If
        End If
        If Not bOk Then
            Call WriteLogEntry("Failed to remove " & Quote(sPath) & ". An error occurred whilst trying to remove a file or subfolder.")
            Call MsgBox("An error occurred whilst trying to remove one or more files from" & vbCrLf & Quote(sPath) & vbCrLf & "You should review the contents of this folder and remove it yourself.", vbOKOnly + vbExclamation, App.Title)
            Call OpenLocation(sPath)
        End If
        'tidy up Launch Base
        Call UpdatePleaseWait("Tidying up Launch Base...")
        If iModType <> TypePlugin Then
            Mods(iMod).ModName = ""
            Call Init_LoadMods_FillModLists
            'clear from modcat
            iUninFile = 1
            Do While iUninFile <= UpdateRecordCount
                If UpdateRecords(iUninFile).CheckModNum = iMod Then
                    If UpdateRecords(iUninFile).ModType = Mods(iMod).ModType Then
                        Call frmModCat.InitialiseRecord(iUninFile)
                        Exit Do
                    End If
                End If
                iUninFile = iUninFile + 1
            Loop
        Else
            Plugins(iMod).PluginName = ""
            sPluginID = Plugins(iMod).PluginID
            Plugins(iMod).PluginID = ""
            Call DisplayPluginDetails(-1)
            If GetLatestPlugin(sPluginID) = -1 Then Call lstMods(TypePlugin).RemoveItem(lstMods(TypePlugin).ListIndex)
            'clear from modcat
            iUninFile = 1
            Do While iUninFile <= UpdateRecordCount
                If UpdateRecords(iUninFile).CheckModNum = iMod Then
                    If UpdateRecords(iUninFile).ModPluginID = Plugins(iMod).PluginID Then
                        Call frmModCat.InitialiseRecord(iUninFile)
                        Exit Do
                    End If
                End If
                iUninFile = iUninFile + 1
            Loop
        End If
        Call HidePleaseWait
    End If
End Sub

Private Sub menu_About_Click()
    Call frmAbout.Show(vbModal)
End Sub

Private Sub menu_disclaimer_Click()
    Call frmDisclaimer.Show(vbModal)
End Sub

Private Sub menu_exit_Click()
    Call Shutdown
End Sub

Private Sub menu_helptopics_Click()
    frmHelp.Show
End Sub

Private Sub menu_livelog_Click()
    If menu_livelog.Checked = False Then
        menu_livelog.Checked = True
        Call frmLiveLog.Show
    Else
        menu_livelog.Checked = False
        Call frmLiveLog.Hide
    End If
End Sub

Private Sub menu_modcat_Click()
    frmMain.Hide
    Call frmModCat.Show
    Call frmModCat.Refresh
    Call frmModCat.UpdateUpdateCat
End Sub

Private Sub menu_history_Click()
    frmMain.Enabled = False
    Call frmHistory.Show
    Call frmHistory.LoadHistory 'can't be run as part of load
End Sub

Private Sub menu_fileman_Click()
    frmMain.Enabled = False
    Call frmFileMan.Show
End Sub

Private Sub menu_options_Click()
    Call frmOptions.Show(vbModal)
End Sub

Private Sub menu_aresoptions_Click()
    frmMain.Enabled = False
    Call frmOptionsAres.Show
End Sub

Private Sub menu_aresini_Click()
    Call frmAresINI.Show(vbModal)
End Sub

Private Sub menu_skin_Click(Index As Integer)
    Call LoadSkin(Index)
End Sub

Private Sub menu_tab_Click(Index As Integer)
    Call SelectTab(Index)
End Sub

Private Sub menu_usertool_Click(Index As Integer)
    Dim process_id
    Call WriteLogEntry("Launching " & Mods(UserToolModNum(Index)).ModName, LogLevel1)
    process_id = Shell(Mods(UserToolModNum(Index)).ModProgram, vbNormalFocus)
End Sub

Private Sub Init_URLs()
    Dim iCounter As Integer
    Dim iCommaPos As Integer
    Dim sWebsite As String
    iCounter = 0
    sWebsite = ReadINIStr("URL", CStr(iCounter), ProgramINI)
    Do While Len(sWebsite) <> 0
        iCommaPos = InStr(1, sWebsite, ",")
        If iCommaPos <> 0 And iCommaPos > 1 And iCommaPos < Len(sWebsite) Then
            If iCounter <> 0 Then
                Load menu_website(iCounter)
            Else
                menu_help_line1.Visible = True
            End If
            menu_website(iCounter).Caption = Left$(sWebsite, iCommaPos - 1)
            menu_website(iCounter).Tag = Mid$(sWebsite, iCommaPos + 1)
            menu_website(iCounter).Visible = True
        End If
        iCounter = iCounter + 1
        sWebsite = ReadINIStr("URL", CStr(iCounter), ProgramINI)
    Loop
End Sub

Private Sub menu_website_Click(Index As Integer)
    If OpenLocation(menu_website(Index).Tag) < 32 Then Call MsgBox("Unable to open " & Quote(menu_website(Index).Tag) & ".", vbOKOnly + vbInformation, App.Title)
End Sub

Private Sub scrollbarBanners_Change()
    Call ScrollBanners(scrollbarBanners.Value)
End Sub

Private Sub skinTabBody_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim iCounter As Integer
    If Index = 4 Then
        For iCounter = (MaxType + 1) To (MaxType + BannerCount)
            lineBannerTop.Visible = False
            lineBannerBottom.Visible = False
            lineBannerLeft.Visible = False
            lineBannerRight.Visible = False
        Next iCounter
    Else
        lblModWebsite(Index).ForeColor = ColorURL
    End If
End Sub

Private Sub txtFA2Folder_LostFocus()
    Call FA2Check(False)
End Sub

Private Sub LoadSkin(ByVal iSkinNum As String)
    Dim iCounter As Integer
    Dim ColorBannerBorder As Long
    Dim sTemp As String
    Select Case iSkinNum
    Case -1
        SelectedSkinPath = JoinPath(JoinPath(EXEDIR, "Skins"), ReadINIStr("General", "SelectedSkin", ProgramINI, DefaultSkinDir))
        If Not FileExists(JoinPath(SelectedSkinPath, "skin.ini")) Then SelectedSkinPath = JoinPath(JoinPath(EXEDIR, "Skins"), DefaultSkinDir)
        If Not FileExists(JoinPath(SelectedSkinPath, "skin.ini")) Then SelectedSkinPath = SkinPath(0)
    Case Else
        SelectedSkinPath = SkinPath(iSkinNum)
    End Select
    'ColorGood = &H4000&
    'ColorBad = &H80&
    'ColorURL = &HC00000
    'ColorURLActive = &HFF&
    sTemp = ReadINIStr("Colors", "TextNeutral", JoinPath(SelectedSkinPath, "skin.ini"), "000000")
    ColorNeutral = RGB(HexToDec(Mid$(sTemp, 1, 2)), HexToDec(Mid$(sTemp, 3, 2)), HexToDec(Mid$(sTemp, 5, 2)))
    sTemp = ReadINIStr("Colors", "TextGood", JoinPath(SelectedSkinPath, "skin.ini"), "009900")
    ColorGood = RGB(HexToDec(Mid$(sTemp, 1, 2)), HexToDec(Mid$(sTemp, 3, 2)), HexToDec(Mid$(sTemp, 5, 2)))
    sTemp = ReadINIStr("Colors", "TextBad", JoinPath(SelectedSkinPath, "skin.ini"), "CC0000")
    ColorBad = RGB(HexToDec(Mid$(sTemp, 1, 2)), HexToDec(Mid$(sTemp, 3, 2)), HexToDec(Mid$(sTemp, 5, 2)))
    sTemp = ReadINIStr("Colors", "TextURL", JoinPath(SelectedSkinPath, "skin.ini"), "0000FF")
    ColorURL = RGB(HexToDec(Mid$(sTemp, 1, 2)), HexToDec(Mid$(sTemp, 3, 2)), HexToDec(Mid$(sTemp, 5, 2)))
    sTemp = ReadINIStr("Colors", "TextURLActive", JoinPath(SelectedSkinPath, "skin.ini"), "FF0000")
    ColorURLActive = RGB(HexToDec(Mid$(sTemp, 1, 2)), HexToDec(Mid$(sTemp, 3, 2)), HexToDec(Mid$(sTemp, 5, 2)))
    sTemp = ReadINIStr("Colors", "TextList", JoinPath(SelectedSkinPath, "skin.ini"), "000000")
    ColorListText = RGB(HexToDec(Mid$(sTemp, 1, 2)), HexToDec(Mid$(sTemp, 3, 2)), HexToDec(Mid$(sTemp, 5, 2)))
    sTemp = ReadINIStr("Colors", "List", JoinPath(SelectedSkinPath, "skin.ini"), "FFFFFF")
    ColorList = RGB(HexToDec(Mid$(sTemp, 1, 2)), HexToDec(Mid$(sTemp, 3, 2)), HexToDec(Mid$(sTemp, 5, 2)))
    sTemp = ReadINIStr("Colors", "BannerBorder", JoinPath(SelectedSkinPath, "skin.ini"), "")
    If Len(sTemp) <> 0 Then
        ColorBannerBorder = RGB(HexToDec(Mid$(sTemp, 1, 2)), HexToDec(Mid$(sTemp, 3, 2)), HexToDec(Mid$(sTemp, 5, 2)))
    Else
        ColorBannerBorder = ColorURLActive
    End If
    lineBannerTop.BorderColor = ColorBannerBorder
    lineBannerBottom.BorderColor = ColorBannerBorder
    lineBannerLeft.BorderColor = ColorBannerBorder
    lineBannerRight.BorderColor = ColorBannerBorder
    For iCounter = 0 To MaxType
        'colors
        lblModDate(iCounter).ForeColor = ColorNeutral
        lblModAuthor(iCounter).ForeColor = ColorNeutral
        lblModSize(iCounter).ForeColor = ColorNeutral
        lblModDescription(iCounter).ForeColor = ColorNeutral
        lblModWebsite(iCounter).ForeColor = ColorURL
        Select Case iCounter
        Case TypeMod, TypeFA2Mod
            Select Case lblModTX(iCounter).Tag
            Case "GOOD": lblModTX(iCounter).ForeColor = ColorGood
            Case "BAD": lblModTX(iCounter).ForeColor = ColorBad
            Case Else: lblModTX(iCounter).ForeColor = ColorNeutral
            End Select
        End Select
        lstMods(iCounter).ForeColor = ColorListText
        lstMods(iCounter).BackColor = ColorList
        picMod(iCounter).BackColor = ColorList
        lblModVersion(iCounter).ForeColor = ColorNeutral
        'tab image
        Set skinTabBody(iCounter).Picture = Nothing
        sTemp = JoinPath(SelectedSkinPath, "tab" & CStr(iCounter) & ".bmp")
        If FileExists(sTemp) Then Set skinTabBody(iCounter).Picture = LoadPicture(sTemp)
        'launch button
        Set cmdLaunch(iCounter).Picture = Nothing
        sTemp = "btnb" & CStr(iCounter)
        Select Case cmdLaunch(iCounter).Tag
        Case "RA2": sTemp = sTemp & "r.bmp"
        Case "YR": sTemp = sTemp & "y.bmp"
        Case Else: If iCounter <> 0 Then sTemp = sTemp & ".bmp" Else sTemp = sTemp & "r.bmp"
        End Select
        If FileExists(JoinPath(SelectedSkinPath, sTemp)) Then
            cmdLaunch(iCounter).Picture = LoadPicture(JoinPath(SelectedSkinPath, sTemp))
        Else
            cmdLaunch(iCounter).Picture = LoadPicture(JoinPath(RESDIR, sTemp))
        End If
        'manual button
        Set cmdManual(iCounter).Picture = Nothing
        sTemp = "btna" & CStr(iCounter) & ".bmp"
        If FileExists(JoinPath(SelectedSkinPath, sTemp)) Then
            cmdManual(iCounter).Picture = LoadPicture(JoinPath(SelectedSkinPath, sTemp))
        Else
            cmdManual(iCounter).Picture = LoadPicture(JoinPath(RESDIR, sTemp))
        End If
    Next iCounter
    'misc labels
    For iCounter = 0 To (lblGeneral.Count - 1)
        lblGeneral(iCounter).ForeColor = ColorNeutral
    Next iCounter
    'TypeMod only
    lblModCampaigns.ForeColor = ColorNeutral
    lblModUsesAres.ForeColor = ColorNeutral
    Set skinTabBody(4).Picture = Nothing
    sTemp = JoinPath(SelectedSkinPath, "tab4" & ".bmp")
    If FileExists(sTemp) Then Set skinTabBody(4).Picture = LoadPicture(sTemp)
    'TypePlugin only
    iCounter = MaxType + 1
    Set cmdLaunch(iCounter).Picture = Nothing
    sTemp = "btnb" & CStr(iCounter)
    Select Case cmdLaunch(iCounter).Tag
    Case "RA2": sTemp = sTemp & "r.bmp"
    Case "YR": sTemp = sTemp & "y.bmp"
    Case Else: sTemp = sTemp & ".bmp"
    End Select
    If FileExists(JoinPath(SelectedSkinPath, sTemp)) Then
        cmdLaunch(iCounter).Picture = LoadPicture(JoinPath(SelectedSkinPath, sTemp))
    Else
        cmdLaunch(iCounter).Picture = LoadPicture(JoinPath(RESDIR, sTemp))
    End If
    'TypeFA2Mod only
    lblTX.ForeColor = ColorNeutral
    txtFA2Folder.BackColor = ColorList
    txtFA2Folder.ForeColor = ColorListText
    Select Case lblModFA2.Tag
    Case "GOOD": lblModFA2.ForeColor = ColorGood
    Case "BAD": lblModFA2.ForeColor = ColorBad
    Case Else: lblModFA2.ForeColor = ColorNeutral
    End Select
    'TypeProgram only
    lblModParams.ForeColor = ColorNeutral
    txtModParams.BackColor = ColorList
    txtModParams.ForeColor = ColorListText
    'For Counter = 0 To MaxType
    '    Set cmdLaunch(Counter).DisabledPicture = Nothing
    '    Set cmdLaunch(Counter).DownPicture = Nothing
    '    Set cmdLaunch(Counter).Picture = Nothing
    '    If FileExists(JoinPath(LoadPath, "tab" & CStr(Counter) & "bla.bmp")) And _
    '      FileExists(JoinPath(LoadPath, "tab" & CStr(Counter) & "blb.bmp")) And _
    '      FileExists(JoinPath(LoadPath, "tab" & CStr(Counter) & "blc.bmp")) Then
    '        Set cmdLaunch(Counter).DisabledPicture = LoadPicture(JoinPath(LoadPath, "tab" & CStr(Counter) & "blc.bmp"))
    '        Set cmdLaunch(Counter).DownPicture = LoadPicture(JoinPath(LoadPath, "tab" & CStr(Counter) & "blb.bmp"))
    '        Set cmdLaunch(Counter).Picture = LoadPicture(JoinPath(LoadPath, "tab" & CStr(Counter) & "bla.bmp"))
    '        cmdLaunch(Counter).Caption = ""
    '    Else
    '        cmdLaunch(Counter).Caption = cmdLaunch(Counter).Tag
    '    End If
    'Next Counter
    'For Counter = 0 To MaxType
    '    Set cmdManual(Counter).DisabledPicture = Nothing
    '    Set cmdManual(Counter).DownPicture = Nothing
    '    Set cmdManual(Counter).Picture = Nothing
    '    If FileExists(JoinPath(LoadPath, "tab" & CStr(Counter) & "bma.bmp")) And _
    '      FileExists(JoinPath(LoadPath, "tab" & CStr(Counter) & "bmb.bmp")) And _
    '      FileExists(JoinPath(LoadPath, "tab" & CStr(Counter) & "bmc.bmp")) Then
    '        Set cmdManual(Counter).DisabledPicture = LoadPicture(JoinPath(LoadPath, "tab" & CStr(Counter) & "bmc.bmp"))
    '        Set cmdManual(Counter).DownPicture = LoadPicture(JoinPath(LoadPath, "tab" & CStr(Counter) & "bmb.bmp"))
    '        Set cmdManual(Counter).Picture = LoadPicture(JoinPath(LoadPath, "tab" & CStr(Counter) & "bma.bmp"))
    '        cmdManual(Counter).Caption = ""
    '    Else
    '        cmdManual(Counter).Caption = cmdManual(Counter).Tag
    '    End If
    'Next Counter
    'update record of what skin is selected
    For iCounter = 0 To (menu_skin.Count - 1)
        If UCase$(GetFileName(SkinPath(iCounter))) = UCase$(GetFileName(SelectedSkinPath)) Then
            menu_skin(iCounter).Checked = True
            Call WriteINIStr("General", "SelectedSkin", GetFileName(SelectedSkinPath), ProgramINI)
        Else
            menu_skin(iCounter).Checked = False
        End If
    Next iCounter
End Sub

Private Sub SelectTab(ByVal Index As Integer)
    Dim iTab As Integer
    For iTab = 0 To 4
        If iTab <> Index Then skinTabBody(iTab).Visible = False
    Next iTab
    Select Case Index
    Case 0, 4: frmMain.Caption = App.Title & ": Mods"
    Case 1: frmMain.Caption = App.Title & ": Plugins"
    Case 2: frmMain.Caption = App.Title & ": FA2 Mods"
    Case 3: frmMain.Caption = App.Title & ": Tools"
    End Select
    skinTabBody(Index).Visible = True
    Select Case SelectedTab
    Case 0: BannerTab = False
    Case 4: BannerTab = True
    End Select
    SelectedTab = Index
End Sub

Private Sub txtModParams_LostFocus()
    Dim iMod As Integer
    iMod = Val(lstMods(TypeProgram).ItemData(lstMods(TypeProgram).ListIndex))
    If txtModParams.Text <> Mods(iMod).ModParams Then
        Mods(iMod).ModParams = txtModParams.Text
        Call WriteINIStr("General", "Params", Mods(iMod).ModParams, JoinPath(Mods(iMod).ModPath, "launcher\userdata.lbu"))
    End If
End Sub


