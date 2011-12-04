VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base Mod Creator"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   7920
   ForeColor       =   &H80000008&
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox filelistbox 
      Height          =   1065
      Left            =   0
      TabIndex        =   117
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.DirListBox dirlistbox 
      Height          =   1215
      Left            =   0
      TabIndex        =   116
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Options"
      TabPicture(0)   =   "main.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameOptions1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frameTX"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frameSecurity"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frameFA2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frameModSounds"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frameGameMode"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Files"
      TabPicture(1)   =   "main.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdUninstFiles"
      Tab(1).Control(1)=   "txtCustomScriptMessage"
      Tab(1).Control(2)=   "cmdBrowseCustomScript"
      Tab(1).Control(3)=   "cmdBrowseProgramDir"
      Tab(1).Control(4)=   "txtCustomScript"
      Tab(1).Control(5)=   "cmdFileErrors"
      Tab(1).Control(6)=   "txtProgramDirectory"
      Tab(1).Control(7)=   "frameMixes"
      Tab(1).Control(8)=   "listFiles"
      Tab(1).Control(9)=   "treeFolders"
      Tab(1).Control(10)=   "lblCustomScriptMessage"
      Tab(1).Control(11)=   "lblCustomScript"
      Tab(1).Control(12)=   "lblFolders"
      Tab(1).Control(13)=   "lblProgramDirectory"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Installer"
      TabPicture(2)   =   "main.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frameInstaller5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "frameInstaller4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "frameInstaller2"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "frameInstaller3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "frameInstaller1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Compression"
      TabPicture(3)   =   "main.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).Control(1)=   "framePatchOptions"
      Tab(3).Control(2)=   "frameUpdateOnly"
      Tab(3).Control(3)=   "frameCompression"
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame1 
         Caption         =   "FLAC"
         Height          =   615
         Left            =   -74760
         TabIndex        =   150
         Top             =   1320
         Width           =   7215
         Begin VB.CheckBox cboxFLAC 
            Caption         =   "Convert WAV files to FLAC before generating installer (Launch Base will convert them back)"
            Height          =   255
            Left            =   240
            TabIndex        =   151
            Top             =   240
            Value           =   1  'Checked
            Width           =   6855
         End
      End
      Begin VB.CommandButton cmdUninstFiles 
         Caption         =   "Edit Uninstall Files List"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -72600
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   6000
         Width           =   2415
      End
      Begin VB.TextBox txtCustomScriptMessage 
         Enabled         =   0   'False
         Height          =   495
         Left            =   -70680
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   146
         Top             =   7200
         Width           =   3135
      End
      Begin VB.CommandButton cmdBrowseCustomScript 
         Caption         =   "..."
         Height          =   275
         Left            =   -67800
         TabIndex        =   145
         Top             =   6855
         Width           =   275
      End
      Begin VB.CommandButton cmdBrowseProgramDir 
         Caption         =   "..."
         Height          =   275
         Left            =   -67800
         TabIndex        =   144
         Top             =   6495
         Width           =   275
      End
      Begin VB.Frame frameGameMode 
         Caption         =   "Default Mode/Map"
         Height          =   975
         Left            =   5520
         TabIndex        =   137
         Top             =   3720
         Width           =   1935
         Begin VB.ComboBox comboGameMode 
            Height          =   315
            ItemData        =   "main.frx":093A
            Left            =   120
            List            =   "main.frx":093C
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtMapIndex 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            MaxLength       =   5
            TabIndex        =   34
            Text            =   "0"
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox txtGameMode 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Text            =   "1"
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblMapIndex 
            Caption         =   "Map:"
            Height          =   255
            Left            =   1440
            TabIndex        =   139
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblGameMode 
            Caption         =   "Game Mode:"
            Height          =   255
            Left            =   120
            TabIndex        =   138
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtCustomScript 
         Height          =   285
         Left            =   -70680
         TabIndex        =   135
         Top             =   6840
         Width           =   2895
      End
      Begin VB.CommandButton cmdFileErrors 
         Caption         =   "View File Errors"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70080
         Style           =   1  'Graphical
         TabIndex        =   134
         Top             =   6000
         Width           =   2415
      End
      Begin VB.TextBox txtProgramDirectory 
         Height          =   285
         Left            =   -70680
         TabIndex        =   130
         Top             =   6480
         Width           =   2895
      End
      Begin VB.Frame frameMixes 
         Caption         =   "File Options"
         Height          =   1695
         Left            =   -74760
         TabIndex        =   125
         Top             =   6000
         Width           =   2055
         Begin VB.CheckBox cboxUseAres 
            Caption         =   "Use Official Ares DLL"
            Height          =   255
            Left            =   120
            TabIndex        =   149
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox comboMixEncrypt 
            Height          =   315
            ItemData        =   "main.frx":093E
            Left            =   120
            List            =   "main.frx":094E
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   720
            Width           =   1815
         End
         Begin VB.ComboBox comboSide3Mix 
            Height          =   315
            ItemData        =   "main.frx":097F
            Left            =   120
            List            =   "main.frx":0989
            Style           =   2  'Dropdown List
            TabIndex        =   126
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label5 
            Caption         =   "Side 3 MIX File:"
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "MIX File Format:"
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame framePatchOptions 
         Caption         =   "Patching Options"
         Enabled         =   0   'False
         Height          =   2655
         Left            =   -74760
         TabIndex        =   107
         Top             =   2040
         Width           =   7215
         Begin VB.Frame framePatchBlockSize 
            Caption         =   "Patch Block Size"
            Enabled         =   0   'False
            Height          =   1335
            Left            =   3720
            TabIndex        =   110
            Top             =   1080
            Width           =   3255
            Begin MSComctlLib.Slider sliderPatchBlockSize 
               Height          =   255
               Left            =   120
               TabIndex        =   59
               Top             =   600
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               Max             =   8
               SelStart        =   2
               Value           =   2
            End
            Begin VB.Label lblPatchBlockSizeDef 
               Alignment       =   2  'Center
               Caption         =   "Default |"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   675
               TabIndex        =   115
               Top             =   360
               Width           =   495
            End
            Begin VB.Label lblPatchDesc2 
               Alignment       =   2  'Center
               Caption         =   "Lower = smaller patches, slower installer."
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   114
               Top             =   960
               Width           =   3015
            End
            Begin VB.Label lblPatchBlockSize 
               Alignment       =   2  'Center
               Caption         =   "64"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   113
               Top             =   360
               Width           =   3015
            End
         End
         Begin VB.Frame framePatchMinSize 
            Caption         =   "Minimum File Size Before Patch"
            Enabled         =   0   'False
            Height          =   1335
            Left            =   240
            TabIndex        =   109
            Top             =   1080
            Width           =   3255
            Begin MSComctlLib.Slider sliderPatchMinSize 
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   600
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   450
               _Version        =   393216
               Enabled         =   0   'False
               Max             =   64
               SelStart        =   16
               TickStyle       =   3
               Value           =   16
               TextPosition    =   1
            End
            Begin VB.Label lblPatchMinSize 
               Alignment       =   2  'Center
               Caption         =   "256 KB"
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   112
               Top             =   360
               Width           =   3015
            End
            Begin VB.Label lblPatchDesc1 
               Alignment       =   2  'Center
               Caption         =   "Files smaller than this will not be patched."
               Enabled         =   0   'False
               Height          =   255
               Left            =   120
               TabIndex        =   111
               Top             =   960
               Width           =   3015
            End
         End
         Begin VB.TextBox txtMarblePath 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   56
            Top             =   600
            Width           =   6495
         End
         Begin VB.CommandButton cmdBrowseMarble 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   275
            Left            =   6720
            TabIndex        =   57
            Top             =   600
            Width           =   275
         End
         Begin VB.Label lblCheckMarble 
            Alignment       =   1  'Right Justify
            Caption         =   "Checking <marble.mix> - please wait..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   4080
            TabIndex        =   118
            Top             =   420
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label lblPatchMarble 
            Caption         =   "Unmodified <marble.mix> (if applicable):"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   108
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame frameModSounds 
         Caption         =   "Launch Base Sounds"
         Height          =   1335
         Left            =   240
         TabIndex        =   100
         Top             =   6360
         Width           =   7215
         Begin VB.PictureBox pboxModLaunchSound 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   705
            Picture         =   "main.frx":09A9
            ScaleHeight     =   330
            ScaleMode       =   0  'User
            ScaleWidth      =   360
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   810
            Width           =   360
         End
         Begin VB.PictureBox pboxModDisplaySound 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   705
            Picture         =   "main.frx":0B33
            ScaleHeight     =   330
            ScaleMode       =   0  'User
            ScaleWidth      =   360
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   330
            Width           =   360
         End
         Begin VB.CommandButton cmdBrowseModLaunchSound 
            Caption         =   "..."
            Height          =   275
            Left            =   6840
            TabIndex        =   25
            Top             =   840
            Width           =   275
         End
         Begin VB.TextBox txtModLaunchSound 
            Height          =   285
            Left            =   1080
            MaxLength       =   252
            TabIndex        =   24
            Top             =   840
            Width           =   5775
         End
         Begin VB.CommandButton cmdBrowseModDisplaySound 
            Caption         =   "..."
            Height          =   275
            Left            =   6840
            TabIndex        =   22
            Top             =   360
            Width           =   275
         End
         Begin VB.TextBox txtModDisplaySound 
            Height          =   285
            Left            =   1080
            MaxLength       =   252
            TabIndex        =   21
            Top             =   360
            Width           =   5775
         End
         Begin VB.Label lblLaunchSound 
            Caption         =   "Launch Sound:"
            Height          =   375
            Left            =   120
            TabIndex        =   102
            Top             =   780
            Width           =   615
         End
         Begin VB.Label lblDisplaySound 
            Caption         =   "Display Sound:"
            Height          =   375
            Left            =   120
            TabIndex        =   101
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame frameInstaller1 
         Caption         =   "Installer Icon"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   96
         Top             =   480
         Width           =   1455
         Begin VB.CheckBox cboxWindowIcon 
            Caption         =   "Window Icon"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   960
            Value           =   1  'Checked
            Width           =   1290
         End
         Begin VB.PictureBox picInstallerIcon 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   480
            ScaleHeight     =   480
            ScaleMode       =   0  'User
            ScaleWidth      =   480
            TabIndex        =   37
            Top             =   360
            Width           =   480
         End
      End
      Begin VB.Frame frameInstaller3 
         Caption         =   "Other Installer Settings"
         Height          =   1935
         Left            =   -74760
         TabIndex        =   95
         Top             =   1920
         Width           =   2655
         Begin VB.CheckBox cboxResetGameConfig 
            Caption         =   "Reset Game Configuration"
            Height          =   255
            Left            =   240
            TabIndex        =   140
            Top             =   1440
            Width           =   2295
         End
         Begin VB.CheckBox cboxCRC 
            Caption         =   "Perform CRC Check"
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   480
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox cboxXPStyle 
            Caption         =   "XP Style"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   960
            Width           =   2295
         End
         Begin VB.CheckBox cboxAutoClose 
            Caption         =   "Close Installer Automatically"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   720
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.CheckBox cboxOldSaves 
            Caption         =   "Move 'saves' to 'saves\old'"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   1200
            Value           =   1  'Checked
            Width           =   2295
         End
      End
      Begin VB.Frame frameInstaller2 
         Caption         =   "Installation Directory"
         Height          =   1335
         Left            =   -73080
         TabIndex        =   94
         Top             =   480
         Width           =   5535
         Begin VB.ComboBox comboDSP 
            Height          =   315
            ItemData        =   "main.frx":0CBD
            Left            =   240
            List            =   "main.frx":0CCA
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox txtINSTDIR 
            Height          =   285
            Left            =   1320
            MaxLength       =   64
            TabIndex        =   39
            Text            =   "Just another YR mod"
            Top             =   360
            Width           =   3975
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Default folder:"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   390
            Width           =   1095
         End
      End
      Begin VB.Frame frameInstaller4 
         Caption         =   "Log Window"
         Height          =   1935
         Left            =   -72000
         TabIndex        =   87
         Top             =   1920
         Width           =   4455
         Begin VB.CheckBox cboxShowInstDetails 
            Caption         =   "Show Log Window"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txtBGColour 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   48
            Text            =   "FF"
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox txtTextColour 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   45
            Text            =   "00"
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtTextColour 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   46
            Text            =   "00"
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtTextColour 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   47
            Text            =   "00"
            Top             =   960
            Width           =   375
         End
         Begin VB.TextBox txtBGColour 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   1
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   49
            Text            =   "FF"
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox txtBGColour 
            Alignment       =   2  'Center
            Height          =   285
            Index           =   2
            Left            =   1920
            MaxLength       =   2
            TabIndex        =   50
            Text            =   "FF"
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label lblLogColours 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "This is how the log window will look."
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   2400
            TabIndex        =   93
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "R"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1200
            TabIndex        =   92
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "G"
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   1560
            TabIndex        =   91
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "B"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1920
            TabIndex        =   90
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Background:"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   1350
            Width           =   975
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Text:"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   990
            Width           =   975
         End
      End
      Begin VB.Frame frameFA2 
         Caption         =   "FinalAlert 2 YR"
         Height          =   975
         Left            =   5520
         TabIndex        =   86
         Top             =   2640
         Width           =   1935
         Begin VB.ComboBox comboFA2 
            Height          =   315
            ItemData        =   "main.frx":0D60
            Left            =   120
            List            =   "main.frx":0D6D
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lblFA2Version 
            Caption         =   "Required version:"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame frameUpdateOnly 
         Caption         =   "Update-Only Installer"
         Height          =   2895
         Left            =   -74760
         TabIndex        =   80
         Top             =   4800
         Width           =   7215
         Begin VB.CheckBox cboxUpdateOnly 
            Caption         =   "Create Update-Only Installer"
            Height          =   255
            Left            =   360
            TabIndex        =   60
            Top             =   1200
            Width           =   2415
         End
         Begin VB.CommandButton cmdUpdateOnlyDestBrowse 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   275
            Left            =   6720
            TabIndex        =   64
            Top             =   2400
            Width           =   275
         End
         Begin VB.CommandButton cmdUpdateOnlySourceBrowse 
            Caption         =   "..."
            Enabled         =   0   'False
            Height          =   275
            Left            =   6720
            TabIndex        =   62
            Top             =   1800
            Width           =   275
         End
         Begin VB.TextBox txtUpdateOnlySource 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   61
            Top             =   1800
            Width           =   6495
         End
         Begin VB.TextBox txtUpdateOnlyDest 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   63
            Top             =   2400
            Width           =   6495
         End
         Begin VB.Label lblUpdateOnlyDesc2 
            Caption         =   $"main.frx":0D90
            Height          =   975
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   6975
         End
         Begin VB.Label lblUpdateOnlyDesc4 
            Caption         =   "Latest Installation:"
            Height          =   255
            Left            =   240
            TabIndex        =   82
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label lblUpdateOnlyDesc3 
            Caption         =   "Previous Installation:"
            Height          =   255
            Left            =   240
            TabIndex        =   81
            Top             =   1560
            Width           =   1575
         End
      End
      Begin VB.Frame frameCompression 
         Caption         =   "Compression Method"
         Height          =   735
         Left            =   -74760
         TabIndex        =   79
         Top             =   480
         Width           =   7215
         Begin VB.CheckBox cboxGenPat 
            Caption         =   "Generate Patches When Possible"
            Height          =   255
            Left            =   4320
            TabIndex        =   55
            Top             =   315
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.ComboBox comboCompressionMethod 
            Height          =   315
            ItemData        =   "main.frx":0F1F
            Left            =   120
            List            =   "main.frx":0F2F
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   300
            Width           =   2055
         End
         Begin VB.CheckBox cboxSolid 
            Caption         =   "Solid Compression"
            Height          =   255
            Left            =   2400
            TabIndex        =   54
            Top             =   315
            Value           =   1  'Checked
            Width           =   1695
         End
      End
      Begin VB.Frame frameSecurity 
         Caption         =   "Plugin Security"
         Height          =   1455
         Left            =   5520
         TabIndex        =   78
         Top             =   4800
         Width           =   1935
         Begin VB.CommandButton cmdKey 
            Caption         =   "Import security.key"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton cmdLock 
            Caption         =   "Export security.lock"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame frameTX 
         Caption         =   "Terrain Expansion"
         Height          =   2055
         Left            =   5520
         TabIndex        =   67
         Top             =   480
         Width           =   1935
         Begin VB.OptionButton optTX 
            Caption         =   "Plugin *is* TX"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton optTX 
            Caption         =   "Not Allowed"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optTX 
            Caption         =   "Required"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton optTX 
            Caption         =   "Not Required"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.TextBox txtTX 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            MaxLength       =   255
            TabIndex        =   30
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lblMinVersionTX 
            Caption         =   "Minimum Version:"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1320
            Width           =   1695
         End
      End
      Begin VB.Frame frameOptions1 
         Caption         =   "Mod Details"
         Height          =   5775
         Left            =   240
         TabIndex        =   69
         Top             =   480
         Width           =   5175
         Begin VB.ComboBox comboPluginID 
            Height          =   315
            ItemData        =   "main.frx":0F64
            Left            =   4080
            List            =   "main.frx":0F77
            TabIndex        =   142
            Top             =   4590
            Width           =   975
         End
         Begin VB.CheckBox cboxShutdownLB 
            Caption         =   "Close Launch Base"
            Height          =   375
            Left            =   3360
            TabIndex        =   141
            Top             =   4560
            Width           =   1680
         End
         Begin VB.CheckBox cboxForRA2 
            Caption         =   "Red Alert 2"
            Height          =   375
            Left            =   3360
            TabIndex        =   15
            Top             =   4560
            Width           =   1695
         End
         Begin VB.TextBox txtModSnapFormat 
            Height          =   285
            Left            =   3600
            MaxLength       =   254
            TabIndex        =   19
            Text            =   "Map%04d.yrm"
            Top             =   5370
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtModScrnFormat 
            Height          =   285
            Left            =   1080
            MaxLength       =   254
            TabIndex        =   18
            Text            =   "SCRN%04d.pcx"
            Top             =   5370
            Width           =   1455
         End
         Begin VB.ComboBox comboModType 
            Height          =   315
            ItemData        =   "main.frx":0FA1
            Left            =   1080
            List            =   "main.frx":0FB1
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   4590
            Width           =   2175
         End
         Begin VB.ComboBox comboModCampaigns 
            Height          =   315
            ItemData        =   "main.frx":0FDA
            Left            =   1080
            List            =   "main.frx":0FEA
            TabIndex        =   13
            Text            =   "Broken Campaigns"
            Top             =   4200
            Width           =   3975
         End
         Begin VB.ComboBox comboModProgram 
            Height          =   315
            ItemData        =   "main.frx":1037
            Left            =   1080
            List            =   "main.frx":1039
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   4980
            Width           =   2175
         End
         Begin VB.CheckBox cboxParams 
            Caption         =   "Optional Paramaters"
            Height          =   375
            Left            =   3360
            TabIndex        =   17
            Top             =   4950
            Width           =   1725
         End
         Begin VB.TextBox txtModUpdateCheck 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   12
            Text            =   "http://www.website.com/mod/mod.upd"
            Top             =   3840
            Width           =   3975
         End
         Begin VB.TextBox txtModWebsite 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   11
            Text            =   "http://www.website.com"
            Top             =   3480
            Width           =   3975
         End
         Begin VB.TextBox txtModAuthor 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   3
            Text            =   "Anonymous"
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox txtModVersion 
            Height          =   285
            Left            =   2280
            MaxLength       =   254
            TabIndex        =   5
            Text            =   "1.0"
            Top             =   1800
            Width           =   2775
         End
         Begin VB.TextBox txtModName 
            Height          =   285
            Left            =   1080
            MaxLength       =   255
            TabIndex        =   2
            Text            =   "Just another YR mod"
            Top             =   1080
            Width           =   3975
         End
         Begin VB.CheckBox cboxModVersion 
            Caption         =   "Automatic"
            Height          =   255
            Left            =   1080
            TabIndex        =   4
            Top             =   1815
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.TextBox txtModDescription 
            Appearance      =   0  'Flat
            Height          =   855
            Left            =   1080
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   10
            Text            =   "main.frx":103B
            Top             =   2520
            Width           =   3975
         End
         Begin VB.TextBox txtModDateYear 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2280
            MaxLength       =   4
            TabIndex        =   7
            Top             =   2160
            Width           =   615
         End
         Begin VB.TextBox txtModDateMonth 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2880
            MaxLength       =   2
            TabIndex        =   8
            Top             =   2160
            Width           =   375
         End
         Begin VB.TextBox txtModDateDay 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   9
            Top             =   2160
            Width           =   375
         End
         Begin VB.CheckBox cboxModDate 
            Caption         =   "Automatic"
            Height          =   255
            Left            =   1080
            TabIndex        =   6
            Top             =   2175
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.PictureBox picModBanner 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   1080
            ScaleHeight     =   735
            ScaleWidth      =   3975
            TabIndex        =   1
            Top             =   240
            Width           =   3975
            Begin VB.Label lblNoBanner 
               BackStyle       =   0  'Transparent
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   120
               TabIndex        =   70
               Top             =   120
               Width           =   3735
            End
         End
         Begin VB.Label lblPluginID 
            Alignment       =   1  'Right Justify
            Caption         =   "Plugin ID:"
            Height          =   255
            Left            =   3240
            TabIndex        =   143
            Tag             =   "ID:"
            Top             =   4635
            Width           =   735
         End
         Begin VB.Label lblModSnapFormat 
            Alignment       =   1  'Right Justify
            Caption         =   "SNAP Format:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   105
            Top             =   5430
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label lblModScrnFormat 
            Alignment       =   1  'Right Justify
            Caption         =   "SCRN Format:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   60
            TabIndex        =   104
            Top             =   5430
            Width           =   945
         End
         Begin VB.Label lblModType 
            Alignment       =   1  'Right Justify
            Caption         =   "Mod Type:"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   4635
            Width           =   855
         End
         Begin VB.Label lblModCampaigns 
            Alignment       =   1  'Right Justify
            Caption         =   "Campaigns:"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   4245
            Width           =   855
         End
         Begin VB.Label lblModProgram 
            Alignment       =   1  'Right Justify
            Caption         =   "Program:"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   5025
            Width           =   855
         End
         Begin VB.Label lblUpdateCheck 
            Alignment       =   1  'Right Justify
            Caption         =   "Update url:"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   3885
            Width           =   855
         End
         Begin VB.Label lblWebsite 
            Alignment       =   1  'Right Justify
            Caption         =   "Website url:"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   3525
            Width           =   855
         End
         Begin VB.Label lblModBannerImage 
            Alignment       =   1  'Right Justify
            Caption         =   "Banner Image:"
            Height          =   495
            Left            =   120
            TabIndex        =   77
            Top             =   255
            Width           =   855
         End
         Begin VB.Label lblModAuthor 
            Alignment       =   1  'Right Justify
            Caption         =   "Author:"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   1485
            Width           =   855
         End
         Begin VB.Label lblModVersion 
            Alignment       =   1  'Right Justify
            Caption         =   "Version:"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   1830
            Width           =   855
         End
         Begin VB.Label lblModName 
            Alignment       =   1  'Right Justify
            Caption         =   "Mod Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   1125
            Width           =   855
         End
         Begin VB.Label lblModDate 
            Alignment       =   1  'Right Justify
            Caption         =   "Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   2190
            Width           =   855
         End
         Begin VB.Label lblModDateExample 
            Caption         =   "[yyyy-mm-dd]"
            Height          =   255
            Left            =   3720
            TabIndex        =   72
            Top             =   2190
            Width           =   975
         End
         Begin VB.Label lblModDescription 
            Alignment       =   1  'Right Justify
            Caption         =   "Description:"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   2550
            Width           =   855
         End
      End
      Begin VB.Frame frameInstaller5 
         Caption         =   "Information/License Page"
         Height          =   3735
         Left            =   -74760
         TabIndex        =   65
         Top             =   3960
         Width           =   7215
         Begin VB.TextBox txtInfoPageButton 
            Height          =   285
            Left            =   5880
            MaxLength       =   32
            TabIndex        =   123
            Text            =   "Information"
            Top             =   600
            Width           =   1095
         End
         Begin VB.TextBox txtInfoPageTitle 
            Height          =   285
            Left            =   960
            MaxLength       =   64
            TabIndex        =   121
            Text            =   "Information"
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtInfoPageText 
            Enabled         =   0   'False
            Height          =   2655
            Left            =   960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   52
            Text            =   "main.frx":1052
            Top             =   960
            Width           =   6015
         End
         Begin VB.ComboBox comboInfoPage 
            Height          =   315
            ItemData        =   "main.frx":10D3
            Left            =   2640
            List            =   "main.frx":10E0
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   210
            Width           =   4335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Back Button:"
            Height          =   255
            Left            =   4800
            TabIndex        =   122
            Top             =   630
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Title:"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   630
            Width           =   735
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Text:"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   990
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView listFiles 
         Height          =   5175
         Left            =   -72720
         TabIndex        =   124
         Top             =   720
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   9128
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.TreeView treeFolders 
         Height          =   5415
         Left            =   -74760
         TabIndex        =   131
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   9551
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label lblCustomScriptMessage 
         Caption         =   "Custom Script Warning:"
         Height          =   255
         Left            =   -72600
         TabIndex        =   147
         Top             =   7230
         Width           =   1935
      End
      Begin VB.Label lblCustomScript 
         Caption         =   "Include Custom Script:"
         Height          =   255
         Left            =   -72600
         TabIndex        =   136
         Top             =   6870
         Width           =   1935
      End
      Begin VB.Label lblFolders 
         Height          =   255
         Left            =   -72600
         TabIndex        =   133
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblProgramDirectory 
         Caption         =   "Include Program Directory:"
         Height          =   255
         Left            =   -72600
         TabIndex        =   132
         Top             =   6510
         Width           =   1935
      End
   End
   Begin MSComDlg.CommonDialog dialogOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open"
      Flags           =   4
   End
   Begin MSComDlg.CommonDialog dialogSave 
      Left            =   600
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "lbp"
      DialogTitle     =   "Save As"
      Filter          =   "Installer Project Files (*.lbp) | *.lbp"
      Flags           =   4
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dialogCreate 
      Left            =   1200
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "exe"
      DialogTitle     =   "Save As"
      Filter          =   "Executable Mod Installer (*.exe) | *.exe"
      Flags           =   4
   End
   Begin MSComDlg.CommonDialog dialogSecurity 
      Left            =   1920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "lock"
      DialogTitle     =   "Save As"
      Flags           =   4
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   0
      X2              =   9360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu menu_filemenu 
      Caption         =   "&File"
      Begin VB.Menu menu_open 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu menu_line1 
         Caption         =   "-"
      End
      Begin VB.Menu menu_save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu menu_saveas 
         Caption         =   "Save &As..."
         Shortcut        =   {F5}
      End
      Begin VB.Menu menu_line2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_create 
         Caption         =   "&Create Installer..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu menu_line3 
         Caption         =   "-"
      End
      Begin VB.Menu menu_recent 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu menu_line4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu menu_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menu_helpmenu 
      Caption         =   "&Help"
      Begin VB.Menu menu_help 
         Caption         =   "&Help Topics..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu menu_about 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const VK_LBUTTON = &H1
Private Const VK_CTRLL = 17
Private Const VK_CTRLR = 17
Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = -1&
Private Const DefaultVersionTX As String = "2.02.01"
Private Const RegFileTypePath As String = "HKCR\LBModCreatorProject\shell\open\command"
Const TextDisabledColour As Long = &H8000000F
Const TextEnabledColour As Long = &H80000005
Private Const MARBLEMD5 As String = "A530E6D32C329B14F96252CF3DB7A054"
Dim RecentFiles(9) As String
Dim RecentMenu(9) As Menu
Dim InvalidNSISChars As String
Dim InvalidFileChars As String
Dim LoopPrevention As Boolean
Public EXEDIR As String
Public ProgramINI As String
Public RESDIR As String
Dim KILNDIR As String
Dim RA2DIR As String
Dim MaxRecentFiles As Integer
Dim MsgBoxResult As VbMsgBoxResult
Dim InstFile() As String
Dim InstFileNodeID() As String
Dim InstFileCount As Integer
Dim DCoderDLL As Boolean

Public Sub GlobalErr(ByVal Subroutine As String, ByRef PassedVars() As Variant)
    Dim FileHandle As Integer
    Dim Entry As String
    Dim Counter As Integer
    Entry = Year(Now()) & "-" & PadNum(Month(Now()), 2) & "-" & PadNum(Day(Now()), 2) & " " & PadNum(Hour(Now()), 2) & ":" & PadNum(Minute(Now()), 2) & ":" & PadNum(Second(Now()), 2)
    Entry = "Internal Error" & vbCrLf & "Date: " & Entry & vbCrLf & "Error: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & "Subroutine: " & Subroutine
    If UBound(PassedVars) <> 0 Then
        For Counter = 1 To UBound(PassedVars)
            Entry = Entry & vbCrLf & "PassedVar" & CStr(Counter) & ": " & PassedVars(Counter)
        Next Counter
    End If
    FileHandle = FreeFile
    Open JoinPath(EXEDIR, "except.txt") For Output As #FileHandle
        Print #FileHandle, Entry
    Close #FileHandle
    If Subroutine <> "SaveSettings" Then Call SaveSettings(JoinPath(EXEDIR, "except.lbp"))
    Call WriteLogEntry(Err.Description, True, True)
End Sub

Public Sub WriteLogEntry(Optional ByVal Entry As String = "", Optional ByVal ForceShutdown As Boolean = False, Optional ByVal InternalError As Boolean = False)
    Dim FileNum As Integer
    Dim MaxSize, MinSize, ActSize As Integer
    If Entry <> "" Then Entry = Year(Now()) & "-" & PadNum(Month(Now()), 2) & "-" & PadNum(Day(Now()), 2) & " " & PadNum(Hour(Now()), 2) & ":" & PadNum(Minute(Now()), 2) & ":" & PadNum(Second(Now()), 2) & "  " & Entry
    'If frmOptions.cboxLogFile.Value = 1 Then
    '    FileNum = FreeFile
    '    Open LOGFILE For Append As #FileNum
    '    If InternalError Then
    '        Print #FileNum, "Internal Error! " & Entry
    '    Else
    '        Print #FileNum, Entry
    '    End If
    '    Close #FileNum
    'End If
    If ForceShutdown Then
        If InternalError Then
            MsgBoxResult = MsgBox(App.Title & " has encountered a problem and needs to close." & vbCrLf & vbCrLf & Entry & vbCrLf & vbCrLf & "Please contact Marshall immediately with the following:" & vbCrLf & "Detailed information about what you were doing at the time." & vbCrLf & "Instructions on how to replicate the problem if you can." & vbCrLf & "The <except.lbp> file that has just been generated in the " & App.Title & " program directory." & vbCrLf & "The <except.txt> file that has just been generated in the " & App.Title & " program directory.", vbOKOnly + vbExclamation, "Internal Error")
        Else
            MsgBoxResult = MsgBox(App.Title & " has encountered a problem and needs to close." & vbCrLf & vbCrLf & Entry, vbOKOnly + vbExclamation, App.Title)
        End If
        Call Shutdown
    End If
End Sub

Private Sub cboxGenPat_Click()
    Call UpdateControls
End Sub

Private Sub cboxUseAres_Click()
    Call UpdateControls
End Sub

Private Sub cmdBrowseProgramDir_Click()
    dialogOpen.FileName = txtProgramDirectory.Text
    dialogOpen.DialogTitle = "Select any file from the desired directory"
    dialogOpen.Filter = "All files|*"
    dialogOpen.DefaultExt = ""
'RetryOpen:
    On Error GoTo CancelOpen
    dialogOpen.ShowOpen
    On Error GoTo 0
    If DirExists(dialogOpen.FileName) Then
        txtProgramDirectory.Text = dialogOpen.FileName
    Else
        txtProgramDirectory.Text = GetFilePath(dialogOpen.FileName)
    End If
    Call txtProgramDirectory.SetFocus
CancelOpen:
End Sub

Private Sub cmdBrowseCustomScript_Click()
    dialogOpen.FileName = txtCustomScript.Text
    dialogOpen.DialogTitle = "Select Custom NSIS Script"
    dialogOpen.Filter = "NSIS scripts (*.nsi)|*.nsi|All files|*"
    dialogOpen.DefaultExt = "nsi"
'RetryOpen:
    On Error GoTo CancelOpen
    dialogOpen.ShowOpen
    On Error GoTo 0
    txtCustomScript.Text = dialogOpen.FileName
    Call txtCustomScript.SetFocus
CancelOpen:
End Sub

Private Sub cmdFileErrors_Click()
    Call frmFileErrors.Show
End Sub

Private Sub cmdUninstFiles_Click()
    Call frmUninstallFiles.Show
End Sub

Private Sub cmdUpdateOnlyDestBrowse_Click()
    dialogOpen.FileName = txtUpdateOnlyDest.Text
    dialogOpen.DialogTitle = "Select Latest Installation Liblist"
    dialogOpen.Filter = "Launch Base Liblists|liblist.gam"
    dialogOpen.DefaultExt = "gam"
'RetryOpen:
    On Error GoTo CancelOpen
    dialogOpen.ShowOpen
    On Error GoTo 0
    txtUpdateOnlyDest.Text = dialogOpen.FileName
CancelOpen:
End Sub

Private Sub cmdUpdateOnlySourceBrowse_Click()
    dialogOpen.FileName = txtUpdateOnlySource.Text
    dialogOpen.DialogTitle = "Select Previous Installation Liblist"
    dialogOpen.Filter = "Launch Base Liblists|liblist.gam"
    dialogOpen.DefaultExt = "gam"
'RetryOpen:
    On Error GoTo CancelOpen
    dialogOpen.ShowOpen
    On Error GoTo 0
    txtUpdateOnlySource.Text = dialogOpen.FileName
CancelOpen:
End Sub

Private Sub comboDSP_Click()
    If comboDSP.ListIndex <> 1 Then
        If comboInfoPage.ListIndex <> 1 Then
            comboInfoPage.ListIndex = 1
        Else
            Call comboInfoPage_Click
        End If
    Else
        Call comboInfoPage_Click
    End If
End Sub

Private Sub cmdKey_Click()
    Dim Counter As Integer
    Dim Ok As Boolean
    Dim DummyStringArray() As String
    dialogSecurity.DialogTitle = "Save As"
    dialogSecurity.DefaultExt = ""
    dialogSecurity.Filter = "Launch Base Security Key (security.key) | security.key"
    On Error GoTo CancelLoadKey
    dialogSecurity.ShowOpen
    On Error GoTo 0
    If FileExists(dialogSecurity.FileName) Then
        If UCase(GetFileName(dialogSecurity.FileName)) = "SECURITY.KEY" Then
            cmdKey.Tag = dialogSecurity.FileName
        Else
            MsgBoxResult = MsgBox("The Security Key file must be named " & Quote("security.key") & ".", vbOKOnly + vbInformation, App.Title)
        End If
    Else
        MsgBoxResult = MsgBox(Quote(dialogSecurity.FileName) & " does not exist!", vbOKOnly + vbInformation, App.Title)
    End If
CancelLoadKey:
End Sub

Private Sub cmdLock_Click()
    Dim TempMD5 As String
    Dim process_id
    Dim process_handle
    Dim FileHandle As Integer
    dialogSecurity.DialogTitle = "Save As"
    dialogSecurity.DefaultExt = "lock"
    dialogSecurity.Filter = "Launch Base Security Lock (*.lock) | *.lock"
    On Error GoTo CancelSaveLock
    dialogSecurity.ShowSave
    On Error GoTo 0
    frmMain.Hide
    frmWait.Show
    frmWait.Label1.Caption = "Generating security lock file..."
    frmWait.Refresh
    If DirExists(KILNDIR) Then Call KillDir(KILNDIR)
    Call MakePath(KILNDIR)
    TempMD5 = SaveScript_MainFiles("", , False, False)
    FileHandle = FreeFile()
    If FileExists(JoinPath(KILNDIR, "authlock.erm")) Then Call Kill(JoinPath(KILNDIR, "authlock.erm"))
    Open JoinPath(KILNDIR, "authlock.erm") For Output As #FileHandle
    Print #FileHandle, TempMD5
    Close #FileHandle
    Open JoinPath(KILNDIR, "fgenlock.erm") For Output As #FileHandle
        Print #FileHandle, "SetCompressor /solid lzma"
        Print #FileHandle, "CRCCheck on"
        If FileExists(dialogSecurity.FileName) Then Call Kill(dialogSecurity.FileName)
        Print #FileHandle, "Outfile " & Quote(dialogSecurity.FileName)
        Print #FileHandle, "Page instfiles"
        Print #FileHandle, "AutoCloseWindow True"
        Print #FileHandle, "Section " & Quote("-main")
        Print #FileHandle, "SetOutPath $EXEDIR"
        Print #FileHandle, "File " & Quote(JoinPath(KILNDIR, "authlock.erm"))
        Print #FileHandle, "File " & Quote(JoinPath(KILNDIR, "authlock.lbp"))
        Print #FileHandle, "SectionEnd"
    Close #FileHandle
    Call SaveSettings(JoinPath(KILNDIR, "authlock.lbp"))
    Call ChDir(RESDIR)
    process_id = Shell(Quote(JoinPath(RESDIR, "makensis.exe")) & " " & Quote(JoinPath(KILNDIR, "fgenlock.erm")), vbHide)
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    Call KillDir(KILNDIR)
    Unload frmWait
    frmMain.Show
    frmMain.SetFocus
    MsgBoxResult = MsgBox("Security lock file generated." & vbCrLf & "Please submit this file to the address mentioned in the Help topics." & vbCrLf & vbCrLf & "Do not make any further changes to this plugin." & vbCrLf & "Any further changes to this plugin will make the security lock file invalid.", vbOKOnly + vbInformation, App.Title)
CancelSaveLock:
End Sub

Private Sub comboCompressionMethod_Change()
    Select Case comboCompressionMethod.ListIndex
    Case 0, 1
        cboxSolid.Value = 0
        cboxSolid.Enabled = False
    Case Else
        cboxSolid.Value = 1
        cboxSolid.Enabled = True
    End Select
End Sub

Private Sub comboMixEncrypt_Click()
    Dim DummyStringArray(0) As String
    Call RefreshInstFiles(DummyStringArray())
End Sub

Private Sub comboModCampaigns_Change()
    If Len(comboModCampaigns.Text) > 255 Then
        'Beep 'With MaxLength, none of the other text boxes beep.
        comboModCampaigns.Text = Left$(comboModCampaigns.Text, 255)
        SendKeys "{End}"
    End If
End Sub

Private Sub comboPluginID_Change()
    If Len(StripNumbers(comboPluginID.Text)) = 0 Then
        comboPluginID.Text = ""
        Call MsgBox("Plugin ID must contain at least one letter.", vbOKOnly + vbInformation, App.Title)
    End If
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    Call Init
End Sub

Public Sub Shutdown()
    Call SaveSettings(JoinPath(EXEDIR, "Recent.lbp"))
    Unload frmFileErrors
    Unload frmUninstallFiles
    Unload frmHelp
    Unload Me
End Sub

Private Sub Init()
    Dim Counter As Integer
    Dim TempString As String
    Dim ErrVars(0) As Variant
    If GetArgByName("noexcept") = "" Then On Error GoTo LocalErr
    InvalidFileChars = Chr(92) & " " & Chr(47) & " " & Chr(58) & " " & Chr(42) & " " & Chr(63) & " " & Chr(34) & " " & Chr(60) & " " & Chr(62) & " " & Chr(124)
    InvalidNSISChars = Chr(34) & " " & Chr(36) & " " & Chr(59) & " " & Chr(96)
    EXEDIR = App.Path
    RESDIR = JoinPath(EXEDIR, "Resource")
    KILNDIR = JoinPath(EXEDIR, "Kiln")
    If LCase(ReadRegStr(RegFileTypePath)) <> LCase((GetShortFileName(JoinPath(EXEDIR, ChangeFileType(App.EXEName, "exe"))) & " %1")) Then Call AssociateFileType("lbp", "LBModCreatorProject", "Launch Base Mod Creator Project File", GetShortFileName(JoinPath(EXEDIR, ChangeFileType(App.EXEName, "exe"))), GetShortFileName(JoinPath(EXEDIR, ChangeFileType(App.EXEName, "exe"))), 1)
    ProgramINI = JoinPath(EXEDIR, "LBModCreator.ini")
    DCoderDLL = FileExists(JoinPath(RESDIR, "dcoder.dll"))
    TempString = ReadINIStr("General", "MaxRecentFiles", ProgramINI)
    If TempString = "" Then
        Call WriteINIStr("General", "MaxRecentFiles", "4", ProgramINI)
        MaxRecentFiles = 4
    Else
        MaxRecentFiles = Val(TempString)
    End If
    Call InitRecentFiles
    treeFolders.OLEDropMode = vbOLEDropManual
    treeFolders.OLEDragMode = vbOLEDragAutomatic
    listFiles.OLEDropMode = vbOLEDropManual
    listFiles.OLEDragMode = vbOLEDragAutomatic
    txtModDisplaySound.OLEDropMode = vbOLEDropManual
    txtModDisplaySound.OLEDragMode = vbOLEDragAutomatic
    txtModLaunchSound.OLEDropMode = vbOLEDropManual
    txtModLaunchSound.OLEDragMode = vbOLEDragAutomatic
    txtProgramDirectory.OLEDropMode = vbOLEDropManual
    txtProgramDirectory.OLEDragMode = vbOLEDragAutomatic
    txtMarblePath.OLEDropMode = vbOLEDropManual
    txtMarblePath.OLEDragMode = vbOLEDragAutomatic
    txtUpdateOnlySource.OLEDropMode = vbOLEDropManual
    txtUpdateOnlySource.OLEDragMode = vbOLEDragAutomatic
    txtUpdateOnlyDest.OLEDropMode = vbOLEDropManual
    txtUpdateOnlyDest.OLEDragMode = vbOLEDragAutomatic
    Call ImageList1.ListImages.Clear
    Call ImageList1.ListImages.Add(1, "ClosedFolder", LoadResPicture("FOLDERC", vbResBitmap))
    Call ImageList1.ListImages.Add(2, "OpenFolder", LoadResPicture("FOLDERO", vbResBitmap))
    Set treeFolders.ImageList = ImageList1
    Call listFiles.ColumnHeaders.Add(, , "Filename", 1440)
    Call listFiles.ColumnHeaders.Add(, , "Type", 540)
    Call listFiles.ColumnHeaders.Add(, , "Size", 705)
    Call listFiles.ColumnHeaders.Add(, , "Path", 2175)
    comboInfoPage.ListIndex = 0
    comboCompressionMethod.ListIndex = 3
    comboModType.ListIndex = 0
    comboDSP.ListIndex = 1
    comboSide3Mix.ListIndex = 0
    comboMixEncrypt.ListIndex = 0
    Set picInstallerIcon.Picture = LoadPicture(JoinPath(RESDIR, "default.ico"), vbLPCustom, , 32, 32)
    picInstallerIcon.ToolTipText = JoinPath(RESDIR, "default.ico")
    Call cboxModVersion_Click
    Call cboxModDate_Click
    lblNoBanner.Caption = txtModName.Text
    Set picModBanner.Picture = LoadResPicture("NOBANNER", vbResBitmap)
    Call txtModDescription_Change
    LoopPrevention = False
    TempString = DeQuote(Command$)
    If TempString <> "" Then
        If Not FileExists(TempString) Then TempString = JoinPath(CurDir, TempString)
        If FileExists(TempString) Then
            Call AddRecentFile(TempString)
            Call LoadSettings(TempString)
            frmMain.Caption = GetFileName(TempString) & " - " & App.Title
            dialogSave.FileName = TempString
        End If
    End If
    Call Me.Show
    Call Me.Refresh
    Call txtModName.SetFocus
    Exit Sub
LocalErr:
    Call GlobalErr("Init", ErrVars())
End Sub

Private Sub SaveScript(ByVal ScriptToSave As String, ByVal ExeToSave As String)
    Dim TempString As String
    Dim TempPatch As String
    Dim FileHandle As Integer
    Dim DirCounter As Integer
    Dim FileCounter As Integer
    Dim Ok As Boolean
    Dim process_id
    Dim process_handle
    Dim UninstallCount As Long
    Dim bNeedFA2 As Boolean
    Dim sMainFiles As String
    Dim ErrVars(2) As Variant
    If GetArgByName("noexcept") = "" Then On Error GoTo LocalErr
    ErrVars(1) = ScriptToSave
    ErrVars(2) = ExeToSave
    If cboxUpdateOnly.Value = 0 Then
        sMainFiles = JoinPath(KILNDIR, "script2.erm")
        Call SaveScript_MainFiles(sMainFiles, UninstallCount, bNeedFA2)
    End If
    FileHandle = FreeFile()
    Open ScriptToSave For Output As FileHandle
        'Header
        Print #FileHandle, "VAR MODNAME"
        Print #FileHandle, "VAR LBDIR"
        Print #FileHandle, "VAR OVERWRITE"
        If cboxUpdateOnly.Value = 1 Then
            Print #FileHandle, "VAR PREVNAME"
            Print #FileHandle, "VAR PREVVERS"
        End If
        If txtCustomScript.Enabled Then
            If FileExists(txtCustomScript.Text) Then
                Print #FileHandle, "VAR UNINSTALLCOUNT"
            End If
        End If
        If bNeedFA2 Then Print #FileHandle, "VAR FA2DIR"
        If comboInfoPage.ListIndex = 2 Then Print #FileHandle, "VAR SKIPPY"
        If cboxSolid.Value = 1 Then TempString = "/SOLID " Else TempString = ""
        Select Case comboCompressionMethod.ListIndex
        Case 0: Print #FileHandle, "SetCompress off"
        Case 1: Print #FileHandle, "SetCompressor " & TempString & "zlib"
        Case 2: Print #FileHandle, "SetCompressor " & TempString & "bzip2"
        Case 3: Print #FileHandle, "SetCompressor " & TempString & "lzma"
        End Select
        Select Case cboxCRC.Value
        Case 0: Print #FileHandle, "CRCCheck off"
        Case 1: Print #FileHandle, "CRCCheck on"
        End Select
        If FileExists(picInstallerIcon.Tag) Then
            Print #FileHandle, "Icon " & Quote(picInstallerIcon.Tag)
        Else
            Print #FileHandle, "Icon " & Quote(JoinPath(RESDIR, "default.ico"))
        End If
        Select Case cboxWindowIcon.Value
        Case 0: Print #FileHandle, "WindowIcon off"
        Case 1: Print #FileHandle, "WindowIcon on"
        End Select
        Select Case cboxXPStyle.Value
        Case 0: Print #FileHandle, "XPStyle off"
        Case 1: Print #FileHandle, "XPStyle on"
        End Select
        Select Case cboxShowInstDetails.Value
        Case 0: Print #FileHandle, "ShowInstDetails nevershow"
        Case 1: Print #FileHandle, "ShowInstDetails show"
        End Select
        Print #FileHandle, "InstProgressFlags smooth colored"
        Select Case cboxAutoClose.Value
        Case 0: Print #FileHandle, "AutoCloseWindow false"
        Case 1: Print #FileHandle, "AutoCloseWindow true"
        End Select
        Print #FileHandle, "SubCaption 0 " & Quote(": " & txtInfoPageTitle.Text)
        Print #FileHandle, "Name " & Quote(txtModName.Text & " Setup")
        Print #FileHandle, "Caption " & Quote(txtModName.Text)
        Print #FileHandle, "Outfile " & Quote(ExeToSave)
        Print #FileHandle, "BrandingText " & Quote("LB Mod Creator " & CStr(App.Major) & "." & PadNum(App.Minor, 2) & IIf(App.Revision = 0, "", "." & PadNum(App.Revision, 2)))
        Print #FileHandle, "InstallColors " & txtTextColour(0).Text & txtTextColour(1).Text & txtTextColour(2).Text & " " & txtBGColour(0).Text & txtBGColour(1).Text & txtBGColour(2).Text
        If comboInfoPage.ListIndex = 2 Then Print #FileHandle, "MiscButtonText " & Quote(txtInfoPageButton.Text)
        'LICENSE PAGE
        If comboInfoPage.ListIndex <> 0 Then
            Print #FileHandle, "PageEx license"
            If comboInfoPage.ListIndex = 2 Then Print #FileHandle, "PageCallbacks skipLicense " & Quote() & " " & Quote()
            Select Case comboDSP.ListIndex
            Case 0, 2: TempString = "Install"
            Case 1: TempString = "Continue"
            End Select
            Print #FileHandle, "LicenseText " & Quote(txtModName.Text & ": " & txtInfoPageTitle.Text) & " " & Quote(TempString)
            Print #FileHandle, "LicenseData info.erm"
            Print #FileHandle, "PageExEnd"
            If comboInfoPage.ListIndex = 2 Then
                Print #FileHandle, "Function skipLicense"
                Print #FileHandle, "IntCmp $SKIPPY 1 NoSkipping 0 0"
                Print #FileHandle, "StrCpy $SKIPPY " & Quote("1")
                Print #FileHandle, "Abort"
                Print #FileHandle, "NoSkipping:"
                Print #FileHandle, "FunctionEnd"
            End If
        End If
        'DIRECTORY PAGE
        Print #FileHandle, "PageEx directory"
        Print #FileHandle, "PageCallbacks " & ".preCheckInstdir " & Quote() & " .validateInstdir "
        Print #FileHandle, "DirText " & "`Welcome to the installation program for " & txtModName.Text & ". " & txtModName.Text & " must be installed to a subfolder inside the Launch Base 'Mods' folder.`" & " ; " & Quote("Select installation directory:")
        Print #FileHandle, "PageExEnd"
        Print #FileHandle, "Page instfiles"
        'Function - Initialisation
        Print #FileHandle, "Function .onInit"
        Print #FileHandle, "StrCpy $MODNAME " & Quote(txtModName.Text)
        If cboxUpdateOnly.Value = 1 Then
            Print #FileHandle, "StrCpy $PREVNAME " & Quote(ReadINIStr("General", "Name", txtUpdateOnlySource.Text))
            Print #FileHandle, "StrCpy $PREVVERS " & Quote(ReadINIStr("General", "Version", txtUpdateOnlySource.Text))
        End If
        Print #FileHandle, "ReadRegStr $LBDIR HKLM " & Quote("SOFTWARE\Marshallx Industries\YR Launch Base") & " " & Quote("InstallPath")
        Print #FileHandle, "StrCmp $LBDIR " & Quote() & " LBNotOkay 0"
        Print #FileHandle, "StrCpy $R1 $LBDIR " & Quote() & " -1"
        Print #FileHandle, "StrCmp $R1 " & Quote("\") & " 0 +2"
        Print #FileHandle, "StrCpy $LBDIR $LBDIR -1"
        Print #FileHandle, "IfFileExists " & Quote("$LBDIR\LaunchBase.ini") & " LBOkay LBNotOkay"
        Print #FileHandle, "LBNotOkay:"
        Print #FileHandle, "MessageBox MB_OK|MB_ICONSTOP " & Quote("Launch Base is not installed! Please install Launch Base before installing " & txtModName.Text & ".")
        Print #FileHandle, "Abort"
        Print #FileHandle, "Quit"
        Print #FileHandle, "LBOkay:"
        Print #FileHandle, "System::Call 'kernel32::CreateMutexA(i 0, i 0, t " & Quote("YRLBMUTEXERM1") & ") i .r1 ?e'"
        Print #FileHandle, "Pop $R0"
        Print #FileHandle, "StrCmp $R0 0 +4"
        Print #FileHandle, "MessageBox MB_OK|MB_ICONSTOP " & Quote("Either Launch Base itself or another installer is already running.")
        Print #FileHandle, "Abort"
        Print #FileHandle, "Quit"
        Select Case txtINSTDIR.Text
        Case "": Print #FileHandle, "StrCpy $INSTDIR $LBDIR\Mods"
        Case Else: Print #FileHandle, "StrCpy $INSTDIR " & Quote("$LBDIR\Mods\" & txtINSTDIR.Text)
        End Select
        If bNeedFA2 Then
            Print #FileHandle, "ReadINIStr $FA2DIR " & Quote("$LBDIR\LaunchBase.ini") & " " & Quote("General") & " " & Quote("FinalAlert2Path")
            Print #FileHandle, "IfFileExists " & Quote("$FA2DIR\marble.mix") & " FA2Okay 0"
            If optTX(3).Value = True Then
                Print #FileHandle, "StrCpy $FA2DIR " & Quote()
                Print #FileHandle, "MessageBox MB_OK|MB_ICONSTOP " & Quote("$MODNAME needs to check some data from your FinalAlert 2 installation if it is to install the $MODNAME FA2 mod.$\nPlease run Launch Base and enter your FinalAlert 2 folder path on the 'FA2 Mods' tab before running this installer again.$\n$\nWould you just like to install the $MODNAME plugin?") & " IDYES FA2Okay"
            Else
                Print #FileHandle, "MessageBox MB_OK|MB_ICONSTOP " & Quote("$MODNAME needs to check some data from your FinalAlert 2 installation.$\nPlease run Launch Base and enter your FinalAlert 2 folder path on the 'FA2 Mods' tab before running this installer again.")
            End If
            Print #FileHandle, "Abort"
            Print #FileHandle, "Quit"
            Print #FileHandle, "FA2Okay:"
        End If
        If comboInfoPage.ListIndex = 2 Then Print #FileHandle, "StrCpy $SKIPPY " & Quote("0")
        Print #FileHandle, "FunctionEnd"
        'Function - Validate directory
        Select Case comboDSP.ListIndex
        Case 0
            Select Case cboxUpdateOnly.Value
            Case 0: Print #FileHandle, "!include " & Quote(JoinPath(RESDIR, "fvaldir0.erm"))
            Case 1: Print #FileHandle, "!include " & Quote(JoinPath(RESDIR, "fvaldir3.erm"))
            End Select
            Print #FileHandle, "Function .validateInstdir"
            Print #FileHandle, "FunctionEnd"
        Case 1
            Select Case cboxUpdateOnly.Value
            Case 0: Print #FileHandle, "!include " & Quote(JoinPath(RESDIR, "fvaldir1.erm"))
            Case 1: Print #FileHandle, "!include " & Quote(JoinPath(RESDIR, "fvaldir4.erm"))
            End Select
            Print #FileHandle, "Function .preCheckInstdir"
            Print #FileHandle, "FunctionEnd"
        Case 2
            Select Case cboxUpdateOnly.Value
            Case 0: Print #FileHandle, "!include " & Quote(JoinPath(RESDIR, "fvaldir1.erm"))
            Case 1: Print #FileHandle, "!include " & Quote(JoinPath(RESDIR, "fvaldir4.erm"))
            End Select
            Select Case cboxUpdateOnly.Value
            Case 0: Print #FileHandle, "!include " & Quote(JoinPath(RESDIR, "fvaldir2.erm"))
            Case 1: Print #FileHandle, "!include " & Quote(JoinPath(RESDIR, "fvaldir5.erm"))
            End Select
        End Select
        'Remove old files
        Print #FileHandle, "Section " & Quote("-main") & " 0"
        If cboxUpdateOnly.Value = 0 Then
            Print #FileHandle, "StrCmp $OVERWRITE " & Quote("TRUE") & " 0 +3"
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR")
            Print #FileHandle, "Goto DestinationClear"
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\cameo")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\fa2files")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\hva")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\ini")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\interface")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\launcher")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\manual")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\map")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\mix")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\screen")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\shp")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\side 1")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\side 2")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\side 3")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\side 4")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\sound")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\speech")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\string table")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\syringe")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\taunts")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\theme")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\video")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\vxl")
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\Kiln")
            If cboxOldSaves.Value = 1 Then
                Print #FileHandle, "IfFileExists " & Quote("$INSTDIR\saves\*.*") & " 0 +3"
                Print #FileHandle, "Rename  " & Quote("$INSTDIR\saves") & " " & Quote("$INSTDIR\newsaves.ran\old")
                Print #FileHandle, "Rename  " & Quote("$INSTDIR\newsaves.ran") & " " & Quote("$INSTDIR\saves")
            End If
            If cboxResetGameConfig.Value = 1 Then
                Print #FileHandle, "IfFileExists " & Quote("$INSTDIR\launcher\userdata.lbu") & " 0 +2"
                Print #FileHandle, "Delete " & Quote("$INSTDIR\launcher\userdata.lbu")
            End If
            Print #FileHandle, "DestinationClear:"
            'Install files
            Print #FileHandle, "!include " & Quote(sMainFiles)
            'CUSTOM SCRIPT
            If txtCustomScript.Enabled Then
                If FileExists(txtCustomScript.Text) Then
                    Print #FileHandle, "StrCpy $UNINSTALLCOUNT " & Quote(CStr(UninstallCount))
                    Print #FileHandle, "MessageBox MB_YESNO|MB_ICONQUESTION " & "`$MODNAME includes a custom install script.$\nYou should only run this script if you trust the source of this installer. However, $MODNAME may not run correctly if this script is not run." & IIf(Len(txtCustomScriptMessage.Text) <> 0, " The following is a message from the author:$\n$\n" & Quote(txtCustomScriptMessage.Text), "") & "$\n$\nWould you like to run this script?` IDNO +2 IDYES 0"
                    Print #FileHandle, "Call funcCustomScript"
                End If
            End If
            Print #FileHandle, "SectionEnd"
            Print #FileHandle, "Function .onInstSuccess"
            Print #FileHandle, "MessageBox MB_YESNO|MB_ICONQUESTION " & Quote("$MODNAME installation complete.$\nWould you like to run Launch Base now?") & " IDYES LaunchLB IDNO GameOver"
            Print #FileHandle, "LaunchLB:"
            Print #FileHandle, "Exec " & Quote("$LBDIR\LaunchBase.exe")
            Print #FileHandle, "GameOver:"
            Print #FileHandle, "FunctionEnd"
            If txtCustomScript.Enabled Then
                If FileExists(txtCustomScript.Text) Then Print #FileHandle, "!include " & Quote(txtCustomScript.Text)
            End If
        Else
            'Update only installer
            Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\Kiln")
            If FileExists(JoinPath(GetFileName(txtUpdateOnlyDest.Text), "sound1.ogg")) Or FileExists(JoinPath(GetFileName(txtUpdateOnlyDest.Text), "sound1.flac")) Then
                Print #FileHandle, "IfFileExists " & Quote("$INSTDIR\launcher\sound1.wav") & " 0 +2"
                Print #FileHandle, "Delete " & Quote("$INSTDIR\launcher\sound1.wav")
            End If
            If FileExists(JoinPath(GetFileName(txtUpdateOnlyDest.Text), "sound2.ogg")) Or FileExists(JoinPath(GetFileName(txtUpdateOnlyDest.Text), "sound2.flac")) Then
                Print #FileHandle, "IfFileExists " & Quote("$INSTDIR\launcher\sound2.wav") & " 0 +2"
                Print #FileHandle, "Delete " & Quote("$INSTDIR\launcher\sound2.wav")
            End If
            If cboxOldSaves.Value = 1 Then
                Print #FileHandle, "IfFileExists " & Quote("$INSTDIR\saves\*.*") & " 0 +3"
                Print #FileHandle, "Rename  " & Quote("$INSTDIR\saves") & " " & Quote("$INSTDIR\newsaves.ran\old")
                Print #FileHandle, "Rename  " & Quote("$INSTDIR\newsaves.ran") & " " & Quote("$INSTDIR\saves")
            End If
            If cboxResetGameConfig.Value = 1 Then
                Print #FileHandle, "IfFileExists " & Quote("$INSTDIR\launcher\userdata.lbu") & " 0 +2"
                Print #FileHandle, "Delete " & Quote("$INSTDIR\launcher\userdata.lbu")
            End If
            Print #FileHandle, "ClearErrors"
            Print #FileHandle, "SetOutPath $INSTDIR"
            dirlistbox.Path = GetFilePath(GetFilePath(txtUpdateOnlyDest.Text))
            Call dirlistbox.Refresh
            For DirCounter = -1 To (dirlistbox.ListCount - 1)
                Select Case DirCounter
                Case -1
                    filelistbox.Path = dirlistbox.Path
                    TempString = GetFilePath(GetFilePath(txtUpdateOnlySource.Text))
                    Print #FileHandle, "SetOutPath " & Quote("$INSTDIR")
                Case Else
                    filelistbox.Path = dirlistbox.List(DirCounter)
                    TempString = JoinPath(GetFilePath(GetFilePath(txtUpdateOnlySource.Text)), GetFileName(dirlistbox.List(DirCounter)))
                    Print #FileHandle, "SetOutPath " & Quote("$INSTDIR\" & GetFileName(dirlistbox.List(DirCounter)))
                End Select
                Call filelistbox.Refresh
                For FileCounter = 0 To filelistbox.ListCount - 1
                    Ok = False
                    If DirExists(TempString) Then
                        If FileExists(JoinPath(TempString, filelistbox.List(FileCounter))) Then
                            If GetFileMD5(JoinPath(TempString, filelistbox.List(FileCounter))) = GetFileMD5(JoinPath(filelistbox.Path, filelistbox.List(FileCounter))) Then
                                Ok = True 'file has not changed so ignore
                            Else
                                If cboxGenPat.Value = 1 Then
                                    If GetFileSize(JoinPath(TempString, filelistbox.List(FileCounter))) >= (sliderPatchMinSize.Value * 16384) Then
                                        TempPatch = GenPat(JoinPath(TempString, filelistbox.List(FileCounter)), JoinPath(filelistbox.Path, filelistbox.List(FileCounter)), , Val(lblPatchBlockSize.Caption))
                                        Print #FileHandle, "File " & Quote(TempPatch)
                                        Print #FileHandle, "vpatch::vpatchfile " & Quote("$OUTDIR\" & GetFileName(TempPatch)) & " " & Quote("$OUTDIR\" & filelistbox.List(FileCounter)) & " " & Quote("$OUTDIR\" & ChangeFileType(filelistbox.List(FileCounter), "zzz"))
                                        Print #FileHandle, "Pop $R0"
                                        Print #FileHandle, "StrCmp $R0 " & Quote("OK") & " +3 0"
                                        Print #FileHandle, "SetErrors"
                                        Print #FileHandle, "Goto +3"
                                        Print #FileHandle, "Delete " & Quote("$OUTDIR\" & filelistbox.List(FileCounter))
                                        Print #FileHandle, "Rename " & Quote("$OUTDIR\" & ChangeFileType(filelistbox.List(FileCounter), "zzz")) & " " & Quote("$OUTDIR\" & filelistbox.List(FileCounter))
                                        Print #FileHandle, "Delete " & Quote("$OUTDIR\" & GetFileName(TempPatch))
                                        Ok = True
                                    End If
                                 End If
                            End If
                        End If
                    End If
                    If Not Ok Then
                        Print #FileHandle, "File " & Quote(JoinPath(filelistbox.Path, filelistbox.List(FileCounter)))
                    End If
                Next FileCounter
            Next DirCounter
            dirlistbox.Path = GetFilePath(GetFilePath(txtUpdateOnlySource.Text))
            Call dirlistbox.Refresh
            For DirCounter = -1 To (dirlistbox.ListCount - 1)
                Select Case DirCounter
                Case -1
                    filelistbox.Path = dirlistbox.Path
                    TempString = GetFilePath(GetFilePath(txtUpdateOnlyDest.Text))
                Case Else
                    filelistbox.Path = dirlistbox.List(DirCounter)
                    TempString = JoinPath(GetFilePath(GetFilePath(txtUpdateOnlyDest.Text)), GetFileName(dirlistbox.List(DirCounter)))
                End Select
                If Not DirExists(TempString) Then
                    Print #FileHandle, "RMDir /r " & Quote("$INSTDIR\" & GetFileName(TempString))
                Else
                    Print #FileHandle, "SetOutPath " & Quote("$INSTDIR\" & GetFileName(dirlistbox.List(DirCounter)))
                    Call filelistbox.Refresh
                    For FileCounter = 0 To filelistbox.ListCount - 1
                        If Not FileExists(JoinPath(TempString, filelistbox.List(FileCounter))) Then
                            Print #FileHandle, "Delete " & Quote("$OUTDIR\" & filelistbox.List(FileCounter))
                        End If
                    Next FileCounter
                End If
            Next DirCounter
            Print #FileHandle, "SectionEnd"
            Print #FileHandle, "Function .onInstSuccess"
            Print #FileHandle, "IfErrors +2 0"
            Print #FileHandle, "MessageBox MB_YESNO|MB_ICONQUESTION " & Quote("$MODNAME installation complete.$\nWould you like to run Launch Base now?") & " IDYES LaunchLB IDNO GameOver"
            If txtModWebsite.Text <> "" Then
                Print #FileHandle, "MessageBox MB_OK|MB_ICONEXCLAMATION " & Quote("One or more files could not be updated!$\nIt is strongly recommended that you reinstall $MODNAME using the full installer.$\nYou can download the latest full installer from " & txtModWebsite.Text) & " IDOK GameOver"
            Else
                Print #FileHandle, "MessageBox MB_OK|MB_ICONEXCLAMATION " & Quote("One or more files could not be updated!$\nIt is strongly recommended that you reinstall $MODNAME using the full installer.") & " IDOK GameOver"
            End If
            Print #FileHandle, "LaunchLB:"
            Print #FileHandle, "Exec " & Quote("$LBDIR\LaunchBase.exe")
            Print #FileHandle, "GameOver:"
            Print #FileHandle, "FunctionEnd"
        End If
    Close #FileHandle
    Exit Sub
LocalErr:
    Call GlobalErr("SaveScript", ErrVars())
End Sub

Public Sub SaveScript_ProgramDir(ByVal sFolder As String, ByVal iScriptHandle As Integer, ByVal sLiblist As String, ByRef iCounter As Long)
    Dim fold As Scripting.Folder
    Dim foldSub As Scripting.Folder
    Dim fil As File
    Dim fso As New FileSystemObject
    Dim strOutput As String
    Dim TempStr As String
    Set fold = fso.GetFolder(sFolder)
    For Each fil In fold.Files
        'to get filename
        TempStr = GetRelativePath(fil.Path, txtProgramDirectory.Text)
        Print #iScriptHandle, "File " & Quote(fil.Path)
        iCounter = iCounter + 1
        Call WriteINIStr("Uninstall", CStr(iCounter), GetRelativePath(fil.Path, txtProgramDirectory.Text), sLiblist)
    Next
    For Each foldSub In fold.SubFolders
        'to get folder name
        TempStr = GetRelativePath(foldSub.Path, txtProgramDirectory.Text)
        'prevent restricted folder paths
        Select Case UCase(TempStr)
        Case "LAUNCHER"
            Call MsgBox("The 'Include Program Directory' you have specified contains a sub-directory named " & Quote(TempStr) & "." & vbCrLf & "This is a reserved folder name used by Launch Base. This sub-directory and any files in it will not be included in your installer.", vbOKOnly + vbInformation, App.Title)
        Case Else
            Print #iScriptHandle, "SetOutPath " & Quote("$INSTDIR\" & TempStr)
            Call SaveScript_ProgramDir(foldSub.Path, iScriptHandle, sLiblist, iCounter)
        End Select
    Next
End Sub

Private Function SaveScript_MainFiles(ByVal sScript As String, Optional ByRef iUninstallCount As Long, Optional ByRef bNeedFA2 As Boolean, Optional ByVal bScript As Boolean = True) As String
'bScript=FALSE means return the archive MD5 instead of generating the script
    Dim hFile As Integer
    Dim sFile As String
    Dim iFolder As Integer
    Dim iFile As Integer
    Dim hExpand As Integer
    Dim hSide1 As Integer
    Dim hSide2 As Integer
    Dim hSide3 As Integer
    Dim hSide4 As Integer
    Dim sExpand As String
    Dim sLiblist As String
    Dim bOk As Boolean
    Dim bNewFolder As Boolean
    Dim bLauncherDone As Boolean
    Dim iExpandPos As Long
    Dim iLiblistPos As Long
    Dim sRunningMD5 As String
    Dim sTempMD5 As String
    Dim ErrVars(0) As Variant
    If GetArgByName("noexcept") = "" Then On Error GoTo LocalErr
    'note, liblist is created in this routine but gets modified later on for some mod types
    sLiblist = JoinPath(KILNDIR, "liblist.gam")
    If bScript Then
        hFile = FreeFile()
        Open sScript For Output As #hFile
    Else
        sRunningMD5 = ""
    End If
    If cboxForRA2.Value = 0 And comboModType.ListIndex <> 1 Then
        If optTX(3).Value = True Then
            sExpand = "expandmd06.mix"
        Else
            sExpand = "expandmd98.mix"
        End If
    Else
        If optTX(3).Value = True Then
            sExpand = "expand06.mix"
        Else
            sExpand = "expand98.mix"
        End If
    End If
    sExpand = JoinPath(KILNDIR, sExpand)
    bLauncherDone = False
    iFolder = 1
    'this adds all the files to the script and generates the MD5 string for plugin keys
    'requires that both the tree nodes and InstFile() are in alphabetical order, otherwise the order of the checksums will be wrong and the key will be worthless
    If bScript Then Print #hFile, "SetOutPath " & Quote("$INSTDIR")
    Call SaveLiblist(sLiblist)
    iUninstallCount = -1
    Do While iFolder <= treeFolders.Nodes.Count
        If Not bLauncherDone Then
            If UCase$(treeFolders.Nodes.Item(iFolder).Key) > "LAUNCHER" Then
                If bScript Then Print #hFile, "SetOutPath " & Quote("$INSTDIR\launcher")
                If FileExists(picModBanner.Tag) Then
                    If bScript Then
                        Print #hFile, "File /oname=$INSTDIR\launcher\banner." & FileType(picModBanner.Tag) & " " & Quote(picModBanner.Tag)
                    Else
                       sRunningMD5 = sRunningMD5 & GetFileMD5(picModBanner.Tag)
                    End If
                End If
                If bScript Then
                    Print #hFile, "File " & Quote(sLiblist)
                Else
                   iLiblistPos = Len(sRunningMD5)
                End If
                If bScript And FileExists(cmdKey.Tag) Then Print #hFile, "File " & Quote(cmdKey.Tag)
                If FileExists(txtModDisplaySound.Text) Then
                    If bScript Then
                        If cboxFLAC.Value = 1 And FileType(txtModDisplaySound.Text) = "WAV" Then
                            Call ConvertWavToFlac(txtModDisplaySound.Text, JoinPath(KILNDIR, "sound1.flac"))
                            Print #hFile, "File /oname=$INSTDIR\launcher\sound1.flac " & Quote(JoinPath(KILNDIR, "sound1.flac"))
                        Else
                            Print #hFile, "File /oname=$INSTDIR\launcher\sound1." & LCase$(FileType(txtModDisplaySound.Text)) & " " & Quote(txtModDisplaySound.Text)
                        End If
                    Else
                       sRunningMD5 = sRunningMD5 & GetFileMD5(txtModDisplaySound.Text)
                    End If
                End If
                If FileExists(txtModLaunchSound.Text) Then
                    If bScript Then
                        If cboxFLAC.Value = 1 And FileType(txtModLaunchSound.Text) = "WAV" Then
                            Call ConvertWavToFlac(txtModLaunchSound.Text, JoinPath(KILNDIR, "sound2.flac"))
                            Print #hFile, "File /oname=$INSTDIR\launcher\sound2.flac " & Quote(JoinPath(KILNDIR, "sound2.flac"))
                        Else
                            Print #hFile, "File /oname=$INSTDIR\launcher\sound2." & LCase$(FileType(txtModLaunchSound.Text)) & " " & Quote(txtModLaunchSound.Text)
                        End If
                    Else
                       sRunningMD5 = sRunningMD5 & GetFileMD5(txtModLaunchSound.Text)
                    End If
                End If
                bLauncherDone = True
            End If
        End If
        If Not Not InstFileNodeID Then
            bOk = False
            iFile = 0
            bNewFolder = True
            Select Case UCase(InstFileNodeID(iFile))
            Case "CAMEO", "HVA", "INI", "INTERFACE", "MAP", "MIX", "SHP", "SPEECH", "TMP", "VXL" 'EXPANDMD98.MIX
                If Not bOk Then
                    bOk = True
                    If hExpand = 0 Then
                        If comboMixEncrypt.ListIndex <> 0 Then
                            If (Not FileAddedByUser("EXPAND")) Then hExpand = DCoderMIXOpen(sExpand)
                        End If
                    End If
                End If
                Do While iFile < InstFileCount
                    If Len(InstFile(iFile)) <> 0 Then
                        If InstFileNodeID(iFile) = treeFolders.Nodes.Item(iFolder).Key Then
                            If FileExists(InstFile(iFile)) Then
                                If bScript And bNewFolder Then
                                    Print #hFile, "SetOutPath " & Quote("$INSTDIR\" & InstFileNodeID(iFile))
                                    bNewFolder = False
                                End If
                                If hExpand <> 0 Then
                                    Call DCoderMIXInsert(hExpand, InstFile(iFile))
                                Else
                                    If bScript Then
                                        If cboxFLAC.Value = 1 And FileType(InstFile(iFile)) = "FLAC" Then
                                            Call ConvertWavToFlac(InstFile(iFile), JoinPath(KILNDIR, ChangeFileType(GetFileName(InstFile(iFile)), "flac")))
                                            Print #hFile, "File " & Quote(JoinPath(KILNDIR, ChangeFileType(GetFileName(InstFile(iFile)), "flac")))
                                        Else
                                            Print #hFile, "File " & Quote(InstFile(iFile))
                                        End If
                                    Else
                                       sRunningMD5 = sRunningMD5 & GetFileMD5(InstFile(iFile))
                                    End If
                                End If
                            End If
                        End If
                    End If
                    iFile = iFile + 1
                Loop
            Case "FA2FILES"
                Do While iFile < InstFileCount
                    If Len(InstFile(iFile)) <> 0 Then
                        If InstFileNodeID(iFile) = treeFolders.Nodes.Item(iFolder).Key Then
                            If FileExists(InstFile(iFile)) Then
                                If bScript Then
                                    If bNewFolder Then Print #hFile, "SetOutPath " & Quote("$INSTDIR\" & InstFileNodeID(iFile))
                                    If UCase(GetFileName(InstFile(iFile))) = "MARBLE.MIX" Then
                                        If cboxGenPat.Value = 1 Then
                                            If Len(txtMarblePath.Text) <> 0 Then
                                                bNeedFA2 = False 'actually we do need it, but assuming for now that we will fail to make a patch
                                                If FileExists(txtMarblePath.Text) Then
                                                    If GetFileMD5(txtMarblePath.Text) = MARBLEMD5 Then
                                                        If GetFileMD5(InstFile(iFile)) <> MARBLEMD5 Then
                                                            Call GenPat(txtMarblePath.Text, InstFile(iFile), JoinPath(KILNDIR, "marble.pat"), Val(lblPatchBlockSize.Caption))
                                                            If FileExists(JoinPath(KILNDIR, "marble.pat")) Then
                                                                Print #hFile, "StrCmp $FA2DIR " & Quote() & " NoFA2Files"
                                                                Print #hFile, "File " & Quote(JoinPath(KILNDIR, "marble.pat"))
                                                                'apply the patch
                                                                Print #hFile, "vpatch::vpatchfile " & Quote("$OUTDIR\marble.pat") & " " & Quote("$FA2DIR\marble.mix") & " " & Quote("$OUTDIR\marble.mix")
                                                                Print #hFile, "Pop $R0"
                                                                Print #hFile, "StrCmp $R0 " & Quote("OK") & " MarblePatchDone 0"
                                                                Print #hFile, "IfFileExists " & Quote("$LBDIR\Backup\marble.mix") & " 0 FA2NotOkay"
                                                                Print #hFile, "vpatch::vpatchfile " & Quote("$OUTDIR\marble.pat") & " " & Quote("$LBDIR\Backup\marble.mix") & " " & Quote("$OUTDIR\marble.mix")
                                                                Print #hFile, "Pop $R0"
                                                                Print #hFile, "StrCmp $R0 " & Quote("OK") & " MarblePatchDone FA2NotOkay"
                                                                Print #hFile, "FA2NotOkay:"
                                                                Print #hFile, "Delete " & Quote("$OUTDIR\" & filelistbox.List(iFile))
                                                                Print #hFile, "MessageBox MB_OK|MB_ICONSTOP " & Quote("This FinalAlert 2 mod needs to check some data from an unmodified FinalAlert 2 installation.$\nPlease reinstall FinalAlert 2 YR, update your FinalAlert 2 folder path in Launch Base and then run this installer again.")
                                                                Print #hFile, "Quit"
                                                                Print #hFile, "MarblePatchDone:"
                                                                Print #hFile, "Delete " & Quote("$OUTDIR\marble.pat")
                                                                bNeedFA2 = True
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                            If Not bNeedFA2 Then
                                                Call MsgBox("Failed to generate a patch file for " & Quote(InstFile(iFile)) & "." & vbCrLf & "The whole file will be included.", vbOKOnly + vbExclamation, App.Title)
                                                Print #hFile, "File " & Quote(InstFile(iFile))
                                            End If
                                        Else
                                            Print #hFile, "File " & Quote(InstFile(iFile))
                                        End If
                                    Else
                                        Print #hFile, "File " & Quote(InstFile(iFile))
                                    End If
                                    If bNewFolder Then
                                        Print #hFile, "NoFA2Files:"
                                        bNewFolder = False
                                    End If
                                Else
                                   sRunningMD5 = sRunningMD5 & GetFileMD5(InstFile(iFile))
                                End If
                            End If
                        End If
                    End If
                    iFile = iFile + 1
                Loop
            Case "VIDEO"
                iExpandPos = -1
                Do While iFile < InstFileCount
                    If Len(InstFile(iFile)) <> 0 Then
                        If InstFileNodeID(iFile) = treeFolders.Nodes.Item(iFolder).Key Then
                            If FileExists(InstFile(iFile)) Then
                                If bNewFolder Then
                                    If bScript Then Print #hFile, "SetOutPath " & Quote("$INSTDIR\" & InstFileNodeID(iFile))
                                    bNewFolder = False
                                End If
                                If iExpandPos = -1 Then
                                    If UCase$(InstFile(iFile)) > UCase$(sExpand) Then
                                        iExpandPos = Len(sRunningMD5)
                                    End If
                                End If
                                If bScript Then
                                    If cboxFLAC.Value = 1 And FileType(InstFile(iFile)) = "FLAC" Then
                                        Call ConvertWavToFlac(InstFile(iFile), JoinPath(KILNDIR, ChangeFileType(GetFileName(InstFile(iFile)), "flac")))
                                        Print #hFile, "File " & Quote(JoinPath(KILNDIR, ChangeFileType(GetFileName(InstFile(iFile)), "flac")))
                                    Else
                                        Print #hFile, "File " & Quote(InstFile(iFile))
                                    End If
                                Else
                                   sRunningMD5 = sRunningMD5 & GetFileMD5(InstFile(iFile))
                                End If
                            End If
                        End If
                    End If
                    iFile = iFile + 1
                Loop
                If iExpandPos = -1 Then iExpandPos = Len(sRunningMD5)
            Case Else
                Do While iFile < InstFileCount
                    If Len(InstFile(iFile)) <> 0 Then
                        If InstFileNodeID(iFile) = treeFolders.Nodes.Item(iFolder).Key Then
                            If FileExists(InstFile(iFile)) Then
                                If bScript Then
                                    If bNewFolder Then
                                        If InstFileNodeID(iFile) <> "Node0" Then
                                            Print #hFile, "SetOutPath " & Quote("$INSTDIR\" & InstFileNodeID(iFile))
                                        Else
                                            Print #hFile, "SetOutPath " & Quote("$INSTDIR")
                                        End If
                                        bNewFolder = False
                                    End If
                                    If UCase(InstFileNodeID(iFile)) <> "MANUAL" And cboxFLAC.Value = 1 And FileType(InstFile(iFile)) = "FLAC" Then
                                        Call ConvertWavToFlac(InstFile(iFile), JoinPath(KILNDIR, ChangeFileType(GetFileName(InstFile(iFile)), "flac")))
                                        Print #hFile, "File " & Quote(JoinPath(KILNDIR, ChangeFileType(GetFileName(InstFile(iFile)), "flac")))
                                    Else
                                        Print #hFile, "File " & Quote(InstFile(iFile))
                                    End If
                                Else
                                   sRunningMD5 = sRunningMD5 & GetFileMD5(InstFile(iFile))
                                End If
                                If comboModType.ListIndex = 2 Then
                                    iUninstallCount = iUninstallCount + 1
                                    Call WriteINIStr("Uninstall", CStr(iUninstallCount), LCase$(InstFileNodeID(iFile)) & "\" & InstFile(iFile), sLiblist)
                                End If
                            End If
                        End If
                    End If
                    iFile = iFile + 1
                Loop
            End Select
            iFolder = iFolder + 1
        Else
            Exit Do
        End If
    Loop
    If comboModType.ListIndex = 2 Then
        'include program directory
        If DirExists(txtProgramDirectory.Text) Then
            Print #hFile, "SetOutPath " & Quote("$INSTDIR")
            Call SaveScript_ProgramDir(txtProgramDirectory.Text, hFile, sLiblist, iUninstallCount)
        End If
        'extra uninstall files
        iFile = 0
        Do While iFile < frmUninstallFiles.lvUninstallFiles.ListItems.Count
            iUninstallCount = iUninstallCount + 1
            Call WriteINIStr("Uninstall", CStr(iUninstallCount), frmUninstallFiles.lvUninstallFiles.ListItems.Item(iFile), sLiblist)
            iFile = iFile + 1
        Loop
    End If
    If Not bScript Then
        sTempMD5 = GetFileMD5(sLiblist)
        If iLiblistPos = Len(sRunningMD5) Then
            sRunningMD5 = Left$(sRunningMD5, iLiblistPos) & sTempMD5
        Else
            sRunningMD5 = Left$(sRunningMD5, iLiblistPos) & sTempMD5 & Mid$(sRunningMD5, iLiblistPos + 1)
        End If
        iExpandPos = iExpandPos + Len(sTempMD5)
    End If
    If hExpand <> 0 Then
        If bScript Then Print #hFile, "SetOutPath " & Quote("$INSTDIR\video")
        Call DCoderMIXWrite(hExpand)
        If bScript Then
            Print #hFile, "File " & Quote(sExpand)
        Else
            sTempMD5 = GetFileMD5(sExpand)
            If iExpandPos = Len(sRunningMD5) Then
                sRunningMD5 = Left$(sRunningMD5, iExpandPos) & sTempMD5
            Else
                sRunningMD5 = Left$(sRunningMD5, iExpandPos) & sTempMD5 & Mid$(sRunningMD5, iExpandPos + 1)
            End If
        End If
    End If
    If bScript Then
        Close #hFile
    Else
        Call Kill(sLiblist)
        If hExpand Then Call Kill(sExpand)
    End If
    SaveScript_MainFiles = sRunningMD5
    Exit Function
LocalErr:
    Call GlobalErr("GenerateArchive", ErrVars())
End Function

Private Sub DCoderMIXClose(ByVal hMIX As Integer)
End Sub

Private Sub DCoderMIXInsert(ByVal hMIX As Integer, ByVal sSourceFile As String)
End Sub

Private Sub DCoderMIXWrite(ByVal hMIX As Integer)
End Sub

Private Function DCoderMIXOpen(ByVal sMIXFile As String) As Integer
    DCoderMIXOpen = 0
End Function

Private Function FileAddedByUser(ByVal FileID As String) As Boolean
    Dim Ok As Boolean
    Dim File1 As String
    Dim File2 As String
    Dim NodeID As String
    Dim Counter As Integer
    Ok = False
    Select Case UCase(FileID)
    Case "ECACHE"
        File1 = "ECACHEMD98.MIX"
        File2 = "ECACHE98.MIX"
        NodeID = "VIDEO"
    Case "EXPAND"
        File1 = "EXPANDMD98.MIX"
        File2 = "EXPAND98.MIX"
        NodeID = "VIDEO"
    Case Else
        File1 = UCase(FileID)
        File2 = UCase(FileID)
        NodeID = "SYRINGE"
    End Select
    Counter = 0
    Do While Counter < InstFileCount
        If UCase(InstFileNodeID(Counter)) = NodeID Then
            Select Case UCase(GetFileName(InstFile(Counter)))
            Case File1, File2
                Ok = True
                Exit Do
            End Select
        End If
        Counter = Counter + 1
    Loop
    FileAddedByUser = Ok
End Function

'Private Function GetUserAddedMarblePath() As String
'    Dim RetVal As String
'    Dim Counter As Integer
'    RetVal = ""
'    Counter = 0
'    Do While Counter < InstFileCount
'        If UCase(InstFileNodeID(Counter)) = "FA2FILES" Then
'            If UCase(GetFileName(InstFile(Counter))) = "MARBLE.MIX" Then
'                RetVal = InstFile(Counter)
'                Exit Do
'            End If
'        End If
'        Counter = Counter + 1
'    Loop
'    GetUserAddedMarblePath = RetVal
'End Function

Private Function GenPat(ByVal SourceFile As String, ByVal DestFile As String, Optional ByVal PatchFile As String = "", Optional ByVal BlockSize As Integer = 64) As String
    Dim process_id
    Dim process_handle
    Dim DotPos As Long
    If PatchFile = "" Then PatchFile = ChangeFileType(SourceFile, "pat")
    Call ChDir(RESDIR)
    process_id = Shell(JoinPath(RESDIR, "genpat.exe") & " /B=" & CStr(BlockSize) & " " & Quote(SourceFile) & " " & Quote(DestFile) & " " & Quote(PatchFile), vbNormal)
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    GenPat = PatchFile
End Function

Private Sub SaveLiblist(ByVal Liblist As String)
    Dim Counter As Integer
    Dim ErrVars(0) As Variant
    If GetArgByName("noexcept") = "" Then If GetArgByName("noexcept") = "" Then On Error GoTo LocalErr
    Call WriteINIStr("General", "Name", txtModName.Text, Liblist)
    Call WriteINIStr("General", "Author", txtModAuthor.Text, Liblist)
    If cboxModVersion.Value = 0 Then
        Call WriteINIStr("General", "Version", txtModVersion.Text, Liblist)
    Else
        Call WriteINIStr("General", "Version", PadNum(Year(Now())) & "." & PadNum(Month(Now()), 2) & "." & PadNum(Day(Now()), 2) & "." & PadNum(Hour(Now()), 2) & "." & PadNum(Minute(Now()), 2), Liblist)
    End If
    Call WriteINIStr("General", "Date", txtModDateYear.Text & "-" & txtModDateMonth.Text & "-" & txtModDateDay.Text, Liblist)
    Call WriteINIStr("General", "Description", txtModDescription.Text, Liblist)
    Call WriteINIStr("General", "Campaigns", comboModCampaigns.Text, Liblist)
    Call WriteINIStr("General", "Website", txtModWebsite.Text, Liblist)
    Call WriteINIStr("General", "UpdateCheckURL", txtModUpdateCheck.Text, Liblist)
    Select Case comboModType.ListIndex
    Case 0
        Call WriteINIStr("General", "ModType", "mod", Liblist)
        Call WriteINIStr("General", "ModScrnFormat", txtModScrnFormat.Text, Liblist)
        Call WriteINIStr("General", "ModSnapFormat", txtModSnapFormat.Text, Liblist)
        If comboSide3Mix.ListIndex = 1 Then Call WriteINIStr("General", "UseYuriUI", "yes", Liblist)
        Call WriteINIStr("General", "GameMode", txtGameMode.Text, Liblist)
        Call WriteINIStr("General", "MapIndex", txtMapIndex.Text, Liblist)
        If cboxUseAres.Value = 1 Then Call WriteINIStr("General", "UseAres", "yes", Liblist)
    Case 1
        Call WriteINIStr("General", "ModType", "fa2mod", Liblist)
        Select Case comboFA2.ListIndex
        Case 1: Call WriteINIStr("General", "FA2Version", "1.0.0.1", Liblist)
        Case 2: Call WriteINIStr("General", "FA2Version", "1.0.0.2", Liblist)
        End Select
    Case 2
        Call WriteINIStr("General", "ModType", "tool", Liblist)
        Call WriteINIStr("General", "Program", comboModProgram.Text, Liblist)
        If cboxParams.Value = 1 Then Call WriteINIStr("General", "ShowParams", "yes", Liblist)
    Case 3
        Call WriteINIStr("General", "ModType", "plugin", Liblist)
        Call WriteINIStr("General", "PluginID", comboPluginID.Text, Liblist)
    End Select
    If optTX(0).Value = True Then Call WriteINIStr("General", "AllowTX", "yes", Liblist)
    If optTX(1).Value = True Then
        Call WriteINIStr("General", "AllowTX", "yes", Liblist)
        Select Case txtTX.Text
        Case "": Call WriteINIStr("General", "TXVersion", DefaultVersionTX, Liblist)
        Case Else: Call WriteINIStr("General", "TXVersion", txtTX.Text, Liblist)
        End Select
    End If
    If optTX(2).Value = True Then Call WriteINIStr("General", "AllowTX", "no", Liblist)
    If cboxForRA2.Value = 1 Then Call WriteINIStr("General", "IsForRA2", "yes", Liblist)
    Exit Sub
LocalErr:
    Call GlobalErr("SaveLiblist", ErrVars())
End Sub

Private Sub LoadSettings(ByVal PathToLoad As String, Optional ByVal UpdateTitleBar As Boolean = True)
    Dim Counter As Long
    Dim MissingFileCount As Long
    Dim TempString As String
    Dim DummyStringArray(0) As String
    Dim ErrVars(1) As Variant
    If GetArgByName("noexcept") = "" Then On Error GoTo LocalErr
    ErrVars(1) = PathToLoad
    If Val(ReadINIStr("LBP", "LBPVersion", PathToLoad, "0")) <> 3 Then
        MsgBoxResult = MsgBox("Invalid LBP file format!" & vbCrLf & "This version of " & App.Title & " cannot read LBP files created with version " & ReadINIStr("LBP", "LBMCVersion", PathToLoad, "0.99") & ".", vbOKOnly + vbInformation, App.Title)
    Else
        LoopPrevention = True
        'Options
        picModBanner.Tag = ReadINIStr("Options", "BannerImage", PathToLoad)
        Call LoadBannerImage(False)
        txtModName.Text = ReadINIStr("Options", "ModName", PathToLoad)
        txtModAuthor.Text = ReadINIStr("Options", "ModAuthor", PathToLoad)
        txtModDateYear.Text = ReadINIStr("Options", "ModDateYear", PathToLoad)
        txtModDateMonth.Text = ReadINIStr("Options", "ModDateMonth", PathToLoad)
        txtModDateDay.Text = ReadINIStr("Options", "ModDateDay", PathToLoad)
        cboxModDate.Value = Val(ReadINIStr("Options", "ModDateBox", PathToLoad))
        cboxModVersion.Value = Val(ReadINIStr("Options", "ModVersionBox", PathToLoad))
        Call cboxModDate_Click
        Call cboxModVersion_Click
        If cboxModVersion.Value = 0 Then
            txtModVersion.Text = ReadINIStr("Options", "ModVersion", PathToLoad)
        End If
        txtModDescription.Text = ReadINIStr("Options", "ModDescription", PathToLoad)
        comboModCampaigns.Text = ReadINIStr("Options", "ModCampaigns", PathToLoad)
        comboModType.ListIndex = Val(ReadINIStr("Options", "ModType", PathToLoad))
        txtModWebsite.Text = ReadINIStr("Options", "ModWebsite", PathToLoad)
        txtModUpdateCheck.Text = ReadINIStr("Options", "ModUpdateCheck", PathToLoad)
        txtModScrnFormat.Text = ReadINIStr("Options", "ModScrnFormat", PathToLoad)
        txtModSnapFormat.Text = ReadINIStr("Options", "ModSnapFormat", PathToLoad)
        txtModDisplaySound.Text = ReadINIStr("Options", "ModDisplaySound", PathToLoad)
        txtModLaunchSound.Text = ReadINIStr("Options", "ModLaunchSound", PathToLoad)
        optTX(Val(ReadINIStr("Options", "TXOpt", PathToLoad))).Value = True
        txtTX.Text = ReadINIStr("Options", "TXVersion", PathToLoad)
        comboFA2.ListIndex = (Val(ReadINIStr("Options", "FA2Opt", PathToLoad, "0")))
        cboxParams.Value = Val(ReadINIStr("Options", "Params", PathToLoad))
        cboxForRA2.Value = Val(ReadINIStr("Options", "IsForRA2", PathToLoad))
        cboxShutdownLB.Value = Val(ReadINIStr("Options", "ShutdownLB", PathToLoad))
        cmdKey.Tag = ReadINIStr("Options", "SecurityKey", PathToLoad)
        txtGameMode.Text = ReadINIStr("Options", "GameMode", PathToLoad, "1")
        txtMapIndex.Text = ReadINIStr("Options", "MapIndex", PathToLoad, "0")
        comboPluginID.Text = ReadINIStr("Options", "PluginID", PathToLoad)
        'Files
        comboMixEncrypt.ListIndex = Val(ReadINIStr("Files", "MixEncrypt", PathToLoad))
        comboSide3Mix.ListIndex = Val(ReadINIStr("Files", "Side3Mix", PathToLoad))
        txtProgramDirectory.Text = ReadINIStr("Files", "ProgramDirectory", PathToLoad)
        txtCustomScript.Text = ReadINIStr("Files", "CustomScript", PathToLoad)
        txtCustomScriptMessage.Text = ReadINIStr("Files", "CustomScriptMessage", PathToLoad)
        Call frmUninstallFiles.lvUninstallFiles.ListItems.Clear
        Counter = 1
        TempString = ReadINIStr("Files", "UninstallFile" & CStr(Counter), PathToLoad)
        Do While Len(TempString) <> 0
            If Len(TempString) <> 0 Then
                If Len(StripInvalidChars(TempString, StripInvalidChars(InvalidFileChars, "\"))) Then
                    Call frmUninstallFiles.lvUninstallFiles.ListItems.Add(, , TempString)
                End If
            End If
            Counter = Counter + 1
            TempString = ReadINIStr("Files", "UninstallFile" & CStr(Counter), PathToLoad)
        Loop
        cboxUseAres.Value = BooleanStringToInteger(ReadINIStr("Files", "UseAres", PathToLoad, "no"))
        'Installer
        picInstallerIcon.Tag = ReadINIStr("Installer", "InstallerIcon", PathToLoad)
        cboxWindowIcon.Value = Val(ReadINIStr("Installer", "WindowIcon", PathToLoad))
        cboxOldSaves.Value = Val(ReadINIStr("Installer", "OldSaves", PathToLoad))
        cboxResetGameConfig.Value = Val(ReadINIStr("Installer", "ResetGameConfig", PathToLoad))
        cboxCRC.Value = Val(ReadINIStr("Installer", "CRC", PathToLoad, "0"))
        Call LoadInstallerIcon(False)
        txtINSTDIR.Text = ReadINIStr("Installer", "Instdir", PathToLoad)
        comboDSP.ListIndex = Val(ReadINIStr("Installer", "DSP", PathToLoad))
        comboInfoPage.ListIndex = Val(ReadINIStr("Installer", "ShowInfoPage", PathToLoad))
        txtInfoPageTitle.Text = ReadINIStr("Installer", "InfoPageTitle", PathToLoad, "Information")
        txtInfoPageButton.Text = ReadINIStr("Installer", "InfoPageButton", PathToLoad, "Information")
        'INFOPAGE TEXT START
        txtInfoPageText.Text = ""
        MissingFileCount = Val(ReadINIStr("InfoPageText", "LastLine", PathToLoad))
        Counter = 0
        TempString = ""
        Do While Counter <= MissingFileCount
            TempString = TempString & ReadINIStr("InfoPageText", CStr(Counter), PathToLoad)
            Counter = Counter + 1
        Loop
        TempString = ReplaceString(TempString, "$\n", vbCrLf)
        TempString = ReplaceString(TempString, "$\t", vbTab)
        txtInfoPageText.Text = TempString
        'INFOPAGE TEXT END
        cboxShowInstDetails.Value = Val(ReadINIStr("Installer", "ShowInstDetails", PathToLoad))
        cboxAutoClose.Value = Val(ReadINIStr("Installer", "AutoClose", PathToLoad))
        cboxXPStyle.Value = Val(ReadINIStr("Installer", "XPStyle", PathToLoad))
        For Counter = 0 To 2
            txtTextColour(Counter).Text = ReadINIStr("Installer", "TextColour" & CStr(Counter), PathToLoad)
            txtBGColour(Counter).Text = ReadINIStr("Installer", "BGColour" & CStr(Counter), PathToLoad)
        Next Counter
        Call ApplyTextColour
        Call ApplyBGColour
        'Compression
        comboCompressionMethod.ListIndex = Val(ReadINIStr("Compression", "CompressionMethod", PathToLoad))
        cboxSolid.Value = Val(ReadINIStr("Compression", "Solid", PathToLoad))
        cboxGenPat.Value = Val(ReadINIStr("Compression", "GenPat", PathToLoad))
        cboxFLAC.Value = Val(ReadINIStr("Compression", "FLAC", PathToLoad, "1"))
        sliderPatchMinSize.Value = Val(ReadINIStr("Compression", "PatchMinSize", PathToLoad))
        sliderPatchBlockSize.Value = Val(ReadINIStr("Compression", "PatchBlockSize", PathToLoad))
        Call sliderPatchMinSize_Scroll
        Call sliderPatchBlockSize_Scroll
        txtMarblePath.Text = ReadINIStr("Compression", "MarblePath", PathToLoad)
        txtMarblePath.Tag = txtMarblePath.Text
        cboxUpdateOnly.Value = Val(ReadINIStr("Compression", "UpdateOnly", PathToLoad))
        txtUpdateOnlySource.Text = ReadINIStr("Compression", "PreviousInst", PathToLoad)
        txtUpdateOnlySource.Tag = txtUpdateOnlySource.Text
        txtUpdateOnlyDest.Text = ReadINIStr("Compression", "LatestInst", PathToLoad)
        txtUpdateOnlyDest.Tag = txtUpdateOnlyDest.Text
        'Update controls
        Call UpdateControls
        'Load files
        InstFileCount = 0
        Counter = 0
        MissingFileCount = 0
        TempString = ReadINIStr("Files", CStr(Counter), PathToLoad)
        Do While TempString <> ""
            Call AddInstFile(TempString, treeFolders.Nodes.Item(ReadINIStr("FileNodeIDs", CStr(Counter), PathToLoad)), MissingFileCount)
            Counter = Counter + 1
            TempString = ReadINIStr("Files", CStr(Counter), PathToLoad)
        Loop
        If MissingFileCount <> 0 Then
            MsgBoxResult = MsgBox(CStr(MissingFileCount) & " files have not been loaded because they could not be found." & vbCrLf & "They have either been moved, renamed or deleted since this project was last saved.", vbOKOnly + vbInformation, App.Title)
        End If
        Call RefreshInstFiles(DummyStringArray())
        Call RefreshGameModeList
        Call RefreshProgramList
        TempString = ReadINIStr("Options", "ProgramText", PathToLoad)
        If comboModProgram.ListCount <> 0 Then
            For Counter = 0 To comboModProgram.ListCount - 1
                If comboModProgram.List(Counter) = TempString Then
                    comboModProgram.ListIndex = Counter
                    Counter = comboModProgram.ListCount - 1
                End If
            Next Counter
        End If
        frmMain.Caption = GetFileName(PathToLoad) & " - " & App.Title
        LoopPrevention = False
    End If
    Exit Sub
LocalErr:
    Call GlobalErr("LoadSettings", ErrVars())
End Sub

Private Sub SaveSettings_QuickSort(ByRef sArray() As String, ByRef sArray2() As String)
    Dim sPivot As String
    Dim iMax As Long
    Dim iPos As Long
    Dim iLess As Long
    Dim iMore As Long
    Dim sLess() As String
    Dim sMore() As String
    Dim sLess2() As String
    Dim sMore2() As String
    If Not Not sArray And Not Not sArray2 Then 'validation
        If UBound(sArray()) = UBound(sArray2()) Then 'validation
            iMax = UBound(sArray())
            If iMax <> 0 Then
                sPivot = sArray(iMax \ 2)
                ReDim sLess(iMax)
                ReDim sLess2(iMax)
                ReDim sMore(iMax)
                ReDim sMore2(iMax)
                iPos = 0
                iLess = -1
                iMore = -1
                Do While iPos <= iMax
                    If sArray(iPos) <= sPivot Then
                        iLess = iLess + 1
                        sLess(iLess) = sArray(iPos)
                        sLess2(iLess) = sArray2(iPos) 'mimic changes in first array
                    Else
                        iMore = iMore + 1
                        sMore(iMore) = sArray(iPos)
                        sMore2(iMore) = sArray2(iPos) 'mimic changes in first array
                    End If
                    iPos = iPos + 1
                Loop
                'sort sLess() and sMore()
                If iLess > 0 And iLess <> iMax Then
                    ReDim Preserve sLess(iLess)
                    ReDim Preserve sLess2(iLess)
                    Call SaveSettings_QuickSort(sLess(), sLess2())
                End If
                If iMore > 0 And iMore <> iMax Then
                    ReDim Preserve sMore(iMore)
                    ReDim Preserve sMore2(iMore)
                    Call SaveSettings_QuickSort(sMore(), sMore2())
                End If
                'now put back into the passed array
                iPos = 0
                Do While iPos <= iLess
                    sArray(iPos) = sLess(iPos)
                    sArray2(iPos) = sLess2(iPos)
                    iPos = iPos + 1
                Loop
                iLess = 0 'hijacking this to represent position in sMore()
                Do While iLess <= iMore
                    sArray(iPos) = sMore(iLess)
                    sArray2(iPos) = sMore2(iLess)
                    iPos = iPos + 1
                    iLess = iLess + 1
                Loop
            End If
        End If 'validation
    End If 'validation
End Sub

Private Sub SaveSettings_RemoveInstFiles(ByRef sArray() As String, ByRef sArray2() As String)
    Dim iCounter As Integer
    Dim iTotal As Integer
    Dim iNode As Integer
    Dim bOk As Boolean
    Dim iNewCounter As Integer
    iCounter = 0
    iTotal = -1
    Do While iCounter < InstFileCount
        If Len(sArray2(iCounter)) <> 0 Then
            iNode = 1
            Do While iNode <= treeFolders.Nodes.Count
                If treeFolders.Nodes.Item(iNode).Key = sArray2(iCounter) Then
                    If Len(sArray(iCounter)) <> 0 Then
                        'this is valid
                        iTotal = iTotal + 1
                        If iTotal <> iCounter Then
                            'we have skipped at least one record
                            sArray(iTotal) = sArray(iCounter)
                            sArray2(iTotal) = sArray2(iCounter)
                        End If
                    End If
                End If
                iNode = iNode + 1
            Loop
        End If
        iCounter = iCounter + 1
    Loop
    iCounter = iCounter - 1 'we've gone one too far
    If iTotal <> iCounter Then
        'some records have been removed
        ReDim Preserve sArray(iTotal)
        ReDim Preserve sArray2(iTotal)
        InstFileCount = iTotal + 1
    End If
End Sub

Private Sub SaveSettings(ByVal PathToSave As String)
    Dim Counter As Long
    Dim Counter2 As Long
    Dim DummyStringArray(0) As String
    Dim StringArray() As String
    Dim TargetNode As Node
    Dim ErrVars(1) As Variant
    If GetArgByName("noexcept") = "" Then On Error GoTo LocalErr
    ErrVars(1) = PathToSave
    'first, reorder the InstFile arrays and strip out blank entries
    'THIS IS NEEDED BY THE GENERATOR ELSE FILES WON'T BE EXTRACTED/CHECKSUM'D IN THE CORRECT ORDER
    Call SaveSettings_RemoveInstFiles(InstFile(), InstFileNodeID())
    Call SaveSettings_QuickSort(InstFile(), InstFileNodeID())
    'now begin saving
    If FileExists(PathToSave) Then Call Kill(PathToSave)
    Call WriteINIStr("LBP", "LBMCVersion", ReadINIStr("General", "Version", JoinPath(EXEDIR, "launcher", "liblist.gam")), PathToSave)
    Call WriteINIStr("LBP", "LBPVersion", "3", PathToSave)
    'Options
    Call WriteINIStr("Options", "BannerImage", picModBanner.Tag, PathToSave)
    Call WriteINIStr("Options", "ModName", txtModName.Text, PathToSave)
    Call WriteINIStr("Options", "ModAuthor", txtModAuthor.Text, PathToSave)
    Call WriteINIStr("Options", "ModVersion", txtModVersion.Text, PathToSave)
    Call WriteINIStr("Options", "ModVersionBox", CStr(cboxModVersion.Value), PathToSave)
    Call WriteINIStr("Options", "ModDateBox", CStr(cboxModDate.Value), PathToSave)
    Call WriteINIStr("Options", "ModDateYear", txtModDateYear.Text, PathToSave)
    Call WriteINIStr("Options", "ModDateMonth", txtModDateMonth.Text, PathToSave)
    Call WriteINIStr("Options", "ModDateDay", txtModDateDay.Text, PathToSave)
    Call WriteINIStr("Options", "ModDescription", txtModDescription.Text, PathToSave)
    Call WriteINIStr("Options", "ModCampaigns", comboModCampaigns.Text, PathToSave)
    Call WriteINIStr("Options", "ModType", CStr(comboModType.ListIndex), PathToSave)
    Call WriteINIStr("Options", "ModWebsite", txtModWebsite.Text, PathToSave)
    Call WriteINIStr("Options", "ModUpdateCheck", txtModUpdateCheck.Text, PathToSave)
    Call WriteINIStr("Options", "ModScrnFormat", txtModScrnFormat.Text, PathToSave)
    Call WriteINIStr("Options", "ModSnapFormat", txtModSnapFormat.Text, PathToSave)
    Call WriteINIStr("Options", "ModDisplaySound", txtModDisplaySound.Text, PathToSave)
    Call WriteINIStr("Options", "ModLaunchSound", txtModLaunchSound.Text, PathToSave)
    Call WriteINIStr("Options", "IsForRA2", CStr(cboxForRA2.Value), PathToSave)
    Call WriteINIStr("Options", "ShutdownLB", CStr(cboxShutdownLB.Value), PathToSave)
    Call WriteINIStr("Options", "PluginID", comboPluginID.Text, PathToSave)
    For Counter = 0 To 3
        If optTX(Counter).Value = True Then Call WriteINIStr("Options", "TXOpt", CStr(Counter), PathToSave)
    Next Counter
    Call WriteINIStr("Options", "TXVersion", txtTX.Text, PathToSave)
    Call WriteINIStr("Options", "FA2Opt", CStr(comboFA2.ListIndex), PathToSave)
    Call WriteINIStr("Options", "ProgramText", comboModProgram.List(comboModProgram.ListIndex), PathToSave)
    Call WriteINIStr("Options", "Params", CStr(cboxParams.Value), PathToSave)
    Call WriteINIStr("Options", "SecurityKey", cmdKey.Tag, PathToSave)
    Call WriteINIStr("Options", "GameMode", txtGameMode.Text, PathToSave)
    Call WriteINIStr("Options", "MapIndex", txtMapIndex.Text, PathToSave)
    'Files
    Counter = 0
    Do While Counter < InstFileCount
        Call WriteINIStr("Files", CStr(Counter), InstFile(Counter), PathToSave)
        Call WriteINIStr("FileNodeIDs", CStr(Counter), InstFileNodeID(Counter), PathToSave)
        Counter = Counter + 1
    Loop
    Call RefreshGameModeList
    Call RefreshProgramList 'this used to be inside RemoveInstFile
    Call RefreshInstFiles(DummyStringArray()) 'in case we have removed any files from memory
    Call WriteINIStr("Files", "MixEncrypt", CStr(comboMixEncrypt.ListIndex), PathToSave)
    Call WriteINIStr("Files", "Side3Mix", CStr(comboSide3Mix.ListIndex), PathToSave)
    Call WriteINIStr("Files", "ProgramDirectory", txtProgramDirectory.Text, PathToSave)
    Call WriteINIStr("Files", "CustomScript", txtCustomScript.Text, PathToSave)
    Call WriteINIStr("Files", "CustomScriptMessage", txtCustomScriptMessage.Text, PathToSave)
    Call WriteINIStr("Files", "UseAres", IntegerToBoolean(cboxUseAres.Value), PathToSave)
    Counter = 1
    Do While Counter <= frmUninstallFiles.lvUninstallFiles.ListItems.Count
        Call WriteINIStr("Files", "UninstallFile" & CStr(Counter), frmUninstallFiles.lvUninstallFiles.ListItems.Item(Counter).Text, PathToSave)
        Counter = Counter + 1
    Loop
    'Installer
    Call WriteINIStr("Installer", "InstallerIcon", picInstallerIcon.Tag, PathToSave)
    Call WriteINIStr("Installer", "WindowIcon", CStr(cboxWindowIcon.Value), PathToSave)
    Call WriteINIStr("Installer", "OldSaves", CStr(cboxOldSaves.Value), PathToSave)
    Call WriteINIStr("Installer", "ResetGameConfig", CStr(cboxResetGameConfig.Value), PathToSave)
    Call WriteINIStr("Installer", "CRC", CStr(cboxCRC.Value), PathToSave)
    Call WriteINIStr("Installer", "Instdir", txtINSTDIR.Text, PathToSave)
    Call WriteINIStr("Installer", "DSP", CStr(comboDSP.ListIndex), PathToSave)
    Call WriteINIStr("Installer", "ShowInfoPage", CStr(comboInfoPage.ListIndex), PathToSave)
    Call WriteINIStr("Installer", "InfoPageTitle", txtInfoPageTitle.Text, PathToSave)
    Call WriteINIStr("Installer", "InfoPageButton", txtInfoPageButton.Text, PathToSave)
    'INFOPAGE TEXT START
    DummyStringArray(0) = ReplaceString(txtInfoPageText.Text, vbCrLf, "$\n")
    DummyStringArray(0) = ReplaceString(DummyStringArray(0), vbTab, "$\t")
    Counter2 = Len(DummyStringArray(0)) \ 254
    Counter = 0
    Call WriteINIStr("InfoPageText", "LastLine", CStr(Counter2), PathToSave)
    Do While Counter < Counter2
        Call WriteINIStr("InfoPageText", CStr(Counter), Mid$(DummyStringArray(0), (Counter * 254) + 1, 254), PathToSave)
        Counter = Counter + 1
    Loop
    If Len(DummyStringArray(0)) Mod 254 <> 0 Then
        Call WriteINIStr("InfoPageText", CStr(Counter), Mid$(DummyStringArray(0), (Counter * 254) + 1), PathToSave)
    Else
        Call WriteINIStr("InfoPageText", CStr(Counter), "", PathToSave)
    End If
    DummyStringArray(0) = ""
    'INFOPAGE TEXT END
    Call WriteINIStr("Installer", "ShowInstDetails", CStr(cboxShowInstDetails.Value), PathToSave)
    Call WriteINIStr("Installer", "AutoClose", CStr(cboxAutoClose.Value), PathToSave)
    Call WriteINIStr("Installer", "XPStyle", CStr(cboxXPStyle.Value), PathToSave)
    For Counter = 0 To 2
        Call WriteINIStr("Installer", "TextColour" & CStr(Counter), txtTextColour(Counter).Text, PathToSave)
        Call WriteINIStr("Installer", "BGColour" & CStr(Counter), txtBGColour(Counter).Text, PathToSave)
    Next Counter
    Call WriteINIStr("Compression", "CompressionMethod", CStr(comboCompressionMethod.ListIndex), PathToSave)
    Call WriteINIStr("Compression", "Solid", CStr(cboxSolid.Value), PathToSave)
    Call WriteINIStr("Compression", "GenPat", CStr(cboxGenPat.Value), PathToSave)
    Call WriteINIStr("Compression", "FLAC", CStr(cboxFLAC.Value), PathToSave)
    Call WriteINIStr("Compression", "PatchMinSize", CStr(sliderPatchMinSize.Value), PathToSave)
    Call WriteINIStr("Compression", "PatchBlockSize", CStr(sliderPatchBlockSize.Value), PathToSave)
    Call WriteINIStr("Compression", "MarblePath", txtMarblePath.Text, PathToSave)
    Call WriteINIStr("Compression", "UpdateOnly", CStr(cboxUpdateOnly.Value), PathToSave)
    Call WriteINIStr("Compression", "PreviousInst", txtUpdateOnlySource.Text, PathToSave)
    Call WriteINIStr("Compression", "LatestInst", txtUpdateOnlyDest.Text, PathToSave)
    Exit Sub
LocalErr:
    Call GlobalErr("SaveSettings", ErrVars())
End Sub

'**********************************************************
'******************** FILES TREE STUFF ********************
'**********************************************************

Private Sub InitNodes()
    Call treeFolders.Nodes.Clear
    Call treeFolders.Nodes.Add(, , "Node0", txtINSTDIR.Text, "ClosedFolder", "OpenFolder")
    treeFolders.Nodes.Item("Node0").Expanded = True
    Select Case comboModType.ListIndex
    Case 0
        Call treeFolders.Nodes.Add("Node0", tvwChild, "cameo", "cameo", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "hva", "hva", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "ini", "ini", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "interface", "interface", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "manual", "manual", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "map", "map", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "mix", "mix", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "screen", "screen", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "shp", "shp", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "side 1", "side 1", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "side 2", "side 2", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "side 3", "side 3", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "side 4", "side 4", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "sound", "sound", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "speech", "speech", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "string table", "string table", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "syringe", "syringe", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "taunts", "taunts", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "theme", "theme", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "video", "video", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "vxl", "vxl", "ClosedFolder", "OpenFolder")
    Case 1
        Call treeFolders.Nodes.Add("Node0", tvwChild, "fa2files", "fa2files", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "hva", "hva", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "ini", "ini", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "manual", "manual", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "map", "map", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "mix", "mix", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "shp", "shp", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "string table", "string table", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "video", "video", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "vxl", "vxl", "ClosedFolder", "OpenFolder")
    Case 2
        Call treeFolders.Nodes.Add("Node0", tvwChild, "manual", "manual", "ClosedFolder", "OpenFolder")
    Case 3
        Call treeFolders.Nodes.Add("Node0", tvwChild, "cameo", "cameo", "ClosedFolder", "OpenFolder")
        If optTX(3).Value = True Then Call treeFolders.Nodes.Add("Node0", tvwChild, "fa2files", "fa2files", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "hva", "hva", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "ini", "ini", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "interface", "interface", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "manual", "manual", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "map", "map", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "mix", "mix", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "screen", "screen", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "shp", "shp", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "side 1", "side 1", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "side 2", "side 2", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "side 3", "side 3", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "side 4", "side 4", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "speech", "speech", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "taunts", "taunts", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "theme", "theme", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "video", "video", "ClosedFolder", "OpenFolder")
        Call treeFolders.Nodes.Add("Node0", tvwChild, "vxl", "vxl", "ClosedFolder", "OpenFolder")
    End Select
    treeFolders.SelectedItem = treeFolders.Nodes.Item("Node0")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then '0 = user clicked X, 1 = programatically, 2 = Windows
        Call Shutdown
    End If
End Sub

Private Sub listFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If listFiles.SortKey <> (ColumnHeader.Index - 1) Then
        listFiles.SortKey = ColumnHeader.Index - 1
    Else
        If listFiles.SortOrder = lvwAscending Then
            listFiles.SortOrder = lvwDescending
        Else
            listFiles.SortOrder = lvwAscending
        End If
    End If
End Sub

Private Sub listFiles_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Counter As Integer
    Dim DummyStringArray(0) As String
    Dim bWeDeletedStuff As Boolean
    If KeyCode = 46 Then
        bWeDeletedStuff = False
        Counter = 1
        Do While Counter <= listFiles.ListItems.Count
            If listFiles.ListItems(Counter).Selected Then
                listFiles.ListItems(Counter).Selected = False
                Call RemoveInstFile(Val(listFiles.ListItems(Counter).Tag))
                bWeDeletedStuff = True
            End If
            Counter = Counter + 1
        Loop
        If bWeDeletedStuff Then
            Call RefreshGameModeList
            Call RefreshProgramList
        End If
        Call RefreshInstFiles(DummyStringArray())
    End If
End Sub

Private Sub RemoveInstFile(ByVal FileNum As Integer)
    'Do While FileNum < (InstFileCount - 1)
    '    InstFile(FileNum) = InstFile(FileNum + 1)
    '    InstFileNodeID(FileNum) = InstFileNodeID(FileNum + 1)
    '    FileNum = FileNum + 1
    'Loop
    'InstFileCount = InstFileCount - 1
    'ReDim Preserve InstFile(InstFileCount - 1)
    'ReDim Preserve InstFileNodeID(InstFileCount - 1)
    InstFile(FileNum) = ""
    InstFileNodeID(FileNum) = ""
End Sub

Private Sub listFiles_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Counter As Integer
    Dim StringArray() As String
    If data.GetFormat(vbCFFiles) Then
        ReDim StringArray(data.Files.Count)
        Counter = 1
        Do While Counter <= data.Files.Count
            StringArray(Counter) = data.Files.Item(Counter)
            Counter = Counter + 1
        Loop
        Call AddInstFiles(StringArray())
    End If
End Sub

Private Sub listFiles_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Call listFiles.SetFocus
End Sub

Private Sub sliderPatchBlockSize_Scroll()
    lblPatchBlockSize.Caption = 2 ^ (4 + sliderPatchBlockSize.Value)
End Sub

Private Sub sliderPatchMinSize_Scroll()
    lblPatchMinSize.Caption = DataSize((sliderPatchMinSize.Value * 16) * 1024)
End Sub

Private Sub treeFolders_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Counter As Integer
    Dim StringArray() As String
    If data.GetFormat(vbCFFiles) Then
        If data.Files.Count > 0 Then
            ReDim StringArray(data.Files.Count)
            For Counter = 1 To data.Files.Count
                StringArray(Counter) = data.Files.Item(Counter)
            Next Counter
            Call AddInstFiles(StringArray())
        End If
    End If
End Sub

Private Sub optTX_Click(Index As Integer)
    Call UpdateControls
    If Index = 3 Then comboPluginID.Text = "TX"
End Sub

Private Sub SSTab1_OLEDragOver(data As TabDlg.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Y <> 0 Then
        If X <> 0 Then
            If Y <= (SSTab1.TabHeight) Then
                If X <= SSTab1.Width \ SSTab1.TabsPerRow Then
                    If SSTab1.Tab <> 0 Then SSTab1.Tab = 0
                Else
                    If X <= (SSTab1.Width * 2) \ SSTab1.TabsPerRow Then
                        If SSTab1.Tab <> 1 Then SSTab1.Tab = 1
                    Else
                        If X <= (SSTab1.Width * 3) \ SSTab1.TabsPerRow Then
                            If SSTab1.Tab <> 2 Then SSTab1.Tab = 2
                        Else
                            If SSTab1.Tab <> 3 Then SSTab1.Tab = 3
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub treeFolders_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim DummyStringArray(0) As String
    Call RefreshInstFiles(DummyStringArray())
End Sub

Private Sub treeFolders_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Dim TargetNode As Node
    Dim Ok As Boolean
    Dim DummyStringArray(0) As String
    If X <> 0 And Y <> 0 Then 'otherwise we're off the control
        Ok = True
        Set TargetNode = treeFolders.HitTest(X, Y)
        If TargetNode Is Nothing Then
            'do nothing
        Else
            If treeFolders.SelectedItem Is Nothing Then
                Ok = False
            Else
                If treeFolders.SelectedItem <> TargetNode Then
                    Ok = False
                End If
            End If
            If Ok = False Then
                treeFolders.SelectedItem = TargetNode
                Call RefreshInstFiles(DummyStringArray())
            End If
        End If
    End If
End Sub

Private Sub AddInstFiles(ByRef StringArray() As String)
    Dim Counter As Integer
    Dim DummyLong As Long
    Dim TargetNode As Node
    Set TargetNode = treeFolders.SelectedItem
    If UBound(StringArray) > 0 Then
        If Not TargetNode Is Nothing Then
            For Counter = 1 To UBound(StringArray)
                If DirExists(StringArray(Counter)) Then
                    MsgBoxResult = MsgBox("You cannot include directories in any category folders.", vbOKOnly + vbInformation, App.Title)
                    Exit Sub
                End If
            Next Counter
            For Counter = 1 To UBound(StringArray)
                Call AddInstFile(StringArray(Counter), TargetNode, DummyLong)
            Next Counter
        End If
        Call RefreshInstFiles(StringArray())
        Call RefreshGameModeList
        Call RefreshProgramList
    End If
End Sub

Private Sub AddInstFile(ByVal FilePath As String, ByVal TargetNode As Node, ByRef MissingFileCount As Long)
    Dim Counter As Integer
    Dim Ok As Boolean
    If Not FileExists(FilePath) Then
        MissingFileCount = MissingFileCount + 1
    Else
        'check if file already present in folder
        Ok = True
        Counter = 0
        Do While Counter < InstFileCount
            If UCase(GetFileName(InstFile(Counter))) = UCase(GetFileName(FilePath)) And InstFileNodeID(Counter) = TargetNode.Key Then
                InstFile(Counter) = FilePath
                Ok = False
                Exit Do
            End If
            Counter = Counter + 1
        Loop
        If Ok Then
            'file not already present so needs adding
            'find an empty slot
            Counter = 0
            Do While Counter < InstFileCount
                If Len(InstFile(Counter)) = 0 Then Exit Do
                Counter = Counter + 1
            Loop
            If Counter = InstFileCount Then
                'no empty slots so expand array
                ReDim Preserve InstFile(InstFileCount)
                ReDim Preserve InstFileNodeID(InstFileCount)
                InstFileCount = InstFileCount + 1
            End If
            InstFile(Counter) = FilePath
            InstFileNodeID(Counter) = TargetNode.Key
        End If
    End If
End Sub

Private Sub RefreshProgramList()
    Dim TempProgram As String
    Dim Counter As Integer
    TempProgram = comboModProgram.Text
    comboModProgram.Clear
    For Counter = 0 To (InstFileCount - 1)
        If InstFileNodeID(Counter) = "Node0" Then
            If FileType(InstFile(Counter)) = "EXE" Then Call comboModProgram.AddItem(GetFileName(InstFile(Counter)))
        End If
    Next Counter
    If DirExists(txtProgramDirectory.Text) Then
        filelistbox.Path = txtProgramDirectory.Text
        filelistbox.Pattern = "*"
        Call filelistbox.Refresh
        For Counter = 0 To (filelistbox.ListCount - 1)
            If FileType(filelistbox.List(Counter)) = "EXE" Then Call comboModProgram.AddItem(filelistbox.List(Counter))
        Next Counter
    End If
    For Counter = 0 To (comboModProgram.ListCount - 1)
        If UCase(comboModProgram.List(Counter)) = UCase(TempProgram) Then
            comboModProgram.ListIndex = Counter
            Counter = comboModProgram.ListCount - 1
        End If
    Next Counter
    If comboModProgram.ListIndex = -1 And comboModProgram.ListCount <> 0 Then comboModProgram.ListIndex = 0
End Sub

Private Sub RefreshGameModeList()
    Dim Counter As Integer
    Dim ModeCount As Integer
    Dim ModeDefn As String
    Dim CommaPos As Integer
    comboGameMode.Clear
    For Counter = 0 To (InstFileCount - 1)
        If InstFileNodeID(Counter) = "ini" Then
            Select Case LCase(GetFileName(InstFile(Counter)))
            Case "mpmodes.ini", "mpmodesmd.ini"
                If FileExists(InstFile(Counter)) Then
                    ModeCount = 1
                    Do
                        ModeDefn = ReadINIStr("Battle", CStr(ModeCount), InstFile(Counter))
                        If Len(ModeDefn) = 0 Then
                            ModeDefn = ReadINIStr("ManBattle", CStr(ModeCount), InstFile(Counter))
                            If Len(ModeDefn) = 0 Then
                                ModeDefn = ReadINIStr("FreeForAll", CStr(ModeCount), InstFile(Counter))
                                If Len(ModeDefn) = 0 Then
                                    ModeDefn = ReadINIStr("Unholy", CStr(ModeCount), InstFile(Counter))
                                    If Len(ModeDefn) = 0 Then
                                        ModeDefn = ReadINIStr("Cooperative", CStr(ModeCount), InstFile(Counter))
                                        If Len(ModeDefn) = 0 Then Exit Do
                                    End If
                                End If
                            End If
                        End If
                        If Len(ModeDefn) >= 5 Then
                            ModeDefn = Mid$(ModeDefn, 5)
                            CommaPos = InStr(1, ModeDefn, ",")
                            If CommaPos > 1 Then
                                ModeDefn = Left$(ModeDefn, CommaPos - 1)
                                Call comboGameMode.AddItem(CStr(ModeCount) & " " & ModeDefn)
                            Else
                                Exit Do
                            End If
                        Else
                            Exit Do
                        End If
                        ModeCount = ModeCount + 1
                    Loop
                    Exit For
                End If
            End Select
        End If
    Next Counter
    If comboGameMode.ListCount <> 0 Then
        txtGameMode.Visible = False
        comboGameMode.Visible = True
        comboGameMode.ListIndex = Val(txtGameMode.Text) - 1
    Else
        comboGameMode.Visible = False
        txtGameMode.Visible = True
    End If
End Sub

Private Function InstFileErrorCheck(ByVal FileNum As Long, Optional ByVal bAddError As Boolean = True) As Boolean
    Dim FileName As String
    Dim RetVal As Boolean
    RetVal = False
    FileName = GetFileName(InstFile(FileNum))
    If FileIsDestructive(FileName) Then
        If bAddError Then Call FileCheckAddError("If you need telling why then you shouldn't be releasing a mod yet.", FileNum)
    ElseIf FileIsReservedMix(FileName) Then
        If bAddError Then Call FileCheckAddError("This file is reserved for third-party community projects. Please use a mix file number higher than 10.", FileNum)
    ElseIf FileIsSoundtrack(FileName) Then
        If bAddError Then Call FileCheckAddError("Official soundtracks should not be distributed by mods.", FileNum)
    ElseIf FileIsUserTheme(FileName) Then
        If bAddError Then Call FileCheckAddError("User themes should not be distributed.", FileNum)
    ElseIf (FileIsOfficialMapPackMap(FileName) And comboModType.ListIndex <> 3) Then
        If bAddError Then Call FileCheckAddError("This is an official map pack map. Official map packs should not be distributed by mods.", FileNum)
    ElseIf FileIsAssaultMapPack(FileName) And comboModType.ListIndex <> 3 Then
        If bAddError Then Call FileCheckAddError("Assault map packs should not be distributed by mods.", FileNum)
    ElseIf StripInvalidChars(InstFile(FileNum), InvalidNSISChars) <> InstFile(FileNum) Then
        If bAddError Then Call FileCheckAddError("The following characters are not allowed in any file paths: " & InvalidNSISChars, FileNum)
    ElseIf comboModType.ListIndex = 1 And FileIsFA2File(InstFile(FileNum)) And UCase(InstFileNodeID(FileNum)) <> "FA2FILES" Then
        If bAddError Then Call FileCheckAddError("FA2 files can only be added to the " & Quote("fa2files") & " category.", FileNum)
    Else
        RetVal = True
        Select Case UCase(InstFileNodeID(FileNum))
        Case "CAMEO", "SHP":
            If Not FileType(FileName) = "SHP" Then
                If LCase(FileName) <> "mouse.sha" Then
                    If bAddError Then Call FileCheckAddError("Files in this category must be of type SHP.", FileNum)
                    RetVal = False
                End If
            End If
        Case "HVA":
            If Not FileType(FileName) = "HVA" Then
                If bAddError Then Call FileCheckAddError("Files in this category must be of type HVA.", FileNum)
                RetVal = False
            End If
        Case "INI":
            If Not FileType(FileName) = "INI" Then
                If bAddError Then Call FileCheckAddError("Files in this category must be of type INI.", FileNum)
                RetVal = False
            End If
        Case "MAP":
            If Not FileIsMap(FileName) Then
                If bAddError Then Call FileCheckAddError("Files in this category must be map files.", FileNum)
                RetVal = False
            End If
        Case "MIX":
            If Not FileType(FileName) = "MIX" Then
                If bAddError Then Call FileCheckAddError("Files in this category must be of type MIX.", FileNum)
                RetVal = False
            End If
        Case "STRING TABLE":
            Select Case FileType(FileName)
            Case "CSF", "TXT"
                RetVal = True
            Case Else
                If bAddError Then Call FileCheckAddError("Files in this category must be of type CSF or TXT.", FileNum)
                RetVal = False
            End Select
        Case "SOUND", "SPEECH", "TAUNTS", "THEME":
            Select Case FileType(FileName)
            Case "WAV", "OGG", "FLAC", "BAG", "IDX"
                RetVal = True
            Case Else
                If bAddError Then Call FileCheckAddError("Files in this category must be of type WAV, OGG, FLAC, BAG or IDX.", FileNum)
                RetVal = False
            End Select
        Case "VXL":
            If Not FileType(FileName) = "VXL" Then
                If bAddError Then Call FileCheckAddError("Files in this category must be of type VXL.", FileNum)
                RetVal = False
            End If
        Case "FA2FILES":
            If Not FileIsFA2File(FileName) And optTX(3).Value = False Then
                If bAddError Then Call FileCheckAddError("Only valid FinalAlert 2 files are allowed in this category: <FAData.ini>, <FALanguage.ini>, <marble.mix>", FileNum)
                RetVal = False
            Else
                If UCase(FileName) = "MARBLE.MIX" And optTX(1).Value = True Then
                    If bAddError Then Call FileCheckAddError("You cannot include <marble.mix> if your FinalAlert 2 mod requires/allows the Terrain Expansion because the Terrain Expansion includes it's own <marble.mix>.", FileNum)
                    RetVal = False
                Else
                    If LCase(FileName) = "marble.pat" Then Call FileCheckAddError("This file has the same filename as a file that will be automatically generated.", FileNum)
                End If
            End If
        Case "SYRINGE"
            Select Case FileType(FileName)
            Case "DLL"
                If cboxUseAres.Value = 1 And UCase$(FileName) = "ARES.DLL" Then
                    If bAddError Then Call FileCheckAddError("You have chosen to use the official Ares DLL but have also included your own version. Either untick 'Use Official Ares DLL' or remove your own version from your mod.", FileNum)
                    RetVal = False
                Else
                    If Not FileAddedByUser(FileName & ".inj") Then
                        If bAddError Then Call FileCheckAddError("You haven't included the corresponding " & Quote(FileName & ".inj") & " file for this DLL file.", FileNum)
                        RetVal = False
                    End If
                End If
            Case "INJ"
                If cboxUseAres.Value = 1 And UCase$(FileName) = "ARES.DLL.INJ" Then
                    If bAddError Then Call FileCheckAddError("You have chosen to use the official Ares DLL but have also included your own version. Either untick 'Use Official Ares DLL' or remove your own version from your mod.", FileNum)
                    RetVal = False
                Else
                    RetVal = False
                    If Len(FileName) > 8 Then
                        If FileType(Left$(FileName, Len(FileName) - 4)) = "DLL" Then
                            If FileAddedByUser(Left$(FileName, Len(FileName) - 4)) Then
                                RetVal = True
                            Else
                                If bAddError Then Call FileCheckAddError("You haven't included the corresponding " & Quote(Left$(FileName, Len(FileName) - 4)) & " file for this INJ file.", FileNum)
                            End If
                        Else
                            If bAddError Then Call FileCheckAddError("Invalid filename. INJ files should end with " & Quote(".dll.inj"), FileNum)
                        End If
                    Else
                        If bAddError Then Call FileCheckAddError("Invalid filename. INJ files should end with " & Quote(".dll.inj"), FileNum)
                    End If
                End If
            Case Else
                If bAddError Then Call FileCheckAddError("Files in this category must be of type DLL or INJ.", FileNum)
                RetVal = False
            End Select
        End Select
    End If
    InstFileErrorCheck = RetVal
End Function

Private Sub FileCheckAddError(ByVal Message As String, ByVal FileNum As Long)
    If LCase(InstFileNodeID(FileNum)) <> "node0" Then
        frmFileErrors.txtFileErrors.Text = frmFileErrors.txtFileErrors.Text & InstFile(FileNum) & " [" & InstFileNodeID(FileNum) & "]" & vbCrLf & Message & vbCrLf & vbCrLf
    Else
        frmFileErrors.txtFileErrors.Text = frmFileErrors.txtFileErrors.Text & InstFile(FileNum) & " [mod root]" & vbCrLf & Message & vbCrLf & vbCrLf
    End If
End Sub

Private Sub RefreshInstFiles(ByRef StringArray() As String)
    Dim Counter As Long
    Dim Counter2 As Long
    Dim FileNum As Long
    Dim MixMD As String
    Dim MixNum As String
    If cboxForRA2.Value = 1 Then MixMD = "" Else MixMD = "md"
    If optTX(3).Value = True Then MixNum = "06" Else MixNum = "98"
    If treeFolders.SelectedItem <> treeFolders.Nodes.Item("Node0") Then
        If comboMixEncrypt.ListIndex <> 0 Then
            Select Case UCase(treeFolders.SelectedItem.Text)
            Case "CAMEO", "SHP", "SPEECH", "TMP" 'ECACHEMD98.MIX
                If Not FileAddedByUser("ECACHE") Then lblFolders.Caption = treeFolders.Nodes.Item("Node0").Text & "\video\ecache" & MixMD & MixNum & ".mix"
            Case "HVA", "INI", "INTERFACE", "MAP", "MIX", "VXL" 'EXPANDMD98.MIX
                If Not FileAddedByUser("EXPAND") Then lblFolders.Caption = treeFolders.Nodes.Item("Node0").Text & "\video\expand" & MixMD & MixNum & ".mix"
            Case Else
                lblFolders.Caption = treeFolders.Nodes.Item("Node0").Text & "\" & treeFolders.SelectedItem.Text
            End Select
        Else
            lblFolders.Caption = treeFolders.Nodes.Item("Node0").Text & "\" & treeFolders.SelectedItem.Text
        End If
    Else
        lblFolders.Caption = treeFolders.Nodes.Item("Node0").Text
    End If
    listFiles.ListItems.Clear
    listFiles.Sorted = False
    frmFileErrors.txtFileErrors.Text = ""
    Counter = 0
    Do While Counter < InstFileCount
        If Len(InstFile(Counter)) <> 0 Then
            If InstFileNodeID(Counter) = treeFolders.SelectedItem.Key Then
                Call listFiles.ListItems.Add(, , GetFileName(InstFile(Counter)))
                listFiles.ListItems(listFiles.ListItems.Count).Tag = CStr(Counter)
                listFiles.ListItems(listFiles.ListItems.Count).SubItems(1) = LCase(FileType(InstFile(Counter)))
                listFiles.ListItems(listFiles.ListItems.Count).SubItems(2) = DataSize(GetFileSize(InstFile(Counter)))
                listFiles.ListItems(listFiles.ListItems.Count).SubItems(3) = GetFilePath(InstFile(Counter))
                If Not InstFileErrorCheck(Counter) Then
                    listFiles.ListItems(listFiles.ListItems.Count).Bold = True
                End If
            End If
        End If
        Counter = Counter + 1
    Loop
    listFiles.Sorted = True
    Counter = 1
    Do While Counter <= listFiles.ListItems.Count
        listFiles.ListItems(Counter).Selected = False
        FileNum = Val(listFiles.ListItems(Counter).Tag)
        Counter2 = 0
        Do While Counter2 <= UBound(StringArray())
            If InstFile(FileNum) = StringArray(Counter2) Then
                listFiles.ListItems(Counter).Selected = True
                Exit Do
            End If
            Counter2 = Counter2 + 1
        Loop
        Counter = Counter + 1
    Loop
    If listFiles.ListItems.Count <> 0 Then
        listFiles.SelectedItem = listFiles.ListItems(1)
        If UBound(StringArray()) = 0 Then listFiles.SelectedItem.Selected = False
    End If
    'alert user to file errors
    If Len(frmFileErrors.txtFileErrors) = 0 Then
        cmdFileErrors.Enabled = False
    Else
        cmdFileErrors.Enabled = True
        For Counter = 1 To 3
            cmdFileErrors.BackColor = &HFF&
            Me.Refresh
            Call Sleep(200)
            cmdFileErrors.BackColor = &H8000000F
            Me.Refresh
            If Counter <> 3 Then Call Sleep(200)
        Next Counter
    End If
End Sub

'**********************************************************
'********************CONTROL GETS FOCUS********************
'**********************************************************

Private Sub comboModCampaigns_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub picInstallerIcon_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then picInstallerIcon.BorderStyle = 1
End Sub

Private Sub picInstallerIcon_LostFocus()
    picInstallerIcon.BorderStyle = 0
End Sub

Private Sub txtCustomScriptMessage_Change()
    Dim sTemp As String
    Dim Counter As Integer
    sTemp = StripInvalidChars(txtCustomScriptMessage.Text, "`")
    If Len(sTemp) <> Len(txtCustomScriptMessage.Text) Then
        Call MsgBox("Custom Script Message must not include any grave accent characters (`).", vbOKOnly + vbInformation, "Invalid Character")
    End If
    sTemp = StripInvalidChars(sTemp, vbCrLf)
    If Len(sTemp) <> Len(txtCustomScriptMessage.Text) Then
        txtCustomScriptMessage.Text = sTemp
        txtCustomScriptMessage.SelStart = Len(txtCustomScriptMessage.Text)
    End If
End Sub

Private Sub txtGameMode_Change()
    Dim Stripped As String
    Stripped = StripNonNumbers(txtGameMode.Text)
    If Stripped <> txtGameMode.Text Then
        Call MsgBox("Game Mode must be set to an integer greater than or equal to 1.", vbOKOnly + vbInformation, "Invalid Character")
        txtGameMode.Text = Stripped
        SendKeys "{End}"
    End If
End Sub

Private Sub txtGameMode_LostFocus()
    txtGameMode.Text = StripLeadingZeroes(txtGameMode.Text)
    If txtGameMode.Text = "0" Then
        Call MsgBox("Game Mode must be greater than or equal to 1.", vbOKOnly + vbInformation, App.Title)
        txtGameMode.Text = "1"
    End If
End Sub

Private Sub txtMapIndex_Change()
    Dim Stripped As String
    Stripped = StripNonNumbers(txtMapIndex.Text)
    If Stripped <> txtMapIndex.Text Then
        MsgBoxResult = MsgBox("Map Index must be set to an integer greater than or equal to 0.", vbOKOnly + vbInformation, "Invalid Character")
        txtMapIndex.Text = Stripped
        SendKeys "{End}"
    End If
End Sub

Private Sub txtMapIndex_LostFocus()
    txtMapIndex.Text = StripLeadingZeroes(txtMapIndex.Text)
    If txtMapIndex.Text <> "0" Then
        MsgBoxResult = MsgBox("Map Index specifies the currently selected map." & vbCrLf & "You must make sure that the specified map index will exist in the users' game and be compatible with the specified game mode." & vbCrLf & "The minimum value of Map Index is zero (i.e. the first map)." & vbCrLf & vbCrLf & "Are you sure you wish to specify a different default map index?", vbYesNo + vbInformation, App.Title)
        If MsgBoxResult = vbNo Then
            txtMapIndex.Text = "0"
        End If
    End If
End Sub

Private Sub txtMarblePath_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtMarblePath_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If data.GetFormat(vbCFFiles) Then
        If data.Files.Count > 0 Then
            txtMarblePath.Text = data.Files.Item(data.Files.Count)
        End If
    End If
    Call txtMarblePath.SetFocus
End Sub

Private Sub txtModDisplaySound_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtModDisplaySound_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If data.GetFormat(vbCFFiles) Then
        If data.Files.Count > 0 Then
            txtModDisplaySound.Text = data.Files.Item(data.Files.Count)
        End If
    End If
    Call txtModDisplaySound.SetFocus
End Sub

Private Sub txtModLaunchSound_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtModLaunchSound_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If data.GetFormat(vbCFFiles) Then
        If data.Files.Count > 0 Then
            txtModLaunchSound.Text = data.Files.Item(data.Files.Count)
        End If
    End If
    Call txtModLaunchSound.SetFocus
End Sub

Private Sub txtModName_LostFocus()
    If txtModName.Text = "" Then
        MsgBoxResult = MsgBox("You must specify a Mod Name.", vbOKOnly + vbInformation, App.Title)
        txtModName.SetFocus
    End If
End Sub

Private Sub txtModScrnFormat_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtModScrnFormat_LostFocus()
    Dim DummyString As String
    Dim ReturnNumFormat As Long
    Call DisectScrnFormat(txtModScrnFormat.Text, DummyString, DummyString, ReturnNumFormat)
    If ReturnNumFormat = -1 Then
        MsgBoxResult = MsgBox("SCRN Format is not valid. Re-edit?", vbOKCancel + vbQuestion, App.Title)
        If MsgBoxResult = vbOK Then
            Call txtModScrnFormat.SetFocus
        Else
            txtModScrnFormat = "SCRN%04d.pcx"
        End If
    End If
End Sub

Private Sub txtModSnapFormat_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtModSnapFormat_LostFocus()
    Dim DummyString As String
    Dim ReturnNumFormat As Long
    Call DisectScrnFormat(txtModSnapFormat.Text, DummyString, DummyString, ReturnNumFormat)
    If ReturnNumFormat = -1 Then
        MsgBoxResult = MsgBox("SNAP Format is not valid. Re-edit?", vbOKCancel + vbQuestion, App.Title)
        If MsgBoxResult = vbOK Then
            Call txtModSnapFormat.SetFocus
        Else
            txtModSnapFormat = "Map%04d.yrm"
        End If
    End If
    Call UpdateControls
End Sub

Private Sub txtModUpdateCheck_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtModWebsite_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtProgramDirectory_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtProgramDirectory_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If data.GetFormat(vbCFFiles) Then
        If data.Files.Count > 0 Then
            txtProgramDirectory.Text = data.Files.Item(data.Files.Count)
        End If
    End If
    Call txtProgramDirectory.SetFocus
End Sub

Private Sub txtCustomScript_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtCustomScript_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If data.GetFormat(vbCFFiles) Then
        If data.Files.Count > 0 Then
            txtCustomScript.Text = data.Files.Item(data.Files.Count)
        End If
    End If
    Call txtCustomScript.SetFocus
End Sub

Private Sub txtTextColour_GotFocus(Index As Integer)
    SendKeys "{home}+{end}"
End Sub

Private Sub txtBGColour_GotFocus(Index As Integer)
    SendKeys "{home}+{end}"
End Sub

Private Sub txtTextColour_Click(Index As Integer)
    SendKeys "{home}+{end}"
End Sub

Private Sub txtBGColour_Click(Index As Integer)
    SendKeys "{home}+{end}"
End Sub

Private Sub txtModDateYear_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtModDescription_GotFocus()
    Dim Counter As Integer
    If Not IsKeyPressed(VK_LBUTTON) Then
        For Counter = 1 To Len(txtModDescription.Text)
            SendKeys "{right}"
        Next Counter
    End If
End Sub

Private Sub txtModName_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtModVersion_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtTX_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtINSTDIR_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtInfoPageText_GotFocus()
    Dim Counter As Integer
    If Not IsKeyPressed(VK_LBUTTON) Then
        For Counter = 1 To Len(txtInfoPageText.Text)
            SendKeys "{right}"
        Next Counter
    End If
End Sub

Private Sub txtInfoPageTitle_GotFocus()
    Dim Counter As Integer
    If Not IsKeyPressed(VK_LBUTTON) Then
        For Counter = 1 To Len(txtInfoPageTitle.Text)
            SendKeys "{right}"
        Next Counter
    End If
End Sub

Private Sub txtInfoPageButton_GotFocus()
    Dim Counter As Integer
    If Not IsKeyPressed(VK_LBUTTON) Then
        For Counter = 1 To Len(txtInfoPageButton.Text)
            SendKeys "{right}"
        Next Counter
    End If
End Sub

Private Sub txtModAuthor_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtModDateDay_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtModDateMonth_GotFocus()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

'**********************************************************
'********************CONTROL VALIDATION********************
'**********************************************************

Private Sub txtProgramDirectory_LostFocus()
    If txtProgramDirectory.Text <> txtProgramDirectory.Tag Then
        If Not DirExists(txtProgramDirectory.Text) Then
            MsgBoxResult = MsgBox(Quote(txtProgramDirectory.Text) & " does not exist!", vbOKOnly + vbInformation, App.Title)
        Else
            Call RefreshProgramList
        End If
    End If
    txtProgramDirectory.Tag = txtProgramDirectory.Text
End Sub

Private Sub txtCustomScript_LostFocus()
    If txtCustomScript.Text <> txtCustomScript.Tag Then
        If Not FileExists(txtCustomScript.Text) Then
            MsgBoxResult = MsgBox(Quote(txtCustomScript.Text) & " does not exist!", vbOKOnly + vbInformation, App.Title)
        End If
    End If
    txtCustomScript.Tag = txtCustomScript.Text
End Sub

Private Sub cboxUpdateOnly_Click()
    If cboxUpdateOnly.Value = 1 Then
        txtUpdateOnlySource.Enabled = True
        txtUpdateOnlySource.BackColor = TextEnabledColour
        cmdUpdateOnlySourceBrowse.Enabled = True
        txtUpdateOnlyDest.Enabled = True
        txtUpdateOnlyDest.BackColor = TextEnabledColour
        cmdUpdateOnlyDestBrowse.Enabled = True
    Else
        txtUpdateOnlySource.Enabled = False
        txtUpdateOnlySource.BackColor = TextDisabledColour
        cmdUpdateOnlySourceBrowse.Enabled = False
        txtUpdateOnlyDest.Enabled = False
        txtUpdateOnlyDest.BackColor = TextDisabledColour
        cmdUpdateOnlyDestBrowse.Enabled = False
    End If
    Call UpdateControls
End Sub

Private Sub comboModType_Click()
    Call UpdateControls
End Sub

Private Sub UpdateControls()
    Dim DummyStringArray(0) As String
    'Options
    comboModCampaigns.Enabled = (comboModType.ListIndex = 0)
    If LTrim$(comboModCampaigns.Text) = "" Then comboModCampaigns.ListIndex = 0
    comboModProgram.Enabled = (comboModType.ListIndex = 2)
    cboxParams.Enabled = comboModProgram.Enabled
    lblModProgram.Enabled = comboModProgram.Enabled
    If cboxParams.Enabled = False Then cboxParams.Value = 0
    If comboModProgram.ListIndex = -1 And comboModProgram.ListCount <> 0 Then comboModProgram.ListIndex = 0
    frameTX.Enabled = (comboModType.ListIndex = 0 Or comboModType.ListIndex = 1 Or comboModType.ListIndex = 3)
    optTX(0).Enabled = frameTX.Enabled
    optTX(1).Enabled = (comboModType.ListIndex = 0 Or comboModType.ListIndex = 1)
    optTX(2).Enabled = (comboModType.ListIndex = 0 Or comboModType.ListIndex = 1)
    optTX(3).Enabled = (comboModType.ListIndex = 3)
    If optTX(1).Value = True And optTX(1).Enabled = False Then optTX(0).Value = True
    If optTX(2).Value = True And optTX(2).Enabled = False Then optTX(0).Value = True
    If optTX(3).Value = True And optTX(3).Enabled = False Then optTX(0).Value = True
    txtTX.Enabled = (optTX(1).Value = True)
    lblMinVersionTX.Enabled = txtTX.Enabled
    txtModSnapFormat.Enabled = (comboModType.ListIndex = 0)
    lblModSnapFormat.Enabled = txtModSnapFormat.Enabled
    txtModScrnFormat.Enabled = (comboModType.ListIndex = 0)
    lblModScrnFormat.Enabled = txtModScrnFormat.Enabled
    frameFA2.Enabled = (comboModType.ListIndex = 1)
    lblFA2Version.Enabled = frameFA2.Enabled
    comboFA2.Enabled = frameFA2.Enabled
    If comboFA2.Enabled = False Then comboFA2.ListIndex = 0
    frameSecurity.Enabled = comboModType.ListIndex = 3
    cmdLock.Enabled = frameSecurity.Enabled
    cmdKey.Enabled = frameSecurity.Enabled
    cboxForRA2.Enabled = (comboModType.ListIndex = 0)
    cboxForRA2.Visible = cboxForRA2.Enabled
    If cboxForRA2.Enabled = False Then cboxForRA2.Value = 0
    cboxShutdownLB.Enabled = (comboModType.ListIndex = 2)
    cboxShutdownLB.Visible = cboxShutdownLB.Enabled
    If cboxShutdownLB.Enabled = False Then cboxShutdownLB.Value = 0
    frameGameMode.Enabled = (comboModType.ListIndex = 0)
    lblGameMode.Enabled = frameGameMode.Enabled
    lblMapIndex.Enabled = frameGameMode.Enabled
    txtGameMode.Enabled = frameGameMode.Enabled
    txtMapIndex.Enabled = frameGameMode.Enabled
    comboGameMode.Enabled = frameGameMode.Enabled
    lblPluginID.Visible = (comboModType.ListIndex = 3)
    comboPluginID.Visible = (comboModType.ListIndex = 3)
    'Files
    frameMixes.Enabled = (comboModType.ListIndex = 0 Or comboModType.ListIndex = 1 Or comboModType.ListIndex = 3)
    comboMixEncrypt.Enabled = frameMixes.Enabled
    comboSide3Mix.Enabled = (comboModType.ListIndex = 0)
    If comboSide3Mix.Enabled = False Then comboSide3Mix.ListIndex = 0
    txtProgramDirectory.Enabled = (comboModType.ListIndex = 2)
    cmdUninstFiles.Enabled = (comboModType.ListIndex = 2)
    cmdBrowseProgramDir.Enabled = txtProgramDirectory.Enabled
    lblProgramDirectory.Enabled = txtProgramDirectory.Enabled
    Select Case comboModType.ListIndex
    Case 0, 2, 3
        If cboxUpdateOnly.Value = 0 Then
            txtCustomScript.Enabled = True
        Else
            txtCustomScript.Enabled = False
        End If
    Case Else
        txtCustomScript.Enabled = False
    End Select
    cmdBrowseCustomScript.Enabled = txtCustomScript.Enabled
    lblCustomScript.Enabled = txtCustomScript.Enabled
    txtCustomScriptMessage.Enabled = txtCustomScript.Enabled
    lblCustomScriptMessage.Enabled = txtCustomScriptMessage.Enabled
    cboxUseAres.Enabled = (comboModType.ListIndex = 0)
    If Not cboxUseAres.Enabled Then cboxUseAres.Value = 0
    'Compression
    framePatchOptions.Enabled = (cboxGenPat.Value = 1)
    txtMarblePath.Enabled = (cboxGenPat.Value = 1 And (comboModType.ListIndex = 1 Or optTX(3).Value = True))
    cmdBrowseMarble.Enabled = txtMarblePath.Enabled
    lblPatchMarble.Enabled = txtMarblePath.Enabled
    framePatchMinSize.Enabled = framePatchOptions.Enabled
    framePatchBlockSize.Enabled = framePatchOptions.Enabled
    lblPatchMinSize.Enabled = framePatchOptions.Enabled
    lblPatchBlockSize.Enabled = framePatchOptions.Enabled
    lblPatchBlockSizeDef.Enabled = framePatchOptions.Enabled
    lblPatchDesc1.Enabled = framePatchOptions.Enabled
    lblPatchDesc2.Enabled = framePatchOptions.Enabled
    sliderPatchMinSize.Enabled = framePatchOptions.Enabled
    sliderPatchBlockSize.Enabled = framePatchOptions.Enabled
    If cboxUpdateOnly.Enabled = False Then cboxUpdateOnly.Value = 0
    frameUpdateOnly.Enabled = cboxUpdateOnly.Enabled
    lblUpdateOnlyDesc2.Enabled = cboxUpdateOnly.Enabled
    lblUpdateOnlyDesc3.Enabled = cboxUpdateOnly.Enabled
    lblUpdateOnlyDesc4.Enabled = cboxUpdateOnly.Enabled
    'Back Colours
    Select Case txtTX.Enabled
    Case False: txtTX.BackColor = TextDisabledColour
    Case True: txtTX.BackColor = TextEnabledColour
    End Select
    Select Case comboModCampaigns.Enabled
    Case False: comboModCampaigns.BackColor = TextDisabledColour
    Case True: comboModCampaigns.BackColor = TextEnabledColour
    End Select
    Select Case comboModProgram.Enabled
    Case False: comboModProgram.BackColor = TextDisabledColour
    Case True: comboModProgram.BackColor = TextEnabledColour
    End Select
    Select Case txtModScrnFormat.Enabled
    Case False: txtModScrnFormat.BackColor = TextDisabledColour
    Case True: txtModScrnFormat.BackColor = TextEnabledColour
    End Select
    Select Case txtModSnapFormat.Enabled
    Case False: txtModSnapFormat.BackColor = TextDisabledColour
    Case True: txtModSnapFormat.BackColor = TextEnabledColour
    End Select
    Select Case comboFA2.Enabled
    Case False: comboFA2.BackColor = TextDisabledColour
    Case True: comboFA2.BackColor = TextEnabledColour
    End Select
    Select Case txtGameMode.Enabled
    Case False: txtGameMode.BackColor = TextDisabledColour
    Case True: txtGameMode.BackColor = TextEnabledColour
    End Select
    Select Case comboGameMode.Enabled
    Case False: comboGameMode.BackColor = TextDisabledColour
    Case True: comboGameMode.BackColor = TextEnabledColour
    End Select
    Select Case txtMapIndex.Enabled
    Case False: txtMapIndex.BackColor = TextDisabledColour
    Case True: txtMapIndex.BackColor = TextEnabledColour
    End Select
    Select Case txtProgramDirectory.Enabled
    Case False: txtProgramDirectory.BackColor = TextDisabledColour
    Case True: txtProgramDirectory.BackColor = TextEnabledColour
    End Select
    Select Case txtCustomScript.Enabled
    Case False: txtCustomScript.BackColor = TextDisabledColour
    Case True: txtCustomScript.BackColor = TextEnabledColour
    End Select
    Select Case txtCustomScriptMessage.Enabled
    Case False: txtCustomScriptMessage.BackColor = TextDisabledColour
    Case True: txtCustomScriptMessage.BackColor = TextEnabledColour
    End Select
    Select Case txtMarblePath.Enabled
    Case False: txtMarblePath.BackColor = TextDisabledColour
    Case True: txtMarblePath.BackColor = TextEnabledColour
    End Select
    Select Case comboSide3Mix.Enabled
    Case False: comboSide3Mix.BackColor = TextDisabledColour
    Case True: comboSide3Mix.BackColor = TextEnabledColour
    End Select
    Select Case comboMixEncrypt.Enabled
    Case False: comboMixEncrypt.BackColor = TextDisabledColour
    Case True: comboMixEncrypt.BackColor = TextEnabledColour
    End Select
    Call InitNodes
    Call RefreshInstFiles(DummyStringArray)
End Sub

Private Sub comboInfoPage_Click()
    If comboInfoPage.ListIndex = 0 Then
        txtInfoPageText.Enabled = False
        txtInfoPageText.BackColor = TextDisabledColour
        txtInfoPageTitle.Enabled = False
        txtInfoPageTitle.BackColor = TextDisabledColour
        txtInfoPageButton.Enabled = False
        txtInfoPageButton.BackColor = TextDisabledColour
    Else
        txtInfoPageText.Enabled = True
        txtInfoPageText.BackColor = TextEnabledColour
        txtInfoPageTitle.Enabled = True
        txtInfoPageTitle.BackColor = TextEnabledColour
        If comboDSP.ListIndex <> 0 Then
            txtInfoPageButton.Enabled = True
            txtInfoPageButton.BackColor = TextEnabledColour
        Else
            txtInfoPageButton.Enabled = False
            txtInfoPageButton.BackColor = TextDisabledColour
        End If
    End If
    If comboDSP.ListIndex <> 1 Then
        If comboInfoPage.ListIndex <> 1 Then
            comboDSP.ListIndex = 1
        End If
    End If
End Sub

Private Sub txtINSTDIR_Change()
    Dim TempString As String
    TempString = StripInvalidChars(StripInvalidChars(txtINSTDIR.Text, InvalidFileChars), InvalidNSISChars)
    If TempString <> txtINSTDIR.Text Then
        MsgBoxResult = MsgBox("Default Directory cannot contain any of the following characters: " & InvalidFileChars & " " & InvalidNSISChars, vbOKOnly + vbInformation, "Invalid Character")
        txtINSTDIR.Text = TempString
        SendKeys "{End}"
    End If
    treeFolders.Nodes.Item("Node0").Text = TempString
    If LTrim$(txtINSTDIR.Text) = "" Then
        comboDSP.ListIndex = 1
        comboDSP.Enabled = False
    Else
        comboDSP.Enabled = True
    End If
End Sub

Private Sub txtINSTDIR_LostFocus()
    Dim DummyStringArray(0) As String
    Select Case UCase(txtINSTDIR.Text)
    Case "ORIGINALYR", "ORIGINALFA2", "ORIGINALRA2", "BACKUP", "HELP", "SKINS", "SETUPS"
        MsgBoxResult = MsgBox(Quote(txtINSTDIR.Text) & " is not permitted as an installation directory name.", vbOKOnly + vbInformation, App.Title)
        txtINSTDIR.Text = ""
    Case Else
        Call RefreshInstFiles(DummyStringArray())
    End Select
End Sub

Private Sub txtModDisplaySound_Change()
    Dim TempString As String
    TempString = StripInvalidChars(txtModDisplaySound.Text, InvalidNSISChars & " *")
    If TempString <> txtModDisplaySound.Text Then
        MsgBoxResult = MsgBox("Mod Display Sound cannot contain any of the following characters: " & InvalidNSISChars & " *", vbOKOnly + vbInformation, "Invalid Character")
        txtModDisplaySound.Text = TempString
        SendKeys "{End}"
    End If
End Sub

Private Sub txtModLaunchSound_Change()
    Dim TempString As String
    TempString = StripInvalidChars(txtModLaunchSound.Text, InvalidNSISChars & " *")
    If TempString <> txtModLaunchSound.Text Then
        MsgBoxResult = MsgBox("Mod Launch Sound cannot contain any of the following characters: " & InvalidNSISChars & " *", vbOKOnly + vbInformation, "Invalid Character")
        txtModLaunchSound.Text = TempString
        SendKeys "{End}"
    End If
End Sub

Private Sub txtModDisplaySound_LostFocus()
    If txtModDisplaySound.Text <> "" Then
        If Not FileExists(txtModDisplaySound.Text) Then MsgBoxResult = MsgBox(Quote(txtModDisplaySound.Text) & " does not exist!", vbOKOnly + vbInformation, App.Title)
    End If
End Sub

Private Sub txtMarblePath_LostFocus()
    If txtMarblePath.Text <> txtMarblePath.Tag Then
        If txtMarblePath.Text <> "" Then
            Call CheckValidMarblePath(txtMarblePath.Text, True)
        End If
        txtMarblePath.Tag = txtMarblePath.Text
    End If
End Sub

Private Sub txtModLaunchSound_LostFocus()
    If txtModLaunchSound.Text <> "" Then
        If Not FileExists(txtModLaunchSound.Text) Then MsgBoxResult = MsgBox(Quote(txtModLaunchSound.Text) & " does not exist!", vbOKOnly + vbInformation, App.Title)
    End If
End Sub

Private Sub txtModDateYear_Change()
    Dim TempString As String
    TempString = StripNonNumbers(txtModDateYear.Text)
    If TempString <> txtModDateYear.Text Then
        MsgBoxResult = MsgBox("Year must be an integer number.", vbOKOnly + vbInformation, "Invalid Character")
        txtModDateYear.Text = TempString
        SendKeys "{End}"
    End If
    If txtModDateYear.SelStart = 4 Then Call txtModDateMonth.SetFocus
    Call txtModDateDay_Change
End Sub

Private Sub txtModDateMonth_Change()
    Dim TempString As String
    TempString = StripNonNumbers(txtModDateMonth.Text)
    If TempString <> txtModDateMonth.Text Then
        MsgBoxResult = MsgBox("Month must be an integer number.", vbOKOnly + vbInformation, "Invalid Character")
        txtModDateMonth.Text = TempString
        SendKeys "{End}"
    End If
    If txtModDateMonth.Text <> "" Then
        If Val(txtModDateMonth.Text) > 12 Then
            txtModDateMonth.Text = "12"
            SendKeys "{End}"
        End If
    End If
    Call txtModDateDay_Change
    If txtModDateMonth.SelStart = 2 Then Call txtModDateDay.SetFocus
End Sub

Private Sub txtModDateDay_Change()
    Dim TempString As String
    TempString = StripNonNumbers(txtModDateDay.Text)
    If TempString <> txtModDateDay.Text Then
        MsgBoxResult = MsgBox("Day must be an integer number.", vbOKOnly + vbInformation, "Invalid Character")
        txtModDateDay.Text = TempString
        SendKeys "{End}"
    End If
    Dim MaxDays As Integer
    Select Case Val(txtModDateMonth.Text)
    Case 9, 4, 6, 11
        MaxDays = 30
    Case 2
        If IsLeapYear(Val(txtModDateYear.Text)) Then
            MaxDays = 29
        Else
            MaxDays = 28
        End If
    Case Else
        MaxDays = 31
    End Select
    If txtModDateDay.Text <> "" Then
        If Val(txtModDateDay.Text) > MaxDays Then
            txtModDateDay.Text = MaxDays
            SendKeys "{End}"
        End If
    End If
End Sub

Private Sub txtModDateYear_LostFocus()
    txtModDateYear.Text = StripLeadingZeroes(txtModDateYear.Text)
    If Len(txtModDateYear.Text) = 2 Then txtModDateYear.Text = "20" & txtModDateYear.Text
End Sub

Private Sub txtModDateMonth_LostFocus()
    txtModDateMonth.Text = PadNum(txtModDateMonth.Text, 2)
End Sub

Private Sub txtModDateDay_LostFocus()
    txtModDateDay.Text = PadNum(txtModDateDay.Text, 2)
End Sub

Private Sub txtModDescription_Change()
    Dim StrPos As Integer
    Dim TempString As String
    TempString = txtModDescription.Text
    StrPos = InStr(1, TempString, vbCrLf)
    Do While StrPos <> 0
        TempString = Left$(TempString, StrPos - 1) & Right$(TempString, Len(TempString) - (StrPos + 1))
        StrPos = InStr(1, TempString, vbCrLf)
    Loop
    If TempString <> txtModDescription.Text Then
        txtModDescription.Text = TempString
        SendKeys "{end}"
    End If
End Sub

Private Sub txtModVersion_Change()
    Dim TempString As String
    If cboxModVersion.Value = 0 Then
        TempString = StripNonFloat(txtModVersion.Text)
        If TempString <> txtModVersion.Text Then
            MsgBoxResult = MsgBox("Version must only include integers or floating points.", vbOKOnly + vbInformation, "Invalid Character")
            txtModVersion.Text = TempString
            SendKeys "{End}"
        End If
    End If
End Sub

Private Sub txtTX_Change()
    Dim TempString As String
    TempString = StripNonFloat(txtTX.Text)
    If TempString <> txtTX.Text Then
        MsgBoxResult = MsgBox("Version must only include integers or floating points.", vbOKOnly + vbInformation, "Invalid Character")
        txtTX.Text = TempString
        SendKeys "{End}"
    End If
End Sub

Private Sub ApplyTextColour()
    If Len(txtTextColour(0)) = 2 And Len(txtTextColour(1)) = 2 And Len(txtTextColour(2)) = 2 Then
        lblLogColours.ForeColor = RGB(HexToDec(txtTextColour(0).Text), HexToDec(txtTextColour(1).Text), HexToDec(txtTextColour(2).Text))
    End If
End Sub

Private Sub ApplyBGColour()
    If Len(txtBGColour(0)) = 2 And Len(txtBGColour(1)) = 2 And Len(txtBGColour(2)) = 2 Then
        lblLogColours.BackColor = RGB(HexToDec(txtBGColour(0).Text), HexToDec(txtBGColour(1).Text), HexToDec(txtBGColour(2).Text))
    End If
End Sub

Private Sub txtBGColour_Change(Index As Integer)
    Dim TempString As String
    Dim TempBoolean As Boolean
    Dim TempBoolean2 As Boolean
    If LoopPrevention = False Then
        LoopPrevention = True
        txtBGColour(Index).Text = UCase(txtBGColour(Index).Text)
        TempString = Mid$(txtBGColour(Index).Text, 1, 1)
        TempBoolean = TempString >= Chr(48) And TempString <= Chr(57)
        TempBoolean = TempBoolean Or (TempString >= Chr(65) And TempString <= Chr(70))
        If Len(txtBGColour(Index).Text) = 2 Then
            TempString = Mid$(txtBGColour(Index).Text, 2, 1)
            TempBoolean2 = TempString >= Chr(48) And TempString <= Chr(57)
            TempBoolean2 = TempBoolean2 Or (TempString >= Chr(65) And TempString <= Chr(70))
            TempBoolean = TempBoolean And TempBoolean2
        End If
        If TempBoolean Then
            txtBGColour(Index).Text = txtBGColour(Index).Text & txtBGColour(Index).Text
            If Index = 2 Then
                txtTextColour(0).SetFocus
            Else
                txtBGColour(Index + 1).SetFocus
            End If
        Else
            If txtBGColour(Index).Text <> "" Then
                MsgBoxResult = MsgBox("Please enter a valid hexadecimal character (0-F).", vbOKOnly + vbInformation, "Invalid Character")
            End If
            txtBGColour(Index).Text = "00"
            SendKeys "{home}+{end}"
        End If
        Call ApplyBGColour
        LoopPrevention = False
    End If
End Sub

Private Sub txtTextColour_Change(Index As Integer)
    Dim TempString As String
    Dim TempBoolean As Boolean
    Dim TempBoolean2 As Boolean
    If LoopPrevention = False Then
        LoopPrevention = True
        txtTextColour(Index).Text = UCase(txtTextColour(Index).Text)
        TempString = Mid$(txtTextColour(Index).Text, 1, 1)
        TempBoolean = TempString >= Chr(48) And TempString <= Chr(57)
        TempBoolean = TempBoolean Or (TempString >= Chr(65) And TempString <= Chr(70))
        If Len(txtTextColour(Index).Text) = 2 Then
            TempString = Mid$(txtTextColour(Index).Text, 2, 1)
            TempBoolean2 = TempString >= Chr(48) And TempString <= Chr(57)
            TempBoolean2 = TempBoolean2 Or (TempString >= Chr(65) And TempString <= Chr(70))
            TempBoolean = TempBoolean And TempBoolean2
        End If
        If TempBoolean Then
            txtTextColour(Index).Text = txtTextColour(Index).Text & txtTextColour(Index).Text
            If Index = 2 Then
                txtBGColour(0).SetFocus
            Else
                txtTextColour(Index + 1).SetFocus
            End If
        Else
            If txtTextColour(Index).Text <> "" Then
                MsgBoxResult = MsgBox("Please enter a valid hexadecimal character (0-F).", vbOKOnly + vbInformation, "Invalid Character")
            End If
            txtTextColour(Index).Text = "00"
            SendKeys "{home}+{end}"
        End If
        Call ApplyTextColour
        LoopPrevention = False
    End If
End Sub

Private Sub txtModName_Change()
    Dim TempString As String
    TempString = StripInvalidChars(txtModName.Text, InvalidNSISChars)
    If TempString <> txtModName.Text Then
        MsgBoxResult = MsgBox("Mod Name cannot contain any of the following characters: " & InvalidNSISChars, vbOKOnly + vbInformation, "Invalid Character")
        txtModName.Text = TempString
        SendKeys "{End}"
    End If
    lblNoBanner.Caption = txtModName.Text
End Sub

Private Sub cboxModVersion_Click()
    If cboxModVersion.Value = 1 Then
        txtModVersion.Enabled = False
        txtModVersion.BackColor = TextDisabledColour
        txtModVersion.Text = "<y.mm.dd.hh.mm>"
    Else
        txtModVersion.Enabled = True
        txtModVersion.BackColor = TextEnabledColour
        txtModVersion.Text = ""
    End If
End Sub

Private Sub cboxModDate_Click()
    If cboxModDate.Value = 1 Then
        txtModDateYear.Enabled = False
        txtModDateMonth.Enabled = False
        txtModDateDay.Enabled = False
        txtModDateYear.BackColor = TextDisabledColour
        txtModDateMonth.BackColor = TextDisabledColour
        txtModDateDay.BackColor = TextDisabledColour
        txtModDateYear.Text = Year(Now())
        txtModDateMonth.Text = PadNum(Month(Now()), 2)
        txtModDateDay.Text = PadNum(Day(Now()), 2)
    Else
        txtModDateYear.Enabled = True
        txtModDateMonth.Enabled = True
        txtModDateDay.Enabled = True
        txtModDateYear.BackColor = TextEnabledColour
        txtModDateMonth.BackColor = TextEnabledColour
        txtModDateDay.BackColor = TextEnabledColour
    End If
End Sub

'**********************************************************
'*******************MISC STUFF/SELECTORS*******************
'**********************************************************

Private Sub lblNoBanner_Click()
    Call picModBanner_Click
End Sub

Private Sub cmdBrowseMarble_Click()
    dialogOpen.FileName = txtMarblePath.Text
    dialogOpen.DialogTitle = "Select <marble.mix>"
    dialogOpen.Filter = "<marble.mix>|marble.mix"
    dialogOpen.DefaultExt = "mix"
'RetryOpen:
    On Error GoTo CancelOpen
    dialogOpen.ShowOpen
    On Error GoTo 0
    txtMarblePath.Text = dialogOpen.FileName
    txtMarblePath.Tag = txtMarblePath.Text
    Call CheckValidMarblePath(dialogOpen.FileName, True)
CancelOpen:
End Sub

Private Function CheckValidMarblePath(ByVal FilePath As String, Optional ByVal PopupMessages As Boolean = False) As Boolean
    Dim Ok As Boolean
    Ok = False
    If FileExists(FilePath) Then
        If UCase(GetFileName(FilePath)) = "MARBLE.MIX" Then
            lblCheckMarble.Visible = True
            Call frmMain.Refresh
            If GetFileMD5(FilePath) = MARBLEMD5 Then
                Ok = True
            Else
                If PopupMessages Then MsgBoxResult = MsgBox(Quote(FilePath) & " is a modified <marble.mix>." & vbCrLf & "Please select an original <marble.mix>.", vbOKOnly + vbInformation, App.Title)
            End If
            lblCheckMarble.Visible = False
            Call frmMain.Refresh
        Else
            If PopupMessages Then MsgBoxResult = MsgBox(Quote(FilePath) & " is not <marble.mix>.", vbOKOnly + vbInformation, App.Title)
        End If
    Else
        If PopupMessages Then MsgBoxResult = MsgBox(Quote(FilePath) & " does not exist!", vbOKOnly + vbInformation, App.Title)
    End If
    CheckValidMarblePath = Ok
End Function

Private Sub cmdBrowseModDisplaySound_Click()
    If txtModDisplaySound.Text = "" Then
        dialogOpen.FileName = txtModLaunchSound.Text
    Else
        dialogOpen.FileName = txtModDisplaySound.Text
    End If
    If DirExists(dialogOpen.FileName) Then dialogOpen.FileName = JoinPath(dialogOpen.FileName, "*.wav;*.ogg;*.flac")
    dialogOpen.DialogTitle = "Select Mod Display Sound"
    dialogOpen.Filter = "All valid mod sound formats (*.wav, *.ogg, *.flac)|*.wav;*.ogg;*.flac|WAVE Files (*.wav)|*.wav|OGG Vorbis Files (*.ogg)|*.ogg|FLAC Files (*.flac)|*.flac"
    dialogOpen.DefaultExt = ""
'RetryOpen:
    On Error GoTo CancelOpen
    dialogOpen.ShowOpen
    On Error GoTo 0
    If FileExists(dialogOpen.FileName) Then
        Select Case FileType(dialogOpen.FileName)
        Case "WAV", "OGG", "FLAC"
            txtModDisplaySound.Text = dialogOpen.FileName
        Case Else
            MsgBoxResult = MsgBox(Quote(dialogOpen.FileName) & " is not a valid sound file. Mod sound files must be in WAV, OGG or FLAC format.", vbOKOnly + vbInformation, App.Title)
        End Select
    Else
        MsgBoxResult = MsgBox(Quote(dialogOpen.FileName) & " does not exist!", vbOKOnly + vbInformation, App.Title)
    End If
CancelOpen:
End Sub

Private Sub cmdBrowseModLaunchSound_Click()
    If txtModLaunchSound.Text = "" Then
        dialogOpen.FileName = txtModDisplaySound.Text
    Else
        dialogOpen.FileName = txtModLaunchSound.Text
    End If
    If DirExists(dialogOpen.FileName) Then dialogOpen.FileName = JoinPath(dialogOpen.FileName, "*.wav;*.ogg;*.flac")
    dialogOpen.DialogTitle = "Select Mod Launch Sound"
    dialogOpen.Filter = "All valid mod sound formats (*.wav, *.ogg, *.flac)|*.wav;*.ogg;*.flac|WAVE Files (*.wav)|*.wav|OGG Vorbis Files (*.ogg)|*.ogg|FLAC Files (*.flac)|*.flac"
    dialogOpen.DefaultExt = ""
'RetryOpen:
    On Error GoTo CancelOpen
    dialogOpen.ShowOpen
    On Error GoTo 0
    If FileExists(dialogOpen.FileName) Then
        Select Case FileType(dialogOpen.FileName)
        Case "WAV", "OGG", "FLAC"
            txtModLaunchSound.Text = dialogOpen.FileName
        Case Else
            MsgBoxResult = MsgBox(Quote(dialogOpen.FileName) & " is not a valid sound file. Mod sound files must be in WAV, OGG or FLAC format.", vbOKOnly + vbInformation, App.Title)
        End Select
    Else
        MsgBoxResult = MsgBox(Quote(dialogOpen.FileName) & " does not exist!", vbOKOnly + vbInformation, App.Title)
    End If
CancelOpen:
End Sub

Private Sub picModBanner_Click()
    Dim Ok As Boolean
    Ok = False
    dialogOpen.FileName = picModBanner.Tag
    dialogOpen.DialogTitle = "Select Banner Image"
    dialogOpen.Filter = "All valid image formats (*.bmp, *.gif, *.jpg, *.jpeg)|*.bmp;*.gif;*.jpg;*.jpeg|Bitmap Files (*.bmp)|*.bmp|GIF Files (*.gif)|*.gif|JPEG Files (*.jpg, *jpeg)|*.jpg;*.jpeg"
    dialogOpen.DefaultExt = ""
'RetryOpen:
    On Error GoTo CancelOpen
    dialogOpen.ShowOpen
    On Error GoTo 0
    picModBanner.Tag = dialogOpen.FileName
    Call LoadBannerImage(True)
CancelOpen:
End Sub

Private Sub LoadBannerImage(Optional ByVal PopupMessages As Boolean = True)
    Dim Ok As Boolean
    Ok = False
    If FileExists(picModBanner.Tag) Then
        Select Case FileType(picModBanner.Tag)
        Case "BMP", "GIF", "JPEG", "JPG"
            Set picModBanner.Picture = Nothing
            lblNoBanner.Visible = False
            Set picModBanner.Picture = LoadPicture(picModBanner.Tag)
            picModBanner.ToolTipText = picModBanner.Tag
            Ok = True
        Case Else
            If PopupMessages Then MsgBoxResult = MsgBox(Quote(picModBanner.Tag) & " is not a valid banner image format!", vbOKOnly + vbInformation, App.Title)
        End Select
    Else
        If PopupMessages Then MsgBoxResult = MsgBox(Quote(picModBanner.Tag) & " does not exist!", vbOKOnly + vbInformation, App.Title)
    End If
    If Not Ok Then
        Set picModBanner.Picture = Nothing
        lblNoBanner.Visible = True
        Set picModBanner.Picture = LoadResPicture("NOBANNER", vbResBitmap)
        picModBanner.ToolTipText = ""
    End If
End Sub

Private Sub picInstallerIcon_Click()
    dialogOpen.FileName = picInstallerIcon.Tag
    dialogOpen.DialogTitle = "Select Installer Icon"
    dialogOpen.Filter = "Icon Files (*.ico)|*.ico"
    dialogOpen.DefaultExt = "ico"
    On Error GoTo CancelOpenIcon
    dialogOpen.ShowOpen
    On Error GoTo 0
    Set picInstallerIcon.Picture = Nothing
    picInstallerIcon.Tag = dialogOpen.FileName
    Call LoadInstallerIcon(True)
CancelOpenIcon:
End Sub

Private Sub LoadInstallerIcon(Optional ByVal PopupMessages As Boolean = True)
    Dim Ok As Boolean
    Ok = False
    If FileExists(picInstallerIcon.Tag) Then
        If FileType(picInstallerIcon.Tag) = "ICO" Then
            Set picInstallerIcon.Picture = LoadPicture(picInstallerIcon.Tag, vbLPLarge, , 32, 32)
            picInstallerIcon.ToolTipText = picInstallerIcon.Tag
            Ok = True
        Else
            If PopupMessages Then MsgBoxResult = MsgBox(Quote(picInstallerIcon.Tag) & " is not an icon!", vbOKOnly + vbInformation, App.Title)
        End If
    Else
        If PopupMessages Then MsgBoxResult = MsgBox(Quote(picInstallerIcon.Tag) & " does not exist!", vbOKOnly + vbInformation, App.Title)
    End If
    If Not Ok Then
        Set picInstallerIcon.Picture = LoadPicture(JoinPath(RESDIR, "default.ico"), vbLPCustom, , 32, 32)
        picInstallerIcon.ToolTipText = JoinPath(RESDIR, "default.ico")
    End If
End Sub

Private Sub pboxModDisplaySound_Click()
    Call PlayModSound(txtModDisplaySound.Text)
End Sub

Private Sub pboxModLaunchSound_Click()
    Call PlayModSound(txtModLaunchSound.Text)
End Sub

Private Sub PlayModSound(ByVal FilePath As String)
    If FilePath <> "" Then
        If FileExists(FilePath) Then
            Select Case FileType(FilePath)
            Case "OGG", "FLAC"
                MsgBoxResult = MsgBox(App.Title & " cannot play " & FileType(FilePath) & " files. " & vbCrLf & Quote(FilePath) & " will be converted to WAV format when installed on the end users computer.", vbOKOnly + vbInformation, App.Title)
            Case "WAV"
                Call PlaySound(FilePath)
            Case Else
                MsgBoxResult = MsgBox(Quote(FilePath) & " is not a valid sound file. Mod sound files must be in WAV, OGG or FLAC format.", vbOKOnly + vbInformation, App.Title)
            End Select
        Else
            MsgBoxResult = MsgBox(Quote(FilePath) & " does not exist!", vbOKOnly + vbInformation, App.Title)
        End If
    End If
End Sub

'**********************************************************
'********************RECENT FILES STUFF********************
'**********************************************************

Private Sub menu_recent_Click(Index As Integer)
    Call OpenRecentFile(Index)
End Sub

Private Sub OpenRecentFile(ByVal Index As Integer)
    Dim TempString As String
    TempString = menu_recent(Index).Tag
    If FileExists(TempString) Then
        Call LoadSettings(TempString)
        frmMain.Caption = GetFileName(TempString) & " - " & App.Title
        dialogSave.FileName = TempString
        Call BringRecentFileToTop(Index)
    Else
        MsgBoxResult = MsgBox(Quote(TempString) & " has been moved/renamed/deleted since it was last accessed.", vbOKOnly + vbInformation, App.Title)
        Do While Index < (menu_recent.Count - 1)
            menu_recent(Index).Caption = "&" & CStr(Index + 1) & " " & menu_recent(Index + 1).Tag
            menu_recent(Index).Tag = menu_recent(Index + 1).Tag
            Index = Index + 1
        Loop
        If menu_recent.Count <> 1 Then
            Unload menu_recent(menu_recent.Count - 1)
        Else
            menu_recent(0).Visible = False
            menu_line4.Visible = False
        End If
    End If
    Call SaveRecentFiles
End Sub

Private Sub SaveRecentFiles()
    Dim Index As Integer
    Index = 0
    If menu_recent(0).Visible = True Then
        Do While Index <= (menu_recent.Count - 1)
            Call WriteINIStr("RecentFiles", CStr(Index), menu_recent(Index).Tag, ProgramINI)
            Index = Index + 1
        Loop
    End If
    Do While Index <= (MaxRecentFiles - 1)
        Call WriteINIStr("RecentFiles", CStr(Index), " ;deleted", ProgramINI)
        Index = Index + 1
    Loop
End Sub

Private Sub InitRecentFiles()
    Dim FileString As String
    Dim Counter As Integer
    For Counter = 0 To MaxRecentFiles
        FileString = ReadINIStr("RecentFiles", CStr(Counter), ProgramINI)
        If FileString <> "" Then
            If Counter = 0 Then
                menu_recent(Counter).Visible = True
                menu_line4.Visible = True
            Else
                Load menu_recent(Counter)
            End If
            menu_recent(Counter).Tag = FileString
            menu_recent(Counter).Caption = "&" & CStr(Counter + 1) & " " & FileString
        Else
            Counter = MaxRecentFiles
        End If
    Next Counter
End Sub

Private Sub AddRecentFile(ByVal NewPath As String)
    Dim Counter As Integer
    Dim Ok As Boolean
    Ok = False
    If menu_recent(0).Visible = True Then
        'Find file in recent file list (if present) and bring it to the top.
        Counter = 0
        Do While Counter <= (menu_recent.Count - 1)
            If UCase(menu_recent(Counter).Tag) = UCase(NewPath) Then
                Call BringRecentFileToTop(Counter)
                Ok = True
                Counter = menu_recent.Count - 1
            End If
            Counter = Counter + 1
        Loop
        If Not Ok Then
            'Shift all files down by one to make way for new file.
            Counter = menu_recent.Count - 1
            If Counter < (MaxRecentFiles - 1) Then
                Counter = Counter + 1
                Load menu_recent(Counter)
            End If
            Do While Counter > 0
                menu_recent(Counter).Caption = "&" & CStr(Counter + 1) & " " & menu_recent(Counter - 1).Tag
                menu_recent(Counter).Tag = menu_recent(Counter - 1).Tag
                Counter = Counter - 1
            Loop
        End If
    Else
        menu_recent(0).Visible = True
        menu_line4.Visible = True
    End If
    menu_recent(0).Caption = "&1 " & NewPath
    menu_recent(0).Tag = NewPath
    Call SaveRecentFiles
End Sub

Private Sub BringRecentFileToTop(ByVal Index As Integer)
    Dim TempTag As String
    TempTag = menu_recent(Index).Tag
    Do While Index <> 0
        menu_recent(Index).Caption = "&" & CStr(Index + 1) & " " & menu_recent(Index - 1).Tag
        menu_recent(Index).Tag = menu_recent(Index - 1).Tag
        Index = Index - 1
    Loop
    menu_recent(0).Caption = "&1 " & TempTag
    menu_recent(0).Tag = TempTag
    Call SaveRecentFiles
End Sub

'**********************************************************
'************************MENU STUFF************************
'**********************************************************

Private Sub menu_open_Click()
    dialogOpen.FileName = dialogSave.FileName
    dialogOpen.DialogTitle = "Open"
    dialogOpen.Filter = dialogSave.Filter
    dialogOpen.DefaultExt = dialogSave.DefaultExt
    On Error GoTo CancelOpen
    dialogOpen.ShowOpen
    On Error GoTo 0
    dialogSave.FileName = dialogOpen.FileName
    If FileExists(dialogOpen.FileName) Then
        Call AddRecentFile(dialogOpen.FileName)
        Call LoadSettings(dialogOpen.FileName)
    Else
        MsgBoxResult = MsgBox(Quote(dialogOpen.FileName) & " does not exist!", vbOKOnly + vbInformation, App.Title)
    End If
CancelOpen:
End Sub

Private Sub menu_save_Click()
    If dialogSave.FileName = "" Then
        Call menu_saveas_Click
    Else
        Call SaveSettings(dialogSave.FileName)
    End If
End Sub

Private Sub menu_saveas_Click()
    On Error GoTo CancelSaveAs
    dialogSave.ShowSave
    On Error GoTo 0
    Call AddRecentFile(dialogSave.FileName)
    Call SaveSettings(dialogSave.FileName)
    frmMain.Caption = GetFileName(dialogSave.FileName) & " - " & App.Title
CancelSaveAs:
End Sub

Private Sub menu_create_Click()
    Dim FileHandle As Integer
    Dim process_id
    Dim process_handle
    Dim iCounter As Integer
    Dim ErrVars(0) As Variant
    If GetArgByName("noexcept") = "" Then On Error GoTo LocalErr
    iCounter = 0
    Do While iCounter < InstFileCount
        If Len(InstFile(iCounter)) <> 0 Then
            If Not InstFileErrorCheck(iCounter, False) Then
                Call MsgBox("There are file errors that you must correct before your installer can be created.", vbOKOnly + vbInformation, App.Title)
                SSTab1.Tab = 1
                GoTo CancelCreate
            End If
        End If
        iCounter = iCounter + 1
    Loop
    If cboxUpdateOnly.Value = 1 Then
        If CheckUpdateOnlyPaths = False Then GoTo CancelCreate
    End If
    If comboMixEncrypt.ListIndex <> 0 And DCoderDLL = False Then
        MsgBoxResult = MsgBox("DCoder DLL is missing! " & App.Title & " cannot create MIX files without it!" & vbCrLf & "Set 'MIX File Format' to " & Quote("None") & " and pre-compile any MIX files before trying to create your installer.", vbOKOnly + vbExclamation, App.Title)
        GoTo CancelCreate
    End If
    FileHandle = FreeFile
RetryCreate:
    On Error GoTo CancelCreate
    dialogCreate.ShowSave
    On Error GoTo 0
    If FileExists(dialogCreate.FileName) Then
        MsgBoxResult = MsgBox(Quote(dialogCreate.FileName) & " already exists." & vbCrLf & "Do you want to overwrite this file?", vbYesNoCancel + vbQuestion, App.Title)
        Select Case MsgBoxResult
        Case vbNo: GoTo RetryCreate
        Case vbCancel: GoTo CancelCreate
        Case vbYes
            If FileExists(dialogCreate.FileName) Then
                On Error GoTo InstallerInUse
                Call Kill(dialogCreate.FileName)
                On Error GoTo 0
            End If
            GoTo InstallerNotInUse
InstallerInUse:
            On Error GoTo 0
            MsgBoxResult = MsgBox("Error deleting " & Quote(dialogCreate.FileName), vbOKOnly + vbExclamation, App.Title)
            GoTo CancelCreate
InstallerNotInUse:
        End Select
    End If
    frmMain.Hide
    frmWait.Show
    frmWait.Label1.Caption = "Compressing files and creating installer..."
    frmWait.Refresh
    If DirExists(KILNDIR) Then Call KillDir(KILNDIR)
    Call MakePath(KILNDIR)
    Call SaveSettings(JoinPath(KILNDIR, "except.erm"))
    Call SaveScript(JoinPath(KILNDIR, "script1.erm"), dialogCreate.FileName)
    If comboInfoPage.ListIndex <> 0 Then
        Open JoinPath(KILNDIR, "info.erm") For Output As #FileHandle
            Print #FileHandle, txtInfoPageText.Text
        Close #FileHandle
    End If
    If FileExists(JoinPath(EXEDIR, "except.lbp")) Then Call Kill(JoinPath(EXEDIR, "except.lbp"))
    If FileExists(JoinPath(EXEDIR, "except.log")) Then Call Kill(JoinPath(EXEDIR, "except.log"))
    If FileExists(JoinPath(EXEDIR, "except1.nsi")) Then Call Kill(JoinPath(EXEDIR, "except1.nsi"))
    If FileExists(JoinPath(EXEDIR, "except2.nsi")) Then Call Kill(JoinPath(EXEDIR, "except2.nsi"))
    Call ChDir(RESDIR)
    'INSTALLER
    process_id = Shell(JoinPath(RESDIR, "makensis.exe") & " /O" & Quote(JoinPath(KILNDIR, "error.log")) & " " & Quote(JoinPath(KILNDIR, "script1.erm")), vbHide)
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    If FileExists(JoinPath(KILNDIR, "info.erm")) Then Call Kill(JoinPath(KILNDIR, "info.erm"))
    Unload frmWait
    If FileExists(dialogCreate.FileName) Then
        MsgBoxResult = MsgBox("Installer successfully created.", vbOKOnly + vbInformation, App.Title)
    Else
        If FileExists(JoinPath(KILNDIR, "script1.erm")) Then Name JoinPath(KILNDIR, "script1.erm") As JoinPath(EXEDIR, "except1.nsi")
        If FileExists(JoinPath(KILNDIR, "script2.erm")) Then Name JoinPath(KILNDIR, "script2.erm") As JoinPath(EXEDIR, "except2.nsi")
        Name JoinPath(KILNDIR, "except.erm") As JoinPath(EXEDIR, "except.lbp")
        If FileExists(JoinPath(KILNDIR, "error.log")) Then
            Name JoinPath(KILNDIR, "error.log") As JoinPath(EXEDIR, "except.log")
            MsgBoxResult = MsgBox("An error has occurred." & vbCrLf & "Please send <except.lbp>, <except.log>, <except1.nsi> and <except2.nsi> to Marshall.", vbOKOnly + vbExclamation, App.Title)
        Else
            MsgBoxResult = MsgBox("An unknown error has occurred." & vbCrLf & "Please send <except.lbp>, <except1.nsi> and <except2.nsi> to Marshall.", vbOKOnly + vbExclamation, App.Title)
        End If
    End If
    Debug.Print CurDir
    Call KillDir(KILNDIR)
    frmMain.Show
CancelCreate:
    Exit Sub
LocalErr:
    Call GlobalErr("menu_create_Click", ErrVars())
End Sub

Private Sub menu_about_Click()
    frmMain.Enabled = False
    frmAbout.Show
End Sub

Private Function CheckUpdateOnlyPaths() As Boolean
    Dim Ok As Boolean
    Ok = False
    If UCase(GetFileName(txtUpdateOnlySource.Text)) <> "LIBLIST.GAM" Then
        MsgBoxResult = MsgBox("Cannot create update-only installer: You must specify the <liblist.gam> of the Previous Installation!", vbOKOnly + vbInformation, App.Title)
    Else
        If UCase(GetFileName(txtUpdateOnlyDest.Text)) <> "LIBLIST.GAM" Then
            MsgBoxResult = MsgBox("Cannot create update-only installer: You must specify the <liblist.gam> of the Latest Installation!", vbOKOnly + vbInformation, App.Title)
        Else
            If Not FileExists(txtUpdateOnlySource.Text) Then
                MsgBoxResult = MsgBox("Cannot create update-only installer: Previous Installation does not exist!", vbOKOnly + vbInformation, App.Title)
            Else
                If Not FileExists(txtUpdateOnlyDest.Text) Then
                    MsgBoxResult = MsgBox("Cannot create update-only installer: Latest Installation does not exist!", vbOKOnly + vbInformation, App.Title)
                Else
                    If ReadINIStr("General", "Version", txtUpdateOnlySource.Text) = ReadINIStr("General", "Version", txtUpdateOnlySource.Text) Then
                        MsgBoxResult = MsgBox("Cannot create update-only installer: Previous Installation and Latest Installation are the same version!", vbOKOnly + vbInformation, App.Title)
                    Else
                        Ok = True
                    End If
                End If
            End If
        End If
    End If
    If Not Ok Then SSTab1.Tab = 3
    CheckUpdateOnlyPaths = Ok
End Function

Private Sub menu_exit_Click()
    Call Shutdown
End Sub

Private Sub menu_help_Click()
    frmHelp.Show
End Sub

Private Sub txtUpdateOnlyDest_Change()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtUpdateOnlyDest_LostFocus()
    If txtUpdateOnlyDest.Text <> txtUpdateOnlyDest.Tag Then
        If txtUpdateOnlyDest.Text <> "" Then
            Call CheckUpdateOnlyPaths
        End If
        txtUpdateOnlyDest.Tag = txtUpdateOnlyDest.Text
    End If
End Sub

Private Sub txtUpdateOnlyDest_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If data.GetFormat(vbCFFiles) Then
        If data.Files.Count > 0 Then
            txtUpdateOnlyDest.Text = data.Files.Item(data.Files.Count)
        End If
    End If
    Call txtUpdateOnlyDest.SetFocus
End Sub

Private Sub txtUpdateOnlySource_Change()
    If Not IsKeyPressed(VK_LBUTTON) Then SendKeys "{home}+{end}"
End Sub

Private Sub txtUpdateOnlySource_LostFocus()
    If txtUpdateOnlySource.Text <> txtUpdateOnlySource.Tag Then
        If txtUpdateOnlySource.Text <> "" Then
            Call CheckUpdateOnlyPaths
        End If
        txtUpdateOnlySource.Tag = txtUpdateOnlySource.Text
    End If
End Sub

Private Sub txtUpdateOnlySource_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If data.GetFormat(vbCFFiles) Then
        If data.Files.Count > 0 Then
            txtUpdateOnlySource.Text = data.Files.Item(data.Files.Count)
        End If
    End If
    Call txtUpdateOnlySource.SetFocus
End Sub

Private Sub ConvertWavToFlac(ByVal SourcePath As String, Optional ByVal DestPath As String = "")
    Dim process_id
    Dim process_handle
    If DestPath = "" Then DestPath = ChangeFileType(SourcePath, "flac")
    If FileExists(DestPath) Then
        'Call WriteLogEntry("Deleting " & Quote(DestPath) & " to make way for new file.")
        Call Kill(DestPath)
    End If
    Call SetCurrentDirectory(GetFilePath(SourcePath)) 'FLAC CLT isn't very clever
    process_id = Shell(Quote(JoinPath(EXEDIR, "Resource", "flac.exe")) & " -8 " & Quote(GetFileName(SourcePath)) & " --output-prefix=" & Quote(GetFilePath(DestPath) & "\/"), vbHide)
    'not using " --keep-foreign-metadata" because operation will fail if there is no foreign metadata
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
    SourcePath = JoinPath(GetFilePath(DestPath), ChangeFileType(GetFileName(SourcePath), "flac"))
    If UCase$(SourcePath) <> UCase$(DestPath) Then
        Name SourcePath As DestPath
    End If
End Sub

