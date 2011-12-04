VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmModCat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base: Check For Updates"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10920
   Icon            =   "frmModCat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Back to Launch Menu"
      Height          =   375
      Left            =   7320
      TabIndex        =   37
      Top             =   4680
      Width           =   3495
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download and Install Selected Mod"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   40
      Top             =   4680
      Width           =   3495
   End
   Begin VB.CommandButton cmdChangeLog 
      Caption         =   "View Change Log for Selected Mod"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   4680
      Width           =   3495
   End
   Begin VB.Frame frameFilter 
      Caption         =   "Filter"
      Height          =   3495
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   1575
      Begin VB.CheckBox cboxFilterUpdates 
         Caption         =   "Installed Mods"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   38
         ToolTipText     =   "Shows mods that you have already installed and already have the latest version."
         Top             =   3120
         Width           =   1425
      End
      Begin VB.CheckBox cboxFilterGame 
         Caption         =   "Yuri's Revenge"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   35
         ToolTipText     =   "Select which games you are interested in."
         Top             =   2040
         Width           =   1425
      End
      Begin VB.CheckBox cboxFilterGame 
         Caption         =   "Red Alert 2"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Select which games you are interested in."
         Top             =   1800
         Width           =   1425
      End
      Begin VB.CheckBox cboxFilterUpdates 
         Caption         =   "Updates"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "Shows mods that you have already installed for which there is a newer version."
         Top             =   2640
         Width           =   1425
      End
      Begin VB.CheckBox cboxFilterUpdates 
         Caption         =   "New Mods"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Shows mods that you haven't installed."
         Top             =   2880
         Width           =   1425
      End
      Begin VB.CheckBox cboxFilterModType 
         Caption         =   "Tools"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Select which mod types you are interested in acquiring/updating."
         Top             =   1200
         Width           =   1395
      End
      Begin VB.CheckBox cboxFilterModType 
         Caption         =   "FA2 Mods"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Select which mod types you are interested in acquiring/updating."
         Top             =   960
         Width           =   1395
      End
      Begin VB.CheckBox cboxFilterModType 
         Caption         =   "Plugins"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Select which mod types you are interested in acquiring/updating."
         Top             =   720
         Width           =   1395
      End
      Begin VB.CheckBox cboxFilterModType 
         Caption         =   "Mods"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Select which mod types you are interested in acquiring/updating."
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblFilterUpdates 
         Caption         =   "Updates:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label lblFilterGame 
         Caption         =   "Games:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblFilterModType 
         Caption         =   "Mod Types:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frameAvailableMods 
      Caption         =   "Available Mods/Updates"
      Height          =   3495
      Left            =   1800
      TabIndex        =   16
      Top             =   1080
      Width           =   4695
      Begin MSComctlLib.ListView listModCat 
         Height          =   3135
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   5530
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Update Check Progress"
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   10695
      Begin MSComctlLib.ProgressBar pbarModCat 
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblModCatFails 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "see which update checks failed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   7320
         MouseIcon       =   "frmModCat.frx":0E42
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   240
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.Label lblCheckProgress 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   10455
      End
   End
   Begin VB.Frame frameModDetails 
      Caption         =   "Selected Mod Details"
      Height          =   3495
      Left            =   6600
      TabIndex        =   0
      Top             =   1080
      Width           =   4215
      Begin VB.Label lblModGame 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Red Alert 2"
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblGame 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Game:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lblModYourVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.001"
         Height          =   195
         Left            =   1440
         TabIndex        =   21
         Top             =   840
         Width           =   405
      End
      Begin VB.Label lblYourVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Installed Version:"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Download:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblModSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KB"
         Height          =   195
         Left            =   1440
         TabIndex        =   18
         Top             =   1800
         Width           =   210
      End
      Begin VB.Label lblModTX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Not Required"
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   1440
         TabIndex        =   13
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label lblTX 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TX Plugin:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label lblModDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2006-11-12"
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label lblModDescription 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmModCat.frx":114C
         Height          =   855
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   3975
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Latest Version:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label lblCampaigns 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Campaigns:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label lblAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   510
      End
      Begin VB.Label lblWebsite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Website:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lblModVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.001"
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lblModAuthor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Westwood Studios"
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblModCampaigns 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Original Campaigns"
         Height          =   195
         Left            =   1440
         TabIndex        =   2
         Top             =   2040
         Width           =   1350
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
         Left            =   1440
         MouseIcon       =   "frmModCat.frx":11DD
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1560
         Width           =   1935
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
Attribute VB_Name = "frmModCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ColorGood = &H4000&
Private Const ColorBad = &H80&
Private Const ColorURL = &HC00000
Private Const ColorURLActive = &HFF&
Public Busy As Boolean

Public Function DownloadProgress(ByVal iBytesReceived As Long, ByVal iBytesExpected As Long)
    'called by internet module. needed to allow user to abort
    DoEvents
End Function

Private Sub RedimAndInit(iNewSize)
    Dim iCounter As Integer
    If iNewSize = 0 Then
        iCounter = 0
    Else
        iCounter = UBound(UpdateRecords()) + 1
    End If
    ReDim Preserve UpdateRecords(iNewSize)
    Do While iCounter <= iNewSize
        Call InitialiseRecord(iCounter)
        iCounter = iCounter + 1
    Loop
End Sub

Public Sub LoadUpdateCat()
    Dim hModCat As Integer
    Dim iRecord As Integer
    Dim iMod As Integer
    Dim bUpdateMod As Boolean
    Call CallStackPush(Me.Name & ".LoadUpdateCat()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call frmMain.WriteLogEntry("Loading mod update records...", LogLevel1)
    Call RedimAndInit(0)
    Call RedimAndInit(ModCount + 20)
    UpdateRecords(0).CheckURL = ModCatUPD
    UpdateRecordCount = 0
    For iMod = 0 To ModCount
        If Len(Mods(iMod).ModName) <> 0 Then 'mod has been deleted if name is blank
            'Check for legacy versions of the same mod
            iRecord = 1
            Do While iRecord <> (UpdateRecordCount + 1)
                If UpdateRecords(iRecord).CheckURL = Mods(iMod).ModUpdateCheckURL Then
                    If CompareVersions(Mods(iMod).ModVersion, ">", UpdateRecords(iRecord).ModUserVersion) Then
                        UpdateRecords(iRecord).CheckModNum = iMod
                        Call InitialiseRecord(UpdateRecordCount, True)
                        Exit Do
                    End If
                End If
                iRecord = iRecord + 1
            Loop
            If iRecord = (UpdateRecordCount + 1) Then
                'we didn't find the mod so create new record
                UpdateRecordCount = UpdateRecordCount + 1
                UpdateRecords(UpdateRecordCount).CheckModNum = iMod
                Call InitialiseRecord(UpdateRecordCount, True)
            End If
        End If
    Next iMod
    'All plugins are in the online catalogue
    Call CallStackPop
End Sub

Public Sub UpdateUpdateCat()
    Dim iRecord As Long
    Dim iFails As Long
    Call CallStackPush(Me.Name & ".UpdateUpdateCat()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call EnableControls(False)
    Call frmMain.WriteLogEntry("Checking for updates...", LogLevel1)
    lblCheckProgress.Caption = "Checking for updates..."
    cmdCancel.Caption = "Abort Update Check"
    frmModCat.pbarModCat.Value = 0
    Call Me.Refresh
    Busy = True
    frmModCat.pbarModCat.Max = UpdateRecordCount + 1
    iFails = 0
    iRecord = 0
    If Not theInternet.Connected Then
        Me.Enabled = False
        Call theInternet.Connect(Me.hWnd)
        Me.Enabled = True
        Me.SetFocus
    End If
    If theInternet.Connected Then
        Do While iRecord <= UpdateRecordCount
            iFails = iFails + UpdateUpdateRecord(iRecord)
            If theInternet.DownloadCancelled Then Exit Do
            iRecord = iRecord + 1
            frmModCat.pbarModCat.Value = frmModCat.pbarModCat.Value + 1
        Loop
    End If
    Busy = False
    If theInternet.Connected Then
        If iFails = 0 Then
            If Not theInternet.DownloadCancelled Then
                Call frmMain.WriteLogEntry("Update check complete.", LogLevel1)
                lblCheckProgress.Caption = "Update check complete."
            Else
                Call frmMain.WriteLogEntry("Update check cancelled by user.", LogLevel1)
                lblCheckProgress.Caption = "Update check cancelled by user."
            End If
        Else
            lblModCatFails.Visible = True
            If Not theInternet.DownloadCancelled Then
                Call frmMain.WriteLogEntry("Update check complete. " & CStr(iFails) & " mods' update checks failed.", LogLevel1)
                lblCheckProgress.Caption = "Update check complete. " & CStr(iFails) & " mods' update checks failed."
            Else
                Call frmMain.WriteLogEntry("Update check cancelled by user. " & CStr(iFails) & " mods' update checks failed.", LogLevel1)
                lblCheckProgress.Caption = "Update check cancelled by user. " & CStr(iFails) & " mods' update checks failed."
            End If
        End If
    Else
        Call frmMain.WriteLogEntry("Update check skipped - no connection to Internet.", LogLevel1)
        lblCheckProgress.Caption = "Update check skipped - no connection to Internet."
    End If
    cmdCancel.Caption = "Back to Launch Menu"
    Call EnableControls(True)
    Call FilterModCat
    Call CallStackPop
End Sub

Public Function UpdateUpdateRecord(ByVal iRecord As Integer) As Integer
    Dim iCounter As Integer
    Dim iLine As Integer
    Dim iPos As Integer
    Dim sBuffer As String
    Dim sArray() As String
    Dim sValue As String
    Dim sFlag As String
    Dim iUpdate As Integer
    Dim iMirror As Integer
    Dim bPluginAuthed As Boolean
    Dim bCentralOk As Boolean
    Dim iFails As Integer
    Call CallStackPush(Me.Name & ".UpdateUpdateRecord(" & CStr(iRecord))
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    iFails = 0
    If iRecord = 0 Then
        'Need to update all central records from the current Central catalogue
        lblCheckProgress.Caption = "Downloading central catalogue data..."
        Call lblCheckProgress.Refresh
        If theInternet.CopyURLToString(ModCatUPD, sBuffer, Me) Then
            sBuffer = ConvertEOL(sBuffer, vbCrLf)
            If Not theInternet.DownloadCancelled Then
                'Process all the URLs from the central catalogue
                sArray = Split(sBuffer, vbCrLf)
                iLine = 0
                Do While iLine <= UBound(sArray())
                    iPos = InStr(1, sArray(iLine), "=")
                    If iPos > 1 Then
                        sValue = Mid$(sArray(iLine), iPos + 1) 'URL
                        sFlag = Left$(sArray(iLine), iPos - 1) 'index
                        'Check if we have already got a record for this one
                        iCounter = 1
                        Do While iCounter <= UpdateRecordCount
                            If Len(StripNumbers(sFlag)) = 0 Then
                                'numeric index so not plugin
                                If UpdateRecords(iCounter).CheckURL = sValue Then
                                    Exit Do
                                End If
                            Else
                                'Non-numeric index so must be a plugin
                                If UpdateRecords(iCounter).ModPluginID = sFlag Then
                                    UpdateRecords(iCounter).CheckURL = sValue
                                    Exit Do
                                End If
                            End If
                            iCounter = iCounter + 1
                        Loop
                        'If we haven't got a record for this one, create a record. We'll download the UPD later though.
                        If iCounter = UpdateRecordCount + 1 Then
                            UpdateRecordCount = UpdateRecordCount + 1
                            pbarModCat.Max = pbarModCat.Max + 1 'we've already set this so need to update it
                            If UBound(UpdateRecords()) < UpdateRecordCount Then Call RedimAndInit(UpdateRecordCount + 20)
                            Call InitialiseRecord(UpdateRecordCount)
                            UpdateRecords(UpdateRecordCount).CheckURL = sValue
                            If Len(StripNumbers(sFlag)) <> 0 Then
                                'Non-numeric index so must be a plugin
                                UpdateRecords(UpdateRecordCount).ModPluginID = sFlag 'although this will get updated from the UPD when we download it later anyway
                            End If
                        End If
                    End If
                    iLine = iLine + 1
                Loop
            End If
        Else
            Call frmMain.WriteLogEntry("Failed to download central mod catalogue " & Quote(ModCatUPD))
        End If
    Else
        'download the UPD and update the record
        UpdateRecords(iRecord).FailReason = ""
        If Len(UpdateRecords(iRecord).CheckURL) <> 0 Then
            If Len(UpdateRecords(iRecord).ModName) <> 0 Then
                lblCheckProgress.Caption = "Checking for updates... " & UpdateRecords(iRecord).ModName
            Else
                lblCheckProgress.Caption = "Checking for updates... " & UpdateRecords(iRecord).CheckURL
            End If
            Call lblCheckProgress.Refresh
            If theInternet.CopyURLToString(UpdateRecords(iRecord).CheckURL, sBuffer, Me) Then
                    sBuffer = ConvertEOL(sBuffer, vbCrLf)
                    UpdateRecords(iRecord).ModName = ReadINIStrMemory(sBuffer, "General", "Name")
                    UpdateRecords(iRecord).ModLatestVersion = ReadINIStrMemory(sBuffer, "General", "LatestVersion")
                    UpdateRecords(iRecord).ModDate = ReadINIStrMemory(sBuffer, "General", "Date")
                    UpdateRecords(iRecord).ModAuthor = ReadINIStrMemory(sBuffer, "General", "Author")
                    UpdateRecords(iRecord).ModWebsite = ReadINIStrMemory(sBuffer, "General", "Website")
                    UpdateRecords(iRecord).ModDescription = ReadINIStrMemory(sBuffer, "General", "Description")
                    UpdateRecords(iRecord).ModCampaigns = ReadINIStrMemory(sBuffer, "General", "Campaigns")
                    UpdateRecords(iRecord).ModTXVersion = ReadINIStrMemory(sBuffer, "General", "TXVersion")
                    UpdateRecords(iRecord).ModFA2Version = ReadINIStrMemory(sBuffer, "General", "FA2Version")
                    UpdateRecords(iRecord).ModGameIsRA2 = BooleanStringToBoolean(ReadINIStrMemory(sBuffer, "General", "IsForRA2", , , "no"))
                    Select Case UCase$(ReadINIStrMemory(sBuffer, "General", "ModType"))
                    Case "MOD": UpdateRecords(iRecord).ModType = TypeMod
                    Case "PLUGIN": UpdateRecords(iRecord).ModType = TypePlugin
                    Case "FA2MOD": UpdateRecords(iRecord).ModType = TypeFA2Mod
                    Case "USERTOOL", "DEVTOOL", "MODTOOL", "TOOL", "program": UpdateRecords(iRecord).ModType = TypeProgram
                    End Select
                    sValue = ReadINIStrMemory(sBuffer, "General", "PluginID")
                    If Len(sValue) <> 0 Then
                        UpdateRecords(iRecord).ModPluginID = sValue
                        UpdateRecords(iRecord).ModType = TypePlugin
                    End If
                    If UpdateRecords(iRecord).ModType = TypePlugin Then
                        'there is only ever one update record for a given plugin because we don't allow plugins to specify their own UPD
                        'just need to find what version the user has
                        iCounter = frmMain.GetLatestPlugin(UpdateRecords(iRecord).ModPluginID)
                        If iCounter <> -1 Then
                            UpdateRecords(iRecord).ModUserVersion = Plugins(iCounter).PluginVersion
                            UpdateRecords(iRecord).CheckModNum = iCounter
                        Else
                            UpdateRecords(iRecord).ModUserVersion = ""
                        End If
                    Else
                        If UpdateRecords(iRecord).CheckModNum = -1 Then
                            'we might already have the mod but didn't know the update check file
                            'however, now we have enough information to search by name and type
                            iCounter = 0
                            Do While iCounter <= ModCount
                                If UpdateRecords(iRecord).ModName = Mods(iCounter).ModName Then
                                    If UpdateRecords(iRecord).ModType = Mods(iCounter).ModType Then Exit Do
                                End If
                                iCounter = iCounter + 1
                            Loop
                            If iCounter <= ModCount Then
                                'we found a match!
                                UpdateRecords(iRecord).CheckModNum = iCounter
                                UpdateRecords(iRecord).ModUserVersion = Mods(iCounter).ModVersion
                                'but we could have a problem here. what if the mod specified a UPD and the central UPD was different?
                                'this would mean that we have a second redundant mod entry in the catalogue
                                iCounter = 1
                                Do While iCounter <= UpdateRecordCount
                                    If iRecord <> iCounter Then
                                        If UpdateRecords(iCounter).CheckModNum = UpdateRecords(iRecord).CheckModNum Then
                                            If CompareVersions(UpdateRecords(iCounter).ModLatestVersion, ">", UpdateRecords(iRecord).ModLatestVersion) Then
                                                'this other entry offers a more recent version so we'll use that
                                                Call InitialiseRecord(iRecord)
                                                iRecord = -1
                                                Exit Do
                                            Else
                                                'this other entry isn't so great
                                                Call InitialiseRecord(iCounter)
                                                Exit Do
                                            End If
                                        End If
                                    End If
                                    iCounter = iCounter + 1
                                Loop
                            End If
                        End If
                    End If
                    If iRecord <> -1 Then
                        UpdateRecords(iRecord).CheckChangeLog = ReadINIStrMemory(sBuffer, "General", "ChangeLogURL")
                        UpdateRecords(iRecord).CheckUpdateOnly = False
                        UpdateRecords(iRecord).CheckDownloadURL = ""
                        UpdateRecords(iRecord).CheckDownloadSize = 0
                        If CompareVersions(UpdateRecords(iRecord).ModLatestVersion, ">", UpdateRecords(iRecord).ModUserVersion) Then
                            If Not OptFullDownloads Then
                                If Len(UpdateRecords(iRecord).ModUserVersion) <> 0 Then
                                    iUpdate = 0
                                    sValue = ReadINIStrMemory(sBuffer, "General", "Update" & CStr(iUpdate))
                                    Do While Len(sValue) <> 0
                                        If CompareVersions(sValue, "=", UpdateRecords(iRecord).ModUserVersion) Then
                                            UpdateRecords(iRecord).CheckUpdateOnly = True
                                            Exit Do
                                        End If
                                        iUpdate = iUpdate + 1
                                        sValue = ReadINIStrMemory(sBuffer, "General", "Update" & CStr(iUpdate))
                                    Loop
                                End If
                            End If
                            iCounter = -1
                            If UpdateRecords(iRecord).CheckUpdateOnly Then
                                iMirror = 0
                                sValue = ReadINIStrMemory(sBuffer, "Downloads", "Update" & CStr(iUpdate) & "Mirror" & CStr(iMirror))
                                Do While Len(sValue) <> 0
                                    iCounter = iCounter + 1
                                    ReDim sArray(iCounter)
                                    sArray(iCounter) = sValue
                                    iMirror = iMirror + 1
                                    sValue = ReadINIStrMemory(sBuffer, "Downloads", "Update" & CStr(iUpdate) & "Mirror" & CStr(iMirror))
                                Loop
                                Do While iCounter <> -1
                                    iMirror = Random(0, iCounter)
                                    UpdateRecords(iRecord).CheckDownloadSize = theInternet.GetRemoteFileSize(sArray(iMirror))
                                    If UpdateRecords(iRecord).CheckDownloadSize <> 0 Then
                                        'found a download!
                                        UpdateRecords(iRecord).CheckDownloadURL = sArray(iMirror)
                                        Exit Do
                                    Else
                                        'failed to get remote file size so won't be able to get the file either
                                        Do While iMirror < iCounter
                                            sArray(iMirror) = sArray(iMirror + 1)
                                            iMirror = iMirror + 1
                                        Loop
                                        iCounter = iCounter - 1
                                    End If
                                Loop
                            End If
                            If iCounter = -1 Then
                                'for whatever reason, an update-only installer will not be downloaded
                                UpdateRecords(iRecord).CheckUpdateOnly = False
                                iMirror = 0
                                sValue = ReadINIStrMemory(sBuffer, "Downloads", "FullMirror" & CStr(iMirror))
                                Do While Len(sValue) <> 0
                                    iCounter = iCounter + 1
                                    ReDim sArray(iCounter)
                                    sArray(iCounter) = sValue
                                    iMirror = iMirror + 1
                                    sValue = ReadINIStrMemory(sBuffer, "Downloads", "FullMirror" & CStr(iMirror))
                                Loop
                                Do While iCounter <> -1
                                    iMirror = Random(0, iCounter)
                                    UpdateRecords(iRecord).CheckDownloadSize = theInternet.GetRemoteFileSize(sArray(iMirror))
                                    If UpdateRecords(iRecord).CheckDownloadSize <> 0 Then
                                        'found a download!
                                        UpdateRecords(iRecord).CheckDownloadURL = sArray(iMirror)
                                        Exit Do
                                    Else
                                        'failed to get remote file size so won't be able to get the file either
                                        Do While iMirror < iCounter
                                            sArray(iMirror) = sArray(iMirror + 1)
                                            iMirror = iMirror + 1
                                        Loop
                                        iCounter = iCounter - 1
                                    End If
                                Loop
                            End If
                            'If Counter = -1 then for some reason or another, a download is not available
                        'Else no point checking for downloads
                        End If
                    Else
                        Call frmMain.WriteLogEntry("Installed mod's update check file offers a more recent version.", LogLevel1)
                    End If
                Else
                    If Not theInternet.DownloadCancelled Then
                        Call frmMain.WriteLogEntry("Failed to download update check file " & Quote(UpdateRecords(iRecord).CheckURL), LogLevel1)
                        UpdateRecords(iRecord).FailReason = "Failed to download update check file."
                        iFails = iFails + 1
                    End If
                End If
        Else
            UpdateRecords(iRecord).FailReason = "No update check URL specified."
            iFails = iFails + 1
        End If
    End If
    Call CallStackPop
    UpdateUpdateRecord = iFails
End Function

Public Sub InitialiseRecord(ByVal iRecord As Integer, Optional ByVal bGetFromMod As Boolean = False)
    Dim iMod As Integer
    UpdateRecords(iRecord).ModLatestVersion = ""
    UpdateRecords(iRecord).ModDate = ""
    UpdateRecords(iRecord).ModAuthor = ""
    UpdateRecords(iRecord).ModDescription = ""
    UpdateRecords(iRecord).ModCampaigns = ""
    UpdateRecords(iRecord).ModTXVersion = ""
    UpdateRecords(iRecord).ModFA2Version = ""
    UpdateRecords(iRecord).ModPluginID = ""
    UpdateRecords(iRecord).CheckChangeLog = ""
    UpdateRecords(iRecord).CheckUpdateOnly = False
    UpdateRecords(iRecord).CheckDownloadURL = ""
    UpdateRecords(iRecord).CheckDownloadSize = -1
    UpdateRecords(iRecord).FailReason = ""
    If bGetFromMod And Len(UpdateRecords(iRecord).ModPluginID) = 0 Then
        iMod = UpdateRecords(iRecord).CheckModNum
        UpdateRecords(iRecord).ModName = Mods(iMod).ModName
        UpdateRecords(iRecord).ModType = Mods(iMod).ModType
        UpdateRecords(iRecord).ModGameIsRA2 = Mods(iMod).ModIsForRA2
        UpdateRecords(iRecord).ModWebsite = Mods(iMod).ModWebsite
        UpdateRecords(iRecord).ModUserVersion = Mods(iMod).ModVersion
        UpdateRecords(iRecord).CheckURL = Mods(iMod).ModUpdateCheckURL
    Else
        UpdateRecords(iRecord).ModName = ""
        UpdateRecords(iRecord).ModType = -1
        UpdateRecords(iRecord).ModGameIsRA2 = False
        UpdateRecords(iRecord).ModWebsite = ""
        UpdateRecords(iRecord).CheckModNum = -1
        UpdateRecords(iRecord).ModUserVersion = ""
        UpdateRecords(iRecord).CheckURL = ""
    End If
End Sub

Private Sub EnableControls(ByVal Enable As Boolean)
    frameAvailableMods.Enabled = Enable
        listModCat.Enabled = Enable
        If Enable Then
            listModCat.BackColor = &H80000005
        Else
            listModCat.BackColor = &H8000000F
        End If
    frameFilter.Enabled = Enable
        lblFilterModType.Enabled = Enable
        cboxFilterModType(0).Enabled = Enable
        cboxFilterModType(1).Enabled = Enable
        cboxFilterModType(2).Enabled = Enable
        cboxFilterModType(3).Enabled = Enable
        lblFilterGame.Enabled = Enable
        cboxFilterGame(0).Enabled = Enable
        cboxFilterGame(1).Enabled = Enable
        lblFilterUpdates.Enabled = Enable
        cboxFilterUpdates(0).Enabled = Enable
        cboxFilterUpdates(1).Enabled = Enable
        cboxFilterUpdates(2).Enabled = Enable
    frameModDetails.Enabled = Enable
        lblGame.Enabled = Enable
        lblModGame.Enabled = Enable
        lblVersion.Enabled = Enable
        lblModVersion.Enabled = Enable
        lblYourVersion.Enabled = Enable
        lblModYourVersion.Enabled = Enable
        lblDate.Enabled = Enable
        lblModDate.Enabled = Enable
        lblAuthor.Enabled = Enable
        lblModAuthor.Enabled = Enable
        lblWebsite.Enabled = Enable
        lblModWebsite.Enabled = Enable
        lblSize.Enabled = Enable
        lblModSize.Enabled = Enable
        lblCampaigns.Enabled = Enable
        lblModCampaigns.Enabled = Enable
        lblTX.Enabled = Enable
        lblModTX.Enabled = Enable
        lblModDescription.Enabled = Enable
        cmdDownload.Enabled = Enable
        cmdChangeLog.Enabled = Enable
End Sub

Private Sub FilterModCat()
    Dim iRecord As Integer
    Dim iIncluded As Integer
    listModCat.Sorted = False
    iIncluded = 0
    Call listModCat.ListItems.Clear
    listModCat.Sorted = False
    iRecord = 1
    Do While iRecord <= UpdateRecordCount
        Do
            'Name
            If Len(UpdateRecords(iRecord).ModName) = 0 Then Exit Do
            'Version
            If Len(UpdateRecords(iRecord).ModLatestVersion) = 0 Then Exit Do
            'Updates
            If UpdateRecords(iRecord).CheckModNum = -1 Then
                If cboxFilterUpdates(1).Value <> 1 Then Exit Do
            Else
                If CompareVersions(UpdateRecords(iRecord).ModLatestVersion, "<=", UpdateRecords(iRecord).ModUserVersion) Then
                    If cboxFilterUpdates(2).Value <> 1 Then Exit Do
                Else
                    If cboxFilterUpdates(0).Value <> 1 Then Exit Do
                End If
            End If
            'Mod Type
            If UpdateRecords(iRecord).ModType = -1 Then Exit Do
            If cboxFilterModType(UpdateRecords(iRecord).ModType).Value <> 1 Then Exit Do
            'Game
            Select Case UpdateRecords(iRecord).ModType
            Case TypeMod, TypePlugin
                If UpdateRecords(iRecord).ModGameIsRA2 Then
                    If cboxFilterGame(0).Value <> 1 Then Exit Do
                Else
                    If cboxFilterGame(1).Value <> 1 Then Exit Do
                End If
            End Select
            'If we're still here then can include this one
            If UpdateRecords(iRecord).CheckURL = Mods(LBModNum).ModUpdateCheckURL Then
                'Launch Base - needs emphasising
                Call listModCat.ListItems.Add(, , " " & UpdateRecords(iRecord).ModName)
                listModCat.ListItems(listModCat.ListItems.Count).Bold = True
            Else
                Call listModCat.ListItems.Add(, , UpdateRecords(iRecord).ModName)
            End If
            Select Case UpdateRecords(iRecord).ModType
            Case TypeMod: listModCat.ListItems(listModCat.ListItems.Count).SubItems(1) = "Mod"
            Case TypeFA2Mod: listModCat.ListItems(listModCat.ListItems.Count).SubItems(1) = "FA2 Mod"
            Case TypeProgram: listModCat.ListItems(listModCat.ListItems.Count).SubItems(1) = "Tool"
            Case TypePlugin: listModCat.ListItems(listModCat.ListItems.Count).SubItems(1) = "Plugin"
            End Select
            listModCat.ListItems(listModCat.ListItems.Count).Tag = CStr(iRecord)
            iIncluded = iIncluded + 1
            Exit Do
        Loop
        iRecord = iRecord + 1
    Loop
    listModCat.Sorted = True
    If iIncluded <> 0 Then
        listModCat.SelectedItem = listModCat.ListItems.Item(1)
        Call listModCat_ItemClick(listModCat.SelectedItem)
    Else
        Call listModCat_ItemClick(Nothing)
    End If
End Sub

Private Sub cboxFilterGame_Click(Index As Integer)
    Call FilterModCat
End Sub

Private Sub cboxFilterModType_Click(Index As Integer)
    Call FilterModCat
End Sub

Private Sub cboxFilterUpdates_Click(Index As Integer)
    Call FilterModCat
End Sub

Private Sub cmdCancel_Click()
    If Busy Then
        Call theInternet.CancelDownload
    Else
        Me.Hide
        frmMain.Show
        frmMain.SetFocus
    End If
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

Private Sub cmdChangeLog_Click()
    Dim iRecord As Integer
    Dim LocalURL As String
    Dim DownloadError As Boolean
    Dim iDownloadSize As Long
    Dim sRemoteURL As String
    Call CallStackPush(Me.Name & ".cmdChangeLog_Click()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    Call ShowPleaseWait("Downloading Change Log...")
    iRecord = Val(listModCat.SelectedItem.Tag)
    Busy = True
    sRemoteURL = UpdateRecords(iRecord).CheckChangeLog
    LocalURL = JoinPath(SETUPDIR, GetFileName(sRemoteURL))
    iDownloadSize = theInternet.GetRemoteFileSize(sRemoteURL)
    If iDownloadSize <> 0 Then
        If FreeDiskSpace(UCase$(Left$(LocalURL, 1))) > OptSafetySpace Then
            Call frmMain.WriteLogEntry("Downloading " & Quote(sRemoteURL) & " to " & Quote(LocalURL), LogLevel1)
            If theInternet.CopyURLToFile(sRemoteURL, LocalURL) Then
                Call HidePleaseWait
                Call frmMain.WriteLogEntry("Download complete. Executing " & Quote(LocalURL), LogLevel1)
                If OpenLocation(LocalURL) < 32 Then Call frmMain.WriteLogEntry("Error opening " & Quote(LocalURL), LogMsgBox)
            Else
                Call HidePleaseWait
                Call frmMain.WriteLogEntry("Failed to download " & sRemoteURL, LogMsgBox)
            End If
        Else
            Call HidePleaseWait
            Call frmMain.WriteLogEntry("Insufficient free disk space to download " & Quote(sRemoteURL) & " to " & Quote(LocalURL), LogMsgBoxExclaim)
        End If
    Else
        Call HidePleaseWait
        Call frmMain.WriteLogEntry("Failed to access " & sRemoteURL, LogMsgBox)
    End If
    Busy = False
    Call CallStackPop
End Sub

Private Sub cmdDownload_Click()
    Dim mbResult As VbMsgBoxResult
    Dim iRecord As Integer
    Dim LocalURL As String
    Dim sRemoteURL As String
    Dim DownloadError As Boolean
    Dim iDownloadSize As Long
    Call CallStackPush(Me.Name & ".cmdDownload_Click()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    iRecord = Val(listModCat.SelectedItem.Tag)
    Busy = True
    sRemoteURL = UpdateRecords(iRecord).CheckDownloadURL
    LocalURL = JoinPath(SETUPDIR, GetFileName(sRemoteURL))
    Call frmMain.WriteLogEntry("Checking for free disk space on drive " & UCase$(Left(EXEDIR, 1)), LogLevel1)
    If FreeDiskSpace(UCase$(Left$(LocalURL, 1))) > OptSafetySpace + UpdateRecords(iRecord).CheckDownloadSize Then
        Call frmMain.WriteLogEntry("Downloading " & Quote(sRemoteURL) & " to " & Quote(LocalURL), LogLevel1)
        Me.Enabled = False
        Call frmDownloading.Show
        Call frmDownloading.DownloadProgress(0, UpdateRecords(iRecord).CheckDownloadSize)
        If theInternet.CopyURLToFile(sRemoteURL, LocalURL, frmDownloading) Then
            Call frmMain.WriteLogEntry("Download complete.", LogLevel1)
            Call frmMain.CheckForUpdate_DownloadHistory(LocalURL, UpdateRecords(iRecord).ModType, UpdateRecords(iRecord).ModName, UpdateRecords(iRecord).ModLatestVersion, IIf(UpdateRecords(iRecord).CheckUpdateOnly, UpdateRecords(iRecord).ModUserVersion, ""))
            Unload frmDownloading
            Me.Enabled = True
            mbResult = MsgBox("Launch Base will now close so that " & UpdateRecords(iRecord).ModName & " can be installed.", vbOKCancel + vbInformation, App.Title)
            If mbResult = vbOK Then
                Call frmMain.WriteLogEntry("Executing " & LocalURL, LogLevel1)
                Call frmMain.Shutdown(True, False)
                Call Sleep(500)
                Call Shell(Quote(LocalURL), vbNormal)
                Unload Me
                Exit Sub
            Else
                Call frmMain.WriteLogEntry("User chose not to execute " & LocalURL, LogLevel1)
            End If
        Else
            If theInternet.DownloadCancelled Then
                Unload frmDownloading
                Me.Enabled = True
                Call frmMain.WriteLogEntry("Download cancelled by user.", LogLevel1)
            Else
                Unload frmDownloading
                Me.Enabled = True
                Call frmMain.WriteLogEntry("Failed to download " & sRemoteURL, LogMsgBoxExclaim)
            End If
        End If
    Else
        Call frmMain.WriteLogEntry("Insufficient free disk space to download " & Quote(sRemoteURL) & " to " & Quote(LocalURL), LogMsgBox)
    End If
    Busy = False
    Call CallStackPop
End Sub

Private Sub Form_Load()
    lblModCatFails.Visible = False
    UpdateRecordCount = -1
    Call listModCat.ColumnHeaders.Add(, , "Name", 3480)
    Call listModCat.ColumnHeaders.Add(, , "Type", 885)
    cboxFilterModType(0).Value = BooleanToInteger(OptModCatFilterModType0)
    cboxFilterModType(1).Value = BooleanToInteger(OptModCatFilterModType1)
    cboxFilterModType(2).Value = BooleanToInteger(OptModCatFilterModType2)
    cboxFilterModType(3).Value = BooleanToInteger(OptModCatFilterModType3)
    cboxFilterGame(0).Value = BooleanToInteger(OptModCatFilterGame0)
    cboxFilterGame(1).Value = BooleanToInteger(OptModCatFilterGame1)
    cboxFilterUpdates(0).Value = BooleanToInteger(OptModCatFilterUpdates0)
    cboxFilterUpdates(1).Value = BooleanToInteger(OptModCatFilterUpdates1)
    cboxFilterUpdates(2).Value = BooleanToInteger(OptModCatFilterUpdates2)
    Call listModCat_ItemClick(Nothing)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = 1
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call theInternet.Disconnect
    Call WriteINIStr("Options", "ModCatFilterModType0", IntegerToYesNo(cboxFilterModType(0).Value), ProgramINI)
    Call WriteINIStr("Options", "ModCatFilterModType1", IntegerToYesNo(cboxFilterModType(1).Value), ProgramINI)
    Call WriteINIStr("Options", "ModCatFilterModType2", IntegerToYesNo(cboxFilterModType(2).Value), ProgramINI)
    Call WriteINIStr("Options", "ModCatFilterModType3", IntegerToYesNo(cboxFilterModType(3).Value), ProgramINI)
    Call WriteINIStr("Options", "ModCatFilterGame0", IntegerToYesNo(cboxFilterGame(0).Value), ProgramINI)
    Call WriteINIStr("Options", "ModCatFilterGame1", IntegerToYesNo(cboxFilterGame(1).Value), ProgramINI)
    Call WriteINIStr("Options", "ModCatFilterUpdates0", IntegerToYesNo(cboxFilterUpdates(0).Value), ProgramINI)
    Call WriteINIStr("Options", "ModCatFilterUpdates1", IntegerToYesNo(cboxFilterUpdates(1).Value), ProgramINI)
    Call WriteINIStr("Options", "ModCatFilterUpdates2", IntegerToYesNo(cboxFilterUpdates(2).Value), ProgramINI)
    If Not Busy Then
        frmModCat.Hide
        frmMain.Show
        frmMain.SetFocus
    End If
End Sub

Private Sub frameModDetails_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblModWebsite.ForeColor = ColorURL
End Sub

Private Sub lblModCatFails_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        frmModCat.Enabled = False
        Call frmModCatFails.Show
        Call frmModCatFails.LoadFails
    End If
End Sub

Private Sub lblModWebsite_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case vbLeftButton
        If OpenLocation(ReplaceString(lblModWebsite.Caption, "&&", "&")) < 32 Then
            Call MsgBox("Unable to open " & Quote(ReplaceString(lblModWebsite.Caption, "&&", "&")) & ".", vbOKOnly + vbInformation, App.Title)
        End If
    Case vbRightButton
        Call PopupMenu(menu_rc2, 2, X + frameModDetails.Left + lblModWebsite.Left, Y + frameModDetails.Top + lblModWebsite.Top)
    End Select
End Sub

Private Sub menu_url_Click()
    Clipboard.Clear
    Clipboard.SetText lblModWebsite.Caption
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblModWebsite.ForeColor = ColorURL
End Sub

Private Sub listModCat_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If listModCat.SortKey <> (ColumnHeader.Index - 1) Then
        listModCat.SortKey = ColumnHeader.Index - 1
    Else
        If listModCat.SortOrder = lvwAscending Then
            listModCat.SortOrder = lvwDescending
        Else
            listModCat.SortOrder = lvwAscending
        End If
    End If
End Sub

Private Sub listModCat_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim iRecord As Integer
    If Item Is Nothing Then
        lblModVersion.Caption = ""
        lblModYourVersion.Caption = ""
        lblModGame.Caption = ""
        lblModWebsite.Caption = ""
        lblModWebsite.Visible = False
        lblModDate.Caption = ""
        lblModAuthor.Caption = ""
        lblModSize.Caption = ""
        lblCampaigns.Caption = ""
        lblModCampaigns.Caption = ""
        lblTX.Visible = False
        lblModTX.Visible = False
        lblModDescription.Caption = ""
        cmdDownload.Enabled = False
        cmdChangeLog.Enabled = False
    Else
        iRecord = Val(Item.Tag)
        lblModAuthor.Caption = DoubleAmpersand(UpdateRecords(iRecord).ModAuthor)
        lblModWebsite.Caption = DoubleAmpersand(UpdateRecords(iRecord).ModWebsite)
        lblModWebsite.Visible = Len(lblModWebsite.Caption) <> 0
        lblModWebsite.ForeColor = ColorURL
        lblModDescription.Caption = DoubleAmpersand(UpdateRecords(iRecord).ModDescription)
        lblModDate.Caption = DoubleAmpersand(UpdateRecords(iRecord).ModDate)
        lblModVersion.Caption = DoubleAmpersand(UpdateRecords(iRecord).ModLatestVersion)
        If Len(UpdateRecords(iRecord).ModUserVersion) <> 0 Then
            lblModYourVersion.Caption = UpdateRecords(iRecord).ModUserVersion
        Else
            lblModYourVersion.Caption = "Not Installed"
        End If
        Select Case UpdateRecords(iRecord).ModType
        Case TypeMod
            lblCampaigns.Caption = "Campaigns:"
            lblModCampaigns.Caption = DoubleAmpersand(UpdateRecords(iRecord).ModCampaigns)
            lblTX.Visible = True
            lblModTX.Visible = True
            lblModTX.Caption = UpdateRecords(iRecord).ModTXVersion
            If UpdateRecords(iRecord).ModGameIsRA2 Then
                lblModGame.Caption = "Red Alert 2"
            Else
                lblModGame.Caption = "Yuri's Revenge"
            End If
        Case TypeFA2Mod
            lblCampaigns.Caption = "FA2 Version:"
            lblModCampaigns.Caption = UpdateRecords(iRecord).ModFA2Version
            lblTX.Visible = True
            lblModTX.Visible = True
            lblModTX.Caption = UpdateRecords(iRecord).ModTXVersion
            lblModGame.Caption = "n/a"
        Case TypePlugin
            lblCampaigns.Caption = ""
            lblModCampaigns.Caption = ""
            lblTX.Visible = False
            lblModTX.Visible = False
            If UpdateRecords(iRecord).ModGameIsRA2 Then
                lblModGame.Caption = "Red Alert 2"
            Else
                lblModGame.Caption = "Yuri's Revenge"
            End If
        Case Else
            lblCampaigns.Caption = ""
            lblModCampaigns.Caption = ""
            lblTX.Visible = False
            lblModTX.Visible = False
            lblModGame.Caption = "n/a"
        End Select
        If CompareVersions(UpdateRecords(iRecord).ModLatestVersion, ">", UpdateRecords(iRecord).ModUserVersion) Then
            If UpdateRecords(iRecord).CheckDownloadSize <> -1 And Len(UpdateRecords(iRecord).CheckDownloadURL) <> 0 Then
                lblModSize.Caption = DataSize(UpdateRecords(iRecord).CheckDownloadSize)
                If FreeDiskSpace(UCase$(Left$(SETUPDIR, 1))) > (OptSafetySpace + UpdateRecords(iRecord).CheckDownloadSize) Then
                    lblModSize.ForeColor = ColorGood
                    cmdDownload.Enabled = True
                Else
                    lblModSize.ForeColor = ColorBad
                    cmdDownload.Enabled = False
                End If
            Else
                lblModSize.ForeColor = ColorBad
                cmdDownload.Enabled = False
                lblModSize.Caption = "Visit website for download."
            End If
        Else
            lblModSize.ForeColor = ColorBad
            cmdDownload.Enabled = False
            lblModSize.Caption = "You already have the latest version."
        End If
        cmdChangeLog.Enabled = Len(UpdateRecords(iRecord).CheckChangeLog) <> 0
        'NOW TX/FA2 version stuff - dont disable download but do mark colours, etc
        lblModTX.ForeColor = ColorGood
        lblModTX.Caption = "Not Required"
        If Len(UpdateRecords(iRecord).ModTXVersion) <> 0 Then
            If UpdateRecords(iRecord).ModTXVersion = "-1" Then
                lblModTX.Caption = "Not Allowed"
            Else
                lblModTX.Caption = "Version " & UpdateRecords(iRecord).ModTXVersion & " Required"
                If Not frmMain.PrerequisiteCheckTX(UpdateRecords(iRecord).ModTXVersion) Then lblModTX.ForeColor = ColorBad
            End If
        End If
        If UpdateRecords(iRecord).ModType = TypeFA2Mod Then
            lblModCampaigns.ForeColor = ColorGood
            If Len(UpdateRecords(iRecord).ModFA2Version) = 0 Then
                If Len(Mods(FA2ModNum).ModVersion) = 0 Then
                    lblModCampaigns.ForeColor = ColorBad
                    lblModCampaigns.Caption = "Any Version Required"
                Else
                    lblModCampaigns.Caption = "Any Version Allowed"
                End If
            Else
                lblModCampaigns.Caption = "Version " & UpdateRecords(iRecord).ModFA2Version & " Required"
                If Mods(FA2ModNum).ModVersion <> UpdateRecords(iRecord).ModFA2Version Then lblModCampaigns.ForeColor = ColorBad
            End If
        End If
    End If
End Sub
