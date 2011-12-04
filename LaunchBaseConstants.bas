Attribute VB_Name = "LaunchBaseConstants"
Option Explicit
Public Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Public Declare Function LBItemFromPt Lib "comctl32" (ByVal hWnd As Long, ByVal ptx As Long, ByVal pty As Long, ByVal bAutoScroll As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const SYNCHRONIZE = &H100000
Public Const INFINITE = -1&
Public Const LB_GETITEMHEIGHT = &H1A1
Public Type POINTAPI
   X As Long
   Y As Long
End Type
Public theInternet As MarshallxInet
Public Type UpdateRecord
    CheckURL As String
    CheckDownloadURL As String
    CheckDownloadSize As Long
    CheckChangeLog As String
    CheckUpdateOnly As Boolean
    CheckModNum As Integer
    FailReason As String
    ModName As String
    ModLatestVersion As String
    ModUserVersion As String
    ModDate As String
    ModAuthor As String
    ModWebsite As String
    ModDescription As String
    ModCampaigns As String
    ModType As Integer
    ModGameIsRA2 As Boolean
    ModPluginID As String
    ModTXVersion As String
    ModFA2Version As String
End Type
Public UpdateRecords() As UpdateRecord
Public UpdateRecordCount As Integer
Public Type ModRecord
    ModName As String
    ModVersion As String
    ModDate As String
    ModAuthor As String
    ModWebsite As String
    ModCampaigns As String
    ModDescription As String
    ModManual As String
    ModBanner As String
    ModSize As Double
    ModAllowTX As Boolean
    ModTXVersion As String
    ModUseAres As Boolean
    ModType As Integer
    ModPath As String
    ModSound1 As String
    ModSound2 As String
    ModUpdateCheckURL As String
    ModIsForRA2 As Boolean
    ModLiblist As String
    ModFA2Version As String
    ModIsUserTool As Boolean
    ModProgram As String
    ModShowParams As Boolean
    ModShutdownLB As Boolean
    ModParams As String
    ModScrnFormat As String
    ModSnapFormat As String
    ModUseYuriUI As Boolean
    ModGameMode As String
    ModMapIndex As String
End Type
Public Type PluginRecord
    PluginPath As String
    PluginName As String
    PluginVersion As String
    PluginDate As String
    PluginAuthor As String
    PluginWebsite As String
    PluginSize As String
    PluginDescription As String
    PluginRPLegacy As String
    PluginManual As String
    PluginID As String
End Type
Public Const DefaultSkinDir As String = "ren_glass"
Public Const MaxType As Integer = 3
Public Const TypeMod As Integer = 0
Public Const TypePlugin As Integer = 1
Public Const TypeFA2Mod As Integer = 2
Public Const TypeProgram As Integer = 3
Public Const LBModNum As Integer = 0
Public Const YRModNum As Integer = 1
Public Const RA2ModNum As Integer = 2
Public Const FA2ModNum As Integer = 3
Public Const HardCodedMods As Integer = 4
Public EXEDIR As String
Public RA2DIR As String
Public RESDIR As String
Public LOGFILE As String
Public LOGDIR As String
Public SETUPDIR As String
Public SetupsINI As String
Public ProgramINI As String
Public ModCatDB As String
Public ModCatLen As String
Public ModCatUPD As String
Public BACKUPDIR As String
Public DCoderDLL As Boolean
Public Enum MxLogTypeConstant
    LogLevel0 = 0
    LogIE = 1
    LogShutdown = 2
    LogMsgBox = 4
    LogMsgBoxExclaim = 8
    LogLevel1 = 16
    LogLevel2 = 32
End Enum
Public LangNames(9) As String
'Options
Public OptLogFile As Boolean
Public OptInitLog As Boolean
Public OptMaxLogSize As Double
Public OptLogLevel As Integer
Public OptLiveLog As Boolean
Public OptAdvancedMode As Boolean
Public OptRecompile As Boolean
Public OptLooseFileMode As Boolean
Public OptPersistentMod As Boolean
Public OptPersistentPlugin As Boolean
Public OptPersistentModBad As Boolean
Public OptPersistentPluginBad As Boolean
Public OptModSound1 As Boolean
Public OptModSound2 As Boolean
Public OptLBSounds As Boolean
Public OptLogAres As Boolean
Public OptCaptureAresDebug As Boolean
Public OptLogExcept As Boolean
Public OptLogExceptDesc As Boolean
Public OptWindowed As Boolean
Public OptRecord As Boolean
Public OptPlay As Boolean
Public OptSkipLogo As Boolean
Public OptUseCheckSums As Boolean
Public OptGameChecksums As Boolean
Public OptVerifyPlugins As Boolean
Public OptCheckModYPLFiles As Boolean
Public OptAutoTX As Boolean
Public OptAutoUpdate As Boolean
Public OptShowRA2 As Boolean
Public OptShowYR As Boolean
Public OptSpeedControl As Boolean
Public OptMPDebug As Boolean
Public OptSafetySpace As Double
Public OptFullDownloads As Boolean
Public OptAutoAresUpdate As Boolean
Public OptAresTester As Boolean
Public OptCustomSwitches As String
Public OptModCatFilterModType0 As Boolean
Public OptModCatFilterModType1 As Boolean
Public OptModCatFilterModType2 As Boolean
Public OptModCatFilterModType3 As Boolean
Public OptModCatFilterModType4 As Boolean
Public OptModCatFilterGame0 As Boolean
Public OptModCatFilterGame1 As Boolean
Public OptModCatFilterUpdates0 As Boolean
Public OptModCatFilterUpdates1 As Boolean
Public OptModCatFilterUpdates2 As Boolean
Public OptRA2Lang As Integer
Public OptYRLang As Integer
Public OptAresRevision As Long
Public OptSyringeRevision As Long
Public OptAresRevisionDataURL As String
Public OptAresRevisionDataURLURL As String
Public OptAresRevisionDataURLHDR As String
Public OptAresBranch As String
Public OptVideoBackBuffer As Boolean
Public OptAllowVRAMSidebar As Boolean
'Command line arguments
Public CL_game As String
Public CL_modnum As Integer
Public CL_playfile As String
Public CL_tx As Boolean
Public CL_dev As Boolean
Public CL_advanced As Boolean
Public CL_noexcept As Boolean
'Mod records
Public ModCount As Integer
Public PluginCount As Integer
Public Mods() As ModRecord
Public Plugins() As PluginRecord
Public SafeFiles() As String

Public Sub Init_Constants()
    Dim rModCat As UpdateRecord
    ModCatLen = Len(rModCat)
    EXEDIR = App.Path
    RESDIR = JoinPath(EXEDIR, "Resource")
    LOGFILE = JoinPath(EXEDIR, "LaunchBase.log")
    BACKUPDIR = JoinPath(EXEDIR, "Backup")
    LOGDIR = JoinPath(EXEDIR, "Logs")
    SETUPDIR = JoinPath(EXEDIR, "Setups")
    ProgramINI = JoinPath(EXEDIR, "LaunchBase.ini")
    ModCatDB = JoinPath(EXEDIR, "ModCat.lbd")
    SetupsINI = JoinPath(SETUPDIR, "Setups.ini")
    Set theInternet = New MarshallxInet
    theInternet.UserAgent = "Launch Base " & App.Major & "." & PadNum(App.Minor, 2) & "." & PadNum(App.Revision)
    LangNames(0) = "US"
    LangNames(2) = "German"
    LangNames(3) = "French"
    LangNames(8) = "Korean"
    LangNames(9) = "Chinese"
End Sub

