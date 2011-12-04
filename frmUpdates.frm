VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUpdates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YR Launch Base: Check For Updates"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "frmUpdates.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView listUpdates 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4048
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel Update Check"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   4320
      Width           =   3255
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download and Install Selected Update"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Update Check Progress"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8775
   End
   Begin VB.CommandButton cmdChangeLog 
      Caption         =   "View Change Log for Selected Update"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   3600
      Width           =   3255
   End
   Begin MSComctlLib.ProgressBar pbarDownload 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   3960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblDownloadPercent 
      Alignment       =   1  'Right Justify
      Caption         =   "(0%)"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDownloadAmount 
      Alignment       =   1  'Right Justify
      Caption         =   "0 / 0"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "frmUpdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Downloading As Boolean
Dim Checking As Boolean
Dim SelfConnect As Boolean
Public CancelDownload As Boolean
Dim UpdateCount As Integer
Dim UpdateDownloadURL() As String
Dim UpdateChangelogURL() As String
Dim UpdateDownloadSize() As Long
Dim UpdateModPath() As String
Dim DownloadSize As Long
Dim ShuttingDown As Boolean

Public Sub CheckForUpdates()
    Dim Counter As Integer
    Dim RModPath As String
    Dim RName As String
    Dim RVersion As String
    Dim RType As Integer
    Dim RTypeStr As String
    Dim RFolder As String
    Dim RUpdateCheckURL As String
    Dim LocalURL As String
    Dim PreviousVersion As String
    Dim LatestVersion As String
    Dim UpdateDownloadURL As String
    Dim UpdateDownloadSize As Long
    Dim DownloadURL As String
    Dim DownloadSize As Long
    Dim ChangeLogURL As String
    Dim DownloadError As Boolean
    Dim ErrVars(0) As Variant
    If Not frmMain.CL_noexcept Then On Error GoTo LocalErr
    UpdateCount = 0
    lblCheckProgress.Caption = "Initializing update check..."
    Call listUpdates.ListItems.Clear
    listUpdates.Enabled = False
    pbarUpdates.Max = frmMain.ModCount + 1
    pbarUpdates.Value = 0
    cmdDownload.Enabled = False
    cmdChangeLog.Enabled = False
    cmdCancel.Caption = "Cancel Update Check"
    frmUpdates.Refresh
    If FreeDiskSpace(GetDriveFromPath(App.Path)) < frmMain.SafetySpace Then
        Call frmMain.WriteLogEntry("Insufficient free space to download update check files. Update check aborted.")
        Call MsgBox("Not enough free space to download update check files. Update check aborted.", vbOKOnly + vbInformation, App.Title)
        Call cmdCancel_Click
    Else
        If Not Connected Then
            lblCheckProgress.Caption = "Establishing Internet connection for update check..."
            Call frmMain.WriteLogEntry("Establishing Internet connection for update check.")
            frmUpdates.Enabled = False
            Call Connect(frmUpdates.hwnd)
            SelfConnect = True
            frmUpdates.Enabled = True
            frmUpdates.SetFocus
        Else
            SelfConnect = False
        End If
        If Not Connected Then
            SelfConnect = False
            Call frmMain.WriteLogEntry("Failed to establish Internet connection. Update check aborted.")
            Call MsgBox("The Check For Updates facility requires an internet connection." & vbCrLf & "Please check your connection settings and try again.", vbOKOnly + vbInformation, App.Title)
            Call cmdCancel_Click
        Else
            Checking = True
            Call frmMain.WriteLogEntry("Checking for updates.")
            CancelDownload = False
            For Counter = 0 To frmMain.ModCount
                Call frmMain.GetModUpdateDetails(Counter, RModPath, RName, RVersion, RType, RTypeStr, RFolder, RUpdateCheckURL)
                If Len(RUpdateCheckURL) <> 0 And Len(RName) <> 0 Then
                    lblCheckProgress.Caption = "Checking for updates to " & RName & "..."
                    frmUpdates.Refresh
                    LocalURL = JoinPath(RFolder, GetFileName(RUpdateCheckURL))
                    DownloadError = Not CopyURLToFile(RUpdateCheckURL, LocalURL, Me)
                    If DownloadError Or CancelDownload Then
                        If FileExists(LocalURL) Then Call Kill(LocalURL)
                        If CancelDownload Then Exit Sub
                    Else
                        If FileExists(LocalURL) Then
                            Call FlushINI(LocalURL)
                            PreviousVersion = ReadINIStr("Update", "UpdateableVersion", LocalURL)
                            LatestVersion = ReadINIStr("Update", "LatestVersion", LocalURL)
                            ChangeLogURL = ReadINIStr("Update", "ChangeLogURL", LocalURL)
                            UpdateDownloadURL = ReadINIStr("Update", "UpdateDownloadURL", LocalURL)
                            UpdateDownloadSize = Val(ReadINIStr("Update", "UpdateDownloadSize", LocalURL, "0"))
                            DownloadURL = ReadINIStr("Update", "FullDownloadURL", LocalURL)
                            DownloadSize = Val(ReadINIStr("Update", "FullDownloadSize", LocalURL, "0"))
                            Call Kill(LocalURL)
                            If CompareVersions(RVersion, "=", PreviousVersion) And Len(UpdateDownloadURL) <> 0 Then
                                DownloadURL = UpdateDownloadURL
                                DownloadSize = UpdateDownloadSize
                            End If
                            If Len(LatestVersion) <> 0 And (CompareVersions(RVersion, "<", LatestVersion)) Then Call AddGridRow(RModPath, RTypeStr, RName, RVersion, LatestVersion, DownloadSize, DownloadURL, ChangeLogURL)
                        End If
                    End If
                End If
                pbarUpdates.Value = pbarUpdates.Value + 1
                frmUpdates.Refresh
            Next Counter
            Select Case UpdateCount
            Case 0
                lblCheckProgress.Caption = "Update check complete. No updates available."
                Call frmMain.WriteLogEntry("Update check complete. No updates available.")
            Case Else
                lblCheckProgress.Caption = "Update check complete. " & PadNum(UpdateCount) & " updates available."
                Call frmMain.WriteLogEntry("Update check complete. " & PadNum(UpdateCount) & " updates available.")
                listUpdates.ListItems.Item(1).Selected = True
                cmdDownload.Enabled = True
                Call listUpdates_ItemClick(listUpdates.ListItems.Item(1))
            End Select
            cmdCancel.Caption = "Back to Launch Menu"
            Checking = False
        End If
    End If
    listUpdates.Enabled = True
    Exit Sub
LocalErr:
    Call frmMain.GlobalErr("CheckForUpdates", ErrVars())
End Sub

Public Sub DownloadProgress(ByVal TotalBytesRead As Long)
    If DownloadSize <> 0 Then
        If TotalBytesRead > DownloadSize Then
            pbarDownload.Value = 100
        Else
            pbarDownload.Value = (TotalBytesRead / DownloadSize) * 100
        End If
        lblDownloadAmount.Caption = "Download Progress:  " & SizeToString(TotalBytesRead, False) & " / " & SizeToString(DownloadSize)
        lblDownloadPercent.Caption = "(" & (Int((TotalBytesRead / DownloadSize) * 100)) & "%)"
        Me.Refresh
    End If
    DoEvents
End Sub

Private Sub AddGridRow(ByVal RModPath As String, ByVal RTypeStr As String, ByVal RName As String, ByVal RVersion As String, ByVal LatestVersion As String, ByVal DownloadSize As Long, ByVal DownloadURL As String, ByVal ChangeLogURL As String)
    Call listUpdates.ListItems.Add(, , RName)
    listUpdates.ListItems(listUpdates.ListItems.Count).SubItems(1) = RTypeStr
    listUpdates.ListItems(listUpdates.ListItems.Count).SubItems(2) = RVersion
    listUpdates.ListItems(listUpdates.ListItems.Count).SubItems(3) = LatestVersion
    listUpdates.ListItems(listUpdates.ListItems.Count).SubItems(4) = SizeToString(DownloadSize)
    UpdateCount = UpdateCount + 1
    ReDim Preserve UpdateDownloadURL(UpdateCount)
    ReDim Preserve UpdateModPath(UpdateCount)
    ReDim Preserve UpdateDownloadSize(UpdateCount)
    ReDim Preserve UpdateChangelogURL(UpdateCount)
    UpdateDownloadURL(UpdateCount) = DownloadURL
    UpdateModPath(UpdateCount) = RModPath
    UpdateDownloadSize(UpdateCount) = DownloadSize
    UpdateChangelogURL(UpdateCount) = ChangeLogURL
End Sub

Private Sub cmdCancel_Click()
    pbarDownload.Visible = False
    cmdDownload.Visible = True
    cmdChangeLog.Visible = True
    lblDownloadAmount.Visible = False
    lblDownloadPercent.Visible = False
    If Checking Then
        Checking = False
        CancelDownload = True
        Call frmMain.WriteLogEntry("Update check cancelled by user.")
        Unload Me
    Else
        If Downloading Then
            Downloading = False
            CancelDownload = True
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub cmdDownload_Click()
    Dim mbResult As VbMsgBoxResult
    Dim LocalURL As String
    Dim DownloadURL As String
    Dim DownloadError As Boolean
    Dim ErrVars(0) As Variant
    If Not frmMain.CL_noexcept Then On Error GoTo LocalErr
    DownloadURL = UpdateDownloadURL(listUpdates.SelectedItem.Index)
    LocalURL = JoinPath(UpdateModPath(listUpdates.SelectedItem.Index), GetFileName(DownloadURL))
    If FreeDiskSpace(GetDriveFromPath(LocalURL)) < (frmMain.SafetySpace + UpdateDownloadSize(listUpdates.SelectedItem.Index)) Then
        Call frmMain.WriteLogEntry("Insufficient free space to download " & DownloadURL & " to " & LocalURL)
        Call MsgBox("Insufficient free space to download " & DownloadURL & " to " & LocalURL, vbOKOnly + vbInformation, App.Title)
    Else
        Downloading = True
        cmdDownload.Visible = False
        cmdChangeLog.Visible = False
        pbarDownload.Value = 0
        pbarDownload.Visible = True
        frmUpdates.Refresh
        listUpdates.Enabled = False
        Call frmMain.WriteLogEntry("Downloading " & Quote(DownloadURL) & " to " & Quote(LocalURL))
        cmdCancel.Caption = "Cancel Download"
        lblDownloadAmount.Visible = True
        lblDownloadPercent.Visible = True
        CancelDownload = False
        DownloadSize = UpdateDownloadSize(listUpdates.SelectedItem.Index)
        Call DownloadProgress(0)
        DownloadError = Not CopyURLToFile(DownloadURL, LocalURL, Me)
        Downloading = False
        If CancelDownload Or DownloadError Then
            If FileExists(LocalURL) Then Call Kill(LocalURL)
            pbarDownload.Value = 0
            pbarDownload.Visible = False
            cmdDownload.Visible = True
            cmdChangeLog.Visible = True
            lblDownloadAmount.Visible = False
            lblDownloadPercent.Visible = False
            cmdCancel.Caption = "Back to Launch Menu"
            listUpdates.Enabled = True
            If CancelDownload Then
                Call frmMain.WriteLogEntry("Download cancelled by user.")
            Else
                Call frmMain.WriteLogEntry("Download error.")
                Call MsgBox("An error occurred downloading the setup program for this mod." & vbCrLf & "The remote file may not exist or your firewall may be blocking the download.", vbOKOnly, App.Title)
            End If
        Else
            If FileExists(LocalURL) Then
                If GetFileSize(LocalURL) < UpdateDownloadSize(listUpdates.SelectedItem.Index) Then
                    Call frmMain.WriteLogEntry("Error downloading " & Quote(DownloadURL) & ". " & Quote(LocalURL) & " has not been saved!")
                    Call Kill(LocalURL)
                    Call MsgBox("Error downloading " & Quote(DownloadURL) & vbCrLf & Quote(LocalURL) & " has not been saved!", vbOKOnly + vbExclamation, App.Title)
                    Call cmdCancel_Click
                Else
                    Call frmMain.WriteLogEntry("Download complete. Executing " & Quote(LocalURL))
                    mbResult = MsgBox("Launch Base will now close so that " & listUpdates.SelectedItem.Text & " can be updated.", vbOKCancel + vbInformation, App.Title)
                    If mbResult = vbOK Then
                        Call frmMain.Shutdown(True)
                        If SelfConnect Then
                            If MsgBox("Disconnect Internet connection now?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                                Call frmMain.WriteLogEntry("Disconnecting Internet connection.")
                                Call Disconnect
                            End If
                        End If
                        ShuttingDown = True
                        Call Shell(Quote(LocalURL), vbNormal)
                        Unload Me
                        Exit Sub
                    Else
                        Call cmdCancel_Click
                    End If
                End If
            Else
                Call frmMain.WriteLogEntry("Error downloading " & Quote(DownloadURL) & ". " & Quote(LocalURL) & " has not been saved!")
                Call MsgBox("Error downloading " & Quote(DownloadURL) & vbCrLf & Quote(LocalURL) & " has not been saved!", vbOKOnly + vbExclamation, App.Title)
                Call cmdCancel_Click
            End If
        End If
    End If
    Exit Sub
LocalErr:
    Call frmMain.GlobalErr("cmdDownload_Click", ErrVars())
End Sub

Private Sub Form_Load()
    UpdateCount = 0
    Call listUpdates.ColumnHeaders.Add(, , "Name", 3930)
    Call listUpdates.ColumnHeaders.Add(, , "Type", 885)
    Call listUpdates.ColumnHeaders.Add(, , "Your Version", 1200)
    Call listUpdates.ColumnHeaders.Add(, , "Latest Version", 1200)
    Call listUpdates.ColumnHeaders.Add(, , "Download Size", 1260)
    Downloading = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Checking Or Downloading Then
        Cancel = 1
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not ShuttingDown Then
        If SelfConnect Then
            If MsgBox("Disconnect Internet connection now?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                Call frmMain.WriteLogEntry("Disconnecting Internet connection.")
                Call Disconnect
            End If
        End If
        frmMain.Show
        frmMain.SetFocus
    End If
End Sub

Private Sub listUpdates_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdChangeLog.Enabled = (Len(UpdateChangelogURL(listUpdates.SelectedItem.Index)) <> 0)
End Sub
