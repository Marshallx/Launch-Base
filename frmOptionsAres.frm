VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptionsAres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ares Update Options"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8280
   Icon            =   "frmOptionsAres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8280
   Begin VB.CommandButton cmdUpdateNow 
      Caption         =   "Update Syringe && Ares Now"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton cmdForce 
      Caption         =   "Force Immediate Ares Update"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton cmdOKOptions 
      Caption         =   "OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelOptions 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox cboxAutoAresUpdate 
      Caption         =   "Automatic Ares Update"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   $"frmOptionsAres.frx":0E42
      Top             =   120
      Width           =   2025
   End
   Begin MSComctlLib.ListView lvAresBranches 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblAresRelease 
      Caption         =   "Select the Ares release to download. Hover over a release's ID to display its description."
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7995
   End
End
Attribute VB_Name = "frmOptionsAres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LVM_FIRST = &H1000

Public Sub UpdateAresBranches()
    Dim sBuffer As String
    Dim sBranches() As String
    Dim sBranch As String
    Dim iBranch As Long
    Dim itemTemp As ListItem
    Dim iSelected As Long
    Dim iBestBranch As Long
    Dim sBestVersion As String
    Dim iListElement As Long
    Dim bPublic As Boolean
    Dim sAgent As String
    Dim sNewRevDataURL As String
    Call CallStackPush(Me.Name & ".UpdateAresBranches()")
    If Not CL_noexcept Then On Error GoTo LocalErr
    GoTo LocalMain
LocalErr:
    Call GlobalErr
LocalMain:
    ReDim Preserve sBranches(0)
    'Renegade wishes to know which header we are requesting so temporarily change the UserAgent
    sAgent = theInternet.UserAgent
    theInternet.UserAgent = sAgent & " (" & OptAresRevisionDataURLHDR & ")"
    sNewRevDataURL = theInternet.GetHeader(OptAresRevisionDataURLHDR, OptAresRevisionDataURLURL)
    If Len(sNewRevDataURL) <> 0 Then
        If sNewRevDataURL <> OptAresRevisionDataURL Then
            Call frmMain.WriteLogEntry(OptAresRevisionDataURLHDR & " header check reveals that AresRevisionDataURL has changed.", LogLevel0)
            OptAresRevisionDataURL = sNewRevDataURL
            Call WriteINIStr("URL", "AresRevisionDataURL", OptAresRevisionDataURL, ProgramINI)
        End If
    End If
    If theInternet.CopyURLToString(OptAresRevisionDataURL, sBuffer) Then
        sBuffer = ConvertEOL(sBuffer)
        Call lvAresBranches.ListItems.Clear
        iBranch = 0
        iListElement = 0
        Do
            sBranch = ReadINIStrMemory(sBuffer, "branches", CStr(iBranch), "")
            If Len(sBranch) <> 0 Then
                iListElement = iListElement + 1
                ReDim Preserve sBranches(iListElement)
                sBranches(iListElement) = sBranch
            Else
                If iBranch <> 0 Then Exit Do
            End If
            iBranch = iBranch + 1
        Loop
        If UBound(sBranches()) = 0 Then
            Call frmMain.WriteLogEntry("Ares revision data does not list any branches.", LogLevel0)
            Set itemTemp = lvAresBranches.ListItems.Add
            itemTemp.Text = "Ares revision data does not list any branches."
            Set itemTemp = Nothing
            SendMessage lvAresBranches.hWnd, LVM_FIRST + 30, lvAresBranches.ColumnHeaders.Item(1).Index - 1, -1
            cmdForce.Enabled = False
            cmdUpdateNow.Enabled = False
            cmdOKOptions.Enabled = False
        Else
            'load branch data onto grid
            iSelected = -1
            iBestBranch = -1
            iListElement = 0
            For iBranch = 1 To UBound(sBranches())
                sBranch = sBranches(iBranch)
                bPublic = BooleanStringToBoolean(ReadINIStrMemory(sBuffer, sBranch, "public", "no"))
                If bPublic Or OptAresTester Then
                    Set itemTemp = lvAresBranches.ListItems.Add
                    iListElement = iListElement + 1
                    itemTemp.Text = sBranch
                    itemTemp.SubItems(1) = ReadINIStrMemory(sBuffer, sBranch, "name", "<unnamed>")
                    itemTemp.SubItems(2) = ReadINIStrMemory(sBuffer, sBranch, "version", "?")
                    itemTemp.SubItems(3) = ReadINIStrMemory(sBuffer, sBranch, "revision", "?")
                    itemTemp.SubItems(4) = ReadINIStrMemory(sBuffer, sBranch, "stability", "?")
                    If OptAresTester Then itemTemp.SubItems(5) = ReadINIStrMemory(sBuffer, sBranch, "public", "?")
                    itemTemp.Tag = ReadINIStrMemory(sBuffer, sBranch, "location", "")
                    itemTemp.ToolTipText = ReadINIStrMemory(sBuffer, sBranch, "description", "No description.")
                    If OptAresBranch = sBranch Then
                        iSelected = iListElement
                    ElseIf LCase$(itemTemp.SubItems(4)) = "stable" Then
                        If CompareVersions(itemTemp.SubItems(2), ">", sBestVersion) Then
                            If bPublic Then
                                iBestBranch = iListElement
                            End If
                        End If
                    End If
                    Set itemTemp = Nothing
                End If
            Next iBranch
            If iSelected <> -1 Then
                lvAresBranches.ListItems.Item(iSelected).Selected = True
            ElseIf iBestBranch <> -1 Then
                lvAresBranches.ListItems.Item(iBestBranch).Selected = True
            End If
            'resizing
            lvAresBranches.Width = 0
            For iBranch = 1 To lvAresBranches.ColumnHeaders.Count
                SendMessage lvAresBranches.hWnd, LVM_FIRST + 30, lvAresBranches.ColumnHeaders.Item(iBranch).Index - 1, -1
                Select Case iBranch
                Case 1: lvAresBranches.ColumnHeaders.Item(iBranch).Width = Max(lvAresBranches.ColumnHeaders.Item(iBranch).Width, 960) 'ID
                Case 2: lvAresBranches.ColumnHeaders.Item(iBranch).Width = Max(lvAresBranches.ColumnHeaders.Item(iBranch).Width, 960) 'Name
                Case 3: lvAresBranches.ColumnHeaders.Item(iBranch).Width = Max(lvAresBranches.ColumnHeaders.Item(iBranch).Width, 960) 'Version
                Case 4: lvAresBranches.ColumnHeaders.Item(iBranch).Width = Max(lvAresBranches.ColumnHeaders.Item(iBranch).Width, 960) 'Revision
                Case 5: lvAresBranches.ColumnHeaders.Item(iBranch).Width = Max(lvAresBranches.ColumnHeaders.Item(iBranch).Width, 960) 'Stability
                Case 6: If OptAresTester Then lvAresBranches.ColumnHeaders.Item(iBranch).Width = Max(lvAresBranches.ColumnHeaders.Item(iBranch).Width, 960) 'Public
                End Select
                lvAresBranches.Width = lvAresBranches.Width + lvAresBranches.ColumnHeaders.Item(iBranch).Width + 45
            Next iBranch
            lvAresBranches.Width = lvAresBranches.Width + 15
            Me.Width = Min(Max(Me.Width, lvAresBranches.Width + 360), 12000)
            lvAresBranches.Width = Max(lvAresBranches.Width, Me.Width - 315)
            'center form on screen, as it will probably have been resized
            Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
            cmdOKOptions.Left = Me.ScaleWidth - cmdOKOptions.Width - 105
            cmdCancelOptions.Left = cmdOKOptions.Left - 105 - cmdCancelOptions.Width
            cmdForce.Enabled = True
            cmdUpdateNow.Enabled = True
            cmdOKOptions.Enabled = True
        End If
    Else
        Call frmMain.WriteLogEntry("Failed to download Ares revision data.", LogLevel0)
        Set itemTemp = lvAresBranches.ListItems.Add
        itemTemp.Text = "Failed to download Ares revision data."
        Set itemTemp = Nothing
        SendMessage lvAresBranches.hWnd, LVM_FIRST + 30, lvAresBranches.ColumnHeaders.Item(1).Index - 1, -1
        cmdForce.Enabled = False
        cmdUpdateNow.Enabled = False
        cmdOKOptions.Enabled = False
    End If
    Call CallStackPop
End Sub

Private Sub cmdForce_Click()
    Call frmMain.UpdateAres(Me, lvAresBranches.SelectedItem.Text, False, True)
    Me.SetFocus
End Sub

Private Sub cmdUpdateNow_Click()
    Call frmMain.UpdateAres(Me, "syringe", False)
    Call frmMain.UpdateAres(Me, lvAresBranches.SelectedItem.Text, False)
    Me.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmMain.Enabled = True
End Sub

Private Sub lvAresBranches_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvAresBranches.SortKey = (ColumnHeader.Index - 1) Then
        If lvAresBranches.SortOrder = lvwAscending Then
            lvAresBranches.SortOrder = lvwDescending
        Else
            lvAresBranches.SortOrder = lvwAscending
        End If
    Else
        lvAresBranches.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub cmdOKOptions_Click()
    If OptAutoAresUpdate <> IntegerToBoolean(cboxAutoAresUpdate.Value) Then
        Call WriteINIStr("Options", "AresPrompt", IntegerToYesNo(cboxAutoAresUpdate.Value), ProgramINI)
        Select Case cboxAutoAresUpdate.Value
        Case 0
            OptAutoAresUpdate = False
            Call frmMain.WriteLogEntry("Ares Update Options: 'Automatic Ares Update' disabled by " & IIf(Not OptAdvancedMode, "Launch Base", "user") & ".", LogLevel1)
        Case 1
            OptAutoAresUpdate = True
            Call frmMain.WriteLogEntry("Ares Update Options: 'Automatic Ares Update' enabled by user.", LogLevel1)
        End Select
    End If
    If OptAresBranch <> lvAresBranches.SelectedItem.Text Then
        OptAresBranch = lvAresBranches.SelectedItem.Text
        Call WriteINIStr("Options", "AresBranch", OptAresBranch, ProgramINI)
        Call frmMain.WriteLogEntry("Ares Update Options: 'Ares Release' changed to " & OptAresBranch & " by user.", LogLevel1)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Call frmMain.WriteLogEntry("Form_Load: frmOptionsAres", LogLevel2)
    cboxAutoAresUpdate.Value = BooleanToInteger(OptAutoAresUpdate)
    Call lvAresBranches.ColumnHeaders.Add(, , "ID")
    Call lvAresBranches.ColumnHeaders.Add(, , "Name")
    Call lvAresBranches.ColumnHeaders.Add(, , "Version")
    Call lvAresBranches.ColumnHeaders.Add(, , "Revision")
    Call lvAresBranches.ColumnHeaders.Add(, , "Stability")
    If OptAresTester Then Call lvAresBranches.ColumnHeaders.Add(, , "Public")
    Call UpdateAresBranches
    cmdForce.Visible = OptAresTester
End Sub

Private Sub cmdCancelOptions_Click()
    Unload Me
End Sub

Public Sub ShowPleaseWait(ByVal Message As String, Optional ByVal Message2 As String = "")
    Me.Enabled = False
    Call frmPleaseWait.Show
    Call UpdatePleaseWait(Message, Message2)
End Sub

Public Sub UpdatePleaseWait(Optional ByVal Message As String, Optional ByVal Message2 As String)
    If Message <> "" Then frmPleaseWait.Label1.Caption = Message
    frmPleaseWait.Label2.Caption = Message2
    Call frmPleaseWait.Refresh
End Sub

Public Sub HidePleaseWait()
    Unload frmPleaseWait
    Me.Enabled = True
End Sub
