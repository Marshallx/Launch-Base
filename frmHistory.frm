VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmHistory 
   Caption         =   "Launch Base: Download History"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4620
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   4620
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbarHistory 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4110
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Selected Installer"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Selected Installer"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3360
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvHistory 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LVM_FIRST = &H1000
Dim ShuttingDown As Boolean

Private Sub cmdDelete_Click()
    Dim sFile As String
    If Not (lvHistory.SelectedItem Is Nothing) Then
        sFile = JoinPath(SETUPDIR, lvHistory.SelectedItem.SubItems(4))
        If MoveToRecycleBin(sFile) Then
            If Not FileExists(sFile) Then
                Call frmMain.WriteLogEntry(Quote(sFile) & " moved to recycle bin.")
                Call lvHistory.ListItems.Remove(lvHistory.SelectedItem.Index)
            End If
        End If
    End If
    Call lvHistory.SetFocus
    If Not (lvHistory.SelectedItem Is Nothing) Then
        Call lvHistory_ItemClick(lvHistory.SelectedItem)
    End If
End Sub

Private Sub cmdRun_Click()
    Dim sFile As String
    If Not (lvHistory.SelectedItem Is Nothing) Then
        sFile = JoinPath(SETUPDIR, lvHistory.SelectedItem.SubItems(4))
        Call frmMain.WriteLogEntry("Preparing to execute " & sFile & "...", LogLevel1)
        Call frmMain.Shutdown
        Call Shell(sFile, vbNormalFocus)
        ShuttingDown = True
        Unload Me
        Exit Sub
    End If
End Sub

Public Sub LoadHistory() 'can't be run on form load else columns aren't added in time.
    Dim sTemp As String
    Dim sTemp2 As String
    Dim itemTemp As ListItem
    Dim iCounter As Integer
    Dim fso As FileSystemObject
    Dim fso_folder As Folder
    Dim fso_file As File
    If Not DirExists(SETUPDIR) Then Call MakePath(SETUPDIR)
    lvHistory.Sorted = False
    Set fso = New FileSystemObject
    Set fso_folder = fso.GetFolder(SETUPDIR)
    For Each fso_file In fso_folder.Files
        If FileType(fso_file.Name) = "EXE" Then
            Set itemTemp = lvHistory.ListItems.Add
            'File
            itemTemp.SubItems(4) = fso_file.Name
            'Name
            sTemp = ReadINIStr(UCase$(fso_file.Name), "Name", SetupsINI)
            If Len(sTemp) = 0 Then sTemp = "?"
            itemTemp.Text = sTemp
            'Version
            sTemp = ReadINIStr(UCase$(fso_file.Name), "NewVersion", SetupsINI)
            If Len(sTemp) <> 0 Then
                sTemp2 = ReadINIStr(UCase$(fso_file.Name), "OldVersion", SetupsINI)
                If Len(sTemp2) <> 0 Then sTemp = sTemp2 & "-->" & sTemp
            Else
                sTemp = "?"
            End If
            itemTemp.SubItems(1) = sTemp
            'Type
            sTemp = ReadINIStr(UCase$(fso_file.Name), "ModType", SetupsINI)
            If Len(sTemp) <> 0 Then
                Select Case Val(sTemp)
                Case TypeMod: sTemp = "Mod"
                Case TypeFA2Mod: sTemp = "FA2 Mod"
                Case TypePlugin: sTemp = "Plugin"
                Case TypeProgram: sTemp = "Tool"
                End Select
            Else
                sTemp = "?"
            End If
            itemTemp.SubItems(2) = sTemp
            'Date
            sTemp = ReadINIStr(UCase$(fso_file.Name), "Date", SetupsINI)
            If Len(sTemp) = 0 Then sTemp = "?"
            itemTemp.SubItems(3) = sTemp
            Set itemTemp = Nothing
        End If
    Next
    'resize the listview
    lvHistory.Width = 0
    iCounter = 1
    Do While iCounter <= lvHistory.ColumnHeaders.Count
        SendMessage lvHistory.hWnd, LVM_FIRST + 30, lvHistory.ColumnHeaders.Item(iCounter).Index - 1, -1
        Select Case iCounter
        Case 1: lvHistory.ColumnHeaders.Item(iCounter).Width = Max(lvHistory.ColumnHeaders.Item(iCounter).Width, 640) 'Name
        Case 2: lvHistory.ColumnHeaders.Item(iCounter).Width = Max(lvHistory.ColumnHeaders.Item(iCounter).Width, 768) 'Version
        Case 3: lvHistory.ColumnHeaders.Item(iCounter).Width = Max(lvHistory.ColumnHeaders.Item(iCounter).Width, 640) 'Type
        Case 4: lvHistory.ColumnHeaders.Item(iCounter).Width = Max(lvHistory.ColumnHeaders.Item(iCounter).Width, 1152) 'Downloaded
        Case 5: lvHistory.ColumnHeaders.Item(iCounter).Width = Max(lvHistory.ColumnHeaders.Item(iCounter).Width, 640) 'File
        End Select
        lvHistory.Width = lvHistory.Width + lvHistory.ColumnHeaders.Item(iCounter).Width + 30
        iCounter = iCounter + 1
    Loop
    lvHistory.Sorted = True
    lvHistory.Width = lvHistory.Width + 45
    frmHistory.Width = lvHistory.Width
    Call lvHistory.SetFocus
    If lvHistory.ListItems.Count <> 0 Then
        lvHistory.SelectedItem = lvHistory.ListItems.Item(1)
        Call lvHistory_ItemClick(lvHistory.SelectedItem)
    Else
        cmdRun.Enabled = False
        cmdDelete.Enabled = False
    End If
    'center form on screen, as it will probably have been resized
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_Load()
    ShuttingDown = False
    Call lvHistory.ColumnHeaders.Add(, , "Name")
    Call lvHistory.ColumnHeaders.Add(, , "Version")
    Call lvHistory.ColumnHeaders.Add(, , "Type")
    Call lvHistory.ColumnHeaders.Add(, , "Downloaded")
    Call lvHistory.ColumnHeaders.Add(, , "File")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not ShuttingDown Then frmMain.Enabled = True
End Sub

Private Sub Form_Resize()
    Dim iTemp As Long
    iTemp = (frmHistory.ScaleHeight - cmdDelete.Height) - sbarHistory.Height
    lvHistory.Width = frmHistory.ScaleWidth
    If iTemp > 0 Then lvHistory.Height = iTemp
    cmdDelete.Top = lvHistory.Height
    cmdRun.Top = lvHistory.Height
    cmdRun.Width = (lvHistory.Width / 2)
    cmdDelete.Width = (lvHistory.Width - cmdRun.Width) + 15
    cmdRun.Left = cmdDelete.Width - 15
    sbarHistory.Panels(1).Width = sbarHistory.Width
End Sub

Private Sub lvHistory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvHistory.SortKey = (ColumnHeader.Index - 1) Then
        If lvHistory.SortOrder = lvwAscending Then
            lvHistory.SortOrder = lvwDescending
        Else
            lvHistory.SortOrder = lvwAscending
        End If
    Else
        lvHistory.SortKey = ColumnHeader.Index - 1
    End If
End Sub

Private Sub lvHistory_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim sFilePath As String
    sFilePath = JoinPath(SETUPDIR, Item.SubItems(4))
    sbarHistory.Panels(1).Text = "Date Modified: " & FileDateTime(sFilePath) & ", Size: " & DataSize(GetFileSize(sFilePath))
End Sub
