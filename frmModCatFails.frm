VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmModCatFails 
   Caption         =   "Launch Base: Failed Update Checks"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5385
   Icon            =   "frmModCatFails.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvModCatFails 
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
Attribute VB_Name = "frmModCatFails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LVM_FIRST = &H1000

Public Sub LoadFails() 'can't be run on form load else columns aren't added in time.
    Dim itemTemp As ListItem
    Dim iRecord As Integer
    lvModCatFails.Sorted = False
    For iRecord = 1 To UpdateRecordCount
        If Len(UpdateRecords(iRecord).FailReason) <> 0 Then
            Set itemTemp = lvModCatFails.ListItems.Add
            If Len(UpdateRecords(iRecord).ModName) <> 0 Then
                itemTemp.Text = UpdateRecords(iRecord).ModName
            Else
                itemTemp.Text = "<undetermined>"
            End If
            itemTemp.SubItems(1) = UpdateRecords(iRecord).ModUserVersion
            Select Case UpdateRecords(iRecord).ModType
            Case TypeMod: itemTemp.SubItems(2) = "Mod"
            Case TypePlugin: itemTemp.SubItems(2) = "Plugin"
            Case TypeFA2Mod: itemTemp.SubItems(2) = "FA2 Mod"
            Case TypeProgram: itemTemp.SubItems(2) = "Tool"
            End Select
            itemTemp.SubItems(3) = UpdateRecords(iRecord).FailReason
            itemTemp.SubItems(4) = UpdateRecords(iRecord).CheckURL
            Set itemTemp = Nothing
        End If
    Next iRecord
    'resize the listview
    lvModCatFails.Width = 0
    iRecord = 1
    Do While iRecord <= lvModCatFails.ColumnHeaders.Count
        SendMessage lvModCatFails.hWnd, LVM_FIRST + 30, lvModCatFails.ColumnHeaders.Item(iRecord).Index - 1, -1
        Select Case iRecord
        Case 1: lvModCatFails.ColumnHeaders.Item(iRecord).Width = Max(lvModCatFails.ColumnHeaders.Item(iRecord).Width, 640) 'Name
        Case 2: lvModCatFails.ColumnHeaders.Item(iRecord).Width = Max(lvModCatFails.ColumnHeaders.Item(iRecord).Width, 768) 'Version
        Case 3: lvModCatFails.ColumnHeaders.Item(iRecord).Width = Max(lvModCatFails.ColumnHeaders.Item(iRecord).Width, 640) 'Type
        Case 4: lvModCatFails.ColumnHeaders.Item(iRecord).Width = Max(lvModCatFails.ColumnHeaders.Item(iRecord).Width, 1152) 'Fail reason
        Case 5: lvModCatFails.ColumnHeaders.Item(iRecord).Width = Max(lvModCatFails.ColumnHeaders.Item(iRecord).Width, 1152) 'Check URL
        End Select
        lvModCatFails.Width = lvModCatFails.Width + lvModCatFails.ColumnHeaders.Item(iRecord).Width + 30
        iRecord = iRecord + 1
    Loop
    lvModCatFails.Sorted = True
    lvModCatFails.Width = lvModCatFails.Width + 45
    frmModCatFails.Width = lvModCatFails.Width
    Call lvModCatFails.SetFocus
    'center form on screen, as it will probably have been resized
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub Form_Load()
    Call lvModCatFails.ColumnHeaders.Add(, , "Name")
    Call lvModCatFails.ColumnHeaders.Add(, , "Version")
    Call lvModCatFails.ColumnHeaders.Add(, , "Type")
    Call lvModCatFails.ColumnHeaders.Add(, , "Failure Reason")
    Call lvModCatFails.ColumnHeaders.Add(, , "Update Check URL")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmModCat.Enabled = True
End Sub

Private Sub Form_Resize()
    Dim iTemp As Long
    iTemp = frmModCatFails.ScaleHeight
    lvModCatFails.Width = frmModCatFails.ScaleWidth
    If iTemp > 0 Then lvModCatFails.Height = iTemp
End Sub

Private Sub lvModCatFails_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lvModCatFails.SortKey = (ColumnHeader.Index - 1) Then
        If lvModCatFails.SortOrder = lvwAscending Then
            lvModCatFails.SortOrder = lvwDescending
        Else
            lvModCatFails.SortOrder = lvwAscending
        End If
    Else
        lvModCatFails.SortKey = ColumnHeader.Index - 1
    End If
End Sub
