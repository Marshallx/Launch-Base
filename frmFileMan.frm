VERSION 5.00
Begin VB.Form frmFileMan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Residual File Manager"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAllUnsafe 
      Caption         =   "<<< All Unsafe <<<"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   4320
      Width           =   1695
   End
   Begin VB.ListBox listSafe 
      Height          =   2400
      Left            =   3600
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CommandButton cmdSafe 
      Caption         =   ">>> Safe >>>"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton cmdUnsafe 
      Caption         =   "<<< Unsafe <<<"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   4320
      Width           =   1695
   End
   Begin VB.ListBox listResidual 
      Height          =   2400
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   $"frmFileMan.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   6855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Action selected file:"
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   """Safe"" files:"
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Residual files:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Action selected file:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   3375
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      Caption         =   $"frmFileMan.frx":0095
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmFileMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const LVM_FIRST = &H1000

Private Sub cmdDelete_Click()
    If FileExists(JoinPath(RA2DIR, listResidual.List(listResidual.ListIndex))) Then
        Call MoveToRecycleBin(JoinPath(RA2DIR, listResidual.List(listResidual.ListIndex)), True, True, True)
        If Not FileExists(JoinPath(RA2DIR, listResidual.List(listResidual.ListIndex))) Then
            Call listResidual.RemoveItem(listResidual.ListIndex)
        End If
    ElseIf FileExists(JoinPath(BACKUPDIR, listResidual.List(listResidual.ListIndex))) Then
        Call MoveToRecycleBin(JoinPath(BACKUPDIR, listResidual.List(listResidual.ListIndex)), True, True, True)
        If Not FileExists(JoinPath(BACKUPDIR, listResidual.List(listResidual.ListIndex))) Then
            Call listResidual.RemoveItem(listResidual.ListIndex)
        End If
    End If
    cmdDelete.Enabled = False
    cmdSafe.Enabled = False
End Sub

Private Sub cmdSafe_Click()
    Call frmMain.SafeFiles_Add(listResidual.List(listResidual.ListIndex))
    Call listSafe.AddItem(listResidual.List(listResidual.ListIndex))
    Call listResidual.RemoveItem(listResidual.ListIndex)
    cmdDelete.Enabled = False
    cmdSafe.Enabled = False
    cmdAllUnsafe.Enabled = True
End Sub

Private Sub cmdUnsafe_Click()
    Call frmMain.SafeFiles_Find(listSafe.List(listSafe.ListIndex), True)
    Call listResidual.AddItem(listSafe.List(listSafe.ListIndex))
    Call listSafe.RemoveItem(listSafe.ListIndex)
    cmdUnsafe.Enabled = False
    If listSafe.ListCount = 0 Then cmdAllUnsafe.Enabled = False
End Sub

Private Sub cmdallUnsafe_Click()
    Do While listSafe.ListCount <> 0
        listSafe.ListIndex = listSafe.ListCount - 1
        Call frmMain.SafeFiles_Find(listSafe.List(listSafe.ListIndex), True)
        Call listResidual.AddItem(listSafe.List(listSafe.ListIndex))
        Call listSafe.RemoveItem(listSafe.ListIndex)
    Loop
    cmdSafe.Enabled = False
    cmdAllUnsafe.Enabled = False
End Sub

Private Sub Form_Load()
    lblIntro.Caption = "Launch Base has identified the following third-party ""residual"" files in your Red Alert 2 directory." & vbCrLf & "Launch Base will temporarily remove these files when launching a mod." & vbCrLf & "You can mark any of these files as safe and thus they will be ignored by Launch Base." & vbCrLf & "Make sure you have read the Help Topics before marking any files as safe."
    Call LoadFiles
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim iFile As Long
    iFile = UBound(SafeFiles())
    Do While iFile <> 0
        If FileExists(JoinPath(BACKUPDIR, SafeFiles(iFile))) Then
            If Not FileExists(JoinPath(RA2DIR, SafeFiles(iFile))) Then
                Call frmMain.WriteLogEntry(Quote(SafeFiles(iFile)) & " marked as safe. Restoring file to Red Alert 2 directory...", LogLevel1)
                Call frmMain.LoggedMove(JoinPath(BACKUPDIR, SafeFiles(iFile)), JoinPath(RA2DIR, SafeFiles(iFile)))
            Else
                Call frmMain.WriteLogEntry("Warning: " & Quote(SafeFiles(iFile)) & " marked as safe but file exists in both " & Quote(BACKUPDIR) & " and " & Quote(RA2DIR) & ".", LogMsgBox)
            End If
        End If
        iFile = iFile - 1
    Loop
    frmMain.Enabled = True
End Sub

Public Sub LoadFiles()
    Dim fso As FileSystemObject
    Dim fso_folder As Folder
    Dim fso_file As File
    Call frmMain.SafeFiles_Refresh 'refresh in case user has deleted anything since starting Launch Base.
    Set fso = New FileSystemObject
    Set fso_folder = fso.GetFolder(RA2DIR)
    For Each fso_file In fso_folder.Files
        If frmMain.FileIsDirty(fso_file.Name) Then
            If frmMain.SafeFiles_Find(fso_file.Name) <> 0 Then
                Call listSafe.AddItem(fso_file.Name)
            Else
                Select Case UCase$(fso_file.Name)
                Case "EXCEPT.TXT", "SYRINGE.LOG", "DEBUG.TXT"
                    'these files are manipulated by LB so must never be marked as safe
                Case Else
                    Call listResidual.AddItem(fso_file.Name)
                End Select
            End If
        End If
    Next
    Set fso_folder = fso.GetFolder(BACKUPDIR)
    For Each fso_file In fso_folder.Files
        If Not FileExists(JoinPath(RA2DIR, fso_file.Name)) Then
            If frmMain.FileIsDirty(fso_file.Name) Then
                Call listResidual.AddItem(fso_file.Name)
            End If
        End If
    Next
    If listSafe.ListCount <> 0 Then
        listSafe.ListIndex = 0
    Else
        cmdUnsafe.Enabled = False
        cmdAllUnsafe.Enabled = False
    End If
    If listResidual.ListCount <> 0 Then
        listResidual.ListIndex = 0
    Else
        cmdDelete.Enabled = False
        cmdSafe.Enabled = False
    End If
End Sub

Private Sub listResidual_Click()
    If listResidual.ListIndex <> 0 Then
        cmdDelete.Enabled = True
        cmdSafe.Enabled = True
    Else
        cmdDelete.Enabled = True
        cmdSafe.Enabled = True
    End If
End Sub

Private Sub listSafe_Click()
    If listSafe.ListIndex <> 0 Then
        cmdUnsafe.Enabled = True
    Else
        cmdUnsafe.Enabled = True
    End If
End Sub
