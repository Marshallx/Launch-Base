VERSION 5.00
Begin VB.Form frmPlayVideo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base: Play Video"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmPlayVideo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ListBox lstVideos 
      Height          =   2205
      ItemData        =   "frmPlayVideo.frx":0E42
      Left            =   120
      List            =   "frmPlayVideo.frx":0E44
      TabIndex        =   0
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmPlayVideo.frx":0E46
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   5415
   End
   Begin VB.Label lblVideos 
      Alignment       =   2  'Center
      Caption         =   $"frmPlayVideo.frx":0EDC
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmPlayVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CancelLaunch As Boolean
Public VideoVersion As String

Public Sub RefreshList(ByVal VideoPath As String, ByVal MVersion As String)
    Dim Counter As Integer
    Dim RPath As String
    Dim RDummy As String
    Dim fso As FileSystemObject
    Dim fso_folder As Folder
    Dim fso_file As File
    VideoVersion = MVersion
    Set fso = New FileSystemObject
    Set fso_folder = fso.GetFolder(VideoPath)
    lstVideos.Tag = VideoPath
    Call lstVideos.Clear
    Call lstVideos.AddItem("Do not play a video, just launch the game.")
    For Each fso_file In fso_folder.Files
        If FileType(fso_file.Name) = "IPB" Then
            If UCase$(fso_file.Name) <> "SESSION.IPB" Or VideoPath <> RA2DIR Then
                Call lstVideos.AddItem(fso_file.Name)
            End If
        End If
    Next
    lstVideos.ListIndex = 0
    CancelLaunch = False
    If lstVideos.ListCount = 1 Then Call cmdOK_Click
End Sub

Private Sub cmdOK_Click()
    Dim mbResult As VbMsgBoxResult
    If lstVideos.ListIndex = 0 Then
        Me.Hide
    Else
        If InStr(1, lstVideos.List(lstVideos.ListIndex), "[" & VideoVersion & "]") = 0 Then
            mbResult = MsgBox("The video you have selected is not recognised as being for the same mod version that you are about to launch." & vbCrLf & "This video is likely to exhibit unusual behaviour, little or no action, and possibly Internal Errors." & vbCrLf & "It is strongly recommended that you do not play this video." & vbCrLf & vbCrLf & "Are you sure that you want to play this video?", vbYesNo + vbQuestion, "Incorrect Mod Version")
        Else
            mbResult = vbYes
        End If
        If mbResult = vbYes Then
            CL_playfile = lstVideos.List(lstVideos.ListIndex)
            Me.Hide
        Else
            Call lstVideos.SetFocus
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    CancelLaunch = True
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        Cancel = 1
        Call cmdCancel_Click
    End If
End Sub

Private Sub lstVideos_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RemoveVideoFromList(KeyCode)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RemoveVideoFromList(KeyCode)
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RemoveVideoFromList(KeyCode)
End Sub

Private Sub cmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    Call RemoveVideoFromList(KeyCode)
End Sub

Private Sub RemoveVideoFromList(KeyCode As Integer)
    Dim OldIndex As Integer
    If KeyCode = 46 Then
        If lstVideos.ListIndex <> 0 Then
            If MoveToRecycleBin(JoinPath(lstVideos.Tag, lstVideos.List(lstVideos.ListIndex)), True) Then
                If Not FileExists(JoinPath(lstVideos.Tag, lstVideos.List(lstVideos.ListIndex))) Then
                    OldIndex = lstVideos.ListIndex
                    Call lstVideos.RemoveItem(OldIndex)
                    If lstVideos.ListCount - 1 > OldIndex Then
                        lstVideos.ListIndex = OldIndex
                    Else
                        lstVideos.ListIndex = lstVideos.ListCount - 1
                    End If
                End If
            End If
        Else
            Beep
        End If
    End If
End Sub
