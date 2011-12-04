VERSION 5.00
Begin VB.Form frmPlaylist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base: Choose Playlist"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   Icon            =   "frmPlaylist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ListBox lstPlaylists 
      Height          =   2205
      ItemData        =   "frmPlaylist.frx":0E42
      Left            =   120
      List            =   "frmPlaylist.frx":0E49
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
   Begin VB.Label lblVideos 
      Alignment       =   2  'Center
      Caption         =   "Select the playlist you would like to hear in-game."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "frmPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CancelLaunch As Boolean
Public Playlist As String

Public Sub RefreshList(ByVal PlaylistPath As String)
    Dim iCounter As Integer
    Dim sPath As String
    Dim fso As FileSystemObject
    Dim fso_folder As Folder
    Dim fso_file As File
    Set fso = New FileSystemObject
    Set fso_folder = fso.GetFolder(PlaylistPath)
    lstPlaylists.Tag = PlaylistPath
    Call lstPlaylists.Clear
    Call lstPlaylists.AddItem("Use the original Yuri's Revenge playlist or the mod's own playlist (if any).")
    sPath = JoinPath(RA2DIR, "thememd.ini")
    If FileExists(sPath) Then
        If Len(ReadINIStr("YRPMOTS", "GameMusic", sPath)) <> 0 Then
            Call lstPlaylists.AddItem("Use the active YR Playlist Manager playlist.")
        End If
    End If
    If PlaylistPath <> RA2DIR Then
        For Each fso_file In fso_folder.Files
            If FileType(fso_file.Name) = "YPL" Then
                Call lstPlaylists.AddItem(fso_file.Name)
            End If
        Next
    End If
    If lstPlaylists.ListCount >= 2 Then
        lstPlaylists.ListIndex = 1
    Else
        lstPlaylists.ListIndex = 0
    End If
    CancelLaunch = False
    If lstPlaylists.ListCount = 1 Then Call cmdOK_Click
End Sub

Private Sub cmdOK_Click()
    If lstPlaylists.ListIndex = 0 Then
        Playlist = ""
    Else
        Playlist = lstPlaylists.List(lstPlaylists.ListIndex)
    End If
    Me.Hide
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


