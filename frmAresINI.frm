VERSION 5.00
Begin VB.Form frmAresINI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base: Ares.ini"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   Icon            =   "frmAresINI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1200
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   41
      Top             =   2520
      Width           =   2055
      Begin VB.OptionButton optSMTile 
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   44
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSMTile 
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   43
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSMTile 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1200
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   37
      Top             =   2280
      Width           =   2055
      Begin VB.OptionButton optSMSidebar 
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   40
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSMSidebar 
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   39
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSMSidebar 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1200
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   33
      Top             =   2040
      Width           =   2055
      Begin VB.OptionButton optSMHidden 
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   36
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSMHidden 
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   35
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSMHidden 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1200
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   29
      Top             =   1800
      Width           =   2055
      Begin VB.OptionButton optSMAlternate 
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   32
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSMAlternate 
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   31
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSMAlternate 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1200
      ScaleHeight     =   255
      ScaleWidth      =   2055
      TabIndex        =   25
      Top             =   1560
      Width           =   2055
      Begin VB.OptionButton optSMComposite 
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   28
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSMComposite 
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   27
         Top             =   0
         Width           =   255
      End
      Begin VB.OptionButton optSMComposite 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdOKOptions 
      Caption         =   "OK"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelOptions 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   4440
      Width           =   1335
   End
   Begin VB.ComboBox comboDirectXForce 
      Height          =   315
      ItemData        =   "frmAresINI.frx":0E42
      Left            =   1200
      List            =   "frmAresINI.frx":0E4F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   2055
   End
   Begin VB.CheckBox cboxF3DComposite 
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CheckBox cboxF3DAlternate 
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CheckBox cboxF3DHidden 
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CheckBox cboxF3DSidebar 
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CheckBox cboxF3DTile 
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      Caption         =   "See the Ares documentation for information about these settings."
      Height          =   495
      Left            =   120
      TabIndex        =   45
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label24 
      Caption         =   "Tile"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "DirectX.Force="
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Surface.$surface.Memory:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label8 
      Caption         =   "Composite"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Alternate"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Hidden"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Sidebar"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "System"
      Height          =   255
      Left            =   1920
      TabIndex        =   15
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label14 
      Caption         =   "VRAM"
      Height          =   255
      Left            =   2640
      TabIndex        =   14
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "Tile"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "$surface:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label17 
      Caption         =   "Default"
      Height          =   255
      Left            =   1200
      TabIndex        =   11
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label18 
      Caption         =   "Surface.$surface.Force3D:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label Label19 
      Caption         =   "Composite"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label20 
      Caption         =   "Alternate"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "Hidden"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label23 
      Caption         =   "Sidebar"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   855
   End
End
Attribute VB_Name = "frmAresINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OriDirectXForce As Integer
Dim OriSM(5) As Integer
Dim OriF3D(5) As Integer

Private Sub cmdCancelOptions_Click()
    Unload Me
End Sub

Private Sub cmdOKOptions_Click()
    Dim iSurface As Integer
    Dim optMemory(2) As OptionButton
    Dim cboxForce3D As CheckBox
    Dim sSurface As String
    Dim sAresIni As String
    sAresIni = JoinPath(RA2DIR, "ares.ini")
    If OriDirectXForce <> comboDirectXForce.ListIndex Then
        Select Case comboDirectXForce.ListIndex
        Case 0
            Call WriteINIStr("Graphics.Advanced", "DirectX.Force", " ;default", sAresIni)
            Call frmMain.WriteLogEntry("Ares.ini, Graphics.Advanced: 'DirectX.Force= ;default' set by user.")
        Case 1
            Call WriteINIStr("Graphics.Advanced", "DirectX.Force", "hardware", sAresIni)
            Call frmMain.WriteLogEntry("Ares.ini, Graphics.Advanced: 'DirectX.Force=hardware' set by user.")
        Case 2
            Call WriteINIStr("Graphics.Advanced", "DirectX.Force", "emulation", sAresIni)
            Call frmMain.WriteLogEntry("Ares.ini, Graphics.Advanced: 'DirectX.Force=emulation' set by user.")
        End Select
        OriDirectXForce = comboDirectXForce.ListIndex
    End If
    For iSurface = 1 To 5
        Select Case iSurface
        Case 1
            sSurface = "Composite"
            Set optMemory(0) = optSMComposite(0)
            Set optMemory(1) = optSMComposite(1)
            Set optMemory(2) = optSMComposite(2)
            Set cboxForce3D = cboxF3DComposite
        Case 2
            sSurface = "Alternate"
            Set optMemory(0) = optSMAlternate(0)
            Set optMemory(1) = optSMAlternate(1)
            Set optMemory(2) = optSMAlternate(2)
            Set cboxForce3D = cboxF3DAlternate
        Case 3
            sSurface = "Hidden"
            Set optMemory(0) = optSMHidden(0)
            Set optMemory(1) = optSMHidden(1)
            Set optMemory(2) = optSMHidden(2)
            Set cboxForce3D = cboxF3DHidden
        Case 4
            sSurface = "Sidebar"
            Set optMemory(0) = optSMSidebar(0)
            Set optMemory(1) = optSMSidebar(1)
            Set optMemory(2) = optSMSidebar(2)
            Set cboxForce3D = cboxF3DSidebar
        Case 5
            sSurface = "Tile"
            Set optMemory(0) = optSMTile(0)
            Set optMemory(1) = optSMTile(1)
            Set optMemory(2) = optSMTile(2)
            Set cboxForce3D = cboxF3DTile
        End Select
        If Not optMemory(OriSM(iSurface)).Value Then
            If optMemory(0).Value Then
                Call WriteINIStr("Graphics.Advanced", "Surface." & sSurface & ".Memory", " ;default", sAresIni)
                Call frmMain.WriteLogEntry("Ares.ini, Graphics.Advanced: 'Surface." & sSurface & ".Memory= ;default' set by user.")
                OriSM(iSurface) = 0
            ElseIf optMemory(1).Value Then
                Call WriteINIStr("Graphics.Advanced", "Surface." & sSurface & ".Memory", "System", sAresIni)
                Call frmMain.WriteLogEntry("Ares.ini, Graphics.Advanced: 'Surface." & sSurface & ".Memory=System' set by user.")
                OriSM(iSurface) = 1
            ElseIf optMemory(2).Value Then
                Call WriteINIStr("Graphics.Advanced", "Surface." & sSurface & ".Memory", "VRAM", sAresIni)
                Call frmMain.WriteLogEntry("Ares.ini, Graphics.Advanced: 'Surface." & sSurface & ".Memory=VRAM' set by user.")
                OriSM(iSurface) = 2
            End If
        End If
        If cboxForce3D.Value <> OriF3D(iSurface) Then
            If cboxForce3D.Value = 1 Then
                Call WriteINIStr("Graphics.Advanced", "Surface." & sSurface & ".Force3D", "true", sAresIni)
                Call frmMain.WriteLogEntry("Ares.ini, Graphics.Advanced: 'Surface." & sSurface & ".Force3D=true' set by user.")
                OriF3D(iSurface) = 1
            Else
                Call WriteINIStr("Graphics.Advanced", "Surface." & sSurface & ".Force3D", "false", sAresIni)
                Call frmMain.WriteLogEntry("Ares.ini, Graphics.Advanced: 'Surface." & sSurface & ".Force3D=false' set by user.")
                OriF3D(iSurface) = 0
            End If
        End If
    Next iSurface
    Unload Me
End Sub

Friend Sub RefreshAresIniFlags()
    Dim sAresIni As String
    Dim sSurface As String
    Dim i As Integer
    Dim force As Integer
    sAresIni = JoinPath(RA2DIR, "ares.ini")
    If FileExists(sAresIni) Then
        Select Case LCase$(ReadINIStr("Graphics.Advanced", "DirectX.Force", sAresIni))
        Case "hardware": comboDirectXForce.ListIndex = 1
        Case "emulation": comboDirectXForce.ListIndex = 2
        Case Else: comboDirectXForce.ListIndex = 0
        End Select
        For i = 1 To 5
            Select Case i
            Case 1: sSurface = "Composite"
            Case 2: sSurface = "Alternate"
            Case 3: sSurface = "Hidden"
            Case 4: sSurface = "Sidebar"
            Case 5: sSurface = "Tile"
            End Select
            Select Case LCase$(ReadINIStr("Graphics.Advanced", "Surface." & sSurface & ".Memory", sAresIni))
            Case "system"
                Select Case i
                Case 1: optSMComposite.Item(1) = True
                Case 2: optSMAlternate.Item(1) = True
                Case 3: optSMHidden.Item(1) = True
                Case 4: optSMSidebar.Item(1) = True
                Case 5: optSMTile.Item(1) = True
                End Select
                OriSM(i) = 1
            Case "vram"
                Select Case i
                Case 1: optSMComposite.Item(2) = True
                Case 2: optSMAlternate.Item(2) = True
                Case 3: optSMHidden.Item(2) = True
                Case 4: optSMSidebar.Item(2) = True
                Case 5: optSMTile.Item(2) = True
                End Select
                OriSM(i) = 2
            Case Else
                Select Case i
                Case 1: optSMComposite.Item(0) = True
                Case 2: optSMAlternate.Item(0) = True
                Case 3: optSMHidden.Item(0) = True
                Case 4: optSMSidebar.Item(0) = True
                Case 5: optSMTile.Item(0) = True
                End Select
                OriSM(i) = 0
            End Select
            force = BooleanStringToInteger(ReadINIStr("Graphics.Advanced", "Surface." & sSurface & ".Force3D", sAresIni))
            Select Case i
            Case 1: cboxF3DComposite.Value = force
            Case 2: cboxF3DAlternate.Value = force
            Case 3: cboxF3DHidden.Value = force
            Case 4: cboxF3DSidebar.Value = force
            Case 5: cboxF3DTile.Value = force
            End Select
            OriF3D(i) = force
        Next i
    Else
        comboDirectXForce.ListIndex = 0
        optSMComposite.Item(0) = True
        optSMAlternate.Item(0) = True
        optSMHidden.Item(0) = True
        optSMSidebar.Item(0) = True
        optSMTile.Item(0) = True
        cboxF3DComposite.Value = 0
        cboxF3DAlternate.Value = 0
        cboxF3DHidden.Value = 0
        cboxF3DSidebar.Value = 0
        cboxF3DTile.Value = 0
        For i = 1 To 5
            OriSM(i) = 0
            OriF3D(i) = 0
        Next i
    End If
    OriDirectXForce = comboDirectXForce.ListIndex
End Sub

Private Sub Form_Load()
    Call RefreshAresIniFlags
End Sub
