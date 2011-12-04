VERSION 5.00
Begin VB.Form frmPleaseWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base: Please Wait"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "This may take a few minutes, depending on the size of the mod."
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Compiling and installing mod files..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmPleaseWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call frmMain.WriteLogEntry("Form_Load: frmPleaseWait", LogLevel2)
End Sub
