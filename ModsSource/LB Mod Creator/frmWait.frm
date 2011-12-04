VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Launch Base Mod Creator"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "frmWait.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "This may take a few minutes, depending on the size of your mod."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Compressing files and creating installer..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
