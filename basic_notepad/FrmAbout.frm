VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Basic Notepad"
   ClientHeight    =   2730
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3315
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3315
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label lblDescription 
      Caption         =   "Sample App developed in VB6 and compiled in RAD Basic"
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   240
      TabIndex        =   1
      Top             =   1125
      Width           =   2805
   End
   Begin VB.Label lblTitle 
      Caption         =   "Basic Notepad"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2805
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0.0"
      Height          =   225
      Left            =   240
      TabIndex        =   3
      Top             =   780
      Width           =   2805
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Rem Option Explicit
Private Sub cmdOK_Click()
  Unload Me
End Sub

