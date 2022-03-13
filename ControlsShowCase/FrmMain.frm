VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Basic Controls Test"
   ClientHeight    =   6780
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butClickMe 
      Caption         =   "Click me!"
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "TextBox: Simple Entry (no multi)"
      Top             =   3240
      Width           =   3135
   End
   Begin VB.CommandButton butStatic 
      Caption         =   "This is a button"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   6120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2295
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton Option4 
         Caption         =   "Other option Button 2 inside a frame"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   3615
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option Button 1 inside a frame"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Value           =   -1  'True
         Width           =   3255
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Other option/radio button"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "This is a option/radio button"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Value           =   -1  'True
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "This is a CheckBox"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "This is a Label"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3135
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "Help"
      Begin VB.Menu MnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub butClickMe_Click()
    MsgBox "You clicked me!"
End Sub

Private Sub MnuAbout_Click()
    MsgBox "Sample application for testing and showcase of basic GUI controls"
End Sub

Private Sub MnuExit_Click()
    Unload Me
End Sub
