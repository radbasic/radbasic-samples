VERSION 5.00
Begin VB.Form FrmFind
   Caption         =   "FrmFind"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3
   Begin VB.CommandButton CmdSearch
      Caption         =   "Search"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox TextSearch
      Height          =   285
      Left            =   960
      MultiLine       =   0   'False
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label LabelSearchText
      Caption         =   "Search:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "FrmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSearch_Click()
    MsgBox "Search text..."
    Unload Me
End Sub
