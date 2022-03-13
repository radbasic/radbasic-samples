VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Basic Notepad"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7935
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   7920
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   11990
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File"
      Begin VB.Menu MenuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu MenuFileSepa1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MenuEdit 
      Caption         =   "Edit"
      Begin VB.Menu MenuEditFind 
         Caption         =   "Find"
      End
   End
   Begin VB.Menu MenuHelp 
      Caption         =   "Help"
      Begin VB.Menu MenuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MenuEditFind_Click()
    FrmFind.Show
End Sub

Private Sub MenuHelpAbout_Click()
    Rem MsgBox "Basic Notepad 1.0"
    FrmAbout.Show
    
    
End Sub
