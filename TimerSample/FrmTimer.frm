VERSION 5.00
Begin VB.Form FrmTimer 
   Caption         =   "Timer Sample"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   2880
   End
   Begin VB.CommandButton CmdStart 
      Caption         =   "Start"
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label lblArray 
      Caption         =   "D"
      Height          =   495
      Index           =   10
      Left            =   4920
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblArray 
      Caption         =   "L"
      Height          =   495
      Index           =   9
      Left            =   4560
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblArray 
      Caption         =   "R"
      Height          =   495
      Index           =   8
      Left            =   4200
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblArray 
      Caption         =   "O"
      Height          =   495
      Index           =   7
      Left            =   3840
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblArray 
      Caption         =   "W"
      Height          =   495
      Index           =   6
      Left            =   3480
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblArray 
      Caption         =   "O"
      Height          =   495
      Index           =   5
      Left            =   2640
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblArray 
      Caption         =   "L"
      Height          =   495
      Index           =   4
      Left            =   2160
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblArray 
      Caption         =   "L"
      Height          =   495
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblArray 
      Caption         =   "E"
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblArray 
      Caption         =   "H"
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblRBInfo 
      Caption         =   "Timer sample from RAD Basic"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   5415
   End
End
Attribute VB_Name = "FrmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1
Dim currVal As Integer
Private Sub CmdStart_Click()
    currVal = 1
    
    CmdStart.Caption = "Timer started!"
    Timer1.Enabled = True
    
End Sub


Private Sub Timer1_Timer()
    lblArray(currVal).Visible = True
    currVal = currVal + 1
    If (currVal > 10) Then
        Timer1.Enabled = False
        CmdStart.Caption = "Timer stopped!"
    End If
        
End Sub
