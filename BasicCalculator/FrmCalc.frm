VERSION 5.00
Begin VB.Form FrmCalc 
   Caption         =   "RB Calculator"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton butC 
      Caption         =   "C"
      Height          =   615
      Left            =   360
      TabIndex        =   16
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton butEqual 
      Caption         =   "="
      Height          =   615
      Left            =   1800
      TabIndex        =   15
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton butDiv 
      Caption         =   "/"
      Height          =   615
      Left            =   2520
      TabIndex        =   14
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton butMult 
      Caption         =   "X"
      Height          =   615
      Left            =   2520
      TabIndex        =   13
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton butSubs 
      Caption         =   "_"
      Height          =   615
      Left            =   2520
      TabIndex        =   12
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton but9 
      Caption         =   "9"
      Height          =   615
      Left            =   1800
      TabIndex        =   11
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton but8 
      Caption         =   "8"
      Height          =   615
      Left            =   1080
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton but7 
      Caption         =   "7"
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton but6 
      Caption         =   "6"
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton but5 
      Caption         =   "5"
      Height          =   615
      Left            =   1080
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton but4 
      Caption         =   "4"
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton butAdd 
      Caption         =   "+"
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   4080
      Width           =   495
   End
   Begin VB.CommandButton but3 
      Caption         =   "3"
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton but2 
      Caption         =   "2"
      Height          =   615
      Left            =   1080
      TabIndex        =   3
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton but1 
      Caption         =   "1"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton but0 
      Caption         =   "0"
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblResult 
      Caption         =   "0"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "FrmCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Basic Calculator
' SAMPLE for RAD Basic
' This could be done in a more elegant way with controls array
' But it is not supported fully in this version
' In further versions it will be rewritten.
Dim curNumber As Double
Dim number1 As Double
Dim number2 As Double
' 1: for add, 2: subs, 3: mult, 4: div
Dim currentOperation As Integer

Private Sub but0_Click()
    curNumber = curNumber * 10 + 0
    lblResult.Caption = curNumber
End Sub

Private Sub but1_Click()
    curNumber = curNumber * 10 + 1
    lblResult.Caption = curNumber
End Sub

Private Sub but2_Click()
    curNumber = curNumber * 10 + 2
    lblResult.Caption = curNumber
End Sub

Private Sub but3_Click()
    curNumber = curNumber * 10 + 3
    lblResult.Caption = curNumber
End Sub

Private Sub but4_Click()
    curNumber = curNumber * 10 + 4
    lblResult.Caption = curNumber
End Sub

Private Sub but5_Click()
    curNumber = curNumber * 10 + 5
    lblResult.Caption = curNumber
End Sub

Private Sub but6_Click()
    curNumber = curNumber * 10 + 6
    lblResult.Caption = curNumber
End Sub

Private Sub but7_Click()
    curNumber = curNumber * 10 + 7
    lblResult.Caption = curNumber
End Sub

Private Sub but8_Click()
    curNumber = curNumber * 10 + 8
    lblResult.Caption = curNumber
End Sub

Private Sub but9_Click()
    curNumber = curNumber * 10 + 9
    lblResult.Caption = curNumber
End Sub


Private Sub butC_Click()
    curNumber = 0
    number1 = 0
    number2 = 0
    
    lblResult.Caption = curNumber
End Sub

Private Sub butEqual_Click()
    number2 = curNumber
    
    ' Do Calculation
    If (currentOperation = 1) Then
        curNumber = number1 + number2
    ElseIf (currentOperation = 2) Then
        curNumber = number1 - number2
    ElseIf (currentOperation = 3) Then
        curNumber = number1 * number2
    ElseIf (currentOperation = 4) Then
        curNumber = number1 / number2
    Else
        MsgBox "Error! Invalid Operation!"
    End If
    
    lblResult.Caption = curNumber
End Sub
Private Sub butAdd_Click()
    currentOperation = 1
    number1 = curNumber
    curNumber = 0
    lblResult.Caption = curNumber
End Sub
Private Sub butSubs_Click()
    currentOperation = 2
    number1 = curNumber
    curNumber = 0
    lblResult.Caption = curNumber
End Sub

Private Sub butMult_Click()
    currentOperation = 3
    number1 = curNumber
    curNumber = 0
    lblResult.Caption = curNumber
End Sub
Private Sub butDiv_Click()
    currentOperation = 4
    number1 = curNumber
    curNumber = 0
    lblResult.Caption = curNumber
End Sub


