VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Factorial"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCalculate 
      Caption         =   "Calculate!"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox TxtResult 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox TxtNumber 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label LblResult 
      Caption         =   "Result:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label LblEnterNumber 
      Caption         =   "Enter Number:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdCalculate_Click()
    Dim usrNumber As Integer
    Dim resFactorial As Integer
    
    ' Explicit conversion from text to numeric value
    usrNumber = CInt(TxtNumber.Text)
    
    resFactorial = Factorial(usrNumber)
    
    TxtResult.Text = resFactorial
    
End Sub

Private Function Factorial(n As Integer) As Integer
    Dim calc As Integer
    
    If n < 2 Then
       Factorial = 1
    Else
       ' Factorial = n * Factorial(n - 1)
       calc = Factorial(n - 1)
       Factorial = n * calc
    End If
End Function

