VERSION 5.00
Begin VB.Form FrmStringSample 
   Caption         =   "Sample: String Operations"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConcatenate 
      Caption         =   "Concatenate 1 with 2"
      Height          =   345
      Left            =   5175
      TabIndex        =   15
      Top             =   2955
      Width           =   1755
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "Right"
      Height          =   345
      Left            =   3930
      TabIndex        =   14
      Top             =   2955
      Width           =   810
   End
   Begin VB.CommandButton cmdMid 
      Caption         =   "Mid"
      Height          =   345
      Left            =   2745
      TabIndex        =   13
      Top             =   2955
      Width           =   810
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "Left"
      Height          =   345
      Left            =   1575
      TabIndex        =   11
      Top             =   2940
      Width           =   810
   End
   Begin VB.TextBox textResult 
      Height          =   315
      Left            =   1350
      TabIndex        =   9
      Top             =   3765
      Width           =   6360
   End
   Begin VB.TextBox TextLength 
      Height          =   285
      Left            =   1095
      TabIndex        =   8
      Text            =   "8"
      Top             =   2085
      Width           =   615
   End
   Begin VB.TextBox TextStart 
      Height          =   285
      Left            =   1095
      TabIndex        =   7
      Text            =   "2"
      Top             =   1770
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "String Data"
      Height          =   1230
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   4260
      Begin VB.TextBox TextString2 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Text            =   "This is a an String number 2"
         Top             =   690
         Width           =   3045
      End
      Begin VB.TextBox textString1 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "This is a an String number 1"
         Top             =   315
         Width           =   3045
      End
      Begin VB.Label Label1 
         Caption         =   "String 2:"
         Height          =   270
         Left            =   210
         TabIndex        =   3
         Top             =   705
         Width           =   630
      End
      Begin VB.Label lblString1 
         Caption         =   "String 1:"
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   330
         Width           =   630
      End
   End
   Begin VB.Label Label4 
      Caption         =   "String Operation:"
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   1290
   End
   Begin VB.Label lblResult 
      Caption         =   "Result:"
      Height          =   315
      Left            =   690
      TabIndex        =   10
      Top             =   3840
      Width           =   600
   End
   Begin VB.Label Label3 
      Caption         =   "Length:"
      Height          =   300
      Left            =   465
      TabIndex        =   6
      Top             =   2115
      Width           =   570
   End
   Begin VB.Label Label2 
      Caption         =   "Start:"
      Height          =   225
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   465
   End
End
Attribute VB_Name = "FrmStringSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdConcatenate_Click()
    Dim str1 As String
    Dim str2 As String
    
    str1 = textString1.Text
    str2 = TextString2.Text
    
    textResult.Text = str1 & str2
    
End Sub

Private Sub cmdLeft_Click()
    Dim str1 As String
    Dim lengthParam As Integer
    
    lengthParam = CInt(TextLength.Text)
    
    str1 = textString1.Text
    str1 = Left(str1, lengthParam)
    
    textResult.Text = str1
    
End Sub

Private Sub cmdMid_Click()
    Dim str1 As String
    Dim startParam As Integer
    Dim lengthParam As Integer
    
    startParam = CInt(TextStart.Text)
    lengthParam = CInt(TextLength.Text)
    
    str1 = textString1.Text
    str1 = Mid(str1, startParam, lengthParam)
    
    textResult.Text = str1
End Sub

Private Sub cmdRight_Click()
    Dim str1 As String
    Dim lengthParam As Integer
    
    lengthParam = CInt(TextLength.Text)
    
    str1 = textString1.Text
    str1 = Right(str1, lengthParam)
    
    textResult.Text = str1
End Sub
