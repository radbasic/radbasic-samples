VERSION 5.00
Begin VB.Form FrmChat 
   Caption         =   "Open AI - Chat"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAPIKey 
      Height          =   285
      Left            =   1320
      Locked          =   0   'False
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   480
      Width           =   7935
   End
   Begin VB.TextBox txtAnswer 
      Height          =   3675
      Left            =   600
      Locked          =   0   'False
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   10755
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   9120
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtQuestion 
      Height          =   525
      Left            =   600
      Locked          =   0   'False
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   8235
   End
   Begin VB.Label lblAPIKey 
      Caption         =   "API Key:"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "AI:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "You:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   375
   End
End
Attribute VB_Name = "FrmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
    Dim textQuestion As String
    Dim textAnswer As String
    Dim paramUserApiKey As String
    Dim result As String
    Dim sBody As String
    Dim urlOpenAI As String
    Dim objHTTP As Object
    Dim userParamsOk As Boolean
    
    
    ' **** Call REST API **** '
    urlOpenAI = "https://api.openai.com/v1/chat/completions"
    
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "POST", urlOpenAI, False
    
    paramUserApiKey = txtAPIKey.Text
    textQuestion = txtQuestion.Text
    
    userParamsOk = True
    If paramUserApiKey = "" Then
        userParamsOk = False
        MsgBox "Error: You have to specify your OPEN AI API Key!"
    End If
    
    If textQuestion = "" Then
        userParamsOk = False
        MsgBox "Error: You have to specify your question request!"
    End If
    
    If userParamsOk Then
        paramUserApiKey = "Bearer " & paramUserApiKey
    
        objHTTP.setRequestHeader "Content-Type", "application/json"
        objHTTP.setRequestHeader "Authorization", paramUserApiKey
        
        ' txtQuestion.Text => Input text from user
        sBody = "{ ""model"": ""gpt-3.5-turbo"", ""messages"": [ {""role"": ""user"", ""content"": """ & txtQuestion.Text & """} ] }"
        
        objHTTP.send (sBody)
        result = objHTTP.responseText
        
        'Set objHTTP = Nothing
        
        ' **** Extract text **** '
        Dim idxStart As Long, idxEnd As Long, strToSearchLength As Long
        Dim strToSearch As String
        
        strToSearch = "content"":"""
        strToSearchLength = Len(strToSearch)
            
        idxStart = InStr(1, result, strToSearch)
    
        idxStart = idxStart + strToSearchLength
        idxEnd = InStr(idxStart, result, """")
        
        textAnswer = Mid(result, idxStart, idxEnd - idxStart)
        textAnswer = Replace(textAnswer, "\n", vbNewLine)
        
        ' **** Show Text **** '
        txtAnswer.Text = textAnswer
    End If
End Sub

