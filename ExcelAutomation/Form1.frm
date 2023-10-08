VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Sample Excel Automation"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdateExcel 
      Caption         =   "Update Excel"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label1"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpdateExcel_Click()
    Dim ObjExcel As Object
    Dim WorkBook As Object
    Dim ExcelFilePath As String
    
    ' Open Excel
    ExcelFilePath = App.Path & "\RBSampleMyWorkbook.xlsx"
    Set ObjExcel = CreateObject("Excel.Application")
    Set WorkBook = ObjExcel.Workbooks.Open(ExcelFilePath)
    
    ' Make same edits
    ObjExcel.Worksheets("Sheet X").Cells(2, 1) = "Edited From RAD Basic"

    ' Save and close excel
    WorkBook.Close SaveChanges:=True
    ObjExcel.Quit
    
    ' Recommended for memory management and ref counter
    Set ObjExcel = Nothing
    
    ' Update to UI
    lblStatus.Caption = "Updated excel file: " & ExcelFilePath
    lblStatus.Visible = True
    
End Sub

