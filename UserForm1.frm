VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Identified Abuse Cases"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19755
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Accept_Click()
    Application.DisplayAlerts = False
        Sheets("temp").Delete
    Application.DisplayAlerts = True
    Me.Hide
End Sub

Private Sub CommandButton1_Click()
    MsgBox "Send Abuse Cases to the Printer.", vbInformation
End Sub

Private Sub CommandButton2_Click()
    MsgBox "Create Jira tickets.", vbInformation
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Okay_Click()
    Me.Hide
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Initialize()
Dim lbtarget As MSForms.ListBox
Dim rngSource As Range
Dim LastRow As Long

'Refresh UsedRange
Worksheets("temp").UsedRange

'Find Last Row
LastRow = Worksheets("temp").Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

'Set reference to the range of data to be filled
Set rngSource = Worksheets("temp").Range("A2:C" & LastRow)

'Fill the listbox
Set lbtarget = Me.ListBox1
With lbtarget
    .Clear
    .ColumnHeads = True
    'Determine number of columns
    .ColumnCount = 3
    'Set column widths
    .ColumnWidths = "100 pt;678 pt;200 pt"
    'Insert the range of data supplied
    .List = rngSource.Cells.value
End With

Exit Sub
ErrorHandle:
MsgBox Err.Description
End Sub
