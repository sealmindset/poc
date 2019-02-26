VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'show Listbox when a data validation cell is selected
Dim oldVal As String
Dim newVal As String
Dim strList As String
Dim ValType As Long
Dim lStart As Long
Dim lEnd As Long
Dim bRun As Boolean
Application.EnableEvents = False

On Error Resume Next
ValType = Target.Validation.Type   'ValType = 3 if data val list
On Error GoTo errHandler

If ValType <> 3 Then GoTo exitHandler
'Next line commented out, if no On/Off cell entered
'If Me.Range("0").Value <> "" Then GoTo exitHandler

  Select Case Target.Row    'where should code run
   Case Is <> 0    'if Case is > zero, all data val cells
    bRun = True    'continue
  Case Else
    bRun = False    'exit
End Select
      
If bRun = False Then
  GoTo exitHandler
End If

strList = Target.Validation.Formula1
strSep = ", "
bPos = True
      
strList = Right(strList, Len(strList) - 1)
      
  Select Case Target.Row    ' multi-select locations
         Case 9    'if Case is > zero, all are single-select
            lMS = 0  'single select
         Case Else
            lMS = 1  'multi-select
      End Select
      
  Select Case Target.Row    'Show New button locations
         Case 9, 12, 15, 18, 21, 24, 27, 30, 33, 36, 39, 42, 45 'if Case is <> zero, all allow add new
            bNew = True 'New button visible
         Case Else
            bNew = False 'New button hidden
      End Select
      
  Select Case Target.Row    'Fill Down locations
         Case 0       'if Case is > zero, all fill down
            bDown = True   'Fill down
         Case Else
            bDown = False    'Same Cell
      End Select
      
  Select Case Target.Row    'Fill Across locations
         Case 0       'if Case is > zero, all fill across
            bAcross = True    'Fill across
         Case Else
            bAcross = False   'Same Cell
      End Select
      
      strDVList = strList
      
      frmDVList.Show

exitHandler:
  Application.EnableEvents = True
  Exit Sub
errHandler:
    Resume exitHandler
End Sub

