VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNew 
   Caption         =   "Add New Item to List"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4110
   OleObjectBlob   =   "frmNew.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Private Sub cmdOK_Click()
Dim ws As Worksheet
Dim nm As Name
Dim rng As Range
Dim i As Long

On Error Resume Next

strNew = Me.txtNew.value
If strNew = "" Then GoTo exitHandler
bFromNew = True

Set nm = ActiveWorkbook.Names(strDVList)
Set ws = Worksheets(nm.RefersToRange.Parent.Name)
Set rng = ws.Range(nm.RefersToRange.Address)

If ws Is Nothing Then
    Set rng = Evaluate(ActiveWorkbook.Names(strDVList).RefersTo)
    Set ws = rng.Parent
End If

If ws Is Nothing Then
    If Application.International(xlListSeparator) = ";" Then
      strDVList = Replace(strDVList, ";", ",")
    End If
    Set rng = Evaluate("=" & strDVList)
    Set ws = rng.Parent
End If

If ws Is Nothing Then
    MsgBox "Could not find the list"
    GoTo exitHandler
End If

If Application.WorksheetFunction _
  .CountIf(rng, strNew) Then
  MsgBox "Item is already in the list"
  GoTo exitHandler
Else
  i = ws.Cells(Rows.Count, rng.Column).End(xlUp).Row + 1
  ws.Cells(i, rng.Column).value = strNew
  
If nm Is Nothing Then
  Set nm = rng.Name
End If
  
  nm.RefersTo = "='" & ws.Name & "'!" _
    & rng.Resize(rng.Rows.Count + 1).Address
 Set rng = nm.RefersToRange
  
  rng.Sort Key1:=ws.Cells(rng.Cells(1, 1).Row, rng.Column), _
    Order1:=xlAscending, Header:=xlYes, _
    OrderCustom:=1, MatchCase:=False, _
    Orientation:=xlTopToBottom
End If

exitHandler:
   frmNew.Hide
   frmDVList.Show
   Exit Sub
End Sub

Private Sub UserForm_Activate()
On Error Resume Next
Me.txtNew.value = ""
Me.txtNew.SetFocus
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
If bPos = True Then
   Me.StartUpPosition = 0
   Me.Top = frmDVList.Top + 5
   Me.Left = frmDVList.Left + 5
End If

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
On Error Resume Next
  frmNew.Hide
  frmDVList.Show
  Exit Sub
End Sub

