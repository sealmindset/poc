VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDVList 
   Caption         =   "Select Items to Add"
   ClientHeight    =   5280
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   4005
   OleObjectBlob   =   "frmDVList.frx":0000
End
Attribute VB_Name = "frmDVList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Dim bInit As Boolean

'===================================
Private Sub cmdALL_Click()
Dim lCountList As Long
On Error Resume Next

With Me.lstDV
   For lCountList = 0 To .ListCount - 1
      .Selected(lCountList) = True
   Next lCountList
   End With
   
exitHandler:
   Exit Sub
   
errHandler:
   MsgBox "Could not select all items"
   Resume exitHandler

End Sub
'===================================
Private Sub cmdClear_Click()
Dim lCountList As Long
On Error Resume Next

With Me.lstDV
    For lCountList = 0 To .ListCount - 1
        .Selected(lCountList) = False
    Next lCountList
End With

exitHandler:
    Exit Sub

errHandler:
    MsgBox "Could not clear all items"
    Resume exitHandler
End Sub
'===================================
Private Sub cmdClose_Click()
On Error Resume Next
  Unload Me
  ActiveCell.Offset(1, 0).Activate  'move down one row
End Sub
'===================================
Private Sub cmdNew_Click()
On Error Resume Next
frmDVList.Hide
frmNew.Show
End Sub
'===================================
Private Sub cmdOK_Click()
Dim strSelItems As String
Dim lCountList As Long
Dim strAdd As String
Dim lOff As Long
Dim rngAdd As Range

On Error Resume Next
'get list data for selected items

lOff = 0

With Me.lstDV
  For lCountList = 0 To .ListCount - 1
  
    If .Selected(lCountList) Then
      strAdd = .List(lCountList)
      lOff = lOff + 1
      If strSelItems = "" Then
        strSelItems = strAdd
      Else
          strSelItems = strSelItems & strSep & strAdd
      End If
    Else
      strAdd = ""
    End If '.Selected(lCountList)
  
  Next lCountList

End With

With ActiveSheet
  If bAcross = True Then
    Set rngAdd = .Range(ActiveCell, ActiveCell.Offset(0, lOff - 1))
    rngAdd.value = Split(strSelItems, strSep)
  Else
    If bDown = True Then
      Set rngAdd = .Range(ActiveCell, ActiveCell.Offset(lOff - 1, 0))
      rngAdd.value = Application.Transpose(Split(strSelItems, strSep))
    Else
      Set rngAdd = Selection
      rngAdd.value = strSelItems
    End If
  End If
End With

Unload Me

exitHandler:
Exit Sub

End Sub

'===================================
Private Sub lstDV_Click()
On Error Resume Next
If bInit = False Then
  'auto close form when selection made
  '   in Single Select columns
  If lMS = 0 Then
     cmdOK_Click
  End If
End If
End Sub

Private Sub UserForm_Activate()
On Error Resume Next
UserForm_Initialize
End Sub

'===================================
Private Sub UserForm_Initialize()
Dim strSelItems As String
Dim lCountList As Long
Dim strAdd As String
Dim strCell As String
Dim strMsg As String
Dim c As Range
Dim i As Long
Dim Ar As Variant
Dim cList As Range
Dim ws As Worksheet
Dim nm As Name
Dim rng As Range
Dim rngList As Range
Dim vArr As Variant
Dim iArr As Long
Dim lIndex As Long

On Error Resume Next

Me.cmdNew.Visible = bNew
bInit = True
If bPos = True Then
   Me.StartUpPosition = 0
   Me.Top = ActiveWindow.Top _
            + ActiveWindow.Height / 2 _
            - Me.Height / 2
   Me.Left = ActiveWindow.Left _
            + ActiveWindow.Width / 2 _
            - Me.Width / 2
Else
   Me.StartUpPosition = 1  'center owner
End If

Me.lstDV.MultiSelect = lMS
If lMS = 0 Then 'single selection
  Me.cmdALL.Visible = False
Else
  Me.cmdALL.Visible = True
End If

If bFromNew = True Then
     With Me.lstDV
      .AddItem strNew
      .Selected(.ListCount - 1) = True
     End With
     bFromNew = False
Else
  Me.lstDV.Clear
  
  Set nm = ActiveWorkbook.Names(strDVList)
  Set ws = Worksheets(nm.RefersToRange.Parent.Name)
  Set rng = ws.Range(nm.RefersToRange.Address)
  
  If ws Is Nothing Then
      Set rng = Evaluate(ActiveWorkbook.Names(strDVList).RefersTo)
      Set ws = rng.Parent
  End If
  
  If Not rng Is Nothing Then
    For Each cList In rng
       Me.lstDV.AddItem cList.value
    Next cList
  Else
      If Application.International(xlListSeparator) = ";" Then
        strDVList = Replace(strDVList, ";", ",")
      End If
      vArr = Evaluate("=" & strDVList)
      '2016-10-05 handle one-cell lists
      If IsArray(vArr) Then
        For lIndex = LBound(vArr) To UBound(vArr)
          Me.lstDV.AddItem vArr(lIndex, 1)
        Next lIndex
      Else
        Me.lstDV.AddItem vArr
      End If
  End If
  
  Set c = ActiveCell
  strCell = c.value
  
  Ar = Split(strCell, strSep)
  
  For i = LBound(Ar) To UBound(Ar)
     With Me.lstDV
        For lCountList = 0 To .ListCount - 1
           If .MultiSelect = fmMultiSelectMulti Then
              If CStr(.List(lCountList)) = CStr(Ar(i)) Then
                 On Error GoTo errHandler
                 .Selected(lCountList) = True
                 Exit For
              End If
           Else
              If CStr(.List(lCountList)) = CStr(Ar(i)) Then
                 On Error GoTo errHandler
                 .ListIndex = lCountList
                 Exit For
              End If
           End If
        Next lCountList
     End With
  Next i
   
  If strNew <> "" Then
     With Me.lstDV
        For lCountList = 0 To .ListCount - 1
           If .MultiSelect = fmMultiSelectMulti Then
              If CStr(.List(lCountList)) = strNew Then
                 On Error GoTo errHandler
                 .Selected(lCountList) = True
                 Exit For
              End If
           Else
              If CStr(.List(lCountList)) = CStr(Ar(i)) Then
                 On Error GoTo errHandler
                 .ListIndex = lCountList
                 Exit For
              End If
           End If
        Next lCountList
     End With
  End If
End If

exitHandler:
    bInit = False
    strNew = ""
    Set c = Nothing
    Exit Sub
    
errHandler:
    MsgBox "Could not select all items"
    Resume exitHandler

End Sub


