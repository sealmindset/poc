Attribute VB_Name = "SecureArch"
Sub NewApp()
    frmAppName.Show
End Sub
Sub SASubmit()
    recSADB
End Sub
Public Sub addDataToTable(ByVal strTableName As String, strData As String)
Dim objWorksheet As Worksheet
Dim objListObject As ListObject
Dim objTable As ListObject
Dim objRowCount As Integer

Set objWorksheet = ActiveWorkbook.Worksheets("Lists")
Set objTable = objWorksheet.Range(strTableName).ListObject
' Count total in ListObject Table
objRowCount = objWorksheet.Range(strTableName).ListObject.ListRows.Count + 1
' Add
objTable.DataBodyRange.Rows(objRowCount) = strData
' Remove Duplicates
RmDupRows objTable:=objTable, strTableName:=strTableName
End Sub
Sub AddNameData()
Dim strValue As String
strValue = frmAppName.txtAddName.Value

If IsAlphaNumeric(strValue) Then
    addDataToTable strTableName:="Applications", strData:=strValue
    Cells(12, 4).Value = strValue
    frmAppName.txtAddName.Value = ""
    frmAppName.Hide
    ActiveWorkbook.Sheets("SA Profiler").Select
Else
    ClearForm
End If

End Sub
Sub RmDupRows(ByVal objTable As ListObject, strTableName As String)
' Remove Duplicates in a ListObject Table
With objTable
    .Range.RemoveDuplicates Columns:=Array( _
        .ListColumns(strTableName).Index), _
        Header:=xlYes
End With
End Sub
Sub FindAllTablesOnSheet()
    Dim oSh As Worksheet
    Dim oLo As ListObject
    Set oSh = ThisWorkbook.Sheets("Lists")
    For Each oLo In oSh.ListObjects
        Application.Goto oLo.Range
        MsgBox "Table found: " & oLo.Name & ", " & oLo.Range.Address
    Next
End Sub
Sub ClearForm()
    frmAppName.txtAddName.Value = ""
End Sub
Function IsAlphaNumeric(strSource As String) As Boolean
    Dim i As Integer
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 48 To 57, 65 To 90, 97 To 122: 'include 32 if you want to include space
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    If Not IsEmpty(strResult) Then
        IsAlphaNumeric = True
    Else
        IsAlphaNumeric = False
    End If
End Function
Sub restForm()
'
' Clear the contents in the cells in a specified range
'
'Sheets(saFrmName).Range(saFrmDQAll).ClearContents
MsgBox saFrmName & " " & saFrmDQAll
End Sub
Sub recSADB()
Dim arrHeader As Variant

On Error GoTo ErrorHandler

arrHeader = Array("Date", "Version", "Portfolios", "Applications", "Architectures", "Hosted", "Platform", "DataTopics", "DataClass", "Authentication", "Authorization", "PrivilegeAcctMgt", "RolesManaged", "RoleTypes", "Boundaries", "Segmented", "SecurityControls", "DataMethods", "AccessMethods", "Cryptography", "Encryption", "TrustModels", "KeyManagement", "DataInTransit", "DataAtRest", "Integrations", "Vendors", "Logging")

'Create Worksheet and fill it in
If Not WorksheetExists(saFrmDB) Then
    'MsgBox strWS + " Does Not Exist and Empty"
    CreateWS strName:=saFrmDB
    Sheets(saFrmDB).Cells(1).Resize(1, 28).Value = arrHeader
End If

ThisWorkbook.Sheets(saFrmName).Range(saFrmDQAll).Copy

'Open Workbook 2 and paste data (transposed) on first available row starting in column B
With ThisWorkbook.Sheets(saFrmDB)
    ' find last row with data in destination workbook "wbDatabase.xlsm"
    DestLastRow = .Cells(.Rows.Count, "A").End(xlUp).Offset(1).row
    'Create Date
    .Range("A" & DestLastRow).Value = Now
    'Version
    .Range("B" & DestLastRow).Value = "1.0"
     'paste special only values, and transpose
    .Range("C" & DestLastRow).PasteSpecial xlValues, Transpose:=True
End With
Application.CutCopyMode = False

ErrorHandler:
Dim strSubName As String
strSubName = "recSADB"
Select Case Err.Number
Case 0
    MsgBox "All fields are now cleared.", vbInformation
    'GoTo ExitOut
Case 1004
    MsgBox Err.Number & " " & Err.Description, vbCritical
    'If Application.Version = "14.3" Then
    '    Call MacInfo(strSubName, Err, Err.Source, Erl, Err.Description)
    'Else
        Call GetInfo(strSubName, Err, Err.Source, Erl, Err.Description)
    'End If
    'If DBStatus = True Then Call ExportData("errlog", "errlog")
    Err.Clear
    'GoTo ExitOut
Case Else
    'MsgBox Err.Number & " " & Err.Description, vbCritical
    'If Application.Version = "14.3" Then
    '    Call MacInfo(strSubName, Err, Err.Source, Erl, Err.Description)
    'Else
        Call GetInfo(strSubName, Err, Err.Source, Erl, Err.Description)
    'End If
    'If DBStatus = True Then Call ExportData("errlog", "errlog")
    Err.Clear
    'GoTo ExitOut
End Select
Resume Next
End Sub
Sub Test()
Dim strVal As String
strVal = ThisWorkbook.Sheets(saFrmName).Range(saFrmDQAll).Select
MsgBox strVal
End Sub
Sub RangeCompare() '(ByVal Range1 As Range, Range2 As Range) As Boolean
' Returns TRUE if the ranges are identical.
' This function is case-sensitive.
' For ranges with fewer than ~1000 cells, cell-by-cell comparison is faster

' WARNING: This function will fail if your range contains error values.

Dim RangeCompare As Boolean

RangeCompare = False

If Sheets("SADB").Range("A2:AB2").Cells.Count <> Sheets("SADB").Range("A3:AB3").Cells.Count Then
    RangeCompare = False
ElseIf Sheets("SADB").Range("A2:AB2").Cells.Count = 1 Then
    RangeCompare = Sheets("SADB").Range("A2:AB2").Value2 = Sheets("SADB").Range("A3:AB3").Value2
Else
    RangeCompare = SHA1HASH(Join2D(Sheets("SADB").Range("A2:AB2").Value2)) = SHA1HASH(Join2D(Sheets("SADB").Range("A3:AB3").Value2))
End If

MsgBox RangeCompare
End Sub
Public Function Join2D(ByVal vArray As Variant, Optional ByVal sWordDelim As String = " ", Optional ByVal sLineDelim As String = vbNewLine) As String
Dim i As Long, j As Long
Dim aReturn() As String
Dim aLine() As String

On Error GoTo ErrorHandler

ReDim aReturn(LBound(vArray, 1) To UBound(vArray, 1))
ReDim aLine(LBound(vArray, 2) To UBound(vArray, 2))

For i = LBound(vArray, 1) To UBound(vArray, 1)
    For j = LBound(vArray, 2) To UBound(vArray, 2)
        'Put the current line into a 1d array
        aLine(j) = vArray(i, j)
    Next j
    'Join the current line into a 1d array
    aReturn(i) = Join(aLine, sWordDelim)
Next i

Join2D = Join(aReturn, sLineDelim)
    
ErrorHandler:
Dim strSubName As String
strSubName = "Join2D"
Select Case Err.Number
Case 0
    MsgBox "All fields are now cleared.", vbInformation
    'GoTo ExitOut
Case 1004
    MsgBox Err.Number & " " & Err.Description, vbCritical
    'If Application.Version = "14.3" Then
    '    Call MacInfo(strSubName, Err, Err.Source, Erl, Err.Description)
    'Else
        Call GetInfo(strSubName, Err, Err.Source, Erl, Err.Description)
    'End If
    'If DBStatus = True Then Call ExportData("errlog", "errlog")
    Err.Clear
    'GoTo ExitOut
Case Else
    'MsgBox Err.Number & " " & Err.Description, vbCritical
    'If Application.Version = "14.3" Then
    '    Call MacInfo(strSubName, Err, Err.Source, Erl, Err.Description)
    'Else
        Call GetInfo(strSubName, Err, Err.Source, Erl, Err.Description)
    'End If
    'If DBStatus = True Then Call ExportData("errlog", "errlog")
    Err.Clear
    'GoTo ExitOut
End Select
Resume Next

End Function

Sub Range_Find_Method()
'Finds the last non-blank cell on a sheet/range.

Dim lRow As Long
Dim lCol As Long
    
    lRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).row
    
    MsgBox "Last Row: " & lRow

End Sub
Sub testme()
Dim strAddr As String
findLatestSA strAppName:="SMS", strVerNum:="2"
End Sub

Function findLatestSA(ByVal strAppName As String, strVerNum As String, _
                    Optional strSheets As String = "SADB", _
                    Optional strRange As String = "B:AB", _
                    Optional strAppCol As String = "D", _
                    Optional strVerCol As String = "B") As String

    Dim rngFound As Range
    Dim strFirst As String
    
    'MsgBox Sheets(strSheets).Range("B1").SpecialCells(xlCellTypeLastCell).Address
    
    Set rngFound = Sheets(strSheets).Range(strRange).Find(What:=strAppName, _
                               LookIn:=xlValues, _
                               LookAt:=xlWhole, _
                               SearchOrder:=xlByRows, _
                               SearchDirection:=xlNext, _
                               MatchCase:=False)
    
    If Not rngFound Is Nothing Then
        strFirst = rngFound.Address
        Do
            If LCase(Cells(rngFound.row, strAppCol).Text) = LCase(strAppName) And Cells(rngFound.row, strVerCol).Text = LCase(strVerNum) Then
                'Found a match
                MsgBox "Found a match at: " & rngFound.Address & Chr(10) & _
                       "Value in column B: " & Cells(rngFound.row, strVerCol).Text & Chr(10) & _
                       "Value in column D: " & Cells(rngFound.row, strAppCol).Text
            End If
            Set rngFound = Columns(strAppCol).Find(strAppName, rngFound, xlValues, xlWhole)
        Loop While rngFound.Address <> strFirst
    End If

    Set rngFound = Nothing

End Function
