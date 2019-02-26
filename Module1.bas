Attribute VB_Name = "Module1"
Sub Submit()
    Application.Run ("ThreatModel")
End Sub
Sub SWSubmit()
    Application.Run ("SWThreatModel")
End Sub
Function WorksheetExists2(WorksheetName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    With wb
        On Error Resume Next
        WorksheetExists2 = (.Sheets(WorksheetName).Name = WorksheetName)
        On Error GoTo 0
    End With
End Function
Function WorksheetExists(sName As String) As Boolean
    WorksheetExists = Not WorksheetFunction.IsErr(Evaluate("'" & sName & "'!A1"))
End Function
Function CheckSheetEmpty(WorksheetName As String, Optional wb As Workbook) As Boolean
    If WorksheetFunction.CountA(ActiveSheet.UsedRange) = 0 And ActiveSheet.Shapes.Count = 0 Then
        CheckSheetEmpty = True
    Else
        CheckSheetEmpty = False
    End If
End Function
Function CreateWS(ByVal strName As String) As String
If WorksheetExists2(strName) Then
    'ThisWorkbook.Sheets(strName).Visible = True
Else
    With ThisWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = strName
    End With
End If
End Function
Function AddTM2WS(ByVal strDest As String, strSource As String) As String
Dim sht As Worksheet, DstSht As Worksheet
Set sht = ThisWorkbook.Sheets(strSource)
Set DstSht = ThisWorkbook.Sheets(strDest)
'Don't update the screen
With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With
'Copy everything that isn't blank in the source worksheet
With sht
    LastRow = .Range("A" & .Rows.Count).End(xlUp).Row
    .Range("A1:D" & LastRow).Copy
End With
'Append to destination worksheet
If IsEmpty(Range("A1").value) = True Then
    DstSht.Range("A" & DstSht.Rows.Count).End(xlUp).PasteSpecial xlPasteValues
Else
    DstSht.Range("A" & DstSht.Rows.Count).End(xlUp).Offset(1).PasteSpecial xlPasteValues
End If
Application.CutCopyMode = False
'Resume updating the screen
With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With
End Function
Function CountExist(ByVal strWS As String, strPattern As String) As Integer
' Select "Developer" tab
' Select "Visual Basic" icon from 'Code' ribbon section
' In "Microsoft Visual Basic for Applications" window select "Tools" from the top menu.
' Select "References"
' Check the box next to "Microsoft VBScript Regular Expressions 5.5" to include in your workbook.
' Click "OK"
Dim i As Long
Dim xCount As Integer
Dim regEx As New RegExp

With regEx
    .Global = True
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = strPattern
End With

For i = 1 To ActiveWorkbook.Sheets.Count
    If regEx.test(Sheets(i).Name) Then
        xCount = xCount + 1
        'MsgBox "Match " & Sheets(I).Name & " - " & xCount
    End If
Next
CountExist = xCount + 1
End Function
Function chkAbuseCase(ByVal strWS As String, strPattern As String) As Variant
' Find all worksheets that are related, add into an array
Dim myArray As Variant
Dim cntArray As String
Dim shtBefore As String, shtAfter As String
' Shove it into an array
myArray = CollectExist(strWS, strPattern)
' Get the worksheets name from the lastest version to the one before it
shtBefore = myArray(UBound(myArray) - 1)
shtAfter = myArray(UBound(myArray))
' Compare the sheets and identify the difference
chkAbuseCase = compareSheets(shtBefore, shtAfter)
End Function
Function CollectExist(ByVal strWS As String, strPattern As String) As Variant
' Collect all of the worksheets that match the pattern and return as an array
Dim i As Long
Dim arrValues As Variant
Dim regEx As New RegExp

With regEx
    .Global = True
    .MultiLine = False
    .IgnoreCase = True
    .Pattern = strPattern
End With

For i = 1 To ActiveWorkbook.Sheets.Count
    If regEx.test(Sheets(i).Name) Then
        If IsEmpty(arrValues) Then
            arrValues = Array(Sheets(i).Name)
        Else
            ReDim Preserve arrValues(UBound(arrValues) + 1)
            arrValues(UBound(arrValues)) = Sheets(i).Name
        End If
    End If
Next
CollectExist = arrValues
End Function
Function rmdup(ByVal strWS As String) As Boolean
Dim a As Long
For a = ActiveWorkbook.Worksheets(strWS).Cells(Rows.Count, 1).End(xlUp).Row To 1 Step -1
    If WorksheetFunction.CountIf(ActiveWorkbook.Worksheets(strWS).Range("A1:A" & a), Cells(a, 1)) > 1 Then Rows(a).Delete
Next
End Function
Function compareSheets(ByVal shtBefore As String, shtAfter As String) As Variant
' Compare worksheets
Dim mycell As Range
Dim mydiffs As Integer
Dim myArray As Variant

Dim xCell As Range
Dim xStr As String
Dim xRow As Long
Dim xCol As Long

On Error Resume Next
CreateWS strName:="temp"
For Each mycell In ActiveWorkbook.Worksheets(shtAfter).UsedRange
    If Not mycell.value = ActiveWorkbook.Worksheets(shtBefore).Cells(mycell.Row, mycell.Column).value Then
        'For xRow = 1 To mycell.Rows.Count
        '    For xCol = 1 To mycell.Columns.Count
        '        xStr = xStr & mycell.Cells(xRow, xCol).value & vbTab
        '    Next
        '    'xStr = xStr & mycell.Cells(xRow, 0).value & vbTab
        '    xStr = xStr & vbCrLf
        'Next
        ' Mark in yellow the diffs
        mycell.EntireRow.Copy Sheets("temp").Range("A" & Rows.Count).End(xlUp).Offset(1, 0)
        'mycell.Interior.Color = vbYellow
        mydiffs = mydiffs + 1
    End If
Next
rmdup strWS:="temp"
compareSheets = mydiffs
End Function
Sub ThreatModel()
' Main App - Create TM based on the information provided in the Profiler form
Dim strWS As String, strWS1 As String, strFrm As String, strPattern As String
Dim strRow As String, strForm As String, strDC As String, strApp As String
Dim sameCnt As Integer
Dim mydiffs As Variant
Dim MyZeroBasedArray As Variant

strRow = "D21"
strApp = "D12"
strDC = "D18"
strForm = "Profiler"
sameCnt = 1
' Pattern based on the worksheets name format as "SYS - Classification version"
strPattern = "^[A-Z]{3}[ ][-][ ][a-zA-Z]*"
' Get Values for Name
strWS = Sheets(strForm).Range(strApp).value + " - " + Sheets(strForm).Range(strDC).value

'Create Worksheet and fill it in
If Not WorksheetExists(strWS) Then
    'MsgBox strWS + " Does Not Exist and Empty"
    CreateWS strName:=strWS
    TM2 strWS:=strWS, strForm:=strForm, strRow:=strRow
    RecTMDB strWS:=strWS
Else
    If WorksheetExists(strWS) And Not CheckSheetEmpty(strWS) Then
        'MsgBox strWS + " Exist and Not Empty"
        'Chk if WS exist based on the 'strPattern'
        sameCnt = CountExist(strWS, strPattern)
        If sameCnt > 0 Then
            'If WS exist give it a version number
            strWS1 = strWS + " v" & sameCnt
            'MsgBox strWS
            CreateWS strName:=strWS1
            TM2 strWS:=strWS1, strForm:=strForm, strRow:=strRow
            RecTMDB strWS:=strWS1
            mydiffs = chkAbuseCase(strWS, strPattern)
            'Chk if there are differences between vA and vB of the WSz
            If mydiffs = 0 Then
                'No diff, delete vB
                MsgBox "No additional Abuse Cases have been identified.", vbInformation
                Application.DisplayAlerts = False
                    If WorksheetExists(strWS1) Then Sheets(strWS1).Delete
                    If WorksheetExists("temp") Then Sheets("temp").Delete
                Application.DisplayAlerts = True
                'Go back to 'strForm'
                ActiveWorkbook.Sheets(strForm).Select
            Else
                rmdup strWS:="temp"
                With UserForm1
                    .Label1 = Worksheets("temp").Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row - 2 & " Abuse Cases were identified."
                    .Show
                End With

                'Go back to 'strForm'
                ActiveWorkbook.Sheets(strForm).Select
            End If
        End If
    End If
End If

If WorksheetExists(strWS) And CheckSheetEmpty(strWS) Then
    MsgBox strWS + " Exist and Empty"
    TM2 strWS:=strWS, strForm:=strForm, strRow:=strRow
    RecTMDB strWS:=strWS
End If

End Sub
Sub TM2(ByVal strWS As String, strForm As String, strRow As String)
Dim strValues As String, strVal As String
Dim arrValues As Variant
Dim strCell As String
Dim i As Integer, J As Integer, k As Integer

strCell = Sheets(strForm).Range(strRow).value

' Read values from Profiler
arrValues = Split(strCell, ", ")
For i = 0 To UBound(arrValues)
    J = InStr(1, arrValues(i), "(")
    k = InStrRev(arrValues(i), ")")
    strVal = Trim(Mid(arrValues(i), J + 1, (k - J) - 1))
    If WorksheetExists(strVal) Then
        With ThisWorkbook.Sheets(strVal)
            'MsgBox "Review the Threat Model for " + strVal + " with Security!"
            AddTM2WS strDest:=strWS, strSource:=strVal
        End With
    Else
        'MsgBox "INFO: The Threat Model for " + strVal + " does not exist!"
        GoTo ContinueForLoop
    End If
ContinueForLoop:
Next i
End Sub
Sub SWThreatModel()

Call RecAPPDB

MsgBox "Review the Threat Model with Security!"
    
End Sub
Sub RecTMDB(ByVal strWS As String)
'Copy (In this case I want to copy range D4:D7 only, and this will be the same every time)
ThisWorkbook.Sheets("Profiler").Range("D9, D12, D15, D18, D21, D24, D27, D30, D33, D36").Copy

'Open Workbook 2 and paste data (transposed) on first available row starting in column B
With ThisWorkbook.Sheets("TMDB")
    ' find last row with data in destination workbook "wbDatabase.xlsm"
    DestLastRow = .Cells(.Rows.Count, "A").End(xlUp).Offset(1).Row
    'Create Date
    .Range("A" & DestLastRow).value = Now
    'Name
    .Range("B" & DestLastRow).value = strWS
     'paste special only values, and transpose
    .Range("C" & DestLastRow).PasteSpecial xlValues, Transpose:=True
End With
Cells(1, 1).Select
Application.CutCopyMode = False
End Sub
Sub RecAPPDB()
'Copy (In this case I want to copy range D4:D7 only, and this will be the same every time)
ThisWorkbook.Sheets("SW Profiler").Range("D9,D12,D15,D18,D21,D24,D27,D30,D33,D36,D39,D42,D45").Copy

'Open Workbook 2 and paste data (transposed) on first available row starting in column B
With ThisWorkbook.Sheets("APPDB")
    ' find last row with data in destination workbook "wbDatabase.xlsm"
    DestLastRow = .Cells(.Rows.Count, "A").End(xlUp).Offset(1).Row
    'Create Date
    .Range("A" & DestLastRow).value = Now
     'paste special only values, and transpose
    .Range("B" & DestLastRow).PasteSpecial xlValues, Transpose:=True
End With
Cells(1, 1).Select
Application.CutCopyMode = False
End Sub
Sub ClrProfiler()
'
' Clear the contents in the cells in a specified range
'
Sheets("Profiler").Range("D9,D12,D15,D18,D21,D24,D27,D30,D33,D36").ClearContents
End Sub
Sub ClrSWProfiler()
'
' Clear the contents in the cells in a specified range
'
Sheets("SW Profiler").Range("D9,D12,D15,D18,D21,D24,D27,D30,D33,D36,D39,D42,D45").ClearContents
End Sub
Private Sub TstCreateTM()
Dim myArr As Variant
Dim a As Integer
Dim strWSName As String, strTemplate As String

strTemplate = "TM Template"
myArr = Array("MFT", "MQ", "EDI", "ETL", "CRM", "EMS", "API", "MOM", "EBS", "JMS")

With Application
    .ScreenUpdating = False
    .EnableEvents = False
End With
For a = 0 To UBound(myArr)
strWSName = myArr(a)
Application.DisplayAlerts = False
    If WorksheetExists2(strWSName) Then Sheets(strWSName).Delete
Application.DisplayAlerts = True
CreateWS strName:=strWSName
    If WorksheetExists2(strWSName) Then
        If IsEmpty(Range("A1").value) = True Then
            'MsgBox "Review the Threat Model for " + strWSName + " with Security!"
            AddTM2WS strDest:=strWSName, strSource:=strTemplate
            Worksheets(strWSName).Range("A1").Formula = "=(MID(CELL(""filename"",A1),FIND(""]"",CELL(""filename"",A1))+1,255)) & "" Threat Model"""
            Worksheets(strWSName).Visible = False
        Else
            GoTo ContinueForLoop
        End If
    End If
With Application
    .ScreenUpdating = True
    .EnableEvents = True
End With
ContinueForLoop:
Next a
End Sub
Private Sub TsThreatModel()
' Author: Robert A. Vance, Jr.
' Version: 1.0
' Date: August 1, 2012
'
' Create a copy of the requirements for the user to follow based on the data classification
'
Dim wkSht As Worksheet, ws As Worksheet
Dim strWrkSht As String
Dim strDC As String
Dim strDP As String
Dim strWS As String
Dim strFileName As String, strFN
Dim LResult As String

' Get Values for Name
strDP = Sheets("Profiler").Range("D12").value
strDC = Sheets("Profiler").Range("D18").value
strWS = strDP + " - " + strDC

Call RecTMDB(strWS)

' For DEMO purpose
' First Delete the Worksheet if it exists
Application.DisplayAlerts = False
On Error Resume Next
ThisWorkbook.Sheets(strWS).Delete
On Error GoTo 0
Application.DisplayAlerts = True

If WorksheetExists2(strDC) Then
    With ThisWorkbook.Sheets(strDC)
    
        .Copy After:=Sheets("Lists")
        ' Rename copy
        ThisWorkbook.Sheets(strDC + " (2)").Name = strWS
        ' Make it visible
        Sheets(strWS).Visible = True
        ' Tell the user about it
        MsgBox "Review the Threat Model for " + strDP + " with Security!"
    
    End With
Else
    MsgBox "INFO: The Threat Model for " + strDC + " does not exist!"
End If

End Sub
