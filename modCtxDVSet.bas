Attribute VB_Name = "modCtxDVSet"
Option Explicit
'Contextures DVMSP Settings
Global strDVList As String
Global strSep As String
Global lMS As Long
Global bNew As Boolean
Global bPos As Boolean
Global bDown As Boolean
Global bAcross As Boolean
Global bFromNew As Boolean
Global strNew As String

Sub DVMSP_RefStyle()
Attribute DVMSP_RefStyle.VB_Description = "Toggle Ref Style -- change column headings to numbers or letters"
Attribute DVMSP_RefStyle.VB_ProcData.VB_Invoke_Func = "R\n14"
'change column headings to
'numbers or letters
'shortcut is Ctrl + Shift + R
On Error Resume Next
Dim strType As String
With Application
    If .ReferenceStyle = xlA1 Then
      .ReferenceStyle = xlR1C1
      strType = "Numbers"
    Else
      .ReferenceStyle = xlA1
      strType = "Letters"
    End If
End With

MsgBox "Column headings were changed to " & strType, _
  vbInformation + vbOKOnly, "Column Headings"

End Sub

