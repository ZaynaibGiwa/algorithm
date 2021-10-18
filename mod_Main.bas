Attribute VB_Name = "mod_Main"
Option Explicit

Sub Main()

Dim cData As Collection
Set cData = RangeToCollection(WS_DATA.Range("A1:I49"))

GenerateBreaks cData, "RISKCLASS"
GenerateBreaks cData, "RISKWEIGHT"

End Sub

Sub GenerateBreaks(cData As Collection, itemName As String)

Dim dDataLine As Object
Dim mappingValue As String
Dim itemMapping As mapping

Set itemMapping = New mapping
itemMapping.SetMapping itemName

For Each dDataLine In cData
    If dDataLine(itemName) = itemName Then
        GoTo next_dDataLine
    End If
    mappingValue = itemMapping.GetMappingValue(dDataLine)
    
    If mappingValue <> dDataLine(itemName) Then
        GenerateBreak itemName, dDataLine, mappingValue
    End If
next_dDataLine:
Next dDataLine

End Sub



Sub GenerateBreak(itemName As String, dDataLine As Object, mappingValue As String)

Dim wsBreak As Worksheet, lastRow As Long, includeHeader As Boolean

dDataLine.Add itemName & " MAPPING", mappingValue

Set wsBreak = setWorksheet(itemName)
lastRow = GetLastRow(wsBreak.Range("A1"))

If lastRow = 1 Then
    includeHeader = True
Else
    lastRow = lastRow + 1
End If

printDictionaryToRange dDataLine, wsBreak.Cells(lastRow, 1), includeHeader

dDataLine.Remove itemName & " MAPPING"

End Sub

Private Function setWorksheet(itemName As String) As Worksheet

Dim tempSheet As Worksheet

On Error Resume Next
    Set tempSheet = ThisWorkbook.Sheets(itemName)
On Error GoTo 0

If tempSheet Is Nothing Then
    Set tempSheet = ThisWorkbook.Worksheets.Add
    tempSheet.Name = itemName
End If

Set setWorksheet = tempSheet

End Function


Private Sub printDictionaryToRange(dict As Object, startCell As Range, includeHeader As Boolean)

Dim key As Variant, rowOffset As Integer, colOffset As Long

If includeHeader Then
    For Each key In dict.keys()
        startCell.Offset(rowOffset, colOffset).Value = CStr(key)
        colOffset = colOffset + 1
    Next key
    rowOffset = rowOffset + 1
End If

colOffset = 0
For Each key In dict.keys()
    startCell.Offset(rowOffset, colOffset).Value = dict(key)
    colOffset = colOffset + 1
Next key

End Sub

