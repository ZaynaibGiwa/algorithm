Attribute VB_Name = "mod_GenericFunctions"
Option Explicit

Function GetLastRow(rng As Range, Optional columnsCount As Long) As Long

Dim iRow As Long, iCol As Long, curLastRow As Long, finalLastRow As Long, curWs As Worksheet, curCellValue As String

Set curWs = rng.Worksheet
If columnsCount < 1 Then columnsCount = rng.Columns.Count

For iCol = 1 To columnsCount
    iRow = curWs.Cells(1048576, rng.Column).Offset(0, iCol - 1).End(xlUp).Row
    curCellValue = GetSafeCellValue(curWs.Cells(iRow, rng.Column).Offset(0, iCol - 1), True)
    
    Do Until curCellValue <> ""
        iRow = iRow - 1
        If iRow < 1 Then GoTo next_iCol
        curCellValue = GetSafeCellValue(curWs.Cells(iRow, rng.Column).Offset(0, iCol - 1), True)
    Loop
    
    curLastRow = iRow
    finalLastRow = WorksheetFunction.Max(curLastRow, finalLastRow)
next_iCol:
Next iCol

If finalLastRow = 0 Then finalLastRow = 1
GetLastRow = finalLastRow

End Function
Function GetSafeCellValue(ByVal rng As Range, Optional OnErrorReturn_ERROR_ As Boolean) As String

Dim tempStr As String

Set rng = rng.Resize(1, 1)

tempStr = "tempCellValue_xxxxxx"
On Error Resume Next
    tempStr = rng.Value
On Error GoTo 0

If tempStr = "tempCellValue_xxxxxx" Then
    If OnErrorReturn_ERROR_ Then
        GetSafeCellValue = "_ERROR_"
    Else
        GetSafeCellValue = ""
    End If
Else
    GetSafeCellValue = tempStr
End If

End Function


Function RangeToCollection(rng As Range) As Collection

Dim rowsCount As Long, colsCount As Long, iRow As Long, iCol As Long, tempCollection As Collection, tempDictionary As Object, header As String, cellValue As String

Set tempCollection = New Collection
rowsCount = GetLastRow(rng) - rng.Row + 1
colsCount = rng.Columns.Count

For iRow = 1 To rowsCount
    Set tempDictionary = Nothing
    Set tempDictionary = CreateObject("Scripting.Dictionary")
    
    For iCol = 1 To colsCount
        header = rng.Cells(1, iCol).Value
        cellValue = GetSafeCellValue(rng.Cells(iRow, iCol))
        tempDictionary.Add header, cellValue
    Next iCol
    
    tempCollection.Add tempDictionary, CStr(iRow)
    
Next iRow

Set RangeToCollection = tempCollection

End Function
