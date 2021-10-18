VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_cMapLines As Collection
Private p_itemName As String
Private p_BestMatchLine As Object

Sub SetMapping(itemName As String)

Dim mapRange As Range
Set mapRange = WS_MAPPING.Range("mapping_" & itemName)
Set mapRange = mapRange.Resize(, mapRange.Columns.Count - 1)

Set p_cMapLines = RangeToCollection(mapRange)
p_itemName = itemName

End Sub

Property Get GetMappingValue(dDataLine As Object) As String

Dim mapLineCheck As mappingLineCheck
Dim dMappingLine As Object
Dim maxPerfectMatch As Integer
Dim maxGenericMatch As Integer

For Each dMappingLine In p_cMapLines
    Set mapLineCheck = New mappingLineCheck
    mapLineCheck.CompareLines dMappingLine, dDataLine, p_itemName
    
    If mapLineCheck.MatchCount(ePerfect) > maxPerfectMatch Then
        Set p_BestMatchLine = dMappingLine
        maxPerfectMatch = mapLineCheck.MatchCount(ePerfect)
        
    ElseIf mapLineCheck.MatchCount(ePerfect) = maxPerfectMatch Then
        If mapLineCheck.MatchCount(eGeneric) > maxGenericMatch Then
            Set p_BestMatchLine = dMappingLine
            maxGenericMatch = mapLineCheck.MatchCount(eGeneric)
        End If
    End If
    
Next dMappingLine

GetMappingValue = p_BestMatchLine(p_itemName)

End Property
