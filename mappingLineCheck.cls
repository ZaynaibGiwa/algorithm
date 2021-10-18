VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mappingLineCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public dMapLine As Object
Public dDataLine As Object
Public MatchType As MatchType
Public Value As String
Private p_countPerfectMatch As Integer
Private p_countGenericMatch As Integer
Private p_countNoMatch As Integer


Private Sub Class_Initialize()
    Set dMapLine = CreateObject("Scripting.Dictionary")
    Set dDataLine = CreateObject("Scripting.Dictionary")
End Sub


Sub CompareLines(mapLine As Object, dataLine As Object, checkType As String)

Dim key As Variant, mapValue As String, dataValue As String, localMatchType As MatchType

For Each key In mapLine.keys()

    mapValue = mapLine(key)
    dataValue = dataLine(key)
    
    If CStr(key) = checkType Then
        Value = CStr(key)
        localMatchType = eNoMatch
        'GoTo ExitSub
    ElseIf mapValue = dataValue Then
        localMatchType = ePerfect
        p_countPerfectMatch = p_countPerfectMatch + 1
    ElseIf InStr(1, mapValue, Chr(42)) > 0 Then
        localMatchType = eGeneric
        p_countGenericMatch = p_countGenericMatch + 1
    Else
        localMatchType = eNoMatch
    End If

    MatchType = WorksheetFunction.Max(MatchType, localMatchType)

Next key
ExitSub:
End Sub

Property Get MatchCount(MatchType As MatchType) As Integer

Select Case MatchType
    Case ePerfect
        MatchCount = p_countPerfectMatch
    Case eGeneric
        MatchCount = p_countGenericMatch
    Case eNoMatch
        MatchCount = p_countNoMatch
End Select

End Property
