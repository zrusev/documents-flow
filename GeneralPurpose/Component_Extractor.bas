Attribute VB_Name = "Component_Extractor"
Option Explicit
Option Private Module

Function GetMatches(ByRef oWdoc As Word.document) As Variant: Set GetMatches = IsMatch(oWdoc.Range.text, FIND_ALL_PATTERN): End Function

Function GetInput() As Scripting.Dictionary
    Dim dict As New Scripting.Dictionary

    Dim member As IObject
    For Each member In objectsCollection.Members
        Dim field As String: field = member.ObjectName
        Dim userInput As String: userInput = member.GetText
        If Not dict.Exists(field) Then dict.Add field, userInput
    Next member
    
    Set GetInput = dict
End Function
