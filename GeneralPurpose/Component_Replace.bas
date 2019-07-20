Attribute VB_Name = "Component_Replace"
Option Explicit
Option Private Module

Private Sub SearchAndReplaceInStory(ByVal rngStory As Word.Range, _
                                   ByVal strSearch As String, _
                                   ByVal strReplace As String)

    With rngStory.Find
      .ClearFormatting
      .Replacement.ClearFormatting
      .text = strSearch
      .Replacement.text = strReplace
      .Wrap = wdFindContinue
      .Execute Replace:=wdReplaceAll
    End With
    
End Sub

Sub FindReplaceAnywhere(ByRef doc As Word.document, ByVal strFind As String, ByVal strRplc As String)

    Dim rngStory As Word.Range
    Dim lngValidate As Long
    Dim oShp As Shape

    'Fix the skipped blank Header/Footer problem.
    lngValidate = doc.Sections(1).headers(1).Range.StoryType
    'Iterate through all story types in the current document.
    For Each rngStory In doc.StoryRanges
      'Iterate through all linked stories.
      Do
        SearchAndReplaceInStory rngStory, strFind, strRplc
        On Error Resume Next
        Select Case rngStory.StoryType
          Case 6, 7, 8, 9, 10, 11
            If rngStory.ShapeRange.Count > 0 Then
              For Each oShp In rngStory.ShapeRange
                If oShp.TextFrame.HasText Then
                  SearchAndReplaceInStory oShp.TextFrame.TextRange, strFind, strRplc
                End If
              Next
            End If
          Case Else
            'Do Nothing
        End Select
        On Error GoTo 0
        'Get next linked story (if any)
        Set rngStory = rngStory.NextStoryRange
      Loop Until rngStory Is Nothing
    Next

End Sub

