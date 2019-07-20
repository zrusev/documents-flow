Attribute VB_Name = "Component_Document"
Option Explicit
Option Private Module

Function MergeDocuments(ByRef oWord As Word.Application, _
                        ByRef files As IFiles) As Word.document

  Dim objDoc As Word.document, objNewDoc As Word.document

  Dim firstFile As IFile: Set firstFile = files.files(1)
  Set objNewDoc = oWord.documents.Open(firstFile.Path)

  Dim isFirstDocument As Boolean: isFirstDocument = True
  Dim i As Long
  For i = 2 To files.files.Count
    
    Set objDoc = oWord.documents.Open(files.files(i).Path)
    
    objDoc.Range.Copy
    objNewDoc.Activate

    With oWord.Selection
      .EndKey Unit:=wdStory
        If isFirstDocument Then
            .InsertBreak Break_Types.wdSectionBreakNextPage
            .Collapse wdCollapseEnd
            isFirstDocument = False
        End If
        Wait 500
      .Paste
      .Collapse wdCollapseEnd
      .InsertBreak Type:=Break_Types.wdSectionBreakContinuous
    End With
 
    objDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set objDoc = Nothing
  Next i
 
  objNewDoc.Activate
  oWord.Selection.EndKey Unit:=wdStory
  oWord.Selection.Delete

  Set MergeDocuments = objNewDoc
  Set objNewDoc = Nothing
  
End Function

Sub SaveDocument(ByRef doc As Word.document, ByVal strName As String, ByVal strPath As String)

    doc.SaveAs2 strPath & Application.PathSeparator & strName, File_Formats.wdFormatDocument
    
End Sub
