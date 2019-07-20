Attribute VB_Name = "Component_Execution"
Option Explicit

Sub ShowForm()

    Dim colFiles As New Collection
    
    RecursiveDir colFiles, MASTER_DIRECTORY, FILE_EXTENSION, True
    
    GenerateForm colFiles

End Sub

Sub ShowFields()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    DesctructAll
      
    Dim dict As New Scripting.Dictionary
    Dim oWord As Word.Application: Set oWord = New Word.Application
    
    Dim rowNumber As Long: rowNumber = 5
    Dim file As IFile
    For Each file In filesCollection.files
        Dim Path As String: Path = file.Path
        Dim oWdoc As Word.document: Set oWdoc = oWord.documents.Open(Path, ReadOnly:=True)
    
        Dim matches As Variant
        Set matches = GetMatches(oWdoc)
        
        oWdoc.Close SaveChanges:=wdDoNotSaveChanges
        Set oWdoc = Nothing

        If Not matches Is Nothing Then
            Dim match As Object
            For Each match In matches
                If Not dict.Exists(match.Value) Then
                    objectsCollection.Add = GenerateShape(DASHBOARD_SHEET, rowNumber, match.Value)
                    rowNumber = rowNumber + 5
                    dict.Add match.Value, match.Value
                End If
            Next match
        End If
    Next file
    
    oWord.Quit
    Set oWord = Nothing

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Sub ShowDocument()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Dim userInput As New Scripting.Dictionary: Set userInput = GetInput

    Dim oWord As Word.Application: Set oWord = New Word.Application
    
    Dim mergedDoc As Word.document
    Set mergedDoc = MergeDocuments(oWord, filesCollection)

    Dim key As Variant
    For Each key In userInput
        FindReplaceAnywhere mergedDoc, key, userInput(key)
    Next key
    
    SaveDocument mergedDoc, SAVE_AS_NAME, GetDesktop
    
    mergedDoc.Close

    Set mergedDoc = Nothing
    
    oWord.Quit
    Set oWord = Nothing

    DesctructAll

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
