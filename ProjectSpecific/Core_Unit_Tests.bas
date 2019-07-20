Attribute VB_Name = "Core_Unit_Tests"
Option Explicit
Option Private Module
'Test_[Component]_[Method]_[Expected_Result]

Private Declare Function GetTickCount Lib "kernel32" () As Long

Sub Test_Component_Match_IsMatch_Should_Not_Return_Nothing()

    Dim text As String
    text = "aaaaa{{document.footer}}bbbbbbbb{{document.main}}ccccccccc{{document.header}}"
    
    Dim results As Object
    Set results = IsMatch(text, FIND_ALL_PATTERN)
    
    #If Debugging Then
        Debug.Assert Not results Is Nothing
    #End If

End Sub

Sub Test_Component_Match_IsMatch_Should_Not_Return_Three_Elements()

    Dim text As String
    text = "aaaaa{{document.footer}}bbbbbbbb{{document.main}}ccccccccc{{document.header}}"
    
    Dim results As Variant
    Set results = IsMatch(text, FIND_ALL_PATTERN)
    
    #If Debugging Then
        Debug.Assert results.Count = 3
    #End If

End Sub

Sub Test_Component_Replace_FindReplaceAnywhere_Should_Replace_Values_Within_Document()

    Dim pattern As String: pattern = "SOMETHING"

    Dim oWord As Word.Application: Set oWord = New Word.Application
    oWord.Visible = True
    
    Dim oWdoc As Word.document: Set oWdoc = oWord.documents.Open(RANDOM_DOC_NAME)

    FindReplaceAnywhere oWdoc, "{{document.header}}", "VALUE"
    
    Dim results As Object
    Set results = IsMatch(oWdoc.Range.text, pattern)

    #If Debugging Then
        Debug.Assert Not results Is Nothing
    #End If
    
End Sub

Sub Test_Component_Replace_FindReplaceAnywhere_Should_Replace_All_Passed_Values()

    Dim oWord As Word.Application: Set oWord = New Word.Application
    
    #If Debugging Then
        oWord.Visible = True
    #End If
    
    Dim oWdoc As Word.document: Set oWdoc = oWord.documents.Open(RANDOM_DOC_NAME, ReadOnly:=True)

    Dim matches As Variant
    Set matches = IsMatch(oWdoc.Range.text, FIND_ALL_PATTERN)

    Dim objMatch As Object
    For Each objMatch In matches
        FindReplaceAnywhere oWdoc, objMatch.Value, "VALUE"
    Next objMatch

    #If Debugging Then
        oWdoc.Close SaveChanges:=wdDoNotSaveChanges
    #Else
        oWdoc.Close
    #End If
    
    Set oWdoc = Nothing
       
    oWord.Quit
    Set oWord = Nothing
    
End Sub

Sub Test_Component_Document_SaveDocument_Should_Save_File()

    Dim oWord As Word.Application: Set oWord = New Word.Application
    
    #If Debugging Then
        oWord.Visible = True
    #End If
    
    Dim oWdoc As Word.document: Set oWdoc = oWord.documents.Open(RANDOM_DOC_NAME, ReadOnly:=True)

    SaveDocument oWdoc, "test", GetDesktop
    
    #If Debugging Then
        oWdoc.Close SaveChanges:=wdDoNotSaveChanges
    #Else
        oWdoc.Close
    #End If
    
    Set oWdoc = Nothing
       
    oWord.Quit
    Set oWord = Nothing
    
End Sub

' ToDo to refactor
Sub Test_Component_Document_MergeDocuments_Should_Combine_Multiple_Files_To_Five_Pages()

    Dim oWord As Word.Application
    Set oWord = New Word.Application
    
    #If Debugging Then
        oWord.Visible = False
    #End If
    
    Dim dict As New Scripting.Dictionary
    Dim i As Long
    For i = LBound(GetDocumentsKVP) To UBound(GetDocumentsKVP)
        If Not dict.Exists(getGetDocumentsKVP(i)(0)) Then
            dict.Add getGetDocumentsKVP(i)(0), getGetDocumentsKVP(i)(1)
        End If
    Next i
    
    Dim mergedDoc As Word.document
    Set mergedDoc = MergeDocuments(oWord, dict)
    
    #If Debugging Then
        Debug.Assert mergedDoc.BuiltinDocumentProperties(wdPropertyPages) = 5
        mergedDoc.Close SaveChanges:=wdDoNotSaveChanges

    #Else
        mergedDoc.Close
    #End If

    Set mergedDoc = Nothing
    
    oWord.Quit
    Set oWord = Nothing
    
End Sub

Sub Test_a_Functions_Wait_Should_Wait_For_500_Miliseconds()
    
    Dim t1 As Long, t2 As Long
    
    t1 = GetTickCount
    Wait 500
    t2 = GetTickCount
    
    #If Debugging Then
        Debug.Assert t2 - t1 = 500
    #End If

End Sub

Sub Test_Component_Recursion_Should_Return_All_Directories()
    
    Dim colFiles As New Collection
    
    RecursiveDir colFiles, MASTER_DIRECTORY, FILE_EXTENSION, True

    #If Debugging Then
        Debug.Assert colFiles.Count = 14
    #End If
End Sub

Sub Test_Component_Recursion_Should_Extract_From_Collection()

    Dim colFiles As New Collection
    
    RecursiveDir colFiles, MASTER_DIRECTORY, FILE_EXTENSION, True
    
    GenerateForm colFiles

    #If Debugging Then
        Debug.Assert Not UserForm1 Is Nothing
    #End If

End Sub

Sub Test_clsSelectedFiles_Should_Return_Collection_Files()

    Dim file As New clsFile
    file.IFile_Name = "master"
    file.IFile_Path = MASTER_DIRECTORY
        
    Dim files As New clsFiles
    files.IFiles_Add = file
    
    #If Debugging Then
        Debug.Assert Not files.IFiles_Files Is Nothing
    #End If
    
End Sub

Sub Test_Component_Loader_FilesCollection_Should_Be_Set_Globally()

    LoadCollection
    
    #If Debugging Then
        Debug.Assert Not filesCollection.files Is Nothing
    #End If
    
End Sub

Sub Test_Component_Loader_Should_Return_All_Files_From_FilesCollection()

    'Call Test_Component_Recursion_Should_Extract_From_Collection
    
    Dim f As IFile
    For Each f In filesCollection.files
        #If Debugging Then
            Debug.Assert f.Name <> ""
            Debug.Assert f.Path <> ""
        #End If
    Next f
    
End Sub

Sub Test_Component_Generator_Should_Generate_Shape()
   
    Dim obj As New clsObject
    Set obj = GenerateShape(DASHBOARD_SHEET, 5, "{{document.header}}")
    
    Dim sh As Shape: Set sh = obj.GetObject
    
    #If Debugging Then
        Debug.Assert VarType(sh) = VarTypes.vbObject
    #End If
        
    #If Debugging Then
        Debug.Assert DesctructShape(obj) = True
    #End If
        
End Sub

Sub Test_clsObjects_Should_Initialize_Collection()

    Dim col As IObjects
    Set col = New clsObjects
    
    col.Add = New clsObject

    #If Debugging Then
        Debug.Assert Not col.Members Is Nothing
    #End If
    
End Sub

Sub Test_Component_Generator_Should_Generate_Object_From_Collection()

    LoadCollection
    
    Dim mocks As Variant
    mocks = GetDocumentsKVP
    
    Dim i As Long
    For i = LBound(mocks) To UBound(mocks)
        Dim file As IFile: Set file = New clsFile
        file.Name = mocks(i)(0)
        file.Path = mocks(i)(1)
        filesCollection.Add = file
    Next i
 
    Dim rowNumber As Long: rowNumber = 5
    Dim obj As Object
    For Each obj In filesCollection.files
        objectsCollection.Add = GenerateShape(DASHBOARD_SHEET, rowNumber, obj.Name)
        rowNumber = rowNumber + 5
    Next
    
    #If Debugging Then
        Debug.Assert ThisWorkbook.Sheets(DASHBOARD_SHEET).Shapes.Count = 4
    #End If

End Sub

Sub Test_Component_Generator_Should_Destruct_Objects_From_Collection()

    Dim obj As Object
    For Each obj In objectsCollection.Members
        #If Debugging Then
            Debug.Assert DesctructShape(obj) = True
        #End If
    Next
    
    #If Debugging Then
        Debug.Assert ThisWorkbook.Sheets(DASHBOARD_SHEET).Shapes.Count = 0
    #End If
    
End Sub

Sub Test_Component_Extractor_GetInput_Should_Return_Inputs()

    Dim dict As New Scripting.Dictionary
    Set dict = GetInput
    
    #If Debugging Then
        Debug.Assert dict.Count = 4
    #End If

End Sub
