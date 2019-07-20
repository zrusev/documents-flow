Attribute VB_Name = "Component_Generator"
Option Explicit
Option Private Module

Function GenerateForm(ByRef colFiles As Collection) As Variant

    Dim previousCellKey As String
    Dim previousCellValue As String
    Dim activeCellKey As String
    Dim activeCellValue As String
    Dim arr As Variant
    Dim coll As Integer
    Dim y As Integer

    LoadCollection
    LoadForm

    For y = 1 To colFiles.Count
        arr = Split(MASTER_FOLDER & Application.PathSeparator & Replace(colFiles(y), TrailingSlash(MASTER_DIRECTORY), "", , , vbTextCompare), Application.PathSeparator)
        
        previousCellKey = ""
        previousCellValue = ""
        activeCellKey = ""
        activeCellValue = ""
        
        For coll = LBound(arr) + 1 To UBound(arr)
            previousCellKey = previousCellKey & Application.PathSeparator & arr(coll - 1)
            previousCellValue = arr(coll - 1)
            activeCellKey = previousCellKey & Application.PathSeparator & arr(coll)
            activeCellValue = arr(coll)

            If activeCellValue <> "" Then
                UserForm1.AddToNode activeCellKey, activeCellValue, previousCellKey, previousCellValue
            End If
        Next coll
    Next y

End Function

Function GenerateObject(ByRef strObjectType As String, _
        ByVal strObjectlocation As String, _
        ByVal strObjectName As String, _
        ByVal strIsVisibleProp As String, _
        ByVal strCaptureText As String, _
        ByVal dblPositionHeight As Double, _
        ByVal dblPositionWidth As Double, _
        ByVal dblPositionTop As Double, _
        ByVal dblPositionLeft As Double, _
        ByVal strFieldName As String) As IObject

    Dim obj As IObject: Set obj = New clsObject
    
    obj.CreateObject strObjectType, strObjectlocation, strObjectName, _
        dblPositionHeight, dblPositionWidth, dblPositionTop, dblPositionLeft

    obj.SetObject strObjectType, strObjectlocation, strObjectName, strIsVisibleProp, strCaptureText, _
        dblPositionHeight, dblPositionWidth, dblPositionTop, dblPositionLeft, strFieldName

    obj.SetText strCaptureText
    
    Set GenerateObject = obj
    
End Function

Function GenerateShape(ByVal shName As String, ByVal rowNumber As Long, ByVal field As String) As IObject

    Dim clLeft As Double
    Dim clTop As Double
    Dim clWidth As Double
    Dim clHeight As Double
    
    ThisWorkbook.Sheets(shName).Range("C" & rowNumber).Select
    Dim cl As Range: Set cl = ThisWorkbook.Sheets(shName).Range(Selection.Address)
    
    clLeft = cl.Left
    clTop = cl.Top
    clHeight = cl.Height
    clWidth = cl.Width
    
    Set GenerateShape = GenerateObject("Shape", shName, field, "Yes", field, 30, 250, clTop + 100, clLeft, "")

End Function
