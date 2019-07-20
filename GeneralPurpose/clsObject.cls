VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IObject

Private oType As ObjectTypes
Private objLoc As String
Private ObjName As String
Private isVisProp As CustomOperators
Private objCapt As String
Private objHeight As Double
Private objWidth As Double
Private objTop As Double
Private objLeft As Double
Private objField As String
Private sh As Shape
Private txtB As TextBox
Private sl As Slicer
'Private tbl As ITable
Private pict As Shape
Private mCells As Range


Private Property Get MergedCells() As Range: Set MergedCells = mCells: End Property
Private Property Let MergedCells(ByVal mArea As Range): Set mCells = mArea: End Property
Private Property Get Picture() As Shape: Set Picture = pict: End Property
Private Property Let Picture(ByVal ePic As Shape): Set pict = ePic: End Property
'Private Property Get AdoTable() As AdoDBTable: Set AdoTable = tbl: End Property
'Private Property Let AdoTable(ByVal t As AdoDBTable): Set tbl = t: End Property
Private Property Get Slicer() As Slicer: Set Slicer = sl: End Property
Private Property Let Slicer(ByVal oSl As Slicer): Set sl = oSl: End Property
Private Property Get TextBox() As TextBox: Set TextBox = txtB: End Property
Private Property Let TextBox(ByVal tBx As TextBox): Set txtB = tBx: End Property
Private Property Get Shape() As Shape: Set Shape = sh: End Property
Private Property Let Shape(ByVal eSh As Shape): Set sh = eSh: End Property
Public Property Get IObject_objectName() As String: IObject_objectName = ObjName: End Property
Private Property Let IObject_objectName(ByVal RHS As String): ObjName = RHS: End Property
Public Property Get IObject_objectLocation() As String: IObject_objectLocation = objLoc: End Property
Private Property Let IObject_objectLocation(ByVal RHS As String): objLoc = RHS: End Property
Public Property Get IObject_objectType() As ObjectTypes: IObject_objectType = oType: End Property
Private Property Let IObject_objectType(ByVal RHS As ObjectTypes): oType = RHS: End Property
Public Property Get IObject_isVisibleProperty() As CustomOperators: IObject_isVisibleProperty = isVisProp: End Property
Public Property Let IObject_isVisibleProperty(ByVal RHS As CustomOperators): isVisProp = RHS: End Property
Public Property Get IObject_objectCapture() As String: IObject_objectCapture = objCapt: End Property
Private Property Let IObject_objectCapture(ByVal RHS As String): objCapt = RHS: End Property
Public Property Get IObject_objectHeight() As Double: IObject_objectHeight = objHeight: End Property
Private Property Let IObject_objectHeight(ByVal RHS As Double): objHeight = RHS: End Property
Public Property Get IObject_objectWidth() As Double: IObject_objectWidth = objWidth: End Property
Private Property Let IObject_objectWidth(ByVal RHS As Double): objWidth = RHS: End Property
Public Property Get IObject_objectPositionLeft() As Double: IObject_objectPositionLeft = objLeft: End Property
Private Property Let IObject_objectPositionLeft(ByVal RHS As Double): objLeft = RHS: End Property
Public Property Get IObject_objectPositionTop() As Double: IObject_objectPositionTop = objTop: End Property
Private Property Let IObject_objectPositionTop(ByVal RHS As Double): objTop = RHS: End Property
Public Property Get IObject_objectFieldName() As String: IObject_objectFieldName = objField: End Property
Private Property Let IObject_objectFieldName(ByVal RHS As String): objField = RHS: End Property

Private Property Get IsTableObjectCreated() As Boolean
    
    If ThisWorkbook.Sheets(IObject_objectLocation).Evaluate("ISREF(" & IObject_objectName & ")") Then IsTableObjectCreated = True

End Property

Private Property Get IsMergedCellsObjectCreated() As Boolean
    
    If ThisWorkbook.Sheets(IObject_objectLocation).Cells(IObject_objectPositionTop, IObject_objectPositionLeft).MergeCells Then IsMergedCellsObjectCreated = True

End Property

Public Property Get GetObject() As Variant

    Select Case IObject_objectType
           Case 1, 4, 5
                Set GetObject = Shape
           Case 2
                Set GetObject = TextBox
           Case 3
                Set GetObject = Slicer
           Case 6
    '           Set GetObject = AdoTable.Table
           Case 7
               Set GetObject = Picture
           Case 8
               Set GetObject = MergedCells
    End Select
    
End Property

Public Property Get IObject_GetText() As String

    Select Case IObject_objectType
           Case 1, 5
                IObject_GetText = Shape.TextFrame.Characters.text
           Case 2
                IObject_GetText = TextBox.text
           Case 3
                IObject_GetText = Slicer.Caption
           Case 4
                With Shape.ControlFormat
                    IObject_GetText = .List(.Value)
                End With
           Case 6, 7, 8
    End Select
    
End Property

Private Sub SetObjectType()

    Select Case IObject_objectType
           Case 1, 4, 5
                Shape = ThisWorkbook.Sheets(IObject_objectLocation).Shapes(IObject_objectName)
                Shape.Fill.ForeColor.RGB = RGB(255, 192, 0)

                With Shape.Line
                    .Weight = 1
                    .ForeColor.RGB = RGB(255, 255, 255)
                End With
           Case 2
                TextBox = ThisWorkbook.Sheets(IObject_objectLocation).TextBoxes(IObject_objectName)
           Case 3
                Slicer = ActiveWorkbook.SlicerCaches(IObject_objectLocation).Slicers(IObject_objectName)
           Case 6
    '            AdoTable = New AdoDBTable
    '            If Not IsTableObjectCreated Then CreateTableObject
    '            AdoTable.ITable_SetObject IObject_objectLocation, IObject_objectName
           Case 7
                'If Not IsPictureObjectCreated(IObject_objectLocation, IObject_objectName) Then PopulateLastTransaction
                Picture = ThisWorkbook.Sheets(IObject_objectLocation).Shapes(IObject_objectName)
           Case 8
                If Not IsMergedCellsObjectCreated Then CreateMergedCellsObject
                MergedCells = ThisWorkbook.Sheets(IObject_objectLocation).Cells(IObject_objectPositionTop, IObject_objectPositionLeft).MergeArea
    End Select

End Sub

Private Sub IObject_IsVisible()

    Select Case IObject_objectType
           Case 1, 4, 5
                Shape.Visible = IObject_isVisibleProperty
           Case 2
                TextBox.Visible = IObject_isVisibleProperty
           Case 3
                Slicer.Shape.Visible = IObject_isVisibleProperty
           Case 6
    '            If Not IsTableObjectCreated Then CreateTableObject
    '            If IObject_isVisibleProperty = No Then DestroyTable
           Case 7
                Picture.Visible = IObject_isVisibleProperty
                'If Not IsPictureObjectCreated(IObject_objectLocation, IObject_objectName) Then PopulateLastTransaction
                'If IObject_isVisibleProperty = No Then DestroyPicture
           Case 8
                If Not IsMergedCellsObjectCreated Then CreateMergedCellsObject
                If IObject_isVisibleProperty = No Then DestroyMergedCells
    End Select
    
End Sub

Private Sub IObject_SetCaption()

    Select Case IObject_objectType
           Case 1
                Shape.TextFrame.Characters.text = IObject_objectCapture
           Case 2
                TextBox.Caption = IObject_objectCapture
           Case 3
                Slicer.Caption = IObject_objectCapture
           Case 4
                Dim captureListItems As Variant: captureListItems = Split(IObject_objectCapture, "|")
                Dim minLines As Integer: If CInt(UBound(captureListItems) / 2) > 8 Then minLines = CInt(UBound(captureListItems) / 2) Else minLines = 8
                With Shape
                    .ControlFormat.RemoveAllItems
                    .OLEFormat.Object.List = captureListItems 'raises an error if the list is empty
                    .ControlFormat.ListIndex = 1
                    .ControlFormat.DropDownLines = minLines
                End With
           Case 5
                'Form Controls do not have such property
           Case 6
                'Used for the table's headers
           Case 7
                'Not applicable
           Case 8
                If IObject_objectCapture <> "" Then SetMergedCellsValidation
    End Select
    
End Sub

Private Sub IObject_SetPosition()

    Select Case IObject_objectType
           Case 1, 4, 5
                Shape.Height = IObject_objectHeight
                Shape.Width = IObject_objectWidth
                Shape.Top = IObject_objectPositionTop
                Shape.Left = IObject_objectPositionLeft
           Case 2
                TextBox.Height = IObject_objectHeight
                TextBox.Width = IObject_objectWidth
                TextBox.Top = IObject_objectPositionTop
                TextBox.Left = IObject_objectPositionLeft
           Case 3
                If Slicer.DisableMoveResizeUI = True Then Slicer.DisableMoveResizeUI = False
                Slicer.Height = IObject_objectHeight
                Slicer.Width = IObject_objectWidth
                Slicer.Top = IObject_objectPositionTop
                Slicer.Left = IObject_objectPositionLeft
                Slicer.DisableMoveResizeUI = True
          Case 6
          Case 7
                Picture.Height = IObject_objectHeight
                Picture.Width = IObject_objectWidth
                Picture.Top = IObject_objectPositionTop
                Picture.Left = IObject_objectPositionLeft
          Case 8
                With MergedCells
                    .Style = ThisWorkbook.Styles.item(1)
                    .VerticalAlignment = xlCenter
                    .HorizontalAlignment = xlLeft
                    With .Borders(xlEdgeLeft)
                        .LineStyle = xlContinuous
                        .Color = vbWhite
                        .Weight = xlThin
                    End With
                    With .Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Color = vbWhite
                        .Weight = xlThin
                    End With
                End With
    End Select
    
End Sub

Private Sub CreateTableObject()

    Dim i As Integer
    Dim headersCollection As Variant
    Dim headers As String
    Dim wdth As Double
    Dim hght As Double
    
    Const tableStyle = "Custom 2"
    
    hght = IObject_objectHeight
    wdth = IObject_objectWidth
    headers = IObject_objectCapture
    headersCollection = Split(headers, "|")
    
    ThisWorkbook.Sheets(IObject_objectLocation) _
                .ListObjects _
                .Add(xlSrcRange, ThisWorkbook.Sheets(IObject_objectLocation).Cells(IObject_objectPositionTop, IObject_objectPositionLeft), , xlYes) _
                .Name = IObject_objectName
    
    With ThisWorkbook.Sheets(IObject_objectLocation).ListObjects(IObject_objectName)
        .tableStyle = tableStyle
        For i = LBound(headersCollection) To UBound(headersCollection) - 1
            .ListColumns.Add
            .ListColumns(.ListColumns.Count - 1).Name = headersCollection(i)
            .ListColumns(.ListColumns.Count - 1).Range.Columns.ColumnWidth = wdth
        Next i
            .ListColumns(.ListColumns.Count).Name = headersCollection(i)
            .ListColumns(.ListColumns.Count).Range.Columns.ColumnWidth = wdth
        .Range.Interior.ColorIndex = 0
        
        .ShowAutoFilter = False
    End With
    
End Sub

Private Sub CreateMergedCellsObject()

    ThisWorkbook.Sheets(IObject_objectLocation).Activate
    ThisWorkbook.Sheets(IObject_objectLocation).Range(Cells(IObject_objectPositionTop, IObject_objectPositionLeft), Cells(IObject_objectHeight, IObject_objectWidth)).Merge

End Sub

'Private Sub DestroyTable()
'Dim tempRange As Range
'
'Set tempRange = GetTableRange
'AdoTable.Table.Unlist
'AdoTable = Nothing
'
'With tempRange
'    .Value = ""
'    .Clear
'    .Interior.Color = RGB(217, 217, 217)
'End With
'
'Set tempRange = Nothing
'End Sub

Private Sub DestroyMergedCells()

    With MergedCells
        .Validation.Delete
        .UnMerge
        .Value = ""
        .Clear
        .Interior.Color = RGB(217, 217, 217)
    End With
    
End Sub

Private Sub SetMergedCellsValidation()

    With MergedCells
        .Interior.Color = vbWhite
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .BorderAround xlContinuous, xlThin
    End With
    
    Select Case IObject_objectCapture
        Case "Date"
            With MergedCells
                With .Validation
                    .Delete
                    .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="01/01/2013", Formula2:="01/01/2020"
                    .ErrorTitle = "System"
                    .ErrorMessage = "Please enter a valid date."
                End With
                .Locked = False
                .Font.Size = 12
            End With
        Case "Hour"
            With MergedCells
                With .Validation
                    .Delete
                    .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="23"
                    .ErrorTitle = "System"
                    .ErrorMessage = "Please enter a valid entry for hour between 00 and 23."
                End With
                .Locked = False
                .Font.Size = 14
                .NumberFormat = "00"
            End With
        Case "Minute"
            With MergedCells
                With .Validation
                    .Delete
                    .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="59"
                    .ErrorTitle = "System"
                    .ErrorMessage = "Please enter a valid entry for minute between 00 and 59."
                End With
                .Locked = False
                .Font.Size = 14
                .NumberFormat = "00"
            End With
    End Select

End Sub

Private Function IsPictureObjectCreated(sht As String, shapename As String) As Boolean

    Dim sh As Shape
    
    For Each sh In ThisWorkbook.Sheets(sht).Shapes
        If sh.Name = shapename Then
            IsPictureObjectCreated = True
            Exit Function
        End If
    Next sh

End Function

Private Function ItemsWithDataCount(sl As SlicerCache) As Integer

    Dim i As Integer
    Dim item As SlicerItem
    
    For Each item In sl.SlicerItems
        If item.HasData Then i = i + 1
    Next item
    
    ItemsWithDataCount = i

End Function

Public Sub IObject_SetText(elementText As String)

    Select Case IObject_objectType
           Case 1
                Shape.TextFrame.Characters.text = elementText
           Case 2
                TextBox.text = elementText
           Case 3
                Slicer.Caption = elementText
           Case 4
                Shape.ControlFormat.ListIndex = CInt(elementText)
           Case 5
           Case 6 'Form Controls do not have such property
           Case 7
           Case 8
                MergedCells.Value = elementText
    End Select

End Sub
Public Sub IObject_CreateObject(strObjectType As String, strObjectlocation As String, strObjectName As String, _
            dblPositionHeight As Double, dblPositionWidth As Double, dblPositionTop As Double, dblPositionLeft As Double)

   Select Case GetObjectTypesValue(strObjectType)
           Case 1
                Shape = ThisWorkbook.Sheets(strObjectlocation).Shapes.AddShape(msoShapeRectangle, dblPositionLeft, dblPositionTop, dblPositionWidth, dblPositionHeight)
                Shape.Name = strObjectName
           Case 2
           Case 3
           Case 4
           Case 5
           Case 6
           Case 7
           Case 8
    End Select

End Sub

Public Sub IObject_SetObject(strObjectType As String, strObjectlocation As String, strObjectName As String, strIsVisibleProp As String, strCaptureText As String, _
                             dblPositionHeight As Double, dblPositionWidth As Double, dblPositionTop As Double, dblPositionLeft As Double, strFieldName As String)

    IObject_objectType = GetObjectTypesValue(strObjectType)
    IObject_objectLocation = strObjectlocation
    IObject_objectName = strObjectName
    IObject_objectHeight = dblPositionHeight
    IObject_objectWidth = dblPositionWidth
    IObject_objectPositionTop = dblPositionTop
    IObject_objectPositionLeft = dblPositionLeft
    IObject_isVisibleProperty = GetCustomOperatorValue(strIsVisibleProp)
    IObject_objectCapture = strCaptureText
    IObject_objectFieldName = strFieldName
    
    SetObjectType
    
End Sub

Private Sub IObject_Destroy()

    Select Case IObject_objectType
           Case 1
                ThisWorkbook.Sheets(IObject_objectLocation).Shapes(IObject_objectName).Delete
           Case 2
           Case 3
           Case 4
           Case 5
           Case 6
           Case 7
           Case 8
    End Select

End Sub

Public Sub LoadObject()
    
    IObject_SetPosition
    IObject_SetCaption
    IObject_IsVisible
    
End Sub

Public Sub DestroyPicture(): ThisWorkbook.Sheets(IObject_objectLocation).Shapes(IObject_objectName).Delete: End Sub

Public Sub ClearText()

    Select Case IObject_objectType
           Case 1, 3, 4, 5, 6, 7, 8
           Case 2
                TextBox.text = ""
    End Select

End Sub

Public Sub ClearSlicer()

    With Slicer
        .SlicerCache.ClearManualFilter
        If .Style = "SlicerOK" Then .Style = "SlicerSelect"
    End With
    
End Sub

Public Sub ChangeSlicerTheme()

    With Slicer
        If .SlicerCache.VisibleSlicerItems.Count = 1 Then
            .Style = "SlicerOK"
        ElseIf ItemsWithDataCount(.SlicerCache) = 1 Then
            .Style = "SlicerOK"
        Else
            .Style = "SlicerSelect"
        End If
    End With
    
End Sub

'Public Function GetTableRange() As Range: Set GetTableRange = AdoTable.Table.Range: End Function

Public Function IsSingleSlicerItemSelected() As Boolean
    
    With Slicer
        If .SlicerCache.VisibleSlicerItems.Count = 1 Then
            IsSingleSlicerItemSelected = True
        ElseIf ItemsWithDataCount(.SlicerCache) = 1 Then
            IsSingleSlicerItemSelected = True
        Else
            IsSingleSlicerItemSelected = False
        End If
    End With
    
End Function
