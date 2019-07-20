VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Property Get ObjectName() As String:: End Property

Property Let ObjectName(ByVal tName As String):: End Property

Property Get ObjectType() As ObjectTypes:: End Property

Property Let ObjectType(ByVal oType As ObjectTypes):: End Property

Property Get Objectlocation() As String:: End Property

Property Let Objectlocation(ByVal oLoc As String):: End Property

Property Get IsVisibleProperty() As CustomOperators:: End Property

Property Let IsVisibleProperty(ByVal isVis As CustomOperators):: End Property

Property Get ObjectCapture() As String:: End Property

Property Let ObjectCapture(ByVal oCapt As String):: End Property

Property Get ObjectHeight() As Double:: End Property

Property Let ObjectHeight(ByVal oHeight As Double):: End Property

Property Get ObjectWidth() As Double:: End Property

Property Let ObjectWidth(ByVal oWidth As Double):: End Property

Property Get ObjectPositionTop() As Double:: End Property

Property Let ObjectPositionTop(ByVal oTop As Double):: End Property

Property Get ObjectPositionLeft() As Double:: End Property

Property Let ObjectPositionLeft(ByVal oLeft As Double):: End Property

Property Get ObjectFieldName() As String:: End Property

Property Let ObjectFieldName(ByVal oField As String):: End Property

Property Get GetText() As String:: End Property

Sub CreateObject(strObjectType As String, strObjectlocation As String, strObjectName As String, _
             dblPositionHeight As Double, dblPositionWidth As Double, dblPositionTop As Double, dblPositionLeft As Double): End Sub

Sub SetObject(ObjectType As String, Objectlocation As String, ObjectName As String, isVisibleProp As String, captureText As String, _
             positionHeight As Double, positionWidth As Double, PositionTop As Double, PositionLeft As Double, ObjectFieldName As String): End Sub

Sub isVisible(): End Sub

Sub SetCaption(): End Sub

Sub SetPosition(): End Sub

Sub SetText(elementText As String): End Sub

Sub Destroy(): End Sub
