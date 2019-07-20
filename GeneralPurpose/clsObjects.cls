VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IObjects

Private pMembers As Collection


Public Property Get IObjects_Members() As Collection: Set IObjects_Members = pMembers: End Property

Private Property Let IObjects_Members(ByVal RHS As Collection): Set pMembers = RHS: End Property

Public Property Let IObjects_Add(ByVal RHS As IObject): pMembers.Add RHS: End Property


Private Sub Class_Initialize(): Set pMembers = New Collection: End Sub

Private Sub Class_Terminate(): Set pMembers = Nothing: End Sub
