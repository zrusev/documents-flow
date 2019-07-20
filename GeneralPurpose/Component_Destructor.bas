Attribute VB_Name = "Component_Destructor"
Option Explicit
Option Private Module

Function DesctructObject(ByRef obj As IObject) As Boolean: obj.Destroy: DesctructObject = True: End Function

Function DesctructShape(ByRef obj As IObject) As Boolean: DesctructShape = DesctructObject(obj): DesctructShape = True: End Function

Sub DesctructAll()
    
    Dim obj As Object
    For Each obj In objectsCollection.Members
        DesctructShape obj
    Next

End Sub
