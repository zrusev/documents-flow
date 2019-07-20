Attribute VB_Name = "Component_Loader"
Option Explicit
Option Private Module

Global filesCollection As IFiles
Global objectsCollection As IObjects

Public Sub LoadForm(): UserForm1.Show False: End Sub

Public Sub LoadCollection()
    
    Set filesCollection = New clsFiles
    Set objectsCollection = New clsObjects
    
End Sub
