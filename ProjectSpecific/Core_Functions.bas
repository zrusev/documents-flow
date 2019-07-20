Attribute VB_Name = "Core_Functions"
Option Explicit

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function Wait(ByVal miliseconds As Long): Sleep miliseconds: End Function

Function GetDesktop() As String
    
    Dim oWSHShell As Object
    
    Set oWSHShell = CreateObject("WScript.Shell")
    GetDesktop = oWSHShell.SpecialFolders("Desktop")
    
    Set oWSHShell = Nothing

End Function

Function GetObjectTypesValue(ObjName As String) As Integer

    Select Case LCase(ObjName)
           Case LCase("Shape")
                GetObjectTypesValue = ObjectTypes.Shape
           Case LCase("TextBox")
                GetObjectTypesValue = ObjectTypes.TextBox
           Case LCase("Slicer")
                GetObjectTypesValue = ObjectTypes.Slicer
           Case LCase("DropDown")
                GetObjectTypesValue = ObjectTypes.DropDown
           Case LCase("SpinButton")
                GetObjectTypesValue = ObjectTypes.SpinButton
           Case LCase("Table")
                GetObjectTypesValue = ObjectTypes.Table
           Case LCase("Picture")
                GetObjectTypesValue = ObjectTypes.Picture
           Case LCase("MergedCells")
                GetObjectTypesValue = ObjectTypes.MergedCells
    End Select
    
End Function

Function GetCustomOperatorValue(IsVisibleProperty As String) As Integer

    Select Case LCase(IsVisibleProperty)
           Case LCase("Yes")
                GetCustomOperatorValue = CustomOperators.Yes
           Case LCase("No")
                GetCustomOperatorValue = CustomOperators.No
    End Select
    
End Function
