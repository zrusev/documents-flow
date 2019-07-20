Attribute VB_Name = "Core_Subroutines"
Option Explicit
Option Private Module

Public Sub CloseWorkbook()

    If Application.Workbooks.Count = 1 Then
        Application.DisplayAlerts = False
        Application.Quit
    Else
        ThisWorkbook.Close False
    End If
    
End Sub
