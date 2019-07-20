Attribute VB_Name = "Component_Match"
Option Explicit
Option Private Module

Function IsMatch(ByVal strItem As String, ByVal strPattern As String) As Object

    Dim objMatches As Object: Set IsMatch = objMatches
    
    Dim objReg As Object
    Set objReg = CreateObject("vbscript.regexp")
    With objReg
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .pattern = strPattern

        If .test(strItem) Then
            Set objMatches = objReg.Execute(strItem)
            Set IsMatch = objMatches
        End If
    End With
    
End Function

Function BuildQuery(Members() As Variant) As String

    Dim builder As New clsStringBuilder
    
    Dim i As Long
    For i = LBound(Members) To UBound(Members)
        builder.Append "(" & Replace(Members(i)(0), ".", "\.") & ")|"
    Next i
    
    BuildQuery = Left(builder.toString, Len(builder.toString) - 1)

End Function
