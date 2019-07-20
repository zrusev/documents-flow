Attribute VB_Name = "Component_Recursion"
Option Explicit
Option Private Module

Function RecursiveDir(ByRef colFiles As Collection, strFolder As String, strFileSpec As String, bIncludeSubfolders As Boolean)

    Dim strTemp As String
    Dim vFolderName As Variant
    Dim colFolders As New Collection
    
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & strFileSpec)
    
    Do While strTemp <> vbNullString
        colFiles.Add strFolder & strTemp
        strTemp = Dir
    Loop
    
    If bIncludeSubfolders Then
    
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If Len(strFolder & strTemp) >= 255 Then
                    MsgBox "The file's name is too long." & vbNewLine & _
                           "The lenght should not exceed 255 symbols." & vbNewLine & _
                           "'" & strFolder & strTemp & "'", vbInformation, "System"
                    End
                End If
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0 Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop
                
        For Each vFolderName In colFolders
            Call RecursiveDir(colFiles, strFolder & vFolderName, strFileSpec, True)
        Next vFolderName
    End If

End Function

Function TrailingSlash(strFolder As String) As String

    If Len(strFolder) > 0 Then
        If Right(strFolder, 1) = Application.PathSeparator Then
            TrailingSlash = strFolder
        Else
            TrailingSlash = strFolder & Application.PathSeparator
        End If
    End If

End Function
