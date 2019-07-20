VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Library"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7725
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

    Unload Me

End Sub

Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)

    Dim objNode As Node
    Dim lngItem As Long, lngCount As Long
    
    Node.ForeColor = vbBlack
    Node.BackColor = RGB(255, 255, 255)
    
    lngCount = Node.Children
    
    If Not lngCount = 0 Then
    
        Set objNode = Node.Child
    
        For lngItem = 1 To lngCount
            If objNode.Expanded = True Then objNode.Expanded = False
            Set objNode = objNode.Next
        Next
    End If
    Set objNode = Nothing

End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)

    Node.BackColor = RGB(0, 143, 255)
    Node.ForeColor = vbWhite

End Sub

' ToDo: change to double click
Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim filePath As String
    
    If InStr(1, Node.key, ".", vbTextCompare) > 0 Then
        filePath = MASTER_PATH & Node.key
        
        #If Debugging Then
            Debug.Print filePath
        #End If
        
        Dim file As New clsFile
        file.IFile_Name = Node.key
        file.IFile_Path = filePath
        
        filesCollection.Add = file
    Else
        Node.Expanded = True
    End If

End Sub

Private Sub UserForm_Initialize()

    Dim imL As New ImageList3
    
    FormatUserForm Me.Caption

    With imL.ListImages
        .Add , "folder", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Folder.bmp"))
        .Add , "word", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Word.bmp"))
        .Add , "excel", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Excel.bmp"))
        .Add , "text", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Text.bmp"))
        .Add , "link", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Link.bmp"))
        .Add , "url", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\Url.bmp"))
        .Add , "pdf", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\PDF.bmp"))
        .Add , "ppt", LoadPicture(FileExists(ThisWorkbook.Path & "\Icons\PPT.bmp"))
    End With
    
    TreeView1.ImageList = imL
    
End Sub

Private Function FileExists(strDir As String) As String

    If Dir(strDir) = "" Then
        MsgBox strDir & " does not exist." & vbNewLine & vbNewLine & "Please make sure you have all icons loaded.", vbInformation, "System"
        End
    End If
    
    FileExists = strDir
    
End Function

Public Sub AddToNode(ByVal activeCellKey As String, ByVal activeCellValue As String, ByVal previousCellKey As String, ByVal previousCellValue As String)

    If NodeExists(previousCellKey) = False Then
        TreeView1.Nodes.Add , , previousCellKey, previousCellValue, GetFileType(previousCellKey)
    End If
    
    If NodeExists(activeCellKey) = False Then
        TreeView1.Nodes.Add previousCellKey, tvwChild, activeCellKey, activeCellValue, GetFileType(activeCellKey)
    End If
    
End Sub

Private Function NodeExists(ByVal strKey As String) As Boolean

    Dim Node As MSComctlLib.Node
    On Error Resume Next
    Set Node = TreeView1.Nodes(strKey)
    Select Case Err.Number
        Case 0
            NodeExists = True
        Case Else
            NodeExists = False
    End Select
    
End Function

Private Function GetFileType(strValue As String) As String

    If InStr(1, strValue, ".doc", vbTextCompare) > 0 Then
        GetFileType = "word"
        Exit Function
    ElseIf InStr(1, strValue, ".dotx", vbTextCompare) > 0 Then
        GetFileType = "word"
        Exit Function
    ElseIf InStr(1, strValue, ".xls", vbTextCompare) > 0 Then
        GetFileType = "excel"
        Exit Function
    ElseIf InStr(1, strValue, ".pdf", vbTextCompare) > 0 Then
        GetFileType = "pdf"
        Exit Function
    ElseIf InStr(1, strValue, ".ppt", vbTextCompare) > 0 Then
        GetFileType = "ppt"
        Exit Function
    ElseIf InStr(1, strValue, ".txt", vbTextCompare) > 0 Then
        GetFileType = "text"
        Exit Function
    ElseIf InStr(1, strValue, ".lnk", vbTextCompare) > 0 Then
        GetFileType = "link"
        Exit Function
    ElseIf InStr(1, strValue, ".url", vbTextCompare) > 0 Then
        GetFileType = "url"
        Exit Function
    End If
    
    GetFileType = "folder"
    
End Function

Private Sub UserForm_Resize()

    With Me.TreeView1
        .Move 0, 0, Me.Width - 10, Me.Height
    End With

End Sub

'Private Sub TreeView1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
'Static nodeKey As String
'Dim nodTemp As Node
'Dim strText As String
'
'If Button = 0 Then
'    Set nodTemp = TreeView1.HitTest(TwipsPerPixelX * x, TwipsPerPixelY * y)
'    If Not nodTemp Is Nothing Then
'          nodTemp.Expanded = True
'    End If
'
'End If
'End Sub
