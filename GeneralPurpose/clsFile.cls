VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IFile

Private fName As String
Private fPath As String


Public Property Get IFile_Name() As String: IFile_Name = fName: End Property

Public Property Let IFile_Name(ByVal RHS As String): fName = RHS: End Property

Public Property Let IFile_Path(ByVal RHS As String): fPath = RHS: End Property

Public Property Get IFile_Path() As String: IFile_Path = fPath: End Property
