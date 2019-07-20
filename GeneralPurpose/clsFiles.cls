VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IFiles

Private pFiles As Collection


Public Property Get IFiles_Files() As Collection: Set IFiles_Files = pFiles: End Property

Private Property Let IFiles_Files(ByVal RHS As Collection): Set pFiles = RHS: End Property

Public Property Let IFiles_Add(ByVal RHS As IFile): pFiles.Add RHS: End Property


Private Sub Class_Initialize(): Set pFiles = New Collection: End Sub

Private Sub Class_Terminate(): Set pFiles = Nothing: End Sub

