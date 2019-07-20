Attribute VB_Name = "Core_Constants"
Option Explicit

Public Const FIND_ALL_PATTERN = "{{(\w+\.{1}\w+)}}"

Public Const RANDOM_DOC_NAME = "C:\Users\{USER}\{PATH}\UnitTestsFiles\main.doc"
Public Const MASTER_PATH = "C:\Users\{USER}\{PATH}\documents-flow"
Public Const MASTER_FOLDER = "UnitTestsFiles"
Public Const MASTER_DIRECTORY = "C:\Users\{USER}\{PATH}\documents-flow\UnitTestsFiles"
Public Const FILE_EXTENSION = "*.*"
Public Const DASHBOARD_SHEET = "Sheet1"
Public Const SAVE_AS_NAME = "output"

Public Function GetDocumentsKVP() As Variant()

GetDocumentsKVP = Array( _
                    Array("header.doc", "C:\Users\{USER}\{PATH}\documents-flow\UnitTestsFiles\header.doc"), _
                    Array("main.doc", "C:\Users\{USER}\{PATH}\documents-flow\UnitTestsFiles\main.doc"), _
                    Array("specifics.doc", "C:\Users\{USER}\{PATH}\documents-flow\UnitTestsFiles\specifics.doc"), _
                    Array("footer.doc", "C:\Users\{USER}\{PATH}\documents-flow\UnitTestsFiles\footer.doc") _
                )
End Function
