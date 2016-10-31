VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Documents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Class Documents; gets document references and associates entities with
' documents

Private db     As ADODB.Connection

Public Sub init()
   
   Set db = CurrentProject.Connection
   
End Sub

Public Function getDocument(docId As Long) As Document
' Return a document reference object

   Dim sql     As String
   Dim rst     As ADODB.Recordset
   
   Set getDocument = New Document
   
End Function

Public Function createDocReference() As Document

   Set createDocReference = New Document
   
End Function


Private Sub Class_Terminate()

   Set db = Nothing
   
End Sub
