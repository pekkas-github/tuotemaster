VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Documents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Class Documents; gets document references and associates entities with
' documents

Private db     As Documents_Db


Public Sub init()
   
   Set db = New Documents_Db
   
End Sub



Public Function getDocument(docId As Long) As Document
' Return a document reference object

   Dim docRst     As ADODB.Recordset
   Dim docRef     As Document
   
   Set docRst = db.getDocument(docId)
   Set docRef = New Document
   
   If docRst.EOF Then Err.Raise recordnotfound, , "Document reference not found"
   
   With docRst
                docRef.setId (docId)
                docRef.setType (!DocumentType_id)
                docRef.setName (!DocumentName)
                docRef.setLanguage (!language)
                docRef.setURL (!URL)
                docRef.isPersistent = True
                docRef.isModified = False
        End With
                        
   Set getDocument = docRef
   
   Set docRst = Nothing
   Set docRef = Nothing
   
End Function



Public Function createDocReference() As Document

   Set createDocReference = New Document
   
End Function


Private Sub Class_Terminate()

   Set db = Nothing
   
End Sub
