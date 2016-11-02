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

Private db     As Documents_Db


Public Sub init()
   
   Set db = new Documents_Db
   
End Sub



Public Function getDocument(docId As Long) As Document
' Return a document reference object

   Dim docRst     As ADODB.Recordset
   dim docRef	  as Documents
   
   set docRst = db.getDocument(docId)
   set docRef = new Documents
   
   if docRst.EOF Then Err.Raise (recordNotFound,,"Document reference not found")
   
   with docRst
   		docRef.setId(!Id)
   		docRef.setType(!Type)
   		docRef.setName(!DocName)
   		docRef.setLanguage(!Lang)
   		docRef.setURL(!URL)
   		docRef.isPersistent = true
   		docRef.isModified = false
   	end with
   			
   Set getDocument = docRef
   
   set docRst = nothing
   set docRef = nothing
   
End Function



Public Function createDocReference() As Document

   Set createDocReference = New Document
   
End Function


Private Sub Class_Terminate()

   Set db = Nothing
   
End Sub
