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

Private db           As Dba_Documents
Private parentCode   As String
Private versionNro   As String


Public Sub init(pCode As String, pVer As String)
   
   Set db = New Dba_Documents
   parentCode = pCode
   versionNro = pVer
   
End Sub

Private Sub Class_Terminate()

   Set db = Nothing
   
End Sub

'---- METHODS -------------------------

Public Function getDocReference(docId As Long) As Document
' Return a document reference object

   Dim docRst     As ADODB.Recordset
   Dim docRef     As Document
   
   Set docRst = db.getDocument(docId)
   Set docRef = new_Document
   
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
                        
   Set getDocReference = docRef
   
   Set docRst = Nothing
   Set docRef = Nothing
   
End Function


Public Function createDocReference() As Document

   Dim docRef      As Document
   
   Set docRef = new_Document
   
   With docRef
      .setName ("New document reference")
      .isPersistent = False
   End With

   Set createDocReference = docRef
   Set docRef = Nothing
   
End Function


Public Sub saveDocReference(docRef As Document)

   Dim app        As Application_API
   Dim versionId  As Long
   
   Set app = New Application_API
   
   versionId = app.Services.getVersionId(parentCode, versionNro)
   
   If docRef.isPersistent Then
      Call db.updateDocReference(docRef)
   Else
      docRef.setId (db.insertDocReference(versionId, docRef))
   End If
   
   Set app = Nothing
   

End Sub


Public Sub deleteDocReference(docRefId)

   db.deleteDocReference (docRefId)
   
End Sub

