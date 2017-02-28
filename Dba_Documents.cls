VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dba_Documents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private app       As Application_API

Private Sub Class_Initialize()

   Set app = new_Application_API
   
End Sub


Public Function getDocument(docId As Long) As ADODB.Recordset
'** Haetaan dokumenttireferenssin tiedot tietokannasta

    Dim sql As String
    Dim rst As New ADODB.Recordset
    
    sql = "SELECT ref.DocumentName, ref.[Language],ref.URL, ref.DocumentType_Id, typ.TypeName " & _
          "FROM " & DOC_REFERENCE & " AS ref " & _
          "INNER JOIN " & DOC_TYPE & " AS typ ON ref.DocumentType_Id = typ.Id " & _
          "WHERE ref.Id = " & docId
    
    rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
    Set getDocument = rst
    Set rst = Nothing


End Function


Public Function insertDocReference(entityVersion As Long, docRef As dom_Document) As Long


   Dim rst As New ADODB.Recordset
   
   rst.Open DOC_REFERENCE, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   rst.AddNew
   rst!ItemVersion_Id = entityVersion
   rst!DocumentType_id = docRef.getType
   rst!DocumentName = docRef.getName
   rst!language = docRef.getLanguage
   rst!URL = docRef.getURL
   rst.Update
   
   insertDocReference = rst!id
   
   Set rst = Nothing

End Function


Public Sub updateDocReference(docRef As dom_Document)

   Dim rst As New ADODB.Recordset
   Dim sql As String
   
   sql = "SELECT * FROM " & DOC_REFERENCE & " WHERE Id = " & docRef.getId
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      !DocumentType_id = docRef.getType
      !DocumentName = docRef.getName
      !language = docRef.getLanguage
      !URL = docRef.getURL
      .Update
   End With
   
   Set rst = Nothing
   
End Sub


Public Sub deleteDocReference(docId As Long)

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT * FROM " & DOC_REFERENCE & " WHERE Id = " & docId
   
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      .Delete
      .Update
      .Close
   End With
   
   Set rst = Nothing
   
End Sub


Private Sub Class_Terminate()
   
   Set app = Nothing

End Sub
