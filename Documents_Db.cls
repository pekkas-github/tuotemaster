VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Documents_Db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private dbconn    As ADODB.Connection


Private Sub Class_Initialize()

        Set dbconn = CurrentProject.Connection
        
End Sub


Public Function getDocument(docId As Long) As ADODB.Recordset
'** Haetaan dokumenttireferenssin tiedot tietokannasta

    Dim sql As String
    Dim rst As New ADODB.Recordset
    
    sql = "SELECT ref.DocumentName, ref.[Language],ref.URL, ref.DocumentType_Id, typ.TypeName " & _
          "FROM " & DOC_REFERENCE & " AS ref " & _
          "INNER JOIN " & DOC_TYPE & " AS typ ON ref.DocumentType_Id = typ.Id " & _
          "WHERE ref.Id = " & docId
    
    rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
    Set getDocument = rst
    Set rst = Nothing


End Function


Public Function insertDocumentReference(entityVersion As Long, documentRef As Document) As Long
' CREATED in ver 6.3.0

   Dim rst As New ADODB.Recordset
   
   rst.Open "DocumentReference", dbconn, adOpenDynamic, adLockPessimistic
   
   rst.AddNew
   rst!ItemVersion_Id = entityVersion
   rst!DocumentType_id = documentRef.getType
   rst!DocumentName = documentRef.getName
   rst!language = documentRef.getLanguage
   rst!URL = documentRef.getURL
   rst.Update
   
   insertDocumentReference = rst!Id
   
   Set rst = Nothing

End Function


Public Sub updateDocumentReference(entityVersion As Long, documentRef As Document)
' CREATED in ver 6.3.0

   Dim rst As New ADODB.Recordset
   Dim sql As String
   
   sql = "SELECT * FROM DocumentReference WHERE Id = " & documentRef.getId
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   With rst
      !DocumentType_id = documentRef.getType
      !DocumentName = documentRef.getName
      !language = documentRef.getLanguage
      !URL = documentRef.getURL
      .Update
   End With
   
   Set rst = Nothing
   
End Sub


Public Sub deleteDocumentReference(docId As Long)

End Sub
