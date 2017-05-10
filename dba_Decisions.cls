VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dba_Decisions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private app As Application_API

Private Sub Class_Initialize()

   Set app = New Application_API
   
End Sub


Private Sub Class_Terminate()

   Set app = Nothing
   
End Sub
Public Function getDecision(decisionId As Long) As ADODB.Recordset
' Return specific decision data from database

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT Authority, DecisionText, ValidFrom " & _
         "FROM " & VERSION_DECISION & _
         " WHERE Id = " & decisionId
         
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic

   Set getDecision = rst
   Set rst = Nothing
   
End Function


Public Function getLastDecision(entityCode As String, versionNro As String) As ADODB.Recordset

   Dim sql              As String
   Dim sqlLastDecision  As String
   Dim rst              As New ADODB.Recordset
   Dim versionId        As Long
   
   versionId = app.Services.getVersionId(entityCode, versionNro)

'  Get id of the last decision of this entity version (may be empty)
   sqlLastDecision = "SELECT ItemVersion_Id AS verID, Max(ValidFrom) AS StartDay " & _
                     "FROM " & VERSION_DECISION & _
                     " WHERE ItemVersion_Id = " & versionId & _
                     " GROUP BY ItemVersion_Id"
                     
'  Get decision data of the last decision of this entity version (may be empty)
   sql = "SELECT VeDe.Id, VeDe.Authority, VeDe.DecisionText, VeDe.ValidFrom " & _
         "FROM " & VERSION_DECISION & " AS VeDe " & _
         "INNER JOIN (" & sqlLastDecision & ") AS LaDe ON (VeDe.ItemVersion_Id = LaDe.verId AND VeDe.ValidFrom = LaDe.StartDay)"
   
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   Set getLastDecision = rst
   Set rst = Nothing

End Function


Public Sub insertDecisionData(newDecision As dom_Decision, versionId As Long)
' Add a new decision into database

   Dim rst        As New ADODB.Recordset
   
   rst.Open VERSION_DECISION, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !DecisionText = Nz(newDecision.getDecisionText)
      !Authority = Nz(newDecision.getAuthority)
      !ValidFrom = newDecision.getStartDate
      !ItemVersion_Id = versionId
      .Update
      newDecision.isNew = False
   End With
   
   Set rst = Nothing
   
End Sub


Public Sub updateDecisionData(modDecision As dom_Decision)
' Update an existing decision in database; change only text and authority

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & VERSION_DECISION & " WHERE Id = " & modDecision.getId
   
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      !DecisionText = Nz(modDecision.getDecisionText)
      !Authority = Nz(modDecision.getAuthority)
      .Update
      modDecision.isNew = False
   End With
   
   Set rst = Nothing
   
End Sub

