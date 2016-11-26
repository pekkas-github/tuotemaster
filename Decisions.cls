VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Decisions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
' Class Decisions

' This is a kind of repository abstaction to the UI client.
' It creates new Decision objects, fetches existing Decisiond data
' and stores new and modified dDecision data.

' CREATED in HOT#005

Private app                As Application_API
Private Services           As Services
Private dbconn             As ADODB.Connection


Public Sub init()

   Set app = New Application_API
   Set Services = New Services
   Set dbconn = CurrentProject.Connection
   
End Sub


Public Function createDecision() As Decision

   Set createDecision = new_Decision(True)
   
End Function


Public Function getDecision(decisionId As Long) As Decision
' Return decision object requested by decision id.


   Dim sql           As String
   Dim rst           As New ADODB.Recordset
   Dim dec           As Decision
   
   sql = "SELECT Authority, DecisionText, ValidFrom, ValidTo " & _
         "FROM " & VERSION_DECISION & _
         " WHERE Id = " & decisionId
         
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   Set dec = new_Decision(False)
   
   With rst
      dec.setId (decisionId)
      dec.setAuthority (Nz(!authority))
      dec.setDecisionText (Nz(!decisionText))
      dec.setStartDate (!validFrom)
      dec.setEndDate (!validTo)
   End With
   
   Set getDecision = dec
   
   Set rst = Nothing
   Set dec = Nothing
   
End Function


Public Function getLastDecision(entityCode As String, versionNro As String) As Decision
' Return data record of the last decision of this entity version
' If there is no last decision return nothing
   
   Dim lastDecision  As Decision
   Dim sql           As String
   Dim rst           As New ADODB.Recordset
   
   sql = "SELECT d.Id, d.Authority, d.DecisionText, d.ValidFrom, d.ValidTo " & _
         "FROM " & VERSION_DECISION & " AS d " & _
         "INNER JOIN " & ITEM_VERSION & " AS v ON v.Id = d.ItemVersion_Id " & _
         "WHERE v.Item_Code =  '" & entityCode & "' AND v.VersionNumber = '" & versionNro & "' " & _
         "AND ValidFrom <= " & TODAY & " AND " & TODAY & " < ValidTo"
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   With rst
         If .EOF Then
         Set lastDecision = Nothing
      Else
         Set lastDecision = new_Decision(False)
         lastDecision.setId (!Id)
         lastDecision.setAuthority (Nz(!authority))
         lastDecision.setDecisionText (Nz(!decisionText))
         lastDecision.setStartDate (!validFrom)
         lastDecision.setEndDate (!validTo)
      End If
   End With
   
   Set getLastDecision = lastDecision
   
   Set rst = Nothing
   Set lastDecision = Nothing
   
End Function


Public Sub saveDecision(entityCode As String, versionNro As String, newDecision As Decision)
' Insert or update decision data and change the end date of the latest decision

   Dim lastDecision  As Decision
   Dim versionId     As Long
   
   versionId = Services.getVersionId(entityCode, versionNro)
   
   If Not newDecision.isNew Then
'     update existing decision
      Call updateDecisionData(newDecision)
   Else
'     is new decision with ..
      Set lastDecision = getLastDecision(entityCode, versionNro)
      If lastDecision Is Nothing Then
'        .. no previous decisions -> ineset
         Call insertDecisionData(newDecision, versionId)
      ElseIf lastDecision.getStartDate = newDecision.getStartDate Then
'        .. the same start date with the latest -> update
         lastDecision.setDecisionText (newDecision.getDecisionText)
         lastDecision.setAuthority (newDecision.getAuthority)
         Call updateDecisionData(lastDecision)
         Else
'           .. not the same start date with the latest
            lastDecision.setEndDate (newDecision.getStartDate - 1)           ' update end date
            Call updateDecisionData(lastDecision)
            Call insertDecisionData(newDecision, versionId)
      End If
   End If
   
   Set lastDecision = Nothing
   
End Sub

   
Private Sub insertDecisionData(newDecision As Decision, versionId As Long)
' Add a new decision into database

   Dim rst        As New ADODB.Recordset
   
   rst.Open VERSION_DECISION, dbconn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !decisionText = Nz(newDecision.getDecisionText)
      !authority = Nz(newDecision.getAuthority)
      !validFrom = newDecision.getStartDate
      !validTo = newDecision.getEndDate
      !ItemVersion_Id = versionId
      .Update
      newDecision.isNew = False
   End With
   
   Set rst = Nothing
   
End Sub


Private Sub updateDecisionData(modDecision As Decision)
' Update an existing decision in database; change only text and authority

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & VERSION_DECISION & " WHERE Id = " & modDecision.getId
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   With rst
      !decisionText = Nz(modDecision.getDecisionText)
      !authority = Nz(modDecision.getAuthority)
      !validTo = modDecision.getEndDate
      .Update
      modDecision.isNew = False
   End With
   
   Set rst = Nothing
   
End Sub


Public Function rsString(entityCode As String, versionNro As String, language As String) As String

   Dim sql  As String
   
   sql = "SELECT d.Id, n.Text AS Status, d.ValidFrom AS Date, d.DecisionText AS Decision, d.Authority " & _
         "FROM ItemVersionStatusDecision AS d " & _
         "INNER JOIN ItemVersion AS v ON v.Id = d.ItemVersion_Id " & _
         "INNER JOIN " & _
         "ItemVersionStatus AS s ON d.ItemVersion_Id = s.ItemVersion_Id AND " & _
               "d.ValidFrom >= s.Valid_From AND " & _
               "d.ValidFrom < s.Valid_To " & _
         "INNER JOIN ItemVersionStatusName AS n ON s.StatusName_Id = n.Id " & _
         "WHERE n.Lang_Code = '" & language & "' AND v.Item_Code = '" & entityCode & "' AND v.VersionNumber = '" & versionNro & "'"
         
   rsString = sql
   
End Function


Private Sub Class_Terminate()

   Set app = Nothing
   Set Services = Nothing
   Set dbconn = Nothing
   
End Sub
