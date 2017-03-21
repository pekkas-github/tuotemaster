VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_Decisions"
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

Private app                As Application_API
Private Services           As Services
Private dbConn             As ADODB.Connection
private oEntityCode		   as string		' Reference to the parent entity
private oVersionNro		   as string         ' Reference to the version of the parent entity


Public Sub init(entityCode as string, versionNro as string)

   Set app = New Application_API
   Set Services = New Services
   Set dbConn = CurrentProject.Connection
   oentitycode = entitycode
   oversionnro = versionnro
   
End Sub


Public Function createDecision() As dom_Decision
' Create a new decision if
'  - there are no decisions yet or
'  - startDate of the last decision is passed
' Edit last decision if
'  - startDate of the last decision is today
'TODO: toteutus on kesken. Katso mallia repo_PriceListEntries

	dim rst	as ADODB.Recordset
	set rst = dba_Decisions.getLastDecision(oentitycode, oversionnro) ' TODO: tässä hae dba_Decisions:sta rst

	if not rst.EOF AND rst!ValidFrom = Date then
   		set createDecision = returnDecision(rst)
	else	
   		Set createDecision = new_Decision(True)
   	end if	
   	
End Function


Public Function getDecision(decisionId As Long) As dom_Decision
' Return decision object requested by decision id.
' TODO: siirrä sql-kysely dba_Decisions:iin

   Dim sql           As String
   Dim rst           As New ADODB.Recordset
   Dim decision      As dom_Decision
   
   sql = "SELECT Authority, DecisionText, ValidFrom " & _
         "FROM " & VERSION_DECISION & _
         " WHERE Id = " & decisionId
         
   rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
   
   Set decision = new_Decision(False)
   
   With decision
      .setId (decisionId)
      .setAuthority (Nz(rst!authority))
      .setDecisionText (Nz(rst!decisionText))
      .setStartDate (rst!ValidFrom)
   End With
   
   Set getDecision = decision
   
   Set rst = Nothing
   Set decision = Nothing
   
End Function


Public Function getLastDecision() As dom_Decision
'TODO: tee dba_Decisions ja siirrä tämä sinne.

' Return data record of the last decision of this entity version
' If there is no last decision return nothing
   
   Dim lastDecision  As dom_Decision
   Dim sql           As String
   Dim rst           As New ADODB.Recordset
   
   sql = "SELECT d.Id, d.Authority, d.DecisionText, d.ValidFrom, d.ValidTo " & _
         "FROM " & VERSION_DECISION & " AS d " & _
         "INNER JOIN " & ITEM_VERSION & " AS v ON v.Id = d.ItemVersion_Id " & _
         "WHERE v.Item_Code =  '" & oentityCode & "' AND v.VersionNumber = '" & oversionNro & "' " & _
         "AND ValidFrom <= " & TODAY & " AND " & TODAY & " < ValidTo"
   
   rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
   
   With rst
         If .EOF Then
         Set lastDecision = Nothing
      Else
         Set lastDecision = new_Decision(False)
         lastDecision.setId (!id)
         lastDecision.setAuthority (Nz(!authority))
         lastDecision.setDecisionText (Nz(!decisionText))
         lastDecision.setStartDate (!ValidFrom)
      End If
   End With
   
   Set getLastDecision = lastDecision
   
   Set rst = Nothing
   Set lastDecision = Nothing
   
End Function


Public Sub saveDecision(entityCode As String, versionNro As String, newDecision As dom_Decision)
' Insert or update decision data and change the end date of the latest decision
' TODO: Tämä pelkistyy huomattavasti, kun päiväyksien kanssa ei enää
' temppuilla. Jos on uusi, talletetaan, jos ei, päivitetään.

   Dim lastDecision  As dom_Decision
   Dim versionID     As Long
   
   versionID = Services.getVersionId(entityCode, versionNro)
   
   If Not newDecision.isNew Then
'     update existing decision
      Call updateDecisionData(newDecision)
   Else
'     is new decision with ..
      Set lastDecision = getLastDecision(entityCode, versionNro)
      If lastDecision Is Nothing Then
'        .. no previous decisions -> ineset
         Call insertDecisionData(newDecision, versionID)
      ElseIf lastDecision.getStartDate = newDecision.getStartDate Then
'        .. the same start date with the latest -> update
         lastDecision.setDecisionText (newDecision.getDecisionText)
         lastDecision.setAuthority (newDecision.getAuthority)
         Call updateDecisionData(lastDecision)
         Else
'           .. not the same start date with the latest
            lastDecision.setEndDate (newDecision.getStartDate - 1)           ' update end date
            Call updateDecisionData(lastDecision)
            Call insertDecisionData(newDecision, versionID)
      End If
   End If
   
   Set lastDecision = Nothing
   
End Sub

   
Private Sub insertDecisionData(newDecision As dom_Decision, versionID As Long)
' Add a new decision into database

   Dim rst        As New ADODB.Recordset
   
   rst.Open VERSION_DECISION, dbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !decisionText = Nz(newDecision.getDecisionText)
      !authority = Nz(newDecision.getAuthority)
      !ValidFrom = newDecision.getStartDate
      !ValidTo = newDecision.getEndDate
      !ItemVersion_Id = versionID
      .Update
      newDecision.isNew = False
   End With
   
   Set rst = Nothing
   
End Sub


Private Sub updateDecisionData(modDecision As dom_Decision)
' Update an existing decision in database; change only text and authority

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & VERSION_DECISION & " WHERE Id = " & modDecision.getId
   
   rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      !decisionText = Nz(modDecision.getDecisionText)
      !authority = Nz(modDecision.getAuthority)
      !ValidTo = modDecision.getEndDate
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
   Set dbConn = Nothing
   
End Sub

