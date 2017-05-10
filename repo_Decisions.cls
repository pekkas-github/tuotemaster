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
Private dba                As dba_Decisions
Private oEntityCode        As String       ' Reference to the parent entity
Private oVersionNro        As String       ' Reference to the version of the parent entity


Public Sub init(entityCode As String, versionNro As String)

   Set app = New Application_API
   Set dba = new_Dba_Decisions
   oEntityCode = entityCode
   oVersionNro = versionNro
   
End Sub

'---- METHODS ------------------------

Public Function createDecision() As dom_Decision
' Create a new decision if
'  - there are no decisions yet or
'  - startDate of the last decision is passed
' Edit last decision if
'  - startDate of the last decision is today

   Dim rst As ADODB.Recordset
   Dim newDecision   As dom_Decision
   Set rst = dba.getLastDecision(oEntityCode, oVersionNro)

   'TODO: jos rst.EOF on true, niin silloin rst!ValidFrom antaa virheen!
   If Not rst.EOF And rst!ValidFrom = Date Then ' Last decision starting today (-> edit current decision)
      Set newDecision = new_Decision(False)     ' Mark decision "not new"
      With newDecision
         .setId (rst!id)
         .setDecisionText (Nz(rst!DecisionText))
         .setAuthority (Nz(rst!Authority))
         .setStartDate (rst!ValidFrom)
      End With
      Set createDecision = newDecision
      Set newDecision = Nothing
   Else
      Set createDecision = new_Decision(True)   'No decisions or last decision starting earlier than today (-> create new decision)
   End If
        
End Function


Public Function getDecision(decisionId As Long) As dom_Decision
' Return decision object requested by decision id.

   Dim sql           As String
   Dim rst           As ADODB.Recordset
   Dim decision      As dom_Decision
   
   Set rst = dba.getDecision(decisionId)
   Set decision = new_Decision(False)
   
   With decision
      .setId (decisionId)
      .setAuthority (Nz(rst!Authority))
      .setDecisionText (Nz(rst!DecisionText))
      .setStartDate (rst!ValidFrom)
   End With
   
   Set getDecision = decision
   
   Set rst = Nothing
   Set decision = Nothing
   
End Function


Public Function getLastDecision() As dom_Decision
' Return data record of the last decision of this entity version
' If there is no last decision return nothing
   
   Dim lastDecision  As dom_Decision
   Dim rst           As ADODB.Recordset
      
   Set rst = dba.getLastDecision(oEntityCode, oVersionNro)
   
   With rst
         If .EOF Then
         Set lastDecision = Nothing
      Else
         Set lastDecision = new_Decision(False)
         lastDecision.setId (!id)
         lastDecision.setAuthority (Nz(!Authority))
         lastDecision.setDecisionText (Nz(!DecisionText))
         lastDecision.setStartDate (!ValidFrom)
      End If
   End With
   
   Set getLastDecision = lastDecision
   
   Set rst = Nothing
   Set lastDecision = Nothing
   
End Function


Public Sub saveDecision(entityCode As String, versionNro As String, newDecision As dom_Decision)
' Insert or update decision data and change the end date of the latest decision

   Dim versionId     As Long
   
   versionId = app.Services.getVersionId(entityCode, versionNro)
   
   If newDecision.isNew Then
      Call dba.insertDecisionData(newDecision, versionId)
   Else
      Call dba.updateDecisionData(newDecision)
   End If
   
   
End Sub



Private Sub Class_Terminate()

   Set app = Nothing
   
End Sub

