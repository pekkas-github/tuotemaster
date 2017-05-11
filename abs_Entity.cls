VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "abs_Entity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Instantiated Entity is always an entity + version object pair.
' Calls that are associated with entity data itself use methods here.
' Calls that are associated with entity related objects are delegated to public repositories.
' These calls are of form Entity.Documents.getDocument(docId).
  
Private app             As Application_API
Private oDescriptions   As repo_Descriptions
Private oStatuses       As repo_Statuses
Private oDecisions      As repo_Decisions
Private oProperties     As repo_Properties
Private oCode           As String
Private oType           As String
Private oVersionNro     As String
Private oNew            As Boolean


Public Sub init(entityCode As String, entityType As String, versionNro As String)
   
   Set app = new_Application_API
   oCode = entityCode
   oType = entityType
   oVersionNro = versionNro

   Set oDescriptions = new_Descriptions(entityCode, versionNro)
   Set oStatuses = new_Statuses
   Set oDecisions = new_Decisions(entityCode, versionNro)
   Set oProperties = new_Properties(entityCode, oType)
      
End Sub


Public Function getCode() As String

   getCode = oCode
   
End Function


Public Function getType() As String

   getType = oType
   
End Function


Public Function getVersionNumber() As String
   
   getVersionNumber = oVersionNro
   
End Function


Public Function getDescription(language As String) As String
'  Return entity version's description text or empty string if there is no description yet

   Dim description   As dom_Description

   Set description = oDescriptions.getDescription(language)
   
   If description Is Nothing Then
      getDescription = ""
   Else
      getDescription = description.getText
   End If
   
   Set description = Nothing
   
End Function


Public Function getCurrentStatus() As dom_Status
' Return status type code of current status of this entity version

   Set getCurrentStatus = oStatuses.getStatus(oCode, oVersionNro)
   
End Function


Public Function getLastDecision() As dom_Decision
' Return the last decision

   Set getLastDecision = oDecisions.getLastDecision(oCode, oVersionNro)
   
End Function

Public Function getDecision(decisionId As Long) As dom_Decision

   Set getDecision = oDecisions.getDecision(decisionId)
   
End Function


Public Function getNewDecision() As dom_Decision
' Return a new decision object with default values

        Set getNewDecision = oDecisions.createDecision
        
End Function


Public Sub setDecision(newDecision As dom_Decision)
' Save a new decision for this entity version
        
   Call oDecisions.saveDecision(oCode, oVersionNro, newDecision)
   
End Sub


Public Sub setDescription(newText As String, language As String)

   Dim newDescription   As dom_Description
   
   Set newDescription = new_Description(language)
   newDescription.setText (Nz(newText))
   
   Call oDescriptions.saveDescription(newDescription)

End Sub


Public Sub setStatus(statusId As Integer, startDate As Date)

   Call oStatuses.saveStatus(oCode, oVersionNro, new_Status(statusId, startDate))
   
End Sub


Public Function getProperty(propertyType As String) As dom_Property

        Set getProperty = oProperties.getProperty(propertyType)

End Function


Public Sub setProperty(value As dom_Property)

        Call oProperties.saveProperty(value)
        
End Sub


Public Property Get isNew() As Boolean
   
   isNew = oNew
   
End Property


Public Property Let isNew(value As Boolean)

   oNew = value
   
End Property


Private Sub Class_Terminate()

   Set app = Nothing
   Set oProperties = Nothing
   Set oDecisions = Nothing
   Set oDescriptions = Nothing
   Set oStatuses = Nothing
   
End Sub
