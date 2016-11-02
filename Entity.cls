VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entity"
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
Private oCode           As String
Private oType           As String
Private oNames          As Names_IF
Private oOwners         As Owners_IF
Private oProperties     As Properties
Private oVersionNro     As String
Private oNew            As Boolean

Public Documents		As Documents

Public Sub init(entityCode As String, entityType As String, versionNro As String, nameBehavior As Names_IF, ownerBehavior As Owners_IF)
   
   Set app = New Application_API
   Set oNames = nameBehavior
   Set oOwners = ownerBehavior
   set Documents = new Documents
   oCode = entityCode
   oType = entityType
   oVersionNro = versionNro

   
   Set oProperties = new_Properties(entityCode, entityType)
      
End Sub


Public Function getCode() As String

   getCode = oCode
   
End Function


Public Function getType() As String

   getType = oType
   
End Function


Public Function getName(language As String) As String
'  Return entity version's name text or empty string if there is no name with the language yet

   Dim nameObject   As EntityName

   Set nameObject = oNames.getName(oCode, language)
   
   If nameObject Is Nothing Then
      getName = ""
   Else
      getName = nameObject.getText
   End If
   
End Function


Public Function getOwner() As Owner

   Set getOwner = oOwners.getOwner(oCode)
   
End Function


Public Function getVersionNumber() As String
   
   getVersionNumber = oVersionNro
   
End Function


Public Function getDescription(language As String) As String
'  Return entity version's description text or empty string if there is no description yet

   Dim description   As description
   
   Set description = app.Descriptions.getDescription(oCode, oVersionNro, language)
   
   If description Is Nothing Then
      getDescription = ""
   Else
      getDescription = description.getText
   End If
   
   Set description = Nothing
   
End Function


Public Function getCurrentStatus() As Status
' Return status type code of current status of this entity version

   Set getCurrentStatus = app.Statuses.getStatus(oCode, oVersionNro)
   
End Function


Public Function getLastDecision() As Decision
' Return the last decision

   Set getLastDecision = app.Decisions.getLastDecision(oCode, oVersionNro)
   
End Function


Public Sub setDecision(objDecision As Decision)

   Call app.Decisions.saveDecision(oCode, oVersionNro, objDecision)
   
End Sub


Public Sub setDescription(newText As String, language As String)

   Dim newDescription   As description
   
   Set newDescription = new_Description(oCode, oVersionNro, language)
   newDescription.setText (Nz(newText))
   
   Call app.Descriptions.saveDescription(newDescription)
   
End Sub
Public Sub setName(newNameText As String, language As String)

   Call oNames.saveName(oCode, new_Name(newNameText, language))
   
End Sub


Public Sub setOwner(personId As Long, startDate As Date)

   Call oOwners.saveOwner(oCode, new_Owner(personId, startDate))
   
End Sub


Public Sub setStatus(statusId As Integer, startDate As Date)

   Call app.Statuses.saveStatus(oCode, oVersionNro, new_Status(statusId, startDate))
   
End Sub
Public Function Properties() As Properties

   Set Properties = oProperties
   
End Function


Public Property Get isNew() As Boolean
   
   isNew = oNew
   
End Property


Public Property Let isNew(value As Boolean)

   oNew = value
   
End Property


Private Sub Class_Terminate()

   Set app = Nothing
   Set oProperties = Nothing
   Set oNames = Nothing
   
End Sub
