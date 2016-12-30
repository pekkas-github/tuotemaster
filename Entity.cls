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
Private oNames          As If_Names
Private oOwners         As If_Owners
Private oProperties     As Properties
Private oVersionNro     As String
Private oNew            As Boolean

Public Documents       As Documents

Public Sub init(entityCode As String, entityType As String, versionNro As String, nameBehavior As If_Names, ownerBehavior As If_Owners)
   
   Set app = New Application_API
   Set oNames = nameBehavior
   Set oOwners = ownerBehavior
   Set Documents = new_Documents(entityCode, versionNro)
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

   getName = oNames.getName(language)
   
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

Public Function getDecision(decisionId As Long) As Decision

   Set getDecision = app.Decisions.getDecision(decisionId)
   
End Function

Public Function getNewDecision() As Decision
' Return a new decision object with default values

        Set getNewDecision = app.Decisions.createDecision
        
End Function

Public Function getPriceLine(priceLineCode As String) As PriceLine

   Dim repo As repo_PriceLines
   
   Set repo = new_PriceLines(oCode, oVersionNro)
   
   If priceLineCode = "" Then
      Set getPriceLine = repo.createPriceLine
   Else
      Set getPriceLine = repo.getPriceLine(priceLineCode)
   End If
   
   Set repo = Nothing
   
End Function

Public Sub setPriceLine(PriceLine As PriceLine)
   
   Dim repo As repo_PriceLines
   
   Set repo = new_PriceLines(oCode, oVersionNro)

   Call repo.savePriceLine(PriceLine)
   
   Set repo = Nothing
   
End Sub

Public Sub setDecision(newDecision As Decision)
' Save a new decision for this entity version
        
   Call app.Decisions.saveDecision(oCode, oVersionNro, newDecision)
   
End Sub


Public Sub setDescription(newText As String, language As String)

   Dim newDescription   As description
   
   Set newDescription = new_Description(oCode, oVersionNro, language)
   newDescription.setText (Nz(newText))
   
   Call app.Descriptions.saveDescription(newDescription)
   
End Sub


Public Sub setName(newNameText As String, language As String)

   Call oNames.saveName(new_Name(newNameText, language))
   
End Sub


Public Sub setOwner(personId As Long, startDate As Date)

   Call oOwners.saveOwner(oCode, new_Owner(personId, startDate))
   
End Sub


Public Sub setStatus(statusId As Integer, startDate As Date)

   Call app.Statuses.saveStatus(oCode, oVersionNro, new_Status(statusId, startDate))
   
End Sub


Public Function getProperty(propertyType As String) As Property

        Set getProperty = oProperties.getProperty(propertyType)

End Function


Public Sub setProperty(value As Property)

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
   Set oNames = Nothing
   
End Sub
