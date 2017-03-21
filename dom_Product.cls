VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dom_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Implements If_Entity

' This class represents a product entity. Common entity methods are delegated to
' super class (abs_Entity). There is product entity specific behavior for
' - names: product name must be unique among products
' - owners: product has its own owners
'
' Documents are refered by passing the calls to Documents repository
'---------------------------

Private super     As abs_Entity
Private oNew      As Boolean
Private oNames    As If_Names
Private oOwners   As If_Owners


'---- CONSTRUCTOR -------------------------------

Public Sub init(productCode As String, versionNro As String, namesBeh As If_Names, ownersBeh As If_Owners)
   
   Set oNames = namesBeh
   Set oOwners = ownersBeh
   Set super = new_abs_Entity(productCode, versionNro)
      
End Sub


Private Sub Class_Terminate()

   Set super = Nothing
   Set oNames = Nothing
   Set oOwners = Nothing
   
End Sub

'---- METHODS ------------------------------------

Public Function getSalesItems() As repo_SalesItems
' Return Sales Item Repository associated to this product
   
   Set getSalesItems = new_SalesItems(super.getCode)
      
End Function


Public Function getDocuments() As repo_Documents
' Return Documents Repository associated to this product

   Set getDocuments = new_Documents(super.getCode, super.getVersionNumber)
   
End Function

'---- GETTERS AND SETTERS ---------------------------

Public Function getCode() As String

   getCode = super.getCode
   
End Function


Public Function getType() As String

   getType = super.getType
   
End Function


Public Function getName(language As String) As String

   getName = oNames.getName(language)
   
End Function

Public Function getOwner() As dom_Owner

   Set getOwner = oOwners.getOwner
   
End Function


Public Function getVersionNumber() As String
   
   getVersionNumber = super.getVersionNumber
   
End Function


Public Function getDescription(language As String) As String
'  Return entity version's description text or empty string if there is no description yet

   getDescription = super.getDescription(language)
   
End Function


Public Function getCurrentStatus() As dom_Status
' Return status type code of current status of this entity version

   Set getCurrentStatus = super.getCurrentStatus
   
End Function


Public Function getLastDecision() As dom_Decision
' Return the last decision

   Set getLastDecision = super.getLastDecision
   
End Function

Public Function getDecision(decisionId As Long) As dom_Decision

   Set getDecision = super.getDecision(decisionId)
   
End Function

Public Function getNewDecision() As dom_Decision
' Return a new decision object with default values

   Set getNewDecision = super.getNewDecision
        
End Function


Public Sub setDecision(newDecision As dom_Decision)
' Save a new decision for this entity version
   
   super.setDecision newDecision
      
End Sub


Public Sub setDescription(newText As String, language As String)

   Call super.setDescription(newText, language)
      
End Sub


Public Sub setName(newNameText As String, language As String)

   oNames.saveName new_Name(newNameText, language), super.getType
   
End Sub


Public Sub setOwner(personId As Long, startDate As Date)

   oOwners.saveOwner new_Owner(personId, startDate)
   
End Sub


Public Sub setStatus(statusId As Integer, startDate As Date)

   super.setStatus statusId, startDate
      
End Sub


Public Function getProperty(propertyType As String) As dom_Property

   Set getProperty = super.getProperty(propertyType)

End Function


Public Sub setProperty(value As dom_Property)

   super.setProperty value
        
End Sub


Public Property Get isNew() As Boolean
   
   isNew = super.isNew
   
End Property


Public Property Let isNew(value As Boolean)

   super.isNew = value
   
End Property


'---- METHODS CALLED VIA if_Entity INTERFACE --------------------

Private Function If_Entity_getCode() As String
   
   If_Entity_getCode = getCode
   
End Function

Private Function If_Entity_getCurrentStatus() As dom_Status
   
   Set If_Entity_getCurrentStatus = getCurrentStatus

End Function

Private Function If_Entity_getDecision(decisionId As Long) As dom_Decision

   Set If_Entity_getDecision = getDecision(decisionId)
   
End Function

Private Function If_Entity_getDescription(language As String) As String

   If_Entity_getDescription = getDescription(language)
   
End Function

Private Function If_Entity_getDocuments() As repo_Documents

' No document action has been defined for Cateory

End Function

Private Function If_Entity_getLastDecision() As dom_Decision

   Set If_Entity_getLastDecision = getLastDecision
   
End Function

Private Function If_Entity_getName(language As String) As String

   If_Entity_getName = getName(language)
   
End Function

Private Function If_Entity_getNewDecision() As dom_Decision

   Set If_Entity_getNewDecision = getNewDecision
   
End Function

Private Function If_Entity_getOwner() As dom_Owner

   Set If_Entity_getOwner = getOwner
   
End Function

Private Function If_Entity_getProperty(propertyType As String) As dom_Property

   Set If_Entity_getProperty = getProperty(propertyType)
   
End Function

' .getType
Private Function If_Entity_getType() As String
   
   If_Entity_getType = getType
   
End Function

Private Function If_Entity_getVersionNumber() As String

   If_Entity_getVersionNumber = getVersionNumber
   
End Function

Private Sub If_Entity_setDecision(newDecision As dom_Decision)

   setDecision newDecision
   
End Sub

Private Sub If_Entity_setDescription(descriptionText As String, language As String)

   setDescription descriptionText, language
   
End Sub

Private Sub If_Entity_setName(newName As String, language As String)

   setName newName, language
   
End Sub

Private Sub If_Entity_setOwner(partyId As Long, startDate As Date)

   setOwner partyId, startDate
   
End Sub

Private Sub If_Entity_setProperty(newProperty As dom_Property)

   setProperty newProperty
   
End Sub

Private Sub If_Entity_setStatus(statusId As Integer, startDate As Date)

   setStatus statusId, startDate
   
End Sub
