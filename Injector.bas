Attribute VB_Name = "Injector"
Option Compare Database
Option Explicit
' This module has functions that simplifies initialization of new objects
' in the main code. It provides a "standard" notation for creating a new
' object by allowing arguments in the same statement:
'
'    Set obj = new_Object (args)
'
' Each target class has a public "init" method that is called from these
' functions and which acts as a constructor in the class.
'
Public Function new_Decision(isNew As Boolean) As Decision

   Dim newDecision   As New Decision
   
   Call newDecision.init(isNew)
   Set new_Decision = newDecision
   
   Set newDecision = Nothing
   
End Function

Public Function new_Decisions() As Decisions
   
   Dim newDecisions  As New Decisions
   
   Call newDecisions.init
   Set new_Decisions = newDecisions
   
   Set newDecisions = Nothing
   
End Function


Public Function new_Description(entityCode As String, versionNro As String, language As String)

   Dim newDescription As New description
   
   Call newDescription.init(entityCode, versionNro, language)
   Set new_Description = newDescription
   
   Set newDescription = Nothing
   
End Function


Public Function new_Descriptions() As Descriptions

   Dim newDescriptions  As New Descriptions
   
   Call newDescriptions.init
   Set new_Descriptions = newDescriptions
   
   Set newDescriptions = Nothing
   
End Function


Public Function new_Entity(entityCode As String, entityType As String, versionNro As String) As Entity
' Construct a new entity with proper behaviors

   Dim newEntity     As New Entity
   Dim nameBehavior  As Object
   Dim ownerBehavior As Object
   
   Select Case entityType
      Case "MN"
         Set nameBehavior = New Names_Any
         Call nameBehavior.init(entityCode, entityType)
         Set ownerBehavior = New Owners_Parent
         Call ownerBehavior.init
      
      Case "GRP"
         Set nameBehavior = New Names_Any
         Call nameBehavior.init(entityCode, entityType)
         Set ownerBehavior = New Owners_Self
         Call ownerBehavior.init
      
      Case Else
         Set nameBehavior = New Names_Unique
         Call nameBehavior.init(entityCode, entityType)
         Set ownerBehavior = New Owners_Self
         Call ownerBehavior.init
   End Select
      
   Call newEntity.init(entityCode, entityType, versionNro, nameBehavior, ownerBehavior)
   Set new_Entity = newEntity
   
   Set newEntity = Nothing
   
End Function


Public Function new_Entities() As Entities

   Dim newEntities As New Entities
   
   Call newEntities.init
   Set new_Entities = newEntities
   
   Set newEntities = Nothing
   
End Function


Public Function new_Name(nameText As String, language As String)
   
   Dim newName As New EntityName
   
   Call newName.init(nameText, language)
   Set new_Name = newName
   
   Set newName = Nothing
   
End Function


Public Function new_Owner(personId As Long, startDate As Date) As Owner

   Dim newOwner As New Owner
   
   Call newOwner.init(personId, startDate)
   Set new_Owner = newOwner
   
   Set newOwner = Nothing
   
End Function


Public Function new_Property(entityType As String, propertyType As String, valueId As String, isNew As Boolean) As Property

   Dim newProperty   As New Property
   
   Call newProperty.init(entityType, propertyType, valueId, isNew)
   Set new_Property = newProperty
   
   Set newProperty = Nothing
   
End Function

Public Function new_Properties(entityCode As String, entityType As String)

   Dim newProperties As New Properties
   
   Call newProperties.init(entityCode, entityType)
   Set new_Properties = newProperties
   
   Set newProperties = Nothing
   
End Function


Public Function new_Documents(entityCode As String, versionNro As String) As Documents

   Dim newDocuments  As New Documents
   
   Call newDocuments.init(entityCode, versionNro)
   Set new_Documents = newDocuments
   
   Set newDocuments = Nothing
   
End Function

Public Function new_Status(statusType As Integer, startDate As Date) As Status
'

   Dim newStatus As New Status
   
      Call newStatus.init(statusType, startDate)
   
   Set new_Status = newStatus
   
   Set newStatus = Nothing
   
End Function

Public Function new_Statuses() As Statuses
   
   Dim newStatuses As New Statuses
   
   Call newStatuses.init
   Set new_Statuses = newStatuses
   
   Set newStatuses = Nothing
   
End Function
