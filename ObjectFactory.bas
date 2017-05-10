Attribute VB_Name = "ObjectFactory"
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

Public Function new_Application_API() As Application_API
' Singleton

   Static app  As Application_API
   
   If app Is Nothing Then
      Set app = New Application_API
   End If
   
   Set new_Application_API = app
   
End Function

Public Function new_abs_Entities() As abs_Entities
' Singleton
   
   Static statAbsEntities  As abs_Entities
   
   If statAbsEntities Is Nothing Then
      Set statAbsEntities = New abs_Entities
   End If
   
   Set new_abs_Entities = statAbsEntities
   
End Function

Public Function new_abs_Entity(entityCode As String, versionNro As String) As abs_Entity

   Dim newEntity  As New abs_Entity
   
   Call newEntity.init(entityCode, versionNro)
   
   Set new_abs_Entity = newEntity
   Set newEntity = Nothing
   
End Function

Public Function new_Abs_Names(entityCode As String) As Abs_Names

   Dim newAbsNames   As New Abs_Names
   
   Call newAbsNames.init(entityCode)
   Set new_Abs_Names = newAbsNames
   
   Set newAbsNames = Nothing
   
End Function

Public Function new_Dba_Documents() As Dba_Documents

   Set new_Dba_Documents = New Dba_Documents
      
End Function


Public Function new_Dba_Decisions() As dba_Decisions

   Set new_Dba_Decisions = New dba_Decisions
   
End Function

Public Function new_Dba_BillingMapper() As Dba_BillingMapper
   
   Set new_Dba_BillingMapper = New Dba_BillingMapper
      
End Function


Public Function new_Dba_GroupMapper() As Dba_GroupMapper

   Set new_Dba_GroupMapper = New Dba_GroupMapper
      
End Function


Public Function new_Dba_Names() As Dba_Names

   Set new_Dba_Names = New Dba_Names
   
End Function


Public Function new_Dba_Owners() As Dba_Owners

   Set new_Dba_Owners = New Dba_Owners
      
End Function


Public Function new_Dba_PriceLines() As Dba_PriceLines

   Set new_Dba_PriceLines = New Dba_PriceLines
      
End Function


Public Function new_Dba_Properties() As Dba_Properties

   Set new_Dba_Properties = New Dba_Properties
   
End Function


Public Function new_Dba_Services() As Dba_Services

   Set new_Dba_Services = New Dba_Services
   
End Function


Public Function new_Category(categoryCode As String, namesBeh As If_Names, ownersBeh As If_Owners) As dom_Category

   Dim newCategory   As New dom_Category
   
   newCategory.init categoryCode, namesBeh, ownersBeh
   
   Set new_Category = newCategory
   Set newCategory = Nothing
   
End Function
Public Function new_Categories() As repo_Categories
'Singleton

   Static repo As repo_Categories
   
   If repo Is Nothing Then
      Set repo = New repo_Categories
   End If
   
   Set new_Categories = repo
   
End Function
Public Function new_PriceLines(salesItemCode As String, salesItemVersion As String) As repo_PriceLines
' Create and return a new Priceline repository associated to the parent Sales Item.

   Dim repo       As New repo_PriceLines
   Dim versionId  As Long
   
   versionId = DLookup("Id", ITEM_VERSION, "Item_Code = '" & salesItemCode & "' AND VersionNumber = '" & salesItemVersion & "'")
   repo.init versionId
   
   Set new_PriceLines = repo
   Set repo = Nothing
   
End Function


Public Function new_PriceListEntry(Optional priceListID As String = "STD") As dom_PriceListEntry
' Create and return a new PriceListEntry object associated to its parent PriceLine and PriceList (default is STD)

   Dim newPrice   As New dom_PriceListEntry
   
   newPrice.init priceListID
   
   Set new_PriceListEntry = newPrice
   Set newPrice = Nothing
   
End Function

Public Function new_PriceListEntries(priceLineCode As String) As repo_PriceListEntries
' Create and return a new Pricelist Entry repository associated to the parent Price Line.

   Dim repo As New repo_PriceListEntries
   
   repo.init priceLineCode
   
   Set new_PriceListEntries = repo
   Set repo = Nothing
   
End Function

Public Function new_SalesItems(productCode As String) As repo_SalesItems
' Create and return a new Sales Item repository associated to the parent product

   Dim repo As repo_SalesItems
   
   Set repo = New repo_SalesItems
   repo.init productCode
   
   Set new_SalesItems = repo
   Set repo = Nothing
   
End Function
Public Function new_Decision(isNew As Boolean) As dom_Decision

   Dim newDecision   As New dom_Decision
   
   Call newDecision.init(isNew)
   Set new_Decision = newDecision
   
   Set newDecision = Nothing
   
End Function

Public Function new_Decisions(entityCode As String, versionNro As String) As repo_Decisions
   
   Dim newDecisions  As New repo_Decisions
   
   newDecisions.init entityCode, versionNro
   Set new_Decisions = newDecisions
   
   Set newDecisions = Nothing
   
End Function


Public Function new_Description(language As String) As dom_Description

   Dim newDescription As New dom_Description
   
   Call newDescription.init(language)
   Set new_Description = newDescription
   
   Set newDescription = Nothing
   
End Function


Public Function new_Descriptions(entityCode As String, versionNro As String) As repo_Descriptions

   Dim newDescriptions  As New repo_Descriptions
   
   newDescriptions.init entityCode, versionNro
   Set new_Descriptions = newDescriptions
   
   Set newDescriptions = Nothing
   
End Function


Public Function new_Document() As dom_Document

   Dim newDocument  As New dom_Document
   
   newDocument.init
   Set new_Document = newDocument
   
   Set newDocument = Nothing
   
End Function


Public Function new_Documents(entityCode As String, versionNro As String) As repo_Documents

   Dim newDocuments  As New repo_Documents
   
   Call newDocuments.init(entityCode, versionNro)
   Set new_Documents = newDocuments
   
   Set newDocuments = Nothing
   
End Function


Public Function new_Name(nameText As String, language As String) As dom_Name
   
   Dim newName As New dom_Name
   
   Call newName.init(nameText, language)
   Set new_Name = newName
   
   Set newName = Nothing
   
End Function

Public Function new_Names_Any(entityCode As String) As repo_Names_Any

   Dim newNames   As New repo_Names_Any
   
   newNames.init entityCode
   
   Set new_Names_Any = newNames
   Set newNames = Nothing
   
End Function

Public Function new_Names_Unique(entityCode As String) As repo_Names_Unique

   Dim newNames   As New repo_Names_Unique
   
   newNames.init entityCode
   
   Set new_Names_Unique = newNames
   Set newNames = Nothing
   
End Function
Public Function new_Owner(personId As Long, startDate As Date) As dom_Owner

   Dim newOwner As New dom_Owner
   
   Call newOwner.init(personId, startDate)
   Set new_Owner = newOwner
   
   Set newOwner = Nothing
   
End Function


Public Function new_Owners_Self(entityCode As String) As repo_Owners_Self
' Sinleton

   Dim newOwners  As repo_Owners_Self
   
   Set newOwners = New repo_Owners_Self
   newOwners.init entityCode
   
   Set new_Owners_Self = newOwners
   Set newOwners = Nothing
   
End Function

Public Function new_Owners_Parent(entityCode As String) As repo_Owners_Parent

   Dim newOwners  As repo_Owners_Parent
   
   Set newOwners = New repo_Owners_Parent
   newOwners.init entityCode
   
   Set new_Owners_Parent = newOwners
   Set newOwners = Nothing
   
End Function

Public Function new_PriceLine() As dom_PriceLine

   Dim newPriceLine  As New dom_PriceLine
   
   Call newPriceLine.init
   Set new_PriceLine = newPriceLine
   
   Set newPriceLine = Nothing
   
End Function

Public Function new_SalesItem(itemCode As String, itemVersion As String) As dom_SalesItem

   Dim salesItem           As New dom_SalesItem
   Dim nameBehavior        As If_Names
   Dim ownerBehavior       As If_Owners
   
   Set nameBehavior = new_Names_Any(itemCode)
   Set ownerBehavior = new_Owners_Parent(itemCode)
   
   salesItem.init itemCode, itemVersion, nameBehavior, ownerBehavior
   
   Set new_SalesItem = salesItem
   Set salesItem = Nothing
   Set nameBehavior = Nothing
   Set ownerBehavior = Nothing
   
End Function
Public Function new_Property(entityType As String, propertyType As String, valueId As String, isNew As Boolean) As dom_Property

   Dim newProperty   As New dom_Property
   
   Call newProperty.init(entityType, propertyType, valueId, isNew)
   Set new_Property = newProperty
   
   Set newProperty = Nothing
   
End Function

Public Function new_Properties(entityCode As String, entityType As String) As repo_Properties

   Dim newProperties As New repo_Properties
   
   Call newProperties.init(entityCode, entityType)
   
   Set new_Properties = newProperties
   Set newProperties = Nothing
   
End Function

Public Function new_Product(productCode As String, versionNro As String) As dom_Product

   Dim newProduct          As New dom_Product
   Dim nameBehavior        As If_Names
   Dim ownerBehavior       As If_Owners

   Set nameBehavior = new_Names_Unique(productCode)
   Set ownerBehavior = new_Owners_Self(productCode)
   
   newProduct.init productCode, versionNro, nameBehavior, ownerBehavior
   
   Set new_Product = newProduct
   Set newProduct = Nothing
   Set nameBehavior = Nothing
   Set ownerBehavior = Nothing
   
End Function


Public Function new_Products() As repo_Products
' Singleton
   
   Static statProducts  As repo_Products
   
   If statProducts Is Nothing Then
      Set statProducts = New repo_Products
   End If
   
   Set new_Products = statProducts
   
End Function

Public Function new_Status(statusType As Integer, startDate As Date) As dom_Status
'

   Dim NewStatus As New dom_Status
   
      Call NewStatus.init(statusType, startDate)
   
   Set new_Status = NewStatus
   
   Set NewStatus = Nothing
   
End Function

Public Function new_Statuses() As repo_Statuses
' Singleton

   Static statStatuses  As repo_Statuses
   
   If statStatuses Is Nothing Then
      Set statStatuses = New repo_Statuses
   End If
   
   Set new_Statuses = statStatuses

End Function

Public Function new_Services() As Services
' Singleton object
   
   Static singletonServices   As Services
   
   If singletonServices Is Nothing Then
      Set singletonServices = New Services
   End If
   
   Set new_Services = singletonServices
   
End Function


Public Function new_GroupMapper() As GroupMapper

        Set new_GroupMapper = New GroupMapper
        
End Function


Public Function new_BillingMapper() As BillingMapper

        Set new_BillingMapper = New BillingMapper
        
End Function

