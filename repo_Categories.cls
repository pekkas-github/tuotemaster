VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_Categories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'  repo_Categories is a virtual repository of all persistent Category and
'  SubCategory entities in the database. You can get existing categories from database.
'  It is also a factory for new Category entities. You can create new entities
'  and save them. New Category is also associated to its parent category.

Private app          As Application_API
Private super        As abs_Entities


Private Sub Class_Initialize()

   Set super = New abs_Entities
   Set app = new_Application_API
   
End Sub

Private Sub Class_Terminate()

   Set app = Nothing
   Set super = Nothing
   
End Sub

'---- METHODS ----------------------

Public Function createCategory(categoryType As String, parentCode As String) As dom_Category
' Create a new Category entity. Category type is CAT or GRP.
' The entity is also saved in database because
' it should be ready to accept name and owner entries which
' are stored directly to database and need reference to the category.
' The new category is also attached to its parent in group hierarchies.

   Dim newCategory         As dom_Category
   Dim parentCategory      As dom_Category
   Dim nameBehavior        As If_Names
   Dim ownerBehavior       As If_Owners
   Dim categoryCode        As String
   
'  Find next available category / group code; it's needed in next actions
   categoryCode = app.Services.getNextEntityCode(categoryType)

'  Define name and owner behavior specific for category/group entities
   If categoryType = "CAT" Then
      Set nameBehavior = new_Names_Unique(categoryCode)
   Else
      Set nameBehavior = new_Names_Any(categoryCode)
   End If
   Set ownerBehavior = new_Owners_Self(categoryCode)

'  Create, construct and save the new category/group
   Set newCategory = new_Category(categoryCode, nameBehavior, ownerBehavior)
   newCategory.isNew = True
   Me.save newCategory
'  Create and save the version part of the product
   super.createVersion categoryCode, "1.0"

'  Add the new category in its parent
   Set parentCategory = getCategory(parentCode)
   
   app.getGroupMapper.map newCategory.getCode, newCategory.getType, parentCategory.getCode, parentCategory.getType

   Set createCategory = newCategory
   
   Set newCategory = Nothing
   Set parentCategory = Nothing
   Set nameBehavior = Nothing
   Set ownerBehavior = Nothing
   
End Function


Public Function getCategory(catCode As String) As dom_Category
' Return a persistent category.

   Dim retCategory      As dom_Category
   Dim nameBehavior     As If_Names
   Dim ownerBehavior    As If_Owners
         
'  Define name and owner behavior
   Set nameBehavior = new_Names_Unique(catCode)
   Set ownerBehavior = new_Owners_Self(catCode)

'  Build a new category entity
   Set retCategory = new_Category(catCode, nameBehavior, ownerBehavior)
   retCategory.isNew = False
   
   Set getCategory = retCategory
   
   Set retCategory = Nothing
   Set nameBehavior = Nothing
   Set ownerBehavior = Nothing
   
End Function


Public Sub save(newCategory As dom_Category)
' Save category and version in database.

   Dim db   As New dba_Categories
   
   db.saveCategory newCategory
   
   Set db = Nothing
   
End Sub
