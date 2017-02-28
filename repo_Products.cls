VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_Products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'  repo_Products is a virtual repository of all persistent Product entities
'  in the database.You can get existing Products from database.
'  It is also a factory for new Product entities. You can create new entities
'  and save them. New Products are also associated to PROD collection.

Private app          As Application_API
Private super        As abs_Entities
Private util         As Services


Private Sub Class_Initialize()

   Set super = new_abs_Entities
   Set app = new_Application_API
   Set util = new_Services
   
End Sub

Private Sub Class_Terminate()

   Set app = Nothing
   Set super = Nothing
   Set util = Nothing
   
End Sub

'---- METHODS ----------------------

Public Function createProduct(Optional versionNro As String = "1.0") As dom_Product
' Create a new product entity. The entity is also saved in database
' because it should be ready to accept name and owner entries which
' are stored directly to database and need reference to the product.
' The new product is also attached to PROD collection in group hierarchies.

   Dim newProduct          As dom_Product
   Dim rootProduct         As dom_Product
   Dim mapper              As GroupMapper
   Dim nameBehavior        As If_Names
   Dim ownerBehavior       As If_Owners
   Dim productCode         As String
   
'  Find next available product code; it's needed in next actions
   productCode = util.getNextEntityCode("MR")

'  Define name and owner behavior specific for Product entities
   Set nameBehavior = new_Names_Unique(productCode)
   Set ownerBehavior = new_Owners_Self

'  Create, construct and save the new product
   Set newProduct = new_Product(productCode, versionNro, nameBehavior, ownerBehavior)
   newProduct.isNew = True
   Me.save newProduct
'  Create and save the version part of the product
   super.createVersion productCode, versionNro

'  Add the new product in PROD collection in group hierarchy
   Set mapper = New GroupMapper
   Set rootProduct = getProduct("PROD")
   mapper.map newProduct, rootProduct
   
   Set createProduct = newProduct
   
   Set newProduct = Nothing
   Set rootProduct = Nothing
   Set nameBehavior = Nothing
   Set ownerBehavior = Nothing
   Set mapper = Nothing
   
End Function


Public Function getProduct(productCode As String, Optional versionNro As String = "last") As dom_Product
' Return a persistent product. If the version number is not specified, return the last version.

   Dim returnProduct    As dom_Product
   Dim nameBehavior     As If_Names
   Dim ownerBehavior    As If_Owners
   
'  Define the version number and get version object
   If versionNro = "last" Then
      versionNro = util.getLastVersionNumber(productCode)
   End If
      
'  Define name and owner behavior
   Set nameBehavior = new_Names_Unique(productCode)
   Set ownerBehavior = new_Owners_Self

'  Build a new product entity
   Set returnProduct = new_Product(productCode, versionNro, nameBehavior, ownerBehavior)
   returnProduct.isNew = False
   
   Set getProduct = returnProduct
   
   Set returnProduct = Nothing
   Set nameBehavior = Nothing
   Set ownerBehavior = Nothing
   
End Function


Public Sub save(newProduct As dom_Product)
' Save product and version in database.

   Dim db   As New dba_Products
   
   db.saveProduct newProduct
   
   Set db = Nothing
   
End Sub



   
