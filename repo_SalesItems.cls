VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_SalesItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'  repo_SalesItems is a virtual repository of all persistent SalesItems entities
'  in the database.You can get existing SalesItems from database.
'  It is also a factory for new SalesItem entities. You can create new entities
'  and save them. New SalesItems are also associated to their parent product.

Private app          As Application_API
Private super        As abs_Entities
Private util         As Services
Private oProductCode As String      ' Reference to the parent product


Public Sub init(productCode As String)

   Set super = New abs_Entities
   Set app = New Application_API
   Set util = new_Services
   oProductCode = productCode
   
End Sub

Private Sub Class_Terminate()

   Set app = Nothing
   Set super = Nothing
   Set util = Nothing
   
End Sub

'---- METHODS ----------------------

Public Function createSalesItem(Optional versionNro As String = "1.0") As dom_SalesItem
' Create a new sales item entity. The entity is also saved in database
' because it should be ready to accept name and owner entries which
' are stored directly to database and need reference to the sales item.
' The new sales item is also attached to its parent product.

   Dim retSalesItem        As dom_SalesItem
   Dim parentProduct       As dom_Product
   Dim mapper              As GroupMapper
   Dim salesItemCode       As String
   
'  Find next available sales item code; it's needed in next actions
   salesItemCode = util.getNextEntityCode("MN")

'  Create, construct and save the new sales item
   Set retSalesItem = new_SalesItem(salesItemCode, versionNro)
   retSalesItem.isNew = True
   Me.save retSalesItem
'  Create and save the version part of the sales item
   super.createVersion salesItemCode, versionNro

'  Associate the new sales item to its parent product in group hierarchy
   Set mapper = New GroupMapper
   Set parentProduct = app.getProducts.getProduct(oProductCode)
   mapper.map retSalesItem.getCode, retSalesItem.getType, parentProduct.getCode, parentProduct.getType
   
   Set createSalesItem = retSalesItem
   
   Set retSalesItem = Nothing
   Set parentProduct = Nothing
   Set mapper = Nothing
   
End Function


Public Function getSalesItem(salesItemCode As String, Optional versionNro As String = "last") As dom_SalesItem
' Return a persistent sales item. If the version number is not specified, return the last version.

   Dim retSalesItem     As dom_SalesItem
   
'  Define the version number and get version object
   If versionNro = "last" Then
      versionNro = util.getLastVersionNumber(salesItemCode)
   End If
      
'  Build a new sales item entity
   Set retSalesItem = new_SalesItem(salesItemCode, versionNro)
   retSalesItem.isNew = False
   
   Set getSalesItem = retSalesItem
   
   Set retSalesItem = Nothing
   
End Function


Public Sub save(newSalesItem As dom_SalesItem)
' Save product and version in database.

   Dim db   As New dba_SalesItems
   
   db.saveSalesItem newSalesItem
   
   Set db = Nothing
   
End Sub



   
