Attribute VB_Name = "Tests"
Option Compare Database
Option Explicit


Public Sub testCategory()

   Dim category   As dom_Category
   Dim names      As If_Names
   Dim owners     As If_Owners
   
   Set names = new_Names_Unique("CAT000005", "CAT")
   Set owners = new_Owners_Self
   
   Set category = New dom_Category
   
   With category
      .init "CAT000005", "CAT", names, owners
      Debug.Print .getName("fin")
      Debug.Print .getOwner.getName
      Debug.Print .getDescription("fin")
      Debug.Print .getCurrentStatus.getType
      Debug.Print .getLastDecision.getDecisionText
   End With
   
   Set category = Nothing
   Set names = Nothing
   Set owners = Nothing
   
End Sub
Public Sub testProduct()

   Dim product As dom_Product
   Dim names   As If_Names
   Dim owners  As If_Owners
   
   Set names = new_Names_Unique("MR000001", "MR")
   
   Set owners = new_Owners_Self
   
   Set product = New dom_Product
   
   With product
      .init "MR000001", "1.0", names, owners
      Debug.Print .getName("fin")
      Debug.Print .getOwner.getName
      Debug.Print .getDescription("fin")
      Debug.Print .getCurrentStatus.getType
      Debug.Print .getLastDecision.getDecisionText
   End With
   
   Set product = Nothing
   Set names = Nothing
   Set owners = Nothing
   
End Sub

Public Sub testSalesItem()

   Dim salesItem  As dom_SalesItem
   Dim names      As If_Names
   Dim owners     As If_Owners
   
   Set names = new_Names_Any("MN000161", "MN")
   
   Set owners = new_Owners_Parent
   
   Set salesItem = new_SalesItem("MN000164", "1.0", names, owners)
   
   With salesItem
      Debug.Print .getName("fin")
      Debug.Print .getOwner.getName
      Debug.Print .getDescription("fin")
      Debug.Print .getCurrentStatus.getType
      Debug.Print .getLastDecision.getDecisionText
   End With
   
   Set salesItem = Nothing
   Set names = Nothing
   Set owners = Nothing

End Sub

Public Sub testCategoryRepository()

   Dim app        As Application_API
   Dim category   As dom_Category
   
   Set app = New Application_API
   
   Set category = app.getCategories.getCategory("CAT000005", "CAT")
   
   With category
      Debug.Print .getName("fin")
      Debug.Print .getOwner.getName
      Debug.Print .getDescription("fin")
      Debug.Print .getCurrentStatus.getType
      Debug.Print .getLastDecision.getDecisionText
   End With
   
   Set category = Nothing
   
End Sub

Public Sub testProductRepository()

   Dim app     As Application_API
   Dim product As dom_Product
   
   Set app = New Application_API
   
   Set product = app.getProducts.getProduct("MR000001")
   
   With product
      Debug.Print .getName("fin")
      Debug.Print .getOwner.getName
      Debug.Print .getDescription("fin")
      Debug.Print .getCurrentStatus.getType
      Debug.Print .getLastDecision.getDecisionText
   End With

   Set product = Nothing
   Set app = Nothing
   
End Sub

Public Sub testSalesItemRepository()

   Dim app        As Application_API
   Dim salesItem  As dom_SalesItem
   
   Set app = New Application_API
   
   Set salesItem = app.getProducts.getProduct("MR000001").getSalesItems.getSalesItem("MN000161")
   
   With salesItem
      Debug.Print .getName("fin")
      Debug.Print .getOwner.getName
      Debug.Print .getDescription("fin")
      Debug.Print .getCurrentStatus.getType
      Debug.Print .getLastDecision.getDecisionText
   End With
   
   Set salesItem = Nothing
   
End Sub

Public Sub testSalesItemFromProduct()

   Dim app        As Application_API
   Dim salesItem  As dom_SalesItem
   
   Set app = New Application_API

   Set salesItem = app.getProducts.getProduct("MR000001").getSalesItems.getSalesItem("MN000161")
   
   With salesItem
      Debug.Print .getName("fin")
      Debug.Print .getOwner.getName
      Debug.Print .getDescription("fin")
      Debug.Print .getCurrentStatus.getType
      Debug.Print .getLastDecision.getDecisionText
   End With
   
   Set salesItem = Nothing
   Set product = Nothing
   
End Sub
