Attribute VB_Name = "FormUtility"
Option Compare Database
Option Explicit

Private myForms   As Collection  ' Collection of open multi-selectable form objects

Public Sub initFormUtility()
   
   Set myForms = New Collection
   
End Sub

Public Sub openFormProduct(selectedProduct As dom_Product)
' Create, open and store reference of a multi-selectable product form

   Set Args = New Collection
   Args.Add selectedProduct, "Product"
   Args.Add getFormID, "FormID"
   
   openForm New Form_Product
End Sub


Public Sub openFormSalesItem(parentProduct As dom_Product, selectedSalesItem As dom_SalesItem, parentForm As Form)
' Create, open and store reference of a multi-selectable sales item form

   Set Args = New Collection
   Args.Add parentProduct, "Product"
   Args.Add selectedSalesItem, "SalesItem"
   Args.Add parentForm, "ParentForm"
   Args.Add getFormID, "FormID"
   
   openForm New Form_SalesItemVersion
   
End Sub

Private Sub openForm(newForm As Form)
' Open a multi-selectable form and store in collection with unique form id.

        myForms.Add newForm, Args("FormID")
        newForm.Visible = True
End Sub



Public Sub closeForm(formId As String)
' Remove a multi-selectable form from the collection.

   On Error Resume Next
   
   myForms.Remove (formId)
   
End Sub


Private Function getFormID() As String
' Generate the next unique formID key for MyForms collection

   Static formId     As String           ' Unique id for the next form in MyForms collection
   
   If formId = "" Then formId = "0"
   formId = CStr(CInt(formId) + 1)       ' Increment by 1 on each run
   
   getFormID = formId
End Function


Public Function showItemCount(list As ListBox) As Integer
' This functions corrects a bug in VBA list count.
' In csaes where list columns are shown the list count includes
' also the header line (countis one too big) except when the list is empty.

    If list.ColumnHeads Then
        If list.ListCount = 0 Then
            showItemCount = 0
        Else
            showItemCount = list.ListCount - 1
        End If
    Else
        showItemCount = list.ListCount
    End If
End Function

