Attribute VB_Name = "FormUtility"
Option Compare Database
Option Explicit

Private myForms   As Collection  ' Collection of open multi-selectable form objects

Public Sub openFormProduct(selectedProduct As dom_Product)
' Create, open and store reference of a multi-selectable form

   Set Args = New Collection
   Args.Add selectedProduct, "Product"
   
   openNewForm New Form_Product

End Sub

Private Sub openNewForm(newForm As Form)

End Sub

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

