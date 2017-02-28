Attribute VB_Name = "FormUtility"
Option Compare Database
Option Explicit

Private myForms   As Collection  ' Collection of open multi-selectable form objects



Public Sub openFormProduct(selectedProduct As dom_Product)
' Create, open and store reference of a multi-selectable form

   Set Args = New Collection
   Args.Add selectedProduct, "Product"
   Args.Add getFormID, "FormID"
   
   openForm New Form_Product
End Sub



Private Sub openForm(newForm As Form)
' Open a multi-selectable form and store in collection with unique form id.

	myForms.add newForm, Args("FormId")
	newForm.Visible	
End Sub



public sub closeForm(formID as string)
' Remove a multi-selectable form from the collection.

	myForms.Remove(formID)
end sub


private function getFormID() As string
' Generate the next unique formID

	Static formID	  as String		 ' Unique id for the next form in MyForms collection

	If formID = "" then formID = "0"
	formID = CStr(CInt(formID) + 1)	 ' Increment by 1 on each run 
	
	getFormID = formID
End function


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
    End if
End Function

