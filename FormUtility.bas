Attribute VB_Name = "FormUtility"
Option Compare Database
Option Explicit


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

