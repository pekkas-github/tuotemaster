VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' GroupMapper handles cutting and pasting categories and subgroups.
' You can cut and paste several subgroups (with their sub hierarchies.)
' Orphan subgroups will be cleaned by garbage collection.

'---- ATTRIBUTES ------------------------------
Private clipBoard() As Variant
Private app          As Application_API
Private db           As Dba_GroupMapper
Private cutPending   As Boolean


'---- INIT AND TERMINATE --------------------
Private Sub Class_Initialize()
    
   Set app = New Application_API
   Set db = New Dba_GroupMapper

   cutPending = False
    
End Sub

Private Sub Class_Terminate()
' Clean the group hierarchies before closing by deleting all orphan elements.

   db.deleteOrphanGroups
   Set db = Nothing
   Set app = Nothing
   
End Sub

'---- METHODS -------------------------------


Public Sub cut(selectedItems As Variant)
' Store all selected child groups to be potentially cut from their parent.
' Item list includes childCode, childType and parentCode of the selected items
' Parent may be ROOT, CAT or GRP.
            
   clipBoard = selectedItems
            
   cutPending = True
        
End Sub



Public Sub copy(selectedItems As Variant)
' Copy all child groups selected on the clipboard.
' Clipboard contents remain until next cut, copy or garbage collection takes place.
    
   clipBoard = selectedItems
   
   cutPending = False
            
End Sub


Public Sub paste(parentCode As String, parentType As String, parentSubType As String)
' Paste the previously copied/cut children on clipboard to this parent
' Before pasting check that it is allowed  child type - parent sub type match.
' Cut items are erased from clipboard after paste. Copied items remain on clipboard.

   On Error GoTo catch
    
   Dim parentSubMemberType As String
   Dim i                   As Integer
   Dim childCode           As String
   Dim childType           As String
   Dim oldParent           As String
      
'  Get parent group's allowed subgroup type
   parentSubMemberType = app.getCategories.getCategory(parentCode).getProperty("SUB_MEM").getValue

'  Check that attaching is allowed
   If parentSubMemberType <> clipBoard(i, 1) Then
      Err.Raise vbObjectError + 513, , "Parent and child items are incompatible."
   End If

   For i = 0 To UBound(clipBoard)
      childCode = CStr(clipBoard(i, 0))
      childType = CStr(clipBoard(i, 1))
      oldParent = CStr(clipBoard(i, 2))
      If cutPending Then
'        Remove group from current parernt
         unMap childCode, oldParent
      End If
'     Attach group to new parent
      map childCode, childType, parentCode, parentType
   Next i
           
   If cutPending Then Erase clipBoard
   cutPending = False
   Exit Sub
   
catch:
   If Err.number = 9 Then
      Err.Raise 9, , "There is nothing to paste."
   Else
      Err.Raise Err.number
   End If
   
End Sub


Public Sub delete(childCode As String, parentCode As String)
'Removes permanenttly the child group from it's parent. Garbage collection removes
'permanently the rest of the detached branch.

   unMap childCode, parentCode
   
End Sub
Public Sub unMap(childCode As String, parentCode As String)
'   Disconnect this child entity from its parent entity

    Call db.deleteFromGroupHierarchy(childCode, parentCode)
    
End Sub



Public Sub map(childCode As String, childType As String, parentCode As String, parentType As String)
'   Connect the child entity under the selected entity
    
    'If item is already mapped then skip
    If db.isMappedToParent(childCode, parentCode) Then Exit Sub
    
    'Otherwise connect item to the parent
    Call db.insertInGroupHierarchy(childCode, childType, parentCode, parentType)
        
End Sub


Public Function isCutPending() As Boolean

    isCutPending = cutPending

End Function
