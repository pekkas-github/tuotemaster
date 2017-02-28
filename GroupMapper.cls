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
Private db                      As Dba_GroupMapper
Private cutPending              As Boolean


'---- INIT AND TERMINATE --------------------
Private Sub Class_Initialize()
    
    Set db = New Dba_GroupMapper

    cutPending = False
    
End Sub

Private Sub Class_Terminate()
' Clean the group hierarchies before closing by deleting all orphan elements.

   db.deleteOrphanGroups
   Set db = Nothing
   
End Sub

'---- METHODS -------------------------------

Public Sub cut(parentEntity As Entity, childEntities As Collection)
' Cut all child groups selected in the collection from their parent.
' Parent may be ROOT, CAT or GRP.
' Collection remains on clipboard until next cut, copy or garbage collection takes place.

    Dim childEntity As Entity
    
    cutPending = True
    Set Globals.clipBoard = New Collection
    
    For Each childEntity In childEntities
        Call unMap(childEntity, parentEntity)
        Globals.clipBoard.Add childEntity
    Next
            
End Sub


Public Sub paste(parentEntity As Entity)
' Paste the previously copied/cut child collection to this parent
' Before pasting check that it is allowed for child type and parent sub type match.
' Cut (or copied) collection remains on clipboard.

    Dim childEntity     As Entity
    Dim parentSubType   As String
    Dim childType       As String
    
    parentSubType = parentEntity.getProperty("SUB_MEM").getValue
    
    For Each childEntity In Globals.clipBoard
        childType = childEntity.getType
        If parentSubType = childType Then
            Call map(childEntity, parentEntity)
        Else
            Err.Raise vbObjectError + 513, , "Parent and child items are incompatible."
        End If
    Next
        
    cutPending = False
    
End Sub


Public Sub copy(childEntities As Collection)
' Copy all child groups selected in the collection.
' Collection remains on clipboard until next cut, copy or garbage collection takes place.

    Dim childEntity As Entity
    
    Set Globals.clipBoard = New Collection
    
    For Each childEntity In childEntities
        Globals.clipBoard.Add childEntity
    Next

        cutPending = False
            
End Sub


Public Sub unMap(childEntity As Entity, parentEntity As Entity)
'   Disconnect this child entity from its parent entity

    Call db.deleteFromGroupHierarchy(childEntity, parentEntity)
    
End Sub


Public Sub map(childEntity As If_Entity, parentEntity As If_Entity)
'   Connect the child entity under the selected entity
    
    'If item is already mapped then skip
    If db.isMappedToParent(childEntity, parentEntity) Then Exit Sub
    
    'Otherwise connect item to the parent
    Call db.insertInGroupHierarchy(childEntity, parentEntity)
        
End Sub


Public Function isCutPending() As Boolean

    isCutPending = cutPending

End Function
