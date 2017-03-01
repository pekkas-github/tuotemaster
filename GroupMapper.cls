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


Public Sub cut(parentCode As String, childCodes As Collection)
' Cut all child groups selected in the collection from their parent.
' Parent may be ROOT, CAT or GRP.
' Collection remains on clipboard until next cut, copy or garbage collection takes place.

    Dim childCode As Variant
    
    cutPending = True
    Set Globals.clipBoard = childCodes
    
    For Each childCode In childCodes
        Call unMap(CStr(childCode), parentCode)
    Next
            
End Sub


Public Sub paste(parentCode As String, parentType As String, parentSubType As String)
' Paste the previously copied/cut child collection to this parent
' Before pasting check that it is allowed for child type and parent sub type match.
' Cut (or copied) collection remains on clipboard.

    Dim childCode       As Variant
    Dim childType       As String
    Dim util            As Services
    
    Set util = new_Services
    
    For Each childCode In Globals.clipBoard
        childType = util.getEntityType(CStr(childCode))
        If parentSubType = childType Then
            Call map(CStr(childCode), childType, parentCode, parentType)
        Else
            Err.Raise vbObjectError + 513, , "Parent and child items are incompatible."
        End If
    Next
        
    cutPending = False
    
End Sub


Public Sub copy(childCodes As Collection)
' Copy all child groups selected in the collection.
' Collection remains on clipboard until next cut, copy or garbage collection takes place.

    
    Set Globals.clipBoard = childCodes
    
    cutPending = False
            
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
