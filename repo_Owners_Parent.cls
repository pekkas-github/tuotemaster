VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_Owners_Parent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Implements If_Owners

' This class represents the owner behavior for those entities
' that are owned by their parents

Private db     As Dba_Owners
'

Private Sub Class_Initialize()

   Set db = new_Dba_Owners
   
End Sub

Private Sub Class_Terminate()

   Set db = Nothing
   
End Sub

'---- METHODS -----------------
Private Function getOwner(entityCode As String) As dom_Owner
' Return current owner of the parent entity.

   Dim rst  As ADODB.Recordset
   
   Set rst = db.getParentEntityOwner(entityCode)
   
   If rst.EOF Then
      Set getOwner = Nothing     ' no owners
   Else
      Set getOwner = new_Owner(rst!id, rst!startDate)
   End If
   
End Function


Private Sub saveOwner(entityCode As String, newOwner As dom_Owner)
' This action is not allowed for entities that don't have owners of theri own
' (ownership is on the parent level)

   Err.Raise e.actionNotAllowed, , "This entity cannot have owners of its own."
   
End Sub

'---- INTERFACE CALLS ---------------------

Private Sub If_Owners_init()
   

End Sub

Private Function If_Owners_getOwner(entityCode As String) As dom_Owner

   Set If_Owners_getOwner = getOwner(entityCode)
   
End Function

Private Sub If_Owners_saveOwner(entityCode As String, newOwner As dom_Owner)

   Call saveOwner(entityCode, newOwner)
   
End Sub
