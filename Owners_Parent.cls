VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Owners_Parent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Implements Owners_IF

' This class represents the owner behavior for those entities
' that are owned by their parents

Private db     As Owners_Db


Public Sub init()

   Set db = New Owners_Db
   
End Sub


Private Function Owners_IF_getOwner(entityCode As String) As Owner
' Return current owner of the parent entity.

   Dim rst  As ADODB.Recordset
   
   Set rst = db.getParentEntityOwner(entityCode)
   
   If rst.EOF Then
      Set Owners_IF_getOwner = Nothing     ' no owners
   Else
      Set Owners_IF_getOwner = new_Owner(rst!Id, rst!startDate)
   End If
   
End Function

Private Function Owners_IF_rsString(entityCode As String) As String
' Return as a rowsource an owner list of the parent entity of this entity.

   Owners_IF_rsString = db.getRowSourceOfParentEntityOwners(entityCode)
      
End Function

Private Function Owners_IF_rsStringAll() As String
' Rowsource that returns a list of the whole owners register

   Owners_IF_rsStringAll = db.getRowSourceOfAllOwners()
   
End Function

Private Sub Owners_IF_saveOwner(entityCode As String, newOwner As Owner)
' This action is not allowed for entities that don't have owners of theri own
' (ownership is on the parent level)

   Err.Raise e.actionNotAllowed, , "This entity cannot have owners of its own."
   
End Sub

Private Sub Class_Terminate()

   Set db = Nothing
   
End Sub
