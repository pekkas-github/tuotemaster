VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_Owners_Self"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Implements If_Owners

' This class represents the owner behavior for those entities
' that have their own explisit ownership

Private oEntityCode  As String
Private db           As Dba_Owners
'

Public Sub init(entityCode As String)

   oEntityCode = entityCode
   Set db = new_Dba_Owners

End Sub


Private Sub Class_Terminate()

   Set db = Nothing
   
End Sub

'---- METHODS ---------------------

Public Function getOwner() As dom_Owner
' Return the current owner of the entity.

   Dim rst  As ADODB.Recordset
   
   Set rst = db.getThisEntityOwner(oEntityCode)
   
   If rst.EOF Then
      Set getOwner = Nothing     ' no owners
   Else
      Set getOwner = new_Owner(rst!id, rst!startDate)
   End If
   
   Set rst = Nothing
   
End Function


Private Sub saveOwner(newOwner As dom_Owner)
' Add a new owner or change the current owner

   Dim currentOwner  As dom_Owner
   
   Set currentOwner = getOwner()
   
   If currentOwner Is Nothing Then
'     no previous owners
      Call db.insertOwner(oEntityCode, newOwner)
   ElseIf newOwner.getStartDate = currentOwner.getStartDate Then
'     Current owner with the same start date -> update...
      currentOwner.setPersonId (newOwner.getPersonId)
      Call db.updateOwner(oEntityCode, currentOwner)
      Else
'        ...or not the same start date -> update current end date
         currentOwner.setEndDate (newOwner.getStartDate - 1)
         Call db.updateOwner(oEntityCode, currentOwner)
         Call db.insertOwner(oEntityCode, newOwner)
   End If
         
   Set currentOwner = Nothing

End Sub

'---- INTERFACE CALLS -----------------------------

Private Sub If_Owners_init(entityCode As String)

   init entityCode
   
End Sub

Private Function If_Owners_getOwner() As dom_Owner

   Set If_Owners_getOwner = getOwner()
   
End Function

Private Sub If_Owners_saveOwner(newOwner As dom_Owner)

   Call saveOwner(newOwner)
   
End Sub
