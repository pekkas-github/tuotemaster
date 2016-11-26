VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Owners_Self"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Implements If_Owners

' This class represents the owner behavior for those entities
' that have their own explisit ownership

Private db  As Dba_Owners
'

Public Sub init()

   Set db = New Dba_Owners

End Sub


Private Sub Class_Terminate()

   Set db = Nothing
   
End Sub

'---- METHODS ---------------------

Private Function getOwner(entityCode As String) As Owner
' Return the current owner of the entity.

   Dim rst  As ADODB.Recordset
   
   Set rst = db.getThisEntityOwner(entityCode)
   
   If rst.EOF Then
      Set getOwner = Nothing     ' no owners
   Else
      Set getOwner = new_Owner(rst!Id, rst!startDate)
   End If
   
   Set rst = Nothing
   
End Function


Private Sub saveOwner(entityCode As String, newOwner As Owner)
' Add a new owner or change the current owner

   Dim currentOwner  As Owner
   
   Set currentOwner = getOwner(entityCode)
   
   If currentOwner Is Nothing Then
'     no previous owners
      Call db.insertOwner(entityCode, newOwner)
   ElseIf newOwner.getStartDate = currentOwner.getStartDate Then
'     Current owner with the same start date -> update...
      currentOwner.setPersonId (newOwner.getPersonId)
      Call db.updateOwner(entityCode, currentOwner)
      Else
'        ...or not the same start date -> update current end date
         currentOwner.setEndDate (newOwner.getStartDate - 1)
         Call db.updateOwner(entityCode, currentOwner)
         Call db.insertOwner(entityCode, newOwner)
   End If
         
   Set currentOwner = Nothing

End Sub

'---- INTERFACE CALLS -----------------------------

Private Function If_Owners_getOwner(entityCode As String) As Owner

   Set If_Owners_getOwner = getOwner(entityCode)
   
End Function

Private Sub If_Owners_saveOwner(entityCode As String, newOwner As Owner)

   Call saveOwner(entityCode, newOwner)
   
End Sub
