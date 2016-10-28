VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Owners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private dbconn       As ADODB.Connection


'constructor
Public Sub init()

   Set dbconn = CurrentProject.Connection
   
End Sub


Public Function getItem(entityCode As String) As Owner
' Return the current owner of the entity.
' Return empty object if there is no owner yet

   Dim sql           As String
   Dim rst           As New ADODB.Recordset
   
   sql = "SELECT * FROM " & ITEM_OWNER & _
         " WHERE Item_Code = '" & entityCode & "' " & _
         "AND Valid_From <= " & TODAY & " AND Valid_To > " & TODAY
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic

   If rst.EOF Then
      Set getItem = Nothing
   Else
      Set getItem = new_Owner(rst!Party_Id, rst!Valid_From)
   End If
   
   Set rst = Nothing
   
End Function


Public Sub saveItem(entityCode As String, newOwner As Owner)
' Add a new owner or change the current owner

   Dim currentOwner  As Owner
   
   Set currentOwner = getItem(entityCode)
   
   If currentOwner Is Nothing Then
'     no previous owners
      Call insertOwner(entityCode, newOwner)
   ElseIf newOwner.getStartDate = currentOwner.getStartDate Then
'     Current owner with the same start date -> update...
      currentOwner.setPersonId (newOwner.getPersonId)
      Call updateOwner(entityCode, currentOwner)
      Else
'        ...or not the same start date -> update current end date
         currentOwner.setEndDate (newOwner.getStartDate - 1)
         Call updateOwner(entityCode, currentOwner)
         Call insertOwner(entityCode, newOwner)
   End If
         
   Set currentOwner = Nothing

End Sub


Private Sub insertOwner(entityCode As String, newOwner As Owner)
' Insert a new owner

   Dim rst  As New ADODB.Recordset
   
   rst.Open "ItemOwner", dbconn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !Item_Code = entityCode
      !Party_Id = newOwner.getPersonId
      !Valid_From = newOwner.getStartDate
      !valid_to = newOwner.getEndDate
      .Update
   End With
   
   Set rst = Nothing
        
End Sub


Private Sub updateOwner(entityCode As String, currentOwner As Owner)
' Update
'     Party_Id (owner is changed during the same day)
'     Valid_To (new owner has replaced the current one)

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT * FROM " & ITEM_OWNER & _
         " WHERE Item_Code = '" & entityCode & "' " & _
         "AND Valid_From <= " & TODAY & " AND Valid_To > " & TODAY
         
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   rst!Party_Id = currentOwner.getPersonId
   rst!valid_to = currentOwner.getEndDate
   rst.Update
   
   Set rst = Nothing
   
End Sub



Public Function rsStringAll() As String
' Rowsource that returns current owners of all entities

   Dim sql As String
   
   sql = "SELECT DISTINCT p.Id, p.PersonName " & _
         "FROM ItemOwner AS o " & _
         "INNER JOIN ltbPerson AS p ON o.Party_Id = p.Id " & _
         "ORDER BY p.PersonName"

   rsStringAll = sql
   
End Function

Public Function rsString(entityCode As String) As String
' Rowsource that returns owner history of a specific entity
   
   Dim sql As String
   
   sql = "SELECT p.Id, p.PersonName, o.Valid_From AS StartDate " & _
         "FROM ItemOwner AS o " & _
         "INNER JOIN ltbPerson AS p ON o.Party_Id = p.Id " & _
         "WHERE o.Item_Code = '" & entityCode & "' " & _
         "ORDER BY o.Id DESC"
               
   rsString = sql

End Function


Private Sub Class_Terminate()

   Set dbconn = Nothing
   
End Sub
