VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Owners_Db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private dbconn As ADODB.Connection


Private Sub Class_Initialize()

   Set dbconn = CurrentProject.Connection
   
End Sub

Public Function getThisEntityOwner(entityCode As String) As ADODB.Recordset

   Dim sql           As String
   Dim rst           As New ADODB.Recordset
   
   sql = "SELECT Party_Id AS id, Valid_From AS startDate FROM " & ITEM_OWNER & _
         " WHERE Item_Code = '" & entityCode & "' " & _
         "AND Valid_From <= " & TODAY & " AND Valid_To > " & TODAY
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic

   Set getThisEntityOwner = rst
   Set rst = Nothing

End Function


Public Function getParentEntityOwner(entityCode As String) As ADODB.Recordset
' Return current owner of the parent entity. Actually works only for Sales Items because
' there is an explcit MR condition in the query.
' Return empty object if there is no owners yet

   Dim sql           As String
   Dim rst           As New ADODB.Recordset
   
   sql = "SELECT Party_Id AS id, Valid_From AS startDate FROM " & ITEM_OWNER & " AS o " & _
         "INNER JOIN " & GROUP_HIERARCHY & " AS h ON (h.Parent_Code = o.Item_Code AND h.Parent_Type = 'MR') " & _
         "WHERE h.Child_Code = '" & entityCode & "' " & _
         "AND o.Valid_From <= " & TODAY & " AND o.Valid_To > " & TODAY
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic

   Set getParentEntityOwner = rst
   Set rst = Nothing

End Function


Public Function getRowSourceOfParentEntityOwners(entityCode As String) As String

   Dim sql As String
   
   sql = "SELECT p.Id, p.PersonName, o.Valid_From AS StartDate " & _
         "FROM " & ITEM_OWNER & " AS o " & _
         "INNER JOIN " & PARTY & " AS p ON o.Party_Id = p.Id " & _
         "INNER JOIN " & GROUP_HIERARCHY & " AS h ON h.Parent_Code = o.Item_Code " & _
         "WHERE h.Child_Code = '" & entityCode & "' AND Parent_Type = 'MR' " & _
         "ORDER BY o.Id DESC"
         
   getRowSourceOfParentEntityOwners = sql

End Function


Public Function getRowSourceOfThisEntityOwners(entityCode As String) As String
   
   Dim sql As String
   
   sql = "SELECT p.Id, p.PersonName, o.Valid_From AS StartDate " & _
         "FROM " & ITEM_OWNER & " AS o " & _
         "INNER JOIN " & PARTY & " AS p ON o.Party_Id = p.Id " & _
         "WHERE o.Item_Code = '" & entityCode & "' " & _
         "ORDER BY o.Id DESC"
               
   getRowSourceOfThisEntityOwners = sql

End Function


Public Function getRowSourceOfAllOwners() As String
   
   Dim sql As String
   
   sql = "SELECT DISTINCT p.Id, p.PersonName " & _
         "FROM " & ITEM_OWNER & " AS o " & _
         "INNER JOIN " & PARTY & " AS p ON o.Party_Id = p.Id " & _
         "ORDER BY p.PersonName"

   getRowSourceOfAllOwners = sql

End Function


Public Sub insertOwner(entityCode As String, newOwner As Owner)
   
   Dim rst  As New ADODB.Recordset
   
   rst.Open ITEM_OWNER, dbconn, adOpenDynamic, adLockPessimistic
   
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

Public Sub updateOwner(entityCode, currentOwner)
   
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


Private Sub Class_Terminate()

   Set dbconn = Nothing
   
End Sub
