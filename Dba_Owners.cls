VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dba_Owners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private app    As Application_API


Private Sub Class_Initialize()

   Set app = new_Application_API
   
End Sub


Public Function getThisEntityOwner(entityCode As String) As ADODB.Recordset

   Dim sql           As String
   Dim rst           As New ADODB.Recordset
   
   sql = "SELECT Party_Id AS id, Valid_From AS startDate FROM " & ITEM_OWNER & _
         " WHERE Item_Code = '" & entityCode & "' " & _
         "AND Valid_From <= " & TODAY & " AND Valid_To > " & TODAY
   
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic

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
   
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic

   Set getParentEntityOwner = rst
   Set rst = Nothing

End Function


Public Function getPersonName(personId As Long) As String
' Return the actual name of the owner

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT PersonName FROM " & PARTY & " WHERE Id = " & personId
   
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   getPersonName = rst(0)
   
   Set rst = Nothing

End Function


Public Sub insertOwner(entityCode As String, newOwner As dom_Owner)
   
   Dim rst  As New ADODB.Recordset
   
   rst.Open ITEM_OWNER, app.getDbConn, adOpenDynamic, adLockPessimistic
   
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
         
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   rst!Party_Id = currentOwner.getPersonId
   rst!valid_to = currentOwner.getEndDate
   rst.Update
   
   Set rst = Nothing
   
End Sub


Private Sub Class_Terminate()

   Set app = Nothing
   
End Sub

