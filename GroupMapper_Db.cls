VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupMapper_Db"
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
Public Function IsMappedToParent(ChildEntity As Object, ParentEntity As Object) As Boolean
' Tarkistaa, onko child-entity jo liitetty parent-ryhm‰‰n.
'**
    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & GROUP_HIERARCHY & _
          " WHERE Child_Code = '" & ChildEntity.getCode & "' AND Parent_Code = '" & ParentEntity.getCode & "'"
          
    rst.Open sql, dbconn, adOpenKeyset
    
    If rst.EOF Then
        IsMappedToParent = False
    Else
        IsMappedToParent = True
    End If
    
    Set rst = Nothing
    
End Function


Public Sub InsertInGroupHierarchy(ChildEntity As Object, ParentEntity As Object)
' Maps the child group or item reference to the parent entity

    Dim rst As New ADODB.Recordset
    
    rst.ActiveConnection = CurrentProject.Connection
    rst.Open GROUP_HIERARCHY, dbconn, adOpenDynamic, adLockPessimistic
    
    rst.AddNew
    rst!Child_Code = ChildEntity.getCode
    rst!Child_Type = ChildEntity.getType
    rst!Parent_Code = ParentEntity.getCode
    rst!Parent_Type = ParentEntity.getType
    
    rst.Update
    rst.Close
    Set rst = Nothing

End Sub


Public Sub DeleteFromGroupHierarchy(ChildEntity As Object, ParentEntity As Object)
'   Cut the parent-child connection. Child branch stays orphan unless it is later
'   connected to another parent group. Orphan groups are cleaned by garbage collection
'   in "DeleteOrphanGroupings"

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM GroupHierarchy " & _
          "WHERE Child_Code = '" & ChildEntity.code & "' AND Parent_Code = '" & ParentEntity.code & "'"
          
    rst.Open sql, dbconn, , adLockPessimistic
    rst.Delete
    rst.Update

    Set rst = Nothing

End Sub

Private Sub Class_Terminate()

   Set dbconn = Nothing
   
End Sub
