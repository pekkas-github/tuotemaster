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


Public Function isMappedToParent(childEntity As Object, parentEntity As Object) As Boolean
' Chack if the child is already mapped to the parent

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & GROUP_HIERARCHY & _
          " WHERE Child_Code = '" & childEntity.getCode & "' AND Parent_Code = '" & parentEntity.getCode & "'"
          
    rst.Open sql, dbconn, adOpenKeyset
    
    If rst.EOF Then
        isMappedToParent = False
    Else
        isMappedToParent = True
    End If
    
    Set rst = Nothing
    
End Function


Public Sub insertInGroupHierarchy(childEntity As Object, parentEntity As Object)
' Maps the child group or item reference to the parent entity

    Dim rst As New ADODB.Recordset
    
    rst.Open GROUP_HIERARCHY, dbconn, adOpenDynamic, adLockPessimistic
    
    rst.AddNew
    rst!Child_Code = childEntity.getCode
    rst!Child_Type = childEntity.getType
    rst!Parent_Code = parentEntity.getCode
    rst!Parent_Type = parentEntity.getType
    
    rst.Update
    rst.Close
    Set rst = Nothing

End Sub


Public Sub deleteFromGroupHierarchy(childEntity As Object, parentEntity As Object)
'   Cut the parent-child connection. Child branch stays orphan unless it is later
'   connected to another parent group. Orphan groups are cleaned by garbage collection
'   in "DeleteOrphanGroupings"

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & Group_Hierarchy & _
          " WHERE Child_Code = '" & childEntity.code & "' AND Parent_Code = '" & parentEntity.code & "'"
          
    rst.Open sql, dbconn, , adLockPessimistic
    rst.Delete
    rst.Update

    Set rst = Nothing

End Sub


Private Sub Class_Terminate()

   Set dbconn = Nothing
   
End Sub
