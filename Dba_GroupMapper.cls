VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dba_GroupMapper"
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


Public Function isMappedToParent(childCode As String, parentCode As String) As Boolean
' Check if the child is already mapped to the parent

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & GROUP_HIERARCHY & _
          " WHERE Child_Code = '" & childCode & "' AND Parent_Code = '" & parentCode & "'"
          
    rst.Open sql, app.getDbConn, adOpenKeyset
    
    If rst.EOF Then
        isMappedToParent = False
    Else
        isMappedToParent = True
    End If
    
    Set rst = Nothing
    
End Function


Public Function isMappedToParent_old(childEntity As Object, parentEntity As Object) As Boolean
' WILL BE DELETED; NAME HAS BEEN CHANGED
' Check if the child is already mapped to the parent

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & GROUP_HIERARCHY & _
          " WHERE Child_Code = '" & childEntity.getCode & "' AND Parent_Code = '" & parentEntity.getCode & "'"
          
    rst.Open sql, app.getDbConn, adOpenKeyset
    
    If rst.EOF Then
        isMappedToParent_old = False
    Else
        isMappedToParent_old = True
    End If
    
    Set rst = Nothing
    
End Function
Public Sub insertInGroupHierarchy(childCode As String, childType As String, parentCode As String, parentType As String)
' Maps the child group or item reference to the parent entity

    Dim rst As New ADODB.Recordset
    
    rst.Open GROUP_HIERARCHY, app.getDbConn, adOpenDynamic, adLockPessimistic
    
    rst.AddNew
    rst!Child_Code = childCode
    rst!Child_Type = childType
    rst!Parent_Code = parentCode
    rst!Parent_Type = parentType
    
    rst.Update
    rst.Close
    Set rst = Nothing

End Sub


Public Sub insertInGroupHierarchy_old(childEntity As Object, parentEntity As Object)
' WILL BE DELETED; NAME HAS BEEN CHANGED
' Maps the child group or item reference to the parent entity

    Dim rst As New ADODB.Recordset
    
    rst.Open GROUP_HIERARCHY, app.getDbConn, adOpenDynamic, adLockPessimistic
    
    rst.AddNew
    rst!Child_Code = childEntity.getCode
    rst!Child_Type = childEntity.getType
    rst!Parent_Code = parentEntity.getCode
    rst!Parent_Type = parentEntity.getType
    
    rst.Update
    rst.Close
    Set rst = Nothing

End Sub

Public Sub deleteFromGroupHierarchy(childCode As String, parentCode As String)
'   Cut the parent-child connection. Child branch stays orphan unless it is later
'   connected to another parent group. Orphan groups are cleaned by garbage collection
'   in "deleteOrphanGroupings" while closing the group mapper.

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & GROUP_HIERARCHY & _
          " WHERE Child_Code = '" & childCode & "' AND Parent_Code = '" & parentCode & "'"
          
    rst.Open sql, app.getDbConn, , adLockPessimistic
    rst.delete
    rst.Update

    Set rst = Nothing

End Sub


Public Sub deleteFromGroupHierarchy_old(childEntity As Object, parentEntity As Object)
' WILL BE DELETED; NAME HAS BEEN CHANGED
'   Cut the parent-child connection. Child branch stays orphan unless it is later
'   connected to another parent group. Orphan groups are cleaned by garbage collection
'   in "deleteOrphanGroupings" while closing the group mapper.

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & GROUP_HIERARCHY & _
          " WHERE Child_Code = '" & childEntity.getCode & "' AND Parent_Code = '" & parentEntity.getCode & "'"
          
    rst.Open sql, app.getDbConn, , adLockPessimistic
    rst.delete
    rst.Update

    Set rst = Nothing

End Sub


Public Sub deleteOrphanGroups()
' Loop iterativly through group hierarchies and delete all orphan elements.
' Orphan means that entity exists as a parent but not as a child (= it doesn't have parent).
' ROOT and PROD are exceptions because they are the most top elements of hierarchies.

   Dim sql  As String
   Dim rst  As ADODB.Recordset
   
   
   sql = "SELECT g1.Id " & _
         "FROM " & GROUP_HIERARCHY & " AS g1 " & _
         "LEFT JOIN " & GROUP_HIERARCHY & " AS g2 ON g1.Parent_Code = g2.Child_Code " & _
         "WHERE g2.Child_Code Is Null AND g1.Parent_Code <> 'ROOT' AND g1.Parent_Code <> 'PROD'"

'  Find and delete orphans repeatly until all orphans are removed from hierarchies.
   Do
      Set rst = New ADODB.Recordset
      rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
      While Not rst.EOF
         rst.delete
         rst.Update
         rst.MoveNext
      Wend
   Loop Until rst.BOF

   rst.Close

'  Delete permanently all group entities that doesn't exist in any hierarchy any more
   sql = "SELECT i.* " & _
         "FROM " & CORE_ITEM & " AS i " & _
         "LEFT JOIN " & GROUP_HIERARCHY & " AS g ON i.Code = g.Child_Code " & _
         "WHERE g.Child_Code Is Null AND (i.Type = 'CAT' OR i.Type = 'GRP')"
         
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   While Not rst.EOF
      Call app.Services.deleteEntityPermanently(rst!code)
      rst.MoveNext
   Wend
   
   Set rst = Nothing
   
End Sub
Private Sub Class_Terminate()

   Set app = Nothing
   
End Sub
