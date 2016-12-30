VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dba_GroupMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private app    As Application_API
Private dbConn As ADODB.Connection


Private Sub Class_Initialize()
'TODO: Replace this with init
   Set app = New Application_API
   Set dbConn = CurrentProject.Connection
   
End Sub

Public Sub init(conn As ADODB.Connection)
   
   Set app = New Application_API
   Set dbConn = conn
   
End Sub

Public Function isMappedToParent(childEntity As Object, parentEntity As Object) As Boolean
' Chack if the child is already mapped to the parent

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & GROUP_HIERARCHY & _
          " WHERE Child_Code = '" & childEntity.getCode & "' AND Parent_Code = '" & parentEntity.getCode & "'"
          
    rst.Open sql, dbConn, adOpenKeyset
    
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
    
    rst.Open GROUP_HIERARCHY, dbConn, adOpenDynamic, adLockPessimistic
    
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
'   in "deleteOrphanGroupings" while closing the group mapper.

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT * FROM " & GROUP_HIERARCHY & _
          " WHERE Child_Code = '" & childEntity.getCode & "' AND Parent_Code = '" & parentEntity.getCode & "'"
          
    rst.Open sql, dbConn, , adLockPessimistic
    rst.Delete
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
   
   Do
      Set rst = New ADODB.Recordset
      rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
      While Not rst.EOF
         rst.Delete
         rst.Update
         rst.MoveNext
      Wend
   Loop Until rst.BOF

   rst.Close
'TODO: poistetaan myös entiteetti ja sen attribuutit (löytyy Item-taulusta, mutta ei löydy
'      GroupHierarchy-taulun childeista. Koskee jälleen vain CAT ja GRP tyyppejä.

   sql = "SELECT i.* " & _
         "FROM " & CORE_ITEM & " AS i " & _
         "LEFT JOIN " & GROUP_HIERARCHY & " AS g ON i.Code = g.Child_Code " & _
         "WHERE g.Child_Code Is Null AND (i.Type = 'CAT' OR i.Type = 'GRP')"
         
   rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
   
   While Not rst.EOF
      Call app.Services.deleteEntityPermanently(rst!code)
      rst.MoveNext
   Wend
   
   Set rst = Nothing
   
End Sub
Private Sub Class_Terminate()

   Set dbConn = Nothing
   Set app = Nothing
   
End Sub
