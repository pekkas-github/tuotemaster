VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Services_Db"
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
Public Function getLastPriceLineCode() As String
' Return the code of the last record in PriceLine table (C0000nnn)
    
    If DCount("Id", PRICE_LINE) = 0 Then   ' No records yet in the table
        getLastPriceLineCode = "C0000000"
    Else
        getLastPriceLineCode = DMax("Id", PRICE_LINE)
    End If
        
End Function


Public Function GetLastEntityCode(entityPrefix) As String
' Return the last code of a given entity type.
    
    If DCount("Code", CORE_ITEM, "Code Like '" & entityPrefix & WILD_CARD & "'") = 0 Then
        GetLastEntityCode = entityPrefix & "000000"
    Else
         GetLastEntityCode = DMax("Code", CORE_ITEM, "Code Like '" & entityPrefix & WILD_CARD & "'")
    End If
    
End Function


Public Function GetVersionNumbers(entityCode As String) As Collection
' Return all version numbers of the entity in a collection

    Dim rst                 As New ADODB.Recordset
    Dim sql                 As String
    Dim number              As String
    Dim colVersionNumbers   As New Collection
    
    sql = "SELECT v.VersionNumber " & _
          "FROM " & CORE_ITEM & " AS i " & _
          "INNER JOIN " & ITEM_VERSION & " AS v ON v.Item_Code = i.Code " & _
          "WHERE i.Code = '" & entityCode & "'"
          
    rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
    
    While Not rst.EOF
        number = rst(0)
        colVersionNumbers.Add number
        rst.MoveNext
    Wend
    
    Set GetVersionNumbers = colVersionNumbers
    
    Set rst = Nothing
    Set colVersionNumbers = Nothing
    
End Function

Public Function getVersionIds(entityCode As String) As Collection
' Return all version ids of the entity in a collection

   Dim rst        As New ADODB.Recordset
   Dim sql        As String
   Dim verId      As Long
   Dim col        As Collection
   
   sql = "SELECT v.Id " & _
         "FROM " & CORE_ITEM & " AS i " & _
         "INNER JOIN " & ITEM_VERSION & " AS v ON v.Item_Code = i.Code " & _
         "WHERE i.Code = '" & entityCode & "'"
          
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   Set col = New Collection
   
   While Not rst.EOF
       verId = rst(0)
       col.Add verId
       rst.MoveNext
   Wend
   
   Set getVersionIds = col
   Set rst = Nothing
   Set col = Nothing
   
End Function


Public Sub deleteNames(entityCode As String)
   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & ITEM_NAME & " WHERE Item_Code = '" & entityCode & "'"
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   While Not rst.EOF
      rst.Delete
      rst.Update
      rst.MoveNext
   Wend
   
   Set rst = Nothing
   
End Sub


Public Sub deleteOwners(entityCode As String)
   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & ITEM_OWNER & " WHERE Item_Code = '" & entityCode & "'"
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   While Not rst.EOF
      rst.Delete
      rst.Update
      rst.MoveNext
   Wend
   
   Set rst = Nothing
   
End Sub


Public Sub deleteDescriptions(verId As Long)
   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & VERSION_DESCRIPTION & " WHERE ItemVersion_Id = " & verId
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   While Not rst.EOF
      rst.Delete
      rst.Update
      rst.MoveNext
   Wend
   
   Set rst = Nothing

End Sub


Public Sub deleteStatuses(verId As Long)
   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & VERSION_STATUS & " WHERE ItemVersion_Id = " & verId
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   While Not rst.EOF
      rst.Delete
      rst.Update
      rst.MoveNext
   Wend
   
   Set rst = Nothing

End Sub


Public Sub deleteDecisions(verId As Long)
   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & VERSION_DECISION & " WHERE ItemVersion_Id = " & verId
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   While Not rst.EOF
      rst.Delete
      rst.Update
      rst.MoveNext
   Wend
   
   Set rst = Nothing

End Sub


Public Sub deleteVersion(verId As Long)
   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & ITEM_VERSION & " WHERE Id = " & verId
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   While Not rst.EOF
      rst.Delete
      rst.Update
      rst.MoveNext
   Wend
   
   Set rst = Nothing

End Sub


Public Sub deleteProperties(entityCode As String)
   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & ITEM_PROPERTY_VALUE & " WHERE Item_Code = '" & entityCode & "'"
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   While Not rst.EOF
      rst.Delete
      rst.Update
      rst.MoveNext
   Wend
   
   Set rst = Nothing

End Sub


Public Sub deleteEntity(entityCode As String)
   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & CORE_ITEM & " WHERE Code = '" & entityCode & "'"
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   While Not rst.EOF
      rst.Delete
      rst.Update
      rst.MoveNext
   Wend
   
   Set rst = Nothing

End Sub
Private Sub Class_Terminate()

   Set dbconn = Nothing
   
End Sub
