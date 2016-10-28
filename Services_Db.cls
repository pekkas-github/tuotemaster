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

Private Sub Class_Terminate()

   Set dbconn = Nothing
   
End Sub
