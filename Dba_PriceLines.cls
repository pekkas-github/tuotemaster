VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dba_PriceLines"
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

Public Function getPriceLine(priceLineCode As String) As ADODB.Recordset

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT p.Id, p.ItemVersion_Id, t.Content, p.BizTypeBasic_Id, p.BizTypeOverride_Id, p.SAPCode " & _
      "FROM " & PRICE_LINE & " AS p " & _
      "INNER JOIN " & PRICE_LINE_TEXT & " AS t ON p.Id = t.PriceLine_Id " & _
      "WHERE p.Id = '" & priceLineCode & "'"
    
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   Set getPriceLine = rst
   Set rst = Nothing

End Function
