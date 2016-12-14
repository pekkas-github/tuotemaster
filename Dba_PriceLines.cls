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
' Return price line data with finnis price line text.
' Business rule is that we use only finnish price line texts.
' Database enables however any language.

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT p.Id, p.ItemVersion_Id, t.Content, p.BizTypeBasic_Id, p.BizTypeOverride_Id, p.SAPCode " & _
      "FROM " & PRICE_LINE & " AS p " & _
      "INNER JOIN " & PRICE_LINE_TEXT & " AS t ON p.Id = t.PriceLine_Id " & _
      "WHERE p.Id = '" & priceLineCode & "' AND t.Lang = 'fin'"
    
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   Set getPriceLine = rst
   Set rst = Nothing

End Function


Public Function getPriceListEntry(id As Long) As ADODB.Recordset

    Dim sql As String
    Dim rst As New ADODB.Recordset

    sql = "SELECT * FROM " & PRICE_LIST_ENTRY & _
          " WHERE Id = " & id
    
    rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
    
    Set getPriceListEntry = rst
    
    Set rst = Nothing

End Function


Public Function getActivePriceListEntries(priceLineId As String, Optional priceListId As String = "NONE") As ADODB.Recordset
'   Returns per each priceList or from a selected priceList a recordset with
'   - no records if there is no effective or coming priceListEntries
'   - all records where validTo date is greater or equal to today i.e. current or coming priceListEntry

    Dim sql As String
    Dim rst As New ADODB.Recordset
    
    sql = "SELECT * FROM " & PRICE_LIST_ENTRY & _
          " WHERE PriceLine_Id = '" & priceLineId & "' " & _
          "AND validTo > " & TODAY & "-1 "
    
'   Get from all priceLists or from a specific one
    If Not priceListId = "NONE" Then
        sql = sql & "AND PriceList_Id = '" & priceListId & "' "
    End If
    
'   Group each priceList with latest priceEntry first
    sql = sql & "ORDER BY PriceList_Id, ValidTo DESC"
          
    rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
    
    Set getActivePriceListEntries = rst
    Set rst = Nothing
    

End Function


Public Sub insertPriceLine(newPriceLine As priceLine, versionId As Long)
    
    Dim rst As New ADODB.Recordset
    
    rst.Open PRICE_LINE, dbconn, adOpenDynamic, adLockPessimistic
    
    With newPriceLine
'     Add new price line
      rst.AddNew
      rst!id = .getObjectId
      rst!ItemVersion_Id = versionId
      rst!BizTypeBasic_Id = .getBusinessTypeBasic
      rst!BizTypeOverride_Id = .getBusinessTypeOverride
      rst!sapCode = .getSapCode
      rst.Update
      rst.Close
'     Add new price line text
      rst.Open PRICE_LINE_TEXT, dbconn, adOpenDynamic, adLockPessimistic
      rst.AddNew
      rst!PriceLine_Id = .getObjectId
      rst!lang = "fin"
      rst!Content = .getPriceLineText
      rst.Update
      rst.Close
    End With
    
    Set rst = Nothing
    
End Sub

Public Sub updatePriceLine(newPriceLine As priceLine)

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
'  Update price line
   sql = "SELECT * FROM " & PRICE_LINE & " WHERE Id = '" & newPriceLine.getObjectId & "'"
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
    
   With newPriceLine
      rst!BizTypeBasic_Id = .getBusinessTypeBasic
      rst!BizTypeOverride_Id = .getBusinessTypeOverride
      rst!sapCode = .getSapCode
      rst.Update
      rst.Close
   End With

'  Update price line text
   sql = "SELECT * FROM " & PRICE_LINE_TEXT & _
         " WHERE PriceLine_Id = '" & newPriceLine.getObjectId & "' AND Lang ='fin'"
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   rst!Content = newPriceLine.getPriceLineText
   rst.Update
   
   Set rst = Nothing

End Sub

