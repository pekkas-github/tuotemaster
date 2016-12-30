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

Private dbConn As ADODB.Connection

Private Sub Class_Initialize()
'TODO: Replace with init

   Set dbConn = CurrentProject.Connection
   
End Sub


Public Sub init(conn As ADODB.Connection)
   Set dbConn = conn

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
    
   rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
   
   Set getPriceLine = rst
   Set rst = Nothing

End Function


Public Function getPriceListEntry(id As Long) As ADODB.Recordset

    Dim sql As String
    Dim rst As New ADODB.Recordset

    sql = "SELECT * FROM " & PRICE_LIST_ENTRY & _
          " WHERE Id = " & id
    
    rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
    
    Set getPriceListEntry = rst
    
    Set rst = Nothing

End Function


Public Function getCurrentPriceListEntry(priceLineCode As String) As ADODB.Recordset
'   Return currently valid price list entry; cases:
'   - no records if there is no priceListEntries
'   - data of the current priceListEntry in STD pricelist

    Dim sql As String
    Dim rst As New ADODB.Recordset
    
    sql = "SELECT * FROM " & PRICE_LIST_ENTRY & _
          " WHERE PriceLine_Id = '" & priceLineCode & "' " & _
          "AND ValidFrom <= " & TODAY & " AND ValidTo > " & TODAY & "-1 "
              
    rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
    
    Set getCurrentPriceListEntry = rst
    Set rst = Nothing
    

End Function


Public Function getLatestPriceListEntry(priceLineCode As String) As ADODB.Recordset
'   Return the latest price list entry; cases:
'   - no records if there is no priceListEntries
'   - data of the current priceListEntry in STD pricelist
'   - data of a future priceListEntry in STD priceList

    Dim sql As String
    Dim rst As New ADODB.Recordset
    
    sql = "SELECT * FROM " & PRICE_LIST_ENTRY & _
          " WHERE PriceLine_Id = '" & priceLineCode & "' " & _
          "ORDER BY Id DESC"

'  The latest will be on the top of the recordset
    rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
    
    Set getLatestPriceListEntry = rst
    Set rst = Nothing

End Function
Public Sub insertPriceLine(newPriceLine As PriceLine, versionId As Long)
    
    Dim rst As New ADODB.Recordset
    
    rst.Open PRICE_LINE, dbConn, adOpenDynamic, adLockPessimistic
    
    With newPriceLine
'     Add new price line
      rst.AddNew
      rst!id = .getCode
      rst!ItemVersion_Id = versionId
      rst!BizTypeBasic_Id = .getBusinessTypeBasic
      rst!BizTypeOverride_Id = .getBusinessTypeOverride
      rst!sapCode = .getSapCode
      rst.Update
      rst.Close
'     Add new price line text
      rst.Open PRICE_LINE_TEXT, dbConn, adOpenDynamic, adLockPessimistic
      rst.AddNew
      rst!PriceLine_Id = .getCode
      rst!lang = "fin"
      rst!Content = .getPriceLineText
      rst.Update
      rst.Close
    End With
    
    Set rst = Nothing
    
End Sub

Public Sub updatePriceLine(newPriceLine As PriceLine)

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
'  Update price line
   sql = "SELECT * FROM " & PRICE_LINE & " WHERE Id = '" & newPriceLine.getCode & "'"
   rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
    
   With newPriceLine
      rst!BizTypeBasic_Id = .getBusinessTypeBasic
      rst!BizTypeOverride_Id = .getBusinessTypeOverride
      rst!sapCode = .getSapCode
      rst.Update
      rst.Close
   End With

'  Update price line text
   sql = "SELECT * FROM " & PRICE_LINE_TEXT & _
         " WHERE PriceLine_Id = '" & newPriceLine.getCode & "' AND Lang ='fin'"
   rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
   
   rst!Content = newPriceLine.getPriceLineText
   rst.Update
   
   Set rst = Nothing

End Sub


Public Sub insertPriceListEntry(priceEntry As PriceListEntry)

   Dim rst  As New ADODB.Recordset
   
   rst.Open PRICE_LIST_ENTRY, dbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !PriceLine_Id = priceEntry.getPriceLineId
      !PriceList_Id = priceEntry.getPriceListId
      !price = priceEntry.getPrice
      !priceUnit = priceEntry.getUnit
      !priceType = priceEntry.getType
      !ValidFrom = priceEntry.getValidFrom
      !ValidTo = priceEntry.getValidTo
      !vma = priceEntry.getVMA
      !VMA_boost = priceEntry.getVMABoost
      !active = priceEntry.getActivity
      !comment = priceEntry.getComment
      .Update
      .Close
   End With
   
   Set rst = Nothing
   
End Sub


Public Sub updatePriceListEntry(priceEntry As PriceListEntry)

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT * FROM " & PRICE_LIST_ENTRY & " WHERE Id = " & priceEntry.getId
   
   rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      !price = priceEntry.getPrice
      !priceUnit = priceEntry.getUnit
      !priceType = priceEntry.getType
      !ValidFrom = priceEntry.getValidFrom
      !ValidTo = priceEntry.getValidTo
      !vma = priceEntry.getVMA
      !VMA_boost = priceEntry.getVMABoost
      !active = priceEntry.getActivity
      !comment = priceEntry.getComment
      .Update
      .Close
   End With
   
   Set rst = Nothing
   
End Sub
