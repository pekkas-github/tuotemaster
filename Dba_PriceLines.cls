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

Private app    As Application_API

Private Sub Class_Initialize()

   Set app = new_Application_API
   
End Sub



Public Function getPriceLine(priceLineCode As String, salesItemVersionId As Long) As ADODB.Recordset
' Return price line data with finnish price line text.
' Business rule is that we use only finnish price line texts.
' Database enables however any language.

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT p.Id, p.ItemVersion_Id, t.Content, p.BizTypeBasic_Id, p.BizTypeOverride_Id, p.SAPCode " & _
      "FROM " & PRICE_LINE & " AS p " & _
      "INNER JOIN " & PRICE_LINE_TEXT & " AS t ON p.Id = t.PriceLine_Id " & _
      "WHERE p.Id = '" & priceLineCode & "' AND p.ItemVersion_Id = " & salesItemVersionId & " AND t.Lang = 'fin'"
    
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   Set getPriceLine = rst
   Set rst = Nothing

End Function


Public Function getPriceListEntry(priceEntryID As Long, priceLineID As String) As ADODB.Recordset
' Return the requested price entry data.
' Ensure that it belongs to a right price line.

    Dim sql As String
    Dim rst As New ADODB.Recordset

    sql = "SELECT * FROM " & PRICE_LIST_ENTRY & _
          " WHERE Id = " & priceEntryID & " AND PriceLine_Id = '" & priceLineID & "'"
    
    rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
    
    Set getPriceListEntry = rst
    
    Set rst = Nothing

End Function

Public Function getCurrentPriceListEntry(priceLineCode As String) As ADODB.Recordset
'   Return currently valid price list entry; cases:
'   - no records if there is no priceListEntries
'   - data of the current priceListEntry in STD pricelist

    Dim sql                As String
    Dim sqlCurrentEntry    As String
    Dim rst As New ADODB.Recordset
    
    sqlCurrentEntry = "SELECT PriceLine_Id AS pLineId, Max(ValidFrom) AS StartDate " & _
                      "FROM " & PRICE_LIST_ENTRY & _
                      " WHERE PriceLine_Id = '" & priceLineCode & "' " & _
                      "GROUP BY PriceLine_Id"
                      
    sql = "SELECT u.Id, u.PriceLne_Id, u.PriceList_Id, u.Price, u.PriceUnit, u.PriceType, u.ValidFrom, u.VMA, u.VMA_boost, u.Active, u.Comment " & _
          "FROM " & PRICE_LIST_ENTRY & " AS u " & _
          " INNER JOIN (" & sqlCurrentEntry & ") AS s " & _
          "ON (u.PriceLine_Id = s.pLineId AND u.ValidFrom = s.StartDate)"
              
    rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
    
    Set getCurrentPriceListEntry = rst
    Set rst = Nothing
    

End Function


Public Function getLatestPriceListEntry(priceLineCode As String) As ADODB.Recordset
'   Return the latest price list entry; cases:
'   - no records if there is no priceListEntries yet
'   - data of the current priceListEntry in STD pricelist
'   - data of a future priceListEntry in STD priceList

    Dim sql As String
    Dim rst As New ADODB.Recordset
    
    sql = "SELECT * FROM " & PRICE_LIST_ENTRY & _
          " WHERE PriceLine_Id = '" & priceLineCode & "' " & _
          "ORDER BY Id DESC"

'  The latest will be on the top of the recordset
    rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
    
    Set getLatestPriceListEntry = rst
    Set rst = Nothing

End Function
Public Sub insertPriceLine(newPriceLine As dom_PriceLine, versionID As Long)
    
    Dim rst As New ADODB.Recordset
    
    rst.Open PRICE_LINE, app.getDbConn, adOpenDynamic, adLockPessimistic
    
    With newPriceLine
'     Add new price line
      rst.AddNew
      rst!id = .getCode
      rst!ItemVersion_Id = versionID
      rst!BizTypeBasic_Id = .getBusinessTypeBasic
      rst!BizTypeOverride_Id = .getBusinessTypeOverride
      rst!sapCode = .getSapCode
      rst.Update
      rst.Close
'     Add new price line text
      rst.Open PRICE_LINE_TEXT, app.getDbConn, adOpenDynamic, adLockPessimistic
      rst.AddNew
      rst!PriceLine_Id = .getCode
      rst!lang = "fin"
      rst!Content = .getPriceLineText
      rst.Update
      rst.Close
    End With
    
    Set rst = Nothing
    
End Sub

Public Sub updatePriceLine(newPriceLine As dom_PriceLine)

    Dim rst As New ADODB.Recordset
    Dim sql As String
    
'  Update price line
   sql = "SELECT * FROM " & PRICE_LINE & " WHERE Id = '" & newPriceLine.getCode & "'"
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
    
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
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   rst!Content = newPriceLine.getPriceLineText
   rst.Update
   
   Set rst = Nothing

End Sub


Public Sub insertPriceListEntry(priceEntry As dom_PriceListEntry, priceLineID As String)

   Dim rst  As New ADODB.Recordset
   
   rst.Open PRICE_LIST_ENTRY, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !PriceLine_Id = priceLineID
      !PriceList_Id = priceEntry.getPriceListId
      !price = priceEntry.getPrice
      !priceUnit = priceEntry.getUnit
      !priceType = priceEntry.getType
      !ValidFrom = priceEntry.getValidFrom
      !ValidTo = priceEntry.getValidTo
      !vma = priceEntry.getVMA
      !VMA_boost = priceEntry.getVMABoost
      !active = priceEntry.getActivity
      !Comment = priceEntry.getComment
      .Update
      .Close
   End With
   
   Set rst = Nothing
   
End Sub


Public Sub updatePriceListEntry(priceEntry As dom_PriceListEntry)

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT * FROM " & PRICE_LIST_ENTRY & " WHERE Id = " & priceEntry.getId
   
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      !price = priceEntry.getPrice
      !priceUnit = priceEntry.getUnit
      !priceType = priceEntry.getType
      !ValidFrom = priceEntry.getValidFrom
      !ValidTo = priceEntry.getValidTo
      !vma = priceEntry.getVMA
      !VMA_boost = priceEntry.getVMABoost
      !active = priceEntry.getActivity
      !Comment = priceEntry.getComment
      .Update
      .Close
   End With
   
   Set rst = Nothing
   
End Sub


Private Sub Class_Terminate()

   Set app = Nothing
   
End Sub
