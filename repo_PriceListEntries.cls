VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_PriceListEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private dba             As Dba_PriceLines
Private oPriceLineCode  As String
'


Public Sub init(priceLineCode As String)
   oPriceLineCode = priceLineCode
   Set dba = new_Dba_PriceLines
   
End Sub


Private Sub Class_Terminate()

   Set dba = Nothing
        
End Sub

'---- METHODS ---------------------


Public Function getPriceListEntry(priceListEntryId As Long) As dom_PriceListEntry
' Return a selected priceListEntry object from database.
        
   Dim rst  As ADODB.Recordset
   Set rst = dba.getPriceListEntry(priceListEntryId)
      
   Set getPriceListEntry = returnPriceListEntry(rst)
   Set rst = Nothing
   
End Function



Public Function createPriceListEntry(priceLineId As String) As dom_PriceListEntry
' Rerurn
'  - a new PriceListEntry object if there is no or no future price list entry
'  - the future PriceListEntry object for updates if there is one (you cannot
'    create a new price object if there is already a coming price object )

   Dim rst  As ADODB.Recordset
   Set rst = dba.getLatestPriceListEntry(priceLineId)
   
'  The first one in the recordset is the latest one (if not empty)
   If rst.EOF Then  ' no prices yet
      Set createPriceListEntry = createNewPriceListEntry(priceLineId)
   ElseIf rst!ValidFrom < Date Then  ' current price, but no coming price
      Set createPriceListEntry = clonePriceListEntry(rst)
   Else  ' coming price
      Set createPriceListEntry = returnPriceListEntry(rst)
   End If
   
   Set rst = Nothing
   
End Function


Private Function returnPriceListEntry(rst As ADODB.Recordset) As dom_PriceListEntry
' Construct a Pricelist Entry object from recordset

   Dim newPriceListEntry   As dom_PriceListEntry
   Set newPriceListEntry = new_PriceListEntry(rst!PriceLine_Id, "STD")
      
   With newPriceListEntry
      .setId (rst!id)
      .setPrice (rst!price)
      .setUnit (rst!priceUnit)
      .setType (rst!priceType)
      .setValidFrom (rst!ValidFrom)
      .setValidTo (rst!ValidTo)
      .setVMA (rst!vma)
      .setVMABoost (rst!VMA_boost)
      .setActivity (rst!active)
      .setComment Nz((rst!Comment))
      .isNew = False
      .isModified = False
   End With
   
   Set returnPriceListEntry = newPriceListEntry
   Set newPriceListEntry = Nothing

End Function


Private Function createNewPriceListEntry(priceLineId As String) As dom_PriceListEntry
' Build a new PriceListEntry object

   Dim newPriceListEntry   As dom_PriceListEntry
   Set newPriceListEntry = new_PriceListEntry(priceLineId, "STD")
   
   With newPriceListEntry
      .setId (0)
      .setPriceLineId (priceLineId)
      .setPrice (0)
      .setUnit ("€/kk")
      .setType ("RC")
      .setValidFrom (Date)
      .setValidTo ("31.12.2050")
      .setVMA (0)
      .setVMABoost (0)
      .setActivity (True)
      .setComment ("")
      .isNew = True
      .isModified = False
   End With
   
   Set createNewPriceListEntry = newPriceListEntry
   Set newPriceListEntry = Nothing
   
End Function


Private Function clonePriceListEntry(rst As ADODB.Recordset) As dom_PriceListEntry
' Return a clone of given priceListEntry as a new (non-persistent) object
   
   Dim newPriceListEntry   As dom_PriceListEntry
   Set newPriceListEntry = new_PriceListEntry(rst!PriceLine_Id, "STD")

   With newPriceListEntry
'      .setId (rst!id); pitäisikö tämän olla .setId (0)
      .setPrice (rst!price)
      .setUnit (rst!priceUnit)
      .setType (rst!priceType)
      .setValidFrom (Date)
      .setValidTo (rst!ValidTo)
      .setVMA (rst!vma)
      .setVMABoost (rst!VMA_boost)
      .setActivity (rst!active)
      .setComment ("")
      .isNew = True
      .isModified = False
   End With
   
   Set clonePriceListEntry = newPriceListEntry
   Set newPriceListEntry = Nothing

End Function


Public Sub savePriceListEnrty(priceEntry As dom_PriceListEntry)
' Save price list entry data in database with use cases
' - new price - no previous prices (insert new)
' - new price - is current price (update current end data, insert new)
' - edit future price (update current end date, update future price)
' - edit current price (update (comment))
' - edit old price (update (comment))

   Dim currentPriceListEntry  As dom_PriceListEntry
   
   If priceEntry.isNew Then
      Set currentPriceListEntry = getCurrentPriceListEntry
      If Not currentPriceListEntry Is Nothing Then   ' change end date of the current price
         currentPriceListEntry.setValidTo (priceEntry.getValidFrom - 1)
         Call dba.updatePriceListEntry(currentPriceListEntry)
      End If
      Call dba.insertPriceListEntry(priceEntry)
   Else
      If priceEntry.getValidFrom >= Date Then
         Set currentPriceListEntry = getCurrentPriceListEntry
         If Not currentPriceListEntry Is Nothing Then
            currentPriceListEntry.setValidTo (priceEntry.getValidFrom - 1)
            Call dba.updatePriceListEntry(currentPriceListEntry)
         End If
      End If
      Call dba.updatePriceListEntry(priceEntry)
   End If
   
End Sub


Private Function getCurrentPriceListEntry() As dom_PriceListEntry

   Dim rst  As ADODB.Recordset
   
   Set rst = dba.getCurrentPriceListEntry(oPriceLineCode)
   
   If rst.EOF Then
      Set getCurrentPriceListEntry = Nothing
   Else
      Set getCurrentPriceListEntry = returnPriceListEntry(rst)
   End If
   
   Set rst = Nothing

End Function
