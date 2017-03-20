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

' PriceListEntry repository is creating, getting and saving PriceListEntries.
' The only price list used currently is STD.

Private dba             As Dba_PriceLines    ' database operations
Private oPriceLineID  As String            ' PriceLine that initiated this repository
'


Public Sub init(priceLineID As String)

   oPriceLineID = priceLineID
   Set dba = new_Dba_PriceLines
   
End Sub


Private Sub Class_Terminate()

   Set dba = Nothing
        
End Sub

'---- METHODS ---------------------


Public Function getPriceListEntry(priceEntryID As Long)
' Return a selected priceListEntry object from database.
' Return empty rst if priceEntry is not associated with this priceLine
   
   Dim rst  As ADODB.Recordset
   Set rst = dba.getPriceListEntry(priceEntryID, oPriceLineID)
   
   If rst.EOF Then
      Err.Raise vbObjectError + 600, , "Price is not associated with this price line."
   End If
   
   Set getPriceListEntry = returnPriceListEntry(rst)
   Set rst = Nothing

End Function


Public Function createPriceListEntry() As dom_PriceListEntry
' Rerurn
'  - a new PriceListEntry object if there is no or no future price list entry
'  - the future PriceListEntry object for updates if there is one (you cannot
'    create a new price object if there is already a coming price object )

   Dim rst  As ADODB.Recordset
   Set rst = dba.getLatestPriceListEntry(oPriceLineID)
   
'  The first one in the recordset is the latest one (if not empty)
   If rst.EOF Then  ' no prices yet (create a new)
      Set createPriceListEntry = createNewPriceListEntry()
   ElseIf rst!ValidFrom < Date Then  ' current price, but no coming price (create a new, copy current values)
      Set createPriceListEntry = clonePriceListEntry(rst)
   Else  ' coming price (open coming price, not new)
      Set createPriceListEntry = returnPriceListEntry(rst)
   End If
   
   Set rst = Nothing
   
End Function



Private Function returnPriceListEntry(rst As ADODB.Recordset) As dom_PriceListEntry
' Construct an existing Pricelist Entry object from recordset

   Dim newPriceListEntry   As dom_PriceListEntry
   Set newPriceListEntry = new_PriceListEntry("STD")
   
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


Private Function createNewPriceListEntry() As dom_PriceListEntry
' Build a new PriceListEntry object

   Dim newPriceListEntry   As dom_PriceListEntry
   Set newPriceListEntry = new_PriceListEntry("STD")
   
   With newPriceListEntry
      .setId (0)
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
' Return a clone of the given priceListEntry (rst) as a new (non-persistent) object
   
   Dim newPriceListEntry   As dom_PriceListEntry
   Set newPriceListEntry = new_PriceListEntry("STD")
   
   With newPriceListEntry
      .setId (0)
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
' - new price - no previous prices or future price
' - edit current, future or old price
   
   If priceEntry.isNew Then
      Call dba.insertPriceListEntry(priceEntry, oPriceLineID)
   Else
      Call dba.updatePriceListEntry(priceEntry)
   End If
   
End Sub


Private Function getCurrentPriceListEntry() As dom_PriceListEntry
' Get price list entry data from database
' - if there is no prices yet for this price line -> return empty rst

   Dim rst  As ADODB.Recordset
   
   Set rst = dba.getCurrentPriceListEntry(oPriceLineID)
   
   If rst.EOF Then
      Set getCurrentPriceListEntry = Nothing
   Else
      Set getCurrentPriceListEntry = returnPriceListEntry(rst)
   End If
   
   Set rst = Nothing

End Function
