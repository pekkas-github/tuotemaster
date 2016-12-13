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


public Function getPriceListEntry(priceLineId as string, action as string) as PriceListEntry
'	Return a new priceListEntry object or an existing object from the database.

end function


public function clonePriceListEntry(priceEntry as PriceListEntry) as PriceListEntry
'	Return a clone of given priceListEntry as a new (non-persistent) object

end function


public sub savePriceListEnrty(priceEntry as PriceListEntry)


end sub