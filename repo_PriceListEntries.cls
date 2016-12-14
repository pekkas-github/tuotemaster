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

private dba	as Dba_PriceLines

private sub Class_Initialize()

	set dba = Dba_PriceLines
	
end sub

private sub Class_Terminate()

	set dba=nothing
	
end sub


public Function getPriceListEntry(priceListEntryId as Long) as PriceListEntry
'	Return a new priceListEntry object or an existing object from the database.
	
	dim newPriceListEntry	as PriceListEntry
	dim rst	as ADODB.Recordset
	
	set rst = dba_getPriceListEntry(priceListEntryId)
	
	with newPriceListEntry
		.setId(rst!Id)
		.setPriceLineId(rst!PriceLine_Id)
		'etc.
	end with
	
end function


public function clonePriceListEntry(priceEntry as PriceListEntry) as PriceListEntry
'	Return a clone of given priceListEntry as a new (non-persistent) object

end function


public sub savePriceListEnrty(priceEntry as PriceListEntry)


end sub