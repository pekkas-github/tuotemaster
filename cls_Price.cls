Option Compare Database
Option Explicit

'Price-luokka tarjoaa tietuejoukkona tuodun hinnan tiedot luettaviksi

private rst	as ADODB.Recordset

'ALUSTUS- JA LOPETUSTOIMET

public sub init(priceData as ADODB.Recordset)
'Talleta hinnan tiedot objektimuuttujaan
	set rst = priceData
end sub


private sub Class_Terminate()
	set rst = nothing
end sub


'METODIT

public function getCode() as String
	getCode = rst!Code	
end function


public function getPrice() as long
	getPrice = rst!Price
end function


public function getUnit() as string
	getPRiceUnit = rst!Unit
end function


public function getType() as string
	getType = rst!PriceType
end function


public function getActive() as boolean
	getActive = rst!Active
end function


public function getVMA() as decimal
	getVMA = rst!VMA
end function


public function getVMA2() as decimal
	getVMA2 = rst!Boost
end function


public function getComment() as string
	getComment = rst!Comment
end function

