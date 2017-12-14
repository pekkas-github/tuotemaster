Option Compare Database
Option Explicit

'Prices-luokka tarjoaa rajapinnan kaikkiin hintoihin tietokannassa


'ALUSTUS- JA LOPETUSTOIMET

'METODIT

public function getList(salesItemCode as string) as string

	dim sql	as string

	sql = "select ,,,FROM " & PRICE & " WHERE PriceLine_Code = '" & salesItemCode & "' ORDER BY StartDate DESC"
	
	getList = sql
	
end function


public function getItem(priceCode as string) as cls_Price

	dim sql 	as string
	dim rst 	as new ADODB.Recordset
	dim price	as cls_Price
	
	sql = "SELECT ,,,FORM " & PRICE & " WHERE Code = '" & priceCode & "'"
	rst.open sql,CurrentProject.Connection
	
	set price = new cls_Price
	call price.init rst
	
	set getItem = price
	set price = nothing
	
end function