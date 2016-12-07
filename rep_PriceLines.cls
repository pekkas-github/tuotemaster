VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rep_PriceLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private oVersionId   As Long
Private db           As Dba_PriceLines


Public Sub init(salesItemCode As String, salesItemVersion As String)

   oVersionId = DLookup("Id", ITEM_VERSION, "Item_Code = '" & salesItemCode & "' AND VersionNumber = '" & salesItemVersion & "'")
   Set db = New Dba_PriceLines
   
End Sub


Public Function createPriceLine() As priceLine

   Set createPriceLine = New priceLine

End Function


Public Function getPriceLine(priceLineCode As String) As priceLine

   Dim rst           As ADODB.Recordset
   Dim objPriceLine  As priceLine
   
   Set rst = db.getPriceLine(priceLineCode)
   
   Set objPriceLine = New priceLine

   With objPriceLine
      .setObjectId (priceLineCode)
      .setPriceLineText (rst!Content)
      .setBusinessTypeBasic (rst!BizTypeBasic_Id)
      .setBusinessTypeOverride (rst!BizTypeOverride_Id)
      .setSapCode (rst!sapCode)
      .isNew = False
   End With
   
   Set getPriceLine = objPriceLine
   Set objPriceLine = Nothing
   
End Function


Public Sub savePriceLine(newPriceLine As priceLine)
   Dim app  As Application_API
   Set app = New Application_API

   If newPriceLine.isNew Then
      newPriceLine.setObjectId (app.Services.getNextPriceLineCode)
      Call db.insertPriceLine(newPriceLine, oVersionId)
   Else
      Call db.updatePriceLine(newPriceLine)
   End If

   Set app = Nothing
   
End Sub

