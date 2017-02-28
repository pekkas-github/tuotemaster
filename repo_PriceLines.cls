VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_PriceLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private oVersionId   As Long
Private db           As Dba_PriceLines
Private util         As Services


Public Sub init(salesItemCode As String, salesItemVersion As String)

   oVersionId = DLookup("Id", ITEM_VERSION, "Item_Code = '" & salesItemCode & "' AND VersionNumber = '" & salesItemVersion & "'")
   Set db = New Dba_PriceLines
   Set util = new_Services
   
End Sub


Public Function createPriceLine() As dom_PriceLine

   Dim newPriceLine  As dom_PriceLine
   Set newPriceLine = new_PriceLine
   
   With newPriceLine
      .setBusinessTypeBasic ("NONE")
      .setBusinessTypeOverride ("NONE_O")
      .isNew = True
      .isModified = False
   End With
   
   Set createPriceLine = newPriceLine

   Set newPriceLine = Nothing
   
End Function


Public Function getPriceLine(priceLineCode As String) As dom_PriceLine

   Dim rst           As ADODB.Recordset
   Dim objPriceLine  As dom_PriceLine
   
   Set rst = db.getPriceLine(priceLineCode)
   
   Set objPriceLine = new_PriceLine

   With objPriceLine
      .setCode (priceLineCode)
      .setPriceLineText (rst!Content)
      .setBusinessTypeBasic (rst!BizTypeBasic_Id)
      .setBusinessTypeOverride (rst!BizTypeOverride_Id)
      .setSapCode (rst!sapCode)
      .isNew = False
   End With
   
   Set getPriceLine = objPriceLine
   Set objPriceLine = Nothing
   
End Function


Public Sub savePriceLine(newPriceLine As dom_PriceLine)
   Dim app  As Application_API
   Set app = new_Application_API

   If newPriceLine.isNew Then
      newPriceLine.setCode (util.getNextPriceLineCode)
      Call db.insertPriceLine(newPriceLine, oVersionId)
   Else
      Call db.updatePriceLine(newPriceLine)
   End If

   Set app = Nothing
   
End Sub

Private Sub Class_Terminate()

   Set util = Nothing
   Set db = Nothing
   
End Sub
