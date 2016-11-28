VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriceLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Class: PriceLine
'
Private oId                      As String
Private oEntityVersionId         As Long     'not needed. Stored in repository
Private oPriceLIneText           As String
Private oBusinessTypeBasic       As String
Private oBusinessTypeOverride    As String
Private oSapCode                 As String
Private oStatusNew               As Boolean
Private oStatusModified          As Boolean

Private oPriceLineTexts          As Collection 'not needed. Käytetään vain suomenkielisiä nimiä
Private app                      As Application_API

Private Sub Class_Initialize()

    Set app = New Application_API
    oStatusNew = True
    oStatusModified = False
    oBusinessTypeBasic = "NONE"
    oBusinessTypeOverride = "NONE_O"
    Set oPriceLineTexts = New Collection

End Sub
Public Sub init(versionId As Long)
'   CREATED 2.9.2015 in ver 6.2.0
'   Constructs a new PriceLine object

    On Error GoTo catch
    
    oEntityVersionId = versionId

exitproc:
    Exit Sub
    
catch:
    Call errorMsg(1, "PriceLine.init")
    Resume exitproc
    Resume
End Sub

Public Sub load(priceLineData As ADODB.Recordset, priceLineTexts As ADODB.Recordset)
'   CREATED 2.10.2015 in ver 6.2.0
'   Populates attributes with values retreived from the database

   On Error GoTo catch
   
   Dim tempText   As PriceLineText
      
'  Populate attributes
   With priceLineData
       oId = !Id
       oEntityVersionId = !ItemVersion_Id
       oBusinessTypeBasic = Nz(!BizTypeBasic_Id)
       oBusinessTypeOverride = Nz(!BizTypeOverride_Id)
       oSapCode = Nz(!SapCode)
   End With
   oStatusNew = False
   oStatusModified = False
   
'  Populate price line text attributes
   With priceLineTexts
        While Not .EOF
           Set tempText = New PriceLineText
           Call tempText.init(!Content, !lang, False)
           oPriceLineTexts.Add tempText, !lang
           .MoveNext
       Wend
   End With
       
   
exitproc:
   Set tempText = Nothing
   Exit Sub
   
catch:
   Call errorMsg(1, "PriceLine.load")
   Resume exitproc
   Resume
        
End Sub

Public Sub save()
'   CREATED 3.10.2015 in ver 6.2.0

    On Error GoTo catch
    
    Call app.ObjectRepository.storePriceLine(Me)

exitproc:
    Exit Sub
    
catch:
    Call errorMsg(1, "PriceLine.save")
    Resume exitproc
    Resume
    
End Sub
Private Sub Class_Terminate()

    Set oPriceLineTexts = Nothing
    Set app = Nothing
    
End Sub

Public Function getPriceListEntry(priceListId As String, action As String) As PriceListEntry
'   CREATED 19.9.2015 in ver 6.2.0
'   Returns a priceListEntry object that is relevant for adding or updating.
'   Requires that the priceLine is already persistent.
'   Returned priceListEntry is either
'   - new priceListEntry with default values    (no active prices, action = any)
'   - new priceListEntry with data from the current price   (effective price, action = new)
'   - current priceListEntry itself (effective pricce, action = change)
'   - coming priceListEntry itself  (coming price, action = any)

    On Error GoTo catch
    
    Dim activePriceListEntries  As Collection
    Dim PriceListEntry          As PriceListEntry
    Dim price                   As PriceListEntry
    Dim done                    As Boolean
    
    If oStatusNew Then Err.Raise 1   ' Price line is not persistent

'   Create a new price object by default
    Set PriceListEntry = New PriceListEntry
    Call PriceListEntry.init(oId, priceListId)
    
    done = False
    Set activePriceListEntries = app.ObjectRepository.findActivePriceListEntries(oId)
    
'   Check the prices in this price list. If there is not any -> new priceListEntry object
    For Each price In activePriceListEntries
      If price.getPriceListId = priceListId Then

         If price.getStatus = "coming" Then  'return the coming price object
            Set PriceListEntry = price
            done = True

         ElseIf Not done And action = "change" Then   'return the current price object
            Set PriceListEntry = price

         ElseIf Not done And action = "new" Then   'return a new cloned price object
            Set PriceListEntry = app.ObjectFactory.clonePriceListEntry(price)
            PriceListEntry.setValidFrom (Date)
            PriceListEntry.setStatus ("coming")
            
         End If
      End If
    Next
    Set getPriceListEntry = PriceListEntry
    
exitproc:
    Set activePriceListEntries = Nothing
    Set PriceListEntry = Nothing
    Exit Function
    
catch:
    If Err.number = 1 Then
        Err.Raise vbObjectError + 400, , "Save the price line first."
    Else
        Call errorMsg(1, "PriceLine.getPriceListentry")
        Resume exitproc
        Resume
    End If
    
End Function

Public Function getPriceLineText() As String
'   CREATED 31.8.2015 in ver. 6.2.0
'   Returns price line's text
        
    getPriceLineText = oPriceLIneText
        
End Function

Public Sub setPriceLineText(PriceLineText As String)
'   CREATED 31.8.2015 in ver. 6.2.0
'   Modify or add a price line text
    
    oPriceLIneText = PriceLineText
    oStatusModified = True
        
End Sub

Public Function getAllTexts() As Collection
    Set getAllTexts = oPriceLineTexts
    
End Function

Public Function getObjectId() As String
    getObjectId = oId
    
End Function

Public Sub setObjectId(value As String)
    oId = value
    
End Sub

Public Function getEntityVersionId() As Long
    getEntityVersionId = oEntityVersionId
    
End Function

Public Function getBusinessTypeBasic() As String
    getBusinessTypeBasic = oBusinessTypeBasic
    
End Function

Public Sub setBusinessTypeBasic(value As String)
    oBusinessTypeBasic = Nz(value)
    oStatusModified = True
    
End Sub

Public Function getBusinessTypeOverride() As String
    getBusinessTypeOverride = oBusinessTypeOverride
    
End Function

Public Sub setBusinessTypeOverride(value As String)
    oBusinessTypeOverride = Nz(value)
    oStatusModified = True
    
End Sub

Public Function getSapCode() As String
    getSapCode = oSapCode
    
End Function

Public Sub setSapCode(value As String)
    oSapCode = Nz(value)
    oStatusModified = True
    
End Sub


Public Property Get isNew() As Boolean
    isNew = oStatusNew
    
End Property

Public Property Let isNew(value As Boolean)
    oStatusNew = value

End Property

Public Property Get isModified() As Boolean
    isModified = oStatusModified
    
End Property

Public Property Let isModified(value As Boolean)
    oStatusModified = value
    
End Property
