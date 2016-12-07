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
Private oPriceLineText           As String
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
'   Constructs a new PriceLine object
    
    oEntityVersionId = versionId

End Sub

Private Sub Class_Terminate()

    Set oPriceLineTexts = Nothing
    Set app = Nothing
    
End Sub

Public Function getPriceListEntry(priceListId As String, action As String) As priceListEntry
'   Returns a priceListEntry object that is relevant for adding or updating.
'   Requires that the priceLine is already persistent.
'   Returned priceListEntry is either
'   - new priceListEntry with default values    (no active prices, action = any)
'   - new priceListEntry with data from the current price   (effective price, action = new)
'   - current priceListEntry itself (effective pricce, action = change)
'   - coming priceListEntry itself  (coming price, action = any)

    On Error GoTo Catch
    
    Dim activePriceListEntries  As Collection
    Dim priceListEntry          As priceListEntry
    Dim price                   As priceListEntry
    Dim done                    As Boolean
    
    If oStatusNew Then Err.Raise 1   ' Price line is not persistent

'   Create a new price object by default
    Set priceListEntry = New priceListEntry
    Call priceListEntry.init(oId, priceListId)
    
    done = False
    Set activePriceListEntries = app.ObjectRepository.findActivePriceListEntries(oId)
    
'   Check the prices in this price list. If there is not any -> new priceListEntry object
    For Each price In activePriceListEntries
      If price.getPriceListId = priceListId Then

         If price.getStatus = "coming" Then  'return the coming price object
            Set priceListEntry = price
            done = True

         ElseIf Not done And action = "change" Then   'return the current price object
            Set priceListEntry = price

         ElseIf Not done And action = "new" Then   'return a new cloned price object
            Set priceListEntry = app.ObjectFactory.clonePriceListEntry(price)
            priceListEntry.setValidFrom (Date)
            priceListEntry.setStatus ("coming")
            
         End If
      End If
    Next
    Set getPriceListEntry = priceListEntry
    
exitproc:
    Set activePriceListEntries = Nothing
    Set priceListEntry = Nothing
    Exit Function
    
Catch:
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
        
    getPriceLineText = oPriceLineText
        
End Function

Public Sub setPriceLineText(PriceLineText As String)
'   CREATED 31.8.2015 in ver. 6.2.0
'   Modify or add a price line text
    
    oPriceLineText = PriceLineText
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
