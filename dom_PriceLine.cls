VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dom_PriceLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' Class: PriceLine
'
Private oCode                    As String
Private oPriceLineText           As String
Private oBusinessTypeBasic       As String
Private oBusinessTypeOverride    As String
Private oSapCode                 As String
Private oStatusNew               As Boolean
Private oStatusModified          As Boolean


Public Sub init()
    
End Sub


Public Function getPriceListEntries() As repo_PriceListEntries

   Set getPriceListEntries = new_PriceListEntries(oCode)
   
End Function

Public Function getPriceLineText() As String
'   Returns price line's text
        
    getPriceLineText = oPriceLineText
        
End Function


Public Sub setPriceLineText(PriceLineText As String)
'   Modify or add a price line text
    
    oPriceLineText = PriceLineText
    oStatusModified = True
        
End Sub


Public Function getCode() As String
    getCode = oCode
    
End Function


Public Sub setCode(value As String)
    oCode = value
    
End Sub


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
