VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dom_PriceListEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit


Private oId             As Long     ' PriceListEntry ID
Private oPriceListId    As String
Private oPrice          As Single
Private oUnit           As String
Private oType           As String
Private oValidFrom      As Date
Private oValidTo        As Date
Private oVMA            As Single
Private oVMABoost       As Single
Private oActive         As Boolean
Private oComment        As String
Private oIsNew          As Boolean
Private oIsModified     As Boolean


Public Sub init(priceListCode As String)
   
   oPriceListId = priceListCode
   
End Sub


Public Function getId() As Long
    getId = oId
    
End Function


Public Sub setId(value As Long)
    oId = value
    
End Sub


Public Function getPriceListId() As String
   getPriceListId = oPriceListId
    
End Function


Public Function getPrice() As Single
   getPrice = oPrice
    
End Function


Public Sub setPrice(value As Single)
   oPrice = value
   oIsModified = True
    
End Sub


Public Function getUnit() As String
    getUnit = oUnit
    
End Function


Public Sub setUnit(value As String)
    oUnit = value
    oIsModified = True
    
End Sub


Public Function getType() As String
    getType = oType
    
End Function


Public Sub setType(value As String)
    oType = value
    oIsModified = True
    
End Sub


Public Function getValidFrom() As Date
    getValidFrom = oValidFrom
    
End Function

Public Sub setValidFrom(value As Date)
    
   oValidFrom = value
   oIsModified = True
    
End Sub

Public Function getValidTo() As Date
    getValidTo = oValidTo
    
End Function

Public Sub setValidTo(value As Date)
   oValidTo = value
   oIsModified = True
    
End Sub

Public Function getVMA() As Single
    getVMA = oVMA
    
End Function

Public Sub setVMA(value As Single)
    oVMA = value
    oIsModified = True
    
End Sub

Public Function getVMABoost() As Single
    getVMABoost = oVMABoost
    
End Function

Public Sub setVMABoost(value As Single)
    oVMABoost = value
    oIsModified = True
    
End Sub

Public Function getActivity() As Boolean
   getActivity = oActive
   
End Function

Public Sub setActivity(value As Boolean)
   oActive = value
   oIsModified = True
   
End Sub
Public Function getComment() As String
    getComment = oComment
    
End Function

Public Sub setComment(value As String)
    oComment = value
    oIsModified = True
    
End Sub

Public Property Get isNew() As Boolean
    isNew = oIsNew
    
End Property

Public Property Let isNew(value As Boolean)
    oIsNew = value
    
End Property

Public Property Get isModified() As Boolean
    isModified = oIsModified
    
End Property

Public Property Let isModified(value As Boolean)
    oIsModified = value
    
End Property
