VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriceListEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'CREATED 10.9.2015 in ver 6.2.0
'MODIFIED 20.3.2016 in ver 6.4.0 (added oActive)

Private oId             As Long
Private oPriceLineId    As String
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
Private Status          As String   ' effective, coming

Private app             As Application_API

Private Sub Class_Initialize()
    Set app = Tuotemaster_Application.getAPI
    
End Sub
Public Sub init(priceLineId As String, priceListId As String, Optional Id As Long = 0, Optional price As Single = 0, Optional priceUnit As String = "€/kk", Optional priceType As String = "RC", Optional startDate As Date = 0, Optional endDate As Date = "31.12.2050", Optional vma As Single = 0, Optional vmaBoost As Single = 0, Optional active As Boolean = True, Optional comment As String = "", Optional isNew As Boolean = True)

'  CREATED 20.10.2015 in ver 6.2.0
'  MODIFIED in ver 6.4.0 (added oActive, oValidFrom not Date+1)

    On Error GoTo Catch
    
   oId = Id
   oPriceLineId = priceLineId
   oPriceListId = priceListId
   oPrice = price
   oUnit = priceUnit
   oType = priceType
   If startDate = 0 Then
      oValidFrom = Date
   Else
      oValidFrom = startDate
   End If
   oValidTo = endDate
   oVMA = vma
   oVMABoost = vmaBoost
   oActive = active
   oComment = comment
   oIsNew = isNew
   oIsModified = False
   If oValidFrom > Date Then
      Status = "coming"
   Else
      Status = "effective"
   End If

exitproc:
    Exit Sub
    
Catch:
    Call errorMsg(1, "PriceListEntry.construct")
    Resume exitproc
    Resume

End Sub

Public Sub populated(rst As ADODB.Recordset)
'   CREATED 5.11.2015 in ver 6.2.0
'  MODIFIED in ver 6.2.1
'  MODIFIED in ver 6.4.0 (added oActive)

    On Error GoTo Catch
    
    oId = rst!Id
    oPriceLineId = rst!PriceLine_Id
    oPriceListId = rst!PriceList_Id
    oPrice = rst!price
    oUnit = Nz(rst!priceUnit)
    oType = Nz(rst!priceType)
    oValidFrom = rst!validFrom
    oValidTo = rst!validTo
    oVMA = rst!vma
    oVMABoost = rst!VMA_boost
    oActive = Nz(rst!active)
    oComment = Nz(rst!comment)
    oIsNew = False
    oIsModified = False
    If oValidFrom > Date Then
        Status = "coming"
    Else
        Status = "effective"
    End If

exitproc:
    Exit Sub
    
Catch:
    Call errorMsg(1, "PriceListEntry.populate")
    Resume exitproc
    Resume
        
End Sub
Public Sub save()
'   CREATED 20.10.2015 in ver 6.2.0

    On Error GoTo Catch
    
    app.ObjectRepository.storePriceListEntry (Me)

exitproc:
    Exit Sub
    
Catch:
    Call errorMsg(1, "PriceListEntry.save")
    Resume exitproc
    Resume

End Sub

Public Function getId() As Long
    getId = oId
    
End Function

Public Sub setId(value As Long)
    oId = value
    
End Sub
Public Function getPriceLineId() As String
    getPriceLineId = oPriceLineId
    
End Function

Public Sub setPriceLineId(value As String)
   oPriceLineId = value
   
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
'  MODIFIED in ver 6.4.0 (start date restriction is removed)
'   If the priceListEntry is effective (start date < today) then the start date cannot be changed
'   If the priceListEntry is new or coming then the start date can be changed
'   If the startDate <= today then status becomes "effective"
    
    If Status = "effective" Then
        'skip
    Else
        If value <= Date Then
         Status = "effective"
      End If
        oValidFrom = value
        oIsModified = True
    End If
    
End Sub

Public Function getValidTo() As Date
    getValidTo = oValidTo
    
End Function

Public Sub setValidTo(value As Date)
' MODIFIED in ver 6.2.2
' MODIFIED in ver 6.4.0 (removed the restriction "date > today"
'   The end date can be changed only if the price is effective.
'   The end date must be greater or equal than start date.

    If Status = "effective" Then
        oValidTo = value
        If oValidTo < oValidFrom Then
            oValidTo = oValidFrom
        End If
        oIsModified = True
    Else
        'skip
    End If
    
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

Public Function getStatus() As String
    getStatus = Status
    
End Function

Public Sub setStatus(value As String)
    Status = value
    
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

Private Sub Class_Terminate()
    Set app = Nothing
    
End Sub
