VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BillingMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'**
'   This class takes care of mapping and unmapping of billing item codes to
'   mapping status groups and sales item version.
'   Business rule is that each legacy code can be mapped only to one item.
'**

Private db  As Dba_BillingMapper


Private Sub Class_Initialize()
   
   Set db = new_Dba_BillingMapper
   
End Sub
Public Sub mapItem(salesItem As Entity, system As String, code As String)
' Create a mapping between billing item and sales item.
' Prevoius mapping is overwritten (deleted)

'  Delete previous mapping of this legacy item if exists
   Call db.deleteMapping(system, code)
   
'  Add new mapping of this legacy item
   Call db.insertMapping(salesItem, system, code)
   
'  Set billing item mapping status to mapped
   Call db.setStatus(system, code, 41)
    
End Sub


Public Sub unMapItem(system As String, code As String)
' Deletes the mapping of this billing item and changes its mapping status to "unmapped"

'  Delete previous mapping of this legacy item
   Call db.deleteMapping(system, code)
   
'  Set legacy item status to unmapped
   Call db.setStatus(system, code, 45)
   
End Sub

Private Sub Class_Terminate()
   
   Set db = Nothing
   
End Sub
