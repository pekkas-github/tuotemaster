VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dba_SalesItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private app    As Application_API

Private Sub Class_Initilize()

   Set app = new_Application_API
   
End Sub

Public Sub saveSalesItem(newSalesItem As dom_SalesItem)

   Dim rst     As New ADODB.Recordset
      
   With rst
      .Open CORE_ITEM, app.getDbConn, adOpenDynamic, adLockPessimistic
      .AddNew
      !code = newSalesItem.getCode
      !Type = newSalesItem.getType
      .Update
      .Close
      newSalesItem.isNew = False
   End With
   
   Set rst = Nothing
   
End Sub


Private Sub Class_Terminate()

   Set app = Nothing
   
End Sub

