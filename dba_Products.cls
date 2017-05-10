VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dba_Products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private app    As Application_API

Private Sub Class_Initilize()

   Set app = new_Application_API
   
End Sub

Public Sub saveProduct(newProduct As dom_Product)

   Dim rst     As New ADODB.Recordset
      
   With rst
      .Open CORE_ITEM, app.getDbConn, adOpenDynamic, adLockPessimistic
      .AddNew
      !code = newProduct.getCode
      !Type = newProduct.getType
      .Update
      .Close
      newProduct.isNew = False
   End With
   
   Set rst = Nothing
   
End Sub


Private Sub Class_Terminate()

   Set app = Nothing
   
End Sub
