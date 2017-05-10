VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dba_BillingMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private app    As Application_API

Private Sub Class_Initialize()

   Set app = new_Application_API
   
End Sub



Public Sub setStatus(system As String, code As String, status As Integer)
' Set staus of the legacy item

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & BILLING_STATUS & " WHERE System = '" & system & "' AND Code = '" & code & "'"
   
   With rst
      .Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
      !StatusName_Id = status
      .Update
      .Close
   End With
   
   Set rst = Nothing
   
End Sub


Public Sub deleteMapping(system As String, code As String)
' Delete mapping recordes from database

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT * FROM " & BILLING_MAPPING & " WHERE Source_System_Id = '" & system & "' AND Source_Item_Id = '" & code & "'"
   
   With rst
      .Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
      
      While Not .EOF
         .delete
         .Update
         .MoveNext
      Wend
      .Close
   End With
   
   Set rst = Nothing
   
End Sub


Public Sub insertMapping(salesItem As Entity, system As String, code As String)
' Add a new mapping record in database

   Dim rst     As New ADODB.Recordset
   Dim verId   As Long
   
'  Find version id in database
   verId = DLookup("Id", ITEM_VERSION, "Item_Code = '" & salesItem.getCode & "' AND VersionNumber = '" & salesItem.getVersionNumber & "'")
   
   With rst
      .Open BILLING_MAPPING, app.getDbConn, adOpenDynamic, adLockPessimistic
      .AddNew
      !Source_System_Id = system
      !Source_Item_Id = code
      !ItemVersion_Id = verId
      .Update
      .Close
   End With
   
   Set rst = Nothing
   
End Sub



Private Sub Class_Terminate()
   
   Set app = Nothing
   
End Sub
