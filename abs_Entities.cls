VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "abs_Entities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'  This class is the abstraction of all entities of a certain type

Private app          As Application_API
Private util         As Services


Private Sub Class_Initialize()

   Set app = new_Application_API
   Set util = new_Services
   
End Sub

Private Sub Class_Terminate()

   Set app = Nothing
   Set util = Nothing
   
End Sub

'---- METHODS ----------------------

Public Sub createVersion(entityCode As String, versionNro As String)
' Saves a new version record linked to entityCode

   Dim rst        As New ADODB.Recordset
   
   rst.Open ITEM_VERSION, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !versionNumber = versionNro
      !Item_Code = entityCode
      .Update
      .Close
   End With

End Sub




   
