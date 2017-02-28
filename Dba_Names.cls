VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dba_Names"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private app    As Application_API


Private Sub Class_Initialize()

   Set app = new_Application_API
   
End Sub


Public Function getName(entityCode As String, language As String) As ADODB.Recordset
' Return the name of a specific entity with requested language as a name object.

' Arguments:
'     entityCode  - code of the entity which name is requested from the database
'     language    - language version of the name that is requested
' Return values:
'     name object - name that was found from the database
'     nothing     - emtpy object if no name match was found

   Dim sql        As String
   Dim rst        As New ADODB.Recordset
   
   sql = "SELECT * FROM " & ITEM_NAME & _
         " WHERE Item_Code = '" & entityCode & "' AND Lang_Code = '" & language & "'"
   
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   Set getName = rst
   Set rst = Nothing

End Function



Public Sub insertName(entityCode As String, newName As dom_Name)
' Insert a new name record into database

   Dim rst  As New ADODB.Recordset
   
   rst.Open ITEM_NAME, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   rst.AddNew
   rst!text = newName.text
   rst!Lang_Code = newName.lang
   rst!Item_Code = entityCode
   rst.Update
      
   Set rst = Nothing
   
End Sub


Public Sub updateName(entityCode As String, currentName As dom_Name)
' Update the name text of an existing name record

   Dim sql  As String
   Dim rst  As New ADODB.Recordset

   sql = "SELECT * FROM " & ITEM_NAME & _
         " WHERE Item_Code = '" & entityCode & "' AND Lang_Code='" & currentName.lang & "'"
         
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   rst!text = currentName.text
   rst.Update
   
   Set rst = Nothing
   
End Sub


Private Sub Class_Terminate()

   Set app = Nothing
   
End Sub

