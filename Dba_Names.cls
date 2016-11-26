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

Private dbconn As ADODB.Connection


Private Sub Class_Initialize()

   Set dbconn = CurrentProject.Connection
   
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
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   Set getName = rst
   Set rst = Nothing

End Function

Public Function isNameUnique(newName As EntityName, entityType As String) As Boolean
' Check uniquenes of the name among entities of this type

   Dim sql  As String
   Dim rst As New ADODB.Recordset

   sql = "SELECT n.Text FROM " & ITEM_NAME & " AS n INNER JOIN " & CORE_ITEM & " AS i ON i.Code = n.Item_Code " & _
         " WHERE i.Type = '" & entityType & "'AND n.Text = '" & newName.getText & "' AND Lang_Code = '" & newName.getLanguage & "'"
         
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   If rst.EOF Then
      isNameUnique = True
   Else
      isNameUnique = False
   End If
   
   Set rst = Nothing
   
End Function


Public Sub insertName(entityCode As String, newName As EntityName)
' Insert a new name record into database

   Dim rst  As New ADODB.Recordset
   
   rst.Open ITEM_NAME, dbconn, adOpenDynamic, adLockPessimistic
   
   rst.AddNew
   rst!text = newName.text
   rst!Lang_Code = newName.lang
   rst!Item_Code = entityCode
   rst.Update
      
   Set rst = Nothing
   
End Sub


Public Sub updateName(entityCode As String, currentName As EntityName)
' Update the name text of an existing name record

   Dim sql  As String
   Dim rst  As New ADODB.Recordset

   sql = "SELECT * FROM " & ITEM_NAME & _
         " WHERE Item_Code = '" & entityCode & "' AND Lang_Code='" & currentName.lang & "'"
         
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   rst!text = currentName.text
   rst.Update
   
   Set rst = Nothing
   
End Sub


Private Sub Class_Terminate()

   Set dbconn = Nothing

End Sub

