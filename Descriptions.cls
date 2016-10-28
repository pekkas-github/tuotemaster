VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Descriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private app          As Application_API
Private dbconn       As ADODB.Connection

Public Sub init()

   Set app = New Application_API
   Set dbconn = CurrentProject.Connection
   
End Sub



Public Function getDescription(entityCode As String, versionNro As String, language As String) As description

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   Dim Item As description
   
   sql = "SELECT d.Description " & _
         "FROM " & VERSION_DESCRIPTION & " AS d INNER JOIN " & ITEM_VERSION & " AS v ON v.Id = d.ItemVersion_Id " & _
         "WHERE v.Item_Code = '" & entityCode & "' AND v.VersionNumber = '" & versionNro & "' AND Lang_Code = '" & language & "'"
         
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   If rst.EOF Then
      Set getDescription = Nothing
   Else
      Set Item = new_Description(entityCode, versionNro, language)
      Item.setText (Nz(rst!description))
      Set getDescription = Item
   End If
   
   Set rst = Nothing
   
End Function


Public Sub saveDescription(newDescription As description)
' Create a new description record or update an existing one

   Dim currentDescription  As description
   Dim versionId           As Long        ' this is utilized in insert and update subroutines
   
   With newDescription
      Set currentDescription = getDescription(.getEntityCode, .getVersionNro, .getLanguage)
      versionId = DLookup("Id", ITEM_VERSION, "Item_Code = '" & .getEntityCode & "' AND VersionNumber = '" & .getVersionNro & "'")
   End With
   
   If currentDescription Is Nothing Then
      Call insertDescription(versionId, newDescription)
   Else
      currentDescription.setText (newDescription.getText)
      Call updateDescription(versionId, currentDescription)
   End If
   
   Set currentDescription = Nothing

End Sub


Private Sub insertDescription(versionId As Long, newDescription As description)

   Dim rst        As New ADODB.Recordset
   
   rst.Open VERSION_DESCRIPTION, dbconn, adOpenDynamic, adLockPessimistic

   rst.AddNew
   rst!description = newDescription.getText
   rst!Lang_Code = newDescription.getLanguage
   rst!ItemVersion_Id = versionId
   rst.Update
   
   Set rst = Nothing
   
End Sub


Private Sub updateDescription(versionId As Long, currentDescription As description)

   Dim sql        As String
   Dim rst        As New ADODB.Recordset
      
   sql = "SELECT * FROM " & VERSION_DESCRIPTION & _
         " WHERE ItemVersion_Id = " & versionId & _
         " AND Lang_Code = '" & currentDescription.getLanguage & "'"
         
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic

   rst!description = currentDescription.getText
   rst.Update
   
   Set rst = Nothing
   
End Sub


Public Function rsString(entityCode As String, versionNro As String) As String
' Return rowsource statement to the user interface

   Dim sql  As String
   
   sql = "SELECT d.Description, d.Lang_Code " & _
         "FROM " & VERSION_DESCRIPTION & " AS d INNER JOIN " & ITEM_VERSION & " AS v ON v.Id = d.ItemVersion_Id " & _
         "WHERE v.Item_Code = '" & entityCode & "' AND v.VersionNumber = '" & versionNro & "'"
   
   rsString = sql
     
End Function


Public Sub Class_Terminate()
   
   Set app = Nothing
   Set dbconn = Nothing
   
End Sub
