VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_Descriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Private app          As Application_API
Private oEntityCode  As String         ' Reference to the parent entity
Private oVersionNro  As String         ' Reference to the version number of the parent entity

Public Sub init(entityCode As String, versionNro As String)

   Set app = new_Application_API
   oEntityCode = entityCode
   oVersionNro = versionNro
   
End Sub



Public Function getDescription(language As String) As dom_Description

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   Dim item As dom_Description
   
   sql = "SELECT d.Description " & _
         "FROM " & VERSION_DESCRIPTION & " AS d INNER JOIN " & ITEM_VERSION & " AS v ON v.Id = d.ItemVersion_Id " & _
         "WHERE v.Item_Code = '" & oEntityCode & "' AND v.VersionNumber = '" & oVersionNro & "' AND Lang_Code = '" & language & "'"
         
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   If rst.EOF Then
      Set getDescription = Nothing
   Else
      Set item = new_Description(language)
      item.setText (Nz(rst!description))
      Set getDescription = item
   End If
   
   Set rst = Nothing
   
End Function


Public Sub saveDescription(newDescription As dom_Description)
' Create a new description record or update an existing one

   Dim currentDescription  As dom_Description
   Dim versionId           As Long        ' this is utilized in insert and update subroutines
   
   With newDescription
      Set currentDescription = getDescription(.getLanguage)
      versionId = app.Services.getVersionId(oEntityCode, oVersionNro)
   End With
   
   If currentDescription Is Nothing Then
      Call insertDescription(versionId, newDescription)
   Else
      currentDescription.setText (newDescription.getText)
      Call updateDescription(versionId, currentDescription)
   End If
   
   Set currentDescription = Nothing

End Sub


Private Sub insertDescription(versionId As Long, newDescription As dom_Description)

   Dim rst        As New ADODB.Recordset
   
   rst.Open VERSION_DESCRIPTION, app.getDbConn, adOpenDynamic, adLockPessimistic

   rst.AddNew
   rst!description = newDescription.getText
   rst!Lang_Code = newDescription.getLanguage
   rst!ItemVersion_Id = versionId
   rst.Update
   
   Set rst = Nothing
   
End Sub


Private Sub updateDescription(versionId As Long, currentDescription As dom_Description)

   Dim sql        As String
   Dim rst        As New ADODB.Recordset
      
   sql = "SELECT * FROM " & VERSION_DESCRIPTION & _
         " WHERE ItemVersion_Id = " & versionId & _
         " AND Lang_Code = '" & currentDescription.getLanguage & "'"
         
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic

   rst!description = currentDescription.getText
   rst.Update
   
   Set rst = Nothing
   
End Sub


Public Function rsString() As String
' Return rowsource statement to the user interface

   Dim sql  As String
   
   sql = "SELECT d.Description, d.Lang_Code " & _
         "FROM " & VERSION_DESCRIPTION & " AS d INNER JOIN " & ITEM_VERSION & " AS v ON v.Id = d.ItemVersion_Id " & _
         "WHERE v.Item_Code = '" & oEntityCode & "' AND v.VersionNumber = '" & oVersionNro & "'"
   
   rsString = sql
     
End Function


Private Sub Class_Terminate()
   
   Set app = Nothing
   
End Sub
