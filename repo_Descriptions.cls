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
Private Services     As Services
Private dbConn       As ADODB.Connection

Public Sub init()

   Set app = new_Application_API
   Set Services = new_Services
   Set dbConn = CurrentProject.Connection
   
End Sub



Public Function getDescription(entityCode As String, versionNro As String, language As String) As dom_Description

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   Dim item As dom_Description
   
   sql = "SELECT d.Description " & _
         "FROM " & VERSION_DESCRIPTION & " AS d INNER JOIN " & ITEM_VERSION & " AS v ON v.Id = d.ItemVersion_Id " & _
         "WHERE v.Item_Code = '" & entityCode & "' AND v.VersionNumber = '" & versionNro & "' AND Lang_Code = '" & language & "'"
         
   rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic
   
   If rst.EOF Then
      Set getDescription = Nothing
   Else
      Set item = new_Description(entityCode, versionNro, language)
      item.setText (Nz(rst!description))
      Set getDescription = item
   End If
   
   Set rst = Nothing
   
End Function


Public Sub saveDescription(newDescription As dom_Description)
' Create a new description record or update an existing one

   Dim currentDescription  As dom_Description
   Dim versionID           As Long        ' this is utilized in insert and update subroutines
   
   With newDescription
      Set currentDescription = getDescription(.getEntityCode, .getVersionNro, .getLanguage)
      versionID = Services.getVersionId(.getEntityCode, .getVersionNro)
   End With
   
   If currentDescription Is Nothing Then
      Call insertDescription(versionID, newDescription)
   Else
      currentDescription.setText (newDescription.getText)
      Call updateDescription(versionID, currentDescription)
   End If
   
   Set currentDescription = Nothing

End Sub


Private Sub insertDescription(versionID As Long, newDescription As dom_Description)

   Dim rst        As New ADODB.Recordset
   
   rst.Open VERSION_DESCRIPTION, dbConn, adOpenDynamic, adLockPessimistic

   rst.AddNew
   rst!description = newDescription.getText
   rst!Lang_Code = newDescription.getLanguage
   rst!ItemVersion_Id = versionID
   rst.Update
   
   Set rst = Nothing
   
End Sub


Private Sub updateDescription(versionID As Long, currentDescription As dom_Description)

   Dim sql        As String
   Dim rst        As New ADODB.Recordset
      
   sql = "SELECT * FROM " & VERSION_DESCRIPTION & _
         " WHERE ItemVersion_Id = " & versionID & _
         " AND Lang_Code = '" & currentDescription.getLanguage & "'"
         
   rst.Open sql, dbConn, adOpenDynamic, adLockPessimistic

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


Private Sub Class_Terminate()
   
   Set app = Nothing
   Set Services = Nothing
   Set dbConn = Nothing
   
End Sub
