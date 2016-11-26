VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dba_Properties"
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

Public Function getProperty(entityCode As String, propertyType As String) As ADODB.Recordset

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT * " & _
         "FROM " & ITEM_PROPERTY_VALUE & _
         " WHERE Item_Code = '" & entityCode & "' " & _
         "AND Property_Id = '" & propertyType & "'"
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   Set getProperty = rst
   Set rst = Nothing
   
End Function


Public Sub insertProperty(entityCode As String, propertyType As String, propertyValue As String)

   Dim rst  As New ADODB.Recordset
   
   rst.Open ITEM_PROPERTY_VALUE, dbconn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !Item_Code = entityCode
      !Property_Id = propertyType
      !value_Id = propertyValue
      .Update
   End With
   
End Sub


Public Sub updateProperty(entityCode As String, propertyType As String, propertyValue As String)

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT * FROM " & ITEM_PROPERTY_VALUE & " WHERE Item_Code = '" & entityCode & "' AND Property_Id = '" & propertyType & "'"
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   rst!value_Id = propertyValue
   rst.Update
   
   Set rst = Nothing
   
End Sub

Public Function getPropertyTypeName(propertyType As String, language As String) As String
   
   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT Id, [Language], LookupText, LineOrder " & _
         "FROM " & PROPERTY_TYPE & _
         " WHERE Id = '" & propertyType & "' AND [Language] = '" & language & "'"
         
   rst.Open sql, dbconn, adOpenForwardOnly, adLockPessimistic
   
   getPropertyTypeName = rst!LookupText
   
   Set rst = Nothing
   
End Function


Public Function getPropertyValueName(propertyValue As String, language As String) As String

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT LookupText FROM " & PROPERTY_VALUE & _
         " WHERE Value_Id = '" & propertyValue & "' AND [Language] = '" & language & "'"
         
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   If rst.EOF Then
      getPropertyValueName = ""
   Else
      getPropertyValueName = rst!LookupText
   End If
   
   Set rst = Nothing
   
End Function

