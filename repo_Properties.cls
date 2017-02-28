VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
'CREATED in ftr#003
' This class is an abstraction of all property values that are asigned
' to the parent entity. Each entity type has its own set of property types.
' Allowed values of a property type are specific for each entity type.

Private app          As Application_API
Private db           As Dba_Properties
Private oEntityCode  As String
Private oEntityType  As String

'constructor
Public Sub init(entityCode As String, entityType As String)
   
   Set app = new_Application_API
   Set db = New Dba_Properties
   oEntityCode = entityCode
   oEntityType = entityType
   
End Sub


Public Function getProperty(propertyType As String) As dom_Property
'  Get property record from database. If no records were returned then return a new
'  property object.

   Dim rst  As ADODB.Recordset

   Set rst = db.getProperty(oEntityCode, propertyType)
   
   If rst.EOF Then   ' return a new one
      Set getProperty = new_Property(oEntityType, propertyType, "", True)
   Else              ' return an existing one
      Set getProperty = new_Property(oEntityType, propertyType, rst!value_Id, False)
   End If
      
   Set rst = Nothing
   
End Function


Public Sub saveProperty(prop As dom_Property)
'  Insert or update the property (value) into database.

   If prop.isNew Then
      Call db.insertProperty(oEntityCode, prop.getType, prop.getValue)
   ElseIf prop.isModified Then
      Call db.updateProperty(oEntityCode, prop.getType, prop.getValue)
   End If
   
End Sub


Public Function rsString(language As String) As String
'  Return rowsource with valid properties and their current values

   Dim sqlDefaultProperties   As String
   Dim sqlItemValues          As String
   Dim sqlRowsource           As String
   
   'default properties for this entity type
   sqlDefaultProperties = "SELECT DISTINCT a.Property_Id, b.LookupText AS Property " & _
                          "FROM ItemTypePropertyValues AS a " & _
                          "INNER JOIN Properties AS b ON a.Property_Id = b.Id " & _
                          "WHERE ItemType = '" & oEntityType & "' AND b.Language = '" & language & "'"
   
   'current property values of this entity
   sqlItemValues = "SELECT a.Property_Id, b.LookupText AS Value " & _
                   "FROM ItemPropertyValues AS a " & _
                   "INNER JOIN PropertyValues AS b ON a.Property_Id = b.Property_Id AND a.Value_Id = b.Value_Id " & _
                   "WHERE Item_Code = '" & oEntityCode & "' AND b.Language = '" & language & "'"
   
   '
   sqlRowsource = "SELECT dp.Property_Id, dp.Property, val.Value " & _
                  "FROM (" & sqlDefaultProperties & ") AS dp " & _
                  "LEFT JOIN (" & sqlItemValues & ") AS val ON dp.Property_Id = val.Property_Id"
                  
   rsString = sqlRowsource
                    
End Function



Public Function rsStringValues(propertyType As String, language As String) As String
'  Return rowsource with valid property values for this entity-property type combination

   Dim sql As String
   
   sql = "SELECT pro.Id, val.Value_Id, pro.LookupText AS Property, val.LookupText AS Value " & _
         "FROM ItemTypePropertyValues AS itpv " & _
         "INNER JOIN Properties AS pro ON itpv.Property_Id = pro.Id " & _
         "INNER JOIN PropertyValues AS val ON itpv.Value_Id = val.Value_Id " & _
         "WHERE itpv.ItemType = '" & oEntityType & "' " & _
         "AND pro.Id = '" & propertyType & "' " & _
         "AND pro.Language = '" & language & "' " & _
         "AND val.Language = '" & language & "' " & _
         "ORDER BY pro.LookupText"
         
   rsStringValues = sql
   
End Function


Private Sub Class_Terminate()
   
   Set app = Nothing
   Set db = Nothing
   
End Sub

