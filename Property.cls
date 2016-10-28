VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Property"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
'CREATED in ftr#003

Private app             As Application_API
Private db              As Properties_Db
Private oEntityType     As String   ' needed for getting valid value list
Private oPropertyType   As String   ' needed for getting valid value list and property name
Private oValue          As String   ' stores the selected value
Private oNew            As Boolean
Private oModified       As Boolean


'constructor
Public Sub init(entityType As String, propertyTypeId As String, valueId As String, isNew As Boolean)
   
   Set app = New Application_API
   Set db = New Properties_Db
   
   oEntityType = entityType
   oPropertyType = propertyTypeId
   oValue = valueId
   oNew = isNew
   oModified = False
   
End Sub


Public Function getName(language As String) As String
'  Return the name of this proerpty type in default langauage

   getName = db.getPropertyTypeName(oPropertyType, language)
   
End Function


Public Function getType() As String
   
   getType = oPropertyType

End Function


Public Function getEntityType() As String
   
   getEntityType = oEntityType
   
End Function
Public Sub setValue(valueId As String)
   
   oValue = valueId
   oModified = True

End Sub


Public Function getValue() As String
   
   getValue = oValue

End Function

Public Property Get isNew() As Boolean
   
   isNew = oNew

End Property


Public Property Get isModified() As Boolean
   
   isModified = oModified

End Property


Private Sub Class_Terminate()
   
   Set app = Nothing
   Set db = Nothing
      
End Sub
