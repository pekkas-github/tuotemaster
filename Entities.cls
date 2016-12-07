VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Entities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'  This class is the abstraction of all entities of a certain type

Private app          As Application_API
Private dbconn       As ADODB.Connection


Public Sub init()

   Set app = New Application_API
   Set dbconn = CurrentProject.Connection
   
End Sub

Private Sub Class_Terminate()

   Set app = Nothing
   Set dbconn = Nothing

End Sub

'---- METHODS ----------------------

Public Function createEntity(entityType As String, Optional versionNumber As String = "1.0") As Entity
' Create and save a new entity and a version associated with it. Default version number is 1.0.
' Saving creates the code for the entity and its repositories can be created

   Dim newEntity        As Entity
   Dim entityCode       As String
   Dim appServices      As Services
   
   Set appServices = app.Services

'  Find next available code for this entity type
   entityCode = appServices.getNextEntityCode(entityType)

'  Create and save a new entity + version
   Set newEntity = new_Entity(entityCode, entityType, versionNumber)
   newEntity.isNew = True
   Call Me.saveEntity(newEntity)
   Call Me.createVersion(entityCode, versionNumber)
      
   Set createEntity = newEntity
   
   Set newEntity = Nothing
   Set appServices = Nothing
   
End Function


Public Sub createVersion(entityCode As String, versionNro As String)
' Saves a new version record linked to entityCode

   Dim rst        As New ADODB.Recordset
   
   rst.Open ITEM_VERSION, dbconn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !versionNumber = versionNro
      !Item_Code = entityCode
      .Update
      .Close
   End With

End Sub


Public Function getEntity(entityCode As String, Optional versionNro As String = "last") As Entity
' Return a specific entity. If the version number is not specified, return the last version.

   Dim sql              As String
   Dim rst              As New ADODB.Recordset
   Dim newEntity        As Entity
   
'  Define the version number and get version object
   If versionNro = "last" Then
      versionNro = app.Services.getLastVersionNumber(entityCode)
   End If
   
'  Get entity data form database and init the entity object
   sql = "SELECT * FROM " & CORE_ITEM & " WHERE Code = '" & entityCode & "'"
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   Set newEntity = new_Entity(entityCode, rst!Type, versionNro)
   newEntity.isNew = False
   
   Set getEntity = newEntity
   
   Set newEntity = Nothing
   Set rst = Nothing
   
End Function


Public Sub saveEntity(newEntity As Entity)
' Save entity and version in database.
' Basically saving takes place only when a new entity is created. There might be
' need for updating (type property). However type property will propably be moved
' into Property table (makes isNew attribute uneccessary).

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   With rst
      If newEntity.isNew Then
         .Open CORE_ITEM, dbconn, adOpenDynamic, adLockPessimistic
         .AddNew
      Else
         sql = "SELECT * FROM " & CORE_ITEM & " WHERE Code = '" & newEntity & "'"
         .Open sql, dbconn, adOpenDynamic, adLockPessimistic
      End If
      
      !code = newEntity.getCode
      !Type = newEntity.getType
      .Update
      .Close
      newEntity.isNew = False
   End With
   
End Sub



   
