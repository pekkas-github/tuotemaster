VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Owner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' This class describes an owner(ship) of an entity. Owner details are in a separate person table.
' An entity has ownership history so ownership has validity period

Private oPersonId    As Long
Private oStartDate   As Date
Private oEndDate     As Date
Private dbconn       As ADODB.Connection

'constructor
Public Sub init(personId As Long, startDate As Date)

   oPersonId = personId
   oStartDate = startDate
   oEndDate = "31.12.2050"
   Set dbconn = CurrentProject.Connection
   
End Sub


Public Function getName() As String
' Return the actual name of the owner (lets put sql in a db object)

   Dim sql  As String
   Dim rst  As New ADODB.Recordset
   
   sql = "SELECT PersonName FROM " & PARTY & " WHERE Id = " & oPersonId
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   getName = rst(0)
   
   Set rst = Nothing
   
End Function


Public Sub setPersonId(value As Long)

   oPersonId = value
   
End Sub

Public Function getPersonId() As Long
   
   getPersonId = oPersonId
   
End Function

Public Function getStartDate() As Date

   getStartDate = oStartDate
   
End Function


Public Function getEndDate() As Date

   getEndDate = oEndDate
   
End Function


Public Sub setEndDate(value As Date)

   oEndDate = value
   
End Sub

Private Sub Class_Terminate()

   Set dbconn = Nothing
   
End Sub
