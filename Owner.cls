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
Private db           As Dba_Owners
'

'constructor
Public Sub init(personId As Long, startDate As Date)

   oPersonId = personId
   oStartDate = startDate
   oEndDate = "31.12.2050"
   Set db = New Dba_Owners
   
End Sub


Public Function getName() As String
' Return the actual name of the owner

   getName = db.getPersonName(oPersonId)
   
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

   Set db = Nothing
   
End Sub
