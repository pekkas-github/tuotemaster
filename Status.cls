VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
'CREATED in HOT#005

Private oType        As Integer
Private oStartDate   As Date
Private oEndDate     As Date

'constructor
Public Sub init(statusType As Integer, startDate As Date)

   oType = statusType
   oEndDate = "31.12.2050"
   oStartDate = startDate

End Sub

Public Function getName(language As String) As String
' Get name string of the status in passed language

   getName = DLookup("Text", STATUS_NAME, "Lang_Code = '" & language & "' AND Id = " & oType)
   
End Function

Public Sub setEndDate(value As Date)
'  End date is changed when a new status relpaces the previous one

   oEndDate = value
End Sub


Public Function getEndDate() As Date
   getEndDate = oEndDate
End Function


Public Function getType() As Integer
   getType = oType
End Function


Public Sub setType(value As Integer)
   oType = value
End Sub


Public Function getStartDate() As Date
   getStartDate = oStartDate
End Function
