VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Names_Any"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements Names_IF

Private db              As Names_db
Private super           As Names_Abstract     ' Super class that implements common methods

'constructor
Public Sub init()

   Set db = New Names_db
   Set super = New Names_Abstract
   super.init
   
End Sub


Private Function Names_IF_getName(entityCode As String, language As String) As EntityName
' Return a name object with actual name text or "empty" name text
   
      Set Names_IF_getName = super.getName(entityCode, language)
   
End Function


Private Sub Names_IF_saveName(entityCode As String, newName As EntityName)
' Create a new name record or update an existing one.
' This procedure allowes to enter also name duplicates.

   Call super.saveName(entityCode, newName)
   
End Sub


Private Function Names_IF_rsString(entityCode As String) As String
' Method implementation is delegated to the super class.

   Names_IF_rsString = db.getRowSourceOfEntityNames(entityCode)

End Function


Private Sub Class_Terminate()
   
   Set super = Nothing
   Set db = Nothing
   
End Sub

