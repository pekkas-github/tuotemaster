VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Names_Unique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements If_Names

Private db              As Dba_Names
Private super           As Abs_Names     ' Super class that implements common methods
Private entityType      As String

'constructor
Public Sub init(parentCode As String, parentType As String)

   Set db = New Dba_Names
   Set super = New Abs_Names
   Call super.init(parentCode, parentType)
   entityType = parentType
   
End Sub

Private Sub Class_Terminate()
   
   Set super = Nothing
   Set db = Nothing
   
End Sub

'---- METHODS -----------------------

Private Function getName(language As String) As String
' Return a name object with actual name text or "empty" name text

      getName = super.getName(language).getText

End Function


Private Sub saveName(newName As EntityName)
' Check that the name is unique among this entity type and
' create a new name record or update an existing one.

   If Not db.isNameUnique(newName, entityType) Then
      Err.Raise e.notUniqueName, , "Name is not unique under this context"
   End If
   
   Call super.saveName(newName)
   
End Sub

'---- INTERFACE CALLS -----------------------

Private Function If_Names_getName(language As String) As String

   If_Names_getName = getName(language)
   
End Function

Private Sub If_Names_saveName(newName As EntityName)

   Call saveName(newName)

End Sub
