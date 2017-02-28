VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_Names_Unique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Implements If_Names

Private super           As Abs_Names     ' Super class that implements common methods
Private util            As Services

'constructor
Public Sub init(parentCode As String)

   Set super = new_Abs_Names(parentCode)
   Set util = new_Services
   
End Sub

Private Sub Class_Terminate()
   
   Set super = Nothing
   Set util = Nothing
   
End Sub

'---- METHODS -----------------------

Private Function getName(language As String) As String
' Return a name object with actual name text or "empty" name text

      getName = super.getName(language).getText

End Function


Private Sub saveName(newName As dom_Name, entityType As String)
' Check that the name is unique among this entity type and
' create a new name record or update an existing one.

   If util.isNameUnique(newName, entityType) Then
      Call super.saveName(newName)
   Else
      Err.Raise e.notUniqueName, , "Name is not unique under this context"
   End If
   
   
End Sub

'---- INTERFACE CALLS -----------------------


Private Sub If_Names_init(entityCode As String)
   
   init entityCode

End Sub

Private Function If_Names_getName(language As String) As String

   If_Names_getName = getName(language)
   
End Function

Private Sub If_Names_saveName(newName As dom_Name, entityType As String)

   Call saveName(newName, entityType)

End Sub
