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

Implements If_Names

Private super           As Abs_Names     ' Super class that implements common methods

'constructor
Public Sub init(parentCode As String, parentType As String)

   Set super = New Abs_Names
   Call super.init(parentCode, parentType)
   
End Sub

Private Sub Class_Terminate()
   
   Set super = Nothing
   
End Sub

'---- METHODS -----------------

Private Function getName(language As String) As String
' Return the name with actual name text or "-" name text

   getName = super.getName(language).getText
   
End Function


Private Sub saveName(newName As EntityName)
' Create a new name record or update an existing one.
' This procedure allowes to enter also name duplicates.

   Call super.saveName(newName)
   
End Sub


'---- INFERFACE CALLS ---------------
Private Function If_Names_getName(language As String) As String

   If_Names_getName = getName(language)
   
End Function


Private Sub If_Names_saveName(newName As EntityName)

   Call saveName(newName)
   
End Sub
