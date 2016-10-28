VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Names_Abstract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' This is a super class for all sub-classes that implement an Names_IF interface.
' Methods that are common for all sub-classes are delegated to this class.

Private dbconn          As ADODB.Connection 'TODO: Something more elegant is needed here
Private db        As Names_db

Public Sub init()
' constructor

   Set dbconn = CurrentProject.Connection
   Set db = New Names_db
   
End Sub


Public Function getName(entityCode As String, language As String) As EntityName
' Return the name of a specific entity with requested language as a name object.
' If no name then return dummy name

   Dim rst        As ADODB.Recordset
   
   Set rst = db.getName(entityCode, language)
   
   If Not rst.EOF Then
      Set getName = new_Name(rst!text, language)
   Else
      Set getName = new_Name("-", language)
   End If
   
   Set rst = Nothing
   
End Function


Public Sub saveName(entityCode As String, newName As EntityName)
' Create a new name record or update an existing one.
' Name validation has been done in concrete class.

   Dim currentName   As EntityName
   
'  Check if there is a current name with this language
   Set currentName = getName(entityCode, newName.lang)
   
   If currentName.getText = "-" Then    ' if no -> insert
      Call db.insertName(entityCode, newName)
   Else                                ' if yes -> update name text
      currentName.setText (newName.getText)
      Call db.updateName(entityCode, currentName)
   End If
   
   Set currentName = Nothing

End Sub


Private Sub Class_Terminate()

   Set dbconn = Nothing

End Sub
   

