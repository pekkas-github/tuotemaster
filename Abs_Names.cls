VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Abs_Names"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' This is a super class for all sub-classes that implement an If_Names interface.
' Methods that are common for all sub-classes are delegated to this class.

Private db           As Dba_Names
Private parentCode   As String
'

Public Sub init(entityCode As String)
' constructor

   Set db = new_Dba_Names
   parentCode = entityCode
   
End Sub

Private Sub Class_Terminate()

   Set db = Nothing
   
End Sub

'---- METHODS -----------------

Public Function getName(language As String) As dom_Name
' Return the name of a specific entity with requested language as a name object.
' If no name then return dummy name

   Dim rst        As ADODB.Recordset
   
   Set rst = db.getName(parentCode, language)
   
   If Not rst.EOF Then
      Set getName = new_Name(rst!text, language)
   Else
      Set getName = new_Name("-", language)
   End If
   
   Set rst = Nothing
   
End Function


Public Sub saveName(newName As dom_Name)
' Create a new name record or update an existing one.
' Name validation has been done in concrete class.

   Dim currentName   As dom_Name
   
'  Check if there is a current name with this language
   Set currentName = getName(newName.lang)
   
   If currentName.getText = "-" Then    ' if no -> insert
      Call db.insertName(parentCode, newName)
   Else                                ' if yes -> update name text
      currentName.setText (newName.getText)
      Call db.updateName(parentCode, currentName)
   End If
   
   Set currentName = Nothing

End Sub

   

