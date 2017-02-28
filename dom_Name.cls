VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dom_Name"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' This class describes an entity name on a specific language

Private oNameText    As String     ' Name text
Private oLanguage    As String     ' Language of the name


'constructor
Public Sub init(nameText As String, language As String)

   oNameText = nameText
   oLanguage = language
   
End Sub


Public Function getText() As String

   getText = oNameText
   
End Function


Public Sub setText(value As String)

   oNameText = value
   
End Sub


Public Function getLanguage() As String

   getLanguage = oLanguage
   
End Function


Public Sub setLanguage(value As String)

   oLanguage = value
   
End Sub


'These properties are replaced by getters/setters above and will be
'deleted later

Public Property Get text() As String
    text = oNameText
    
End Property
Public Property Let text(value As String)
   oNameText = value

End Property
Public Property Get lang() As String
    lang = oLanguage
    
End Property
Public Property Let lang(value As String)
    oLanguage = value
    
End Property


