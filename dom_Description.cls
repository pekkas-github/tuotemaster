VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dom_Description"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' This class describes a description of an entity.

Private oDescriptionText   As String
Private oLanguage          As String


'constructor
Public Sub init(language As String)

   oLanguage = language
   
End Sub

Public Function getText() As String

   getText = oDescriptionText
   
End Function

Public Sub setText(value As String)

   oDescriptionText = value
   
End Sub

Public Function getLanguage() As String

   getLanguage = oLanguage
   
End Function
