VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' This class represents a single document reference

Private oId          As Long
Private oType        As String
Private oName        As String
Private oLanguage    As String
Private oURL         As String
Private oPersistent  As Boolean
Private oModified    As Boolean

Public Function openDocument() As Boolean
' Open the actual document that this object refers to.
' Return false value, if document is not found from URL location.
' Use VBA-class Application
    
On Error GoTo catch
    
   Application.FollowHyperlink oURL
   openDocument = True
   Exit Function

catch:
    ' Error if the document is not found
    openDocument = False

End Function


Public Function getId() As Long
   getId = oId
   
End Function

Public Sub setId(value As Long)
   oId = value
   
End Sub
Public Function getType() As String
   getType = oType
   
End Function

Public Sub setType(value As String)
   oType = value
   oModified = True
   
End Sub

Public Function getName() As String
   getName = oName
   
End Function

Public Sub setName(value As String)
   oName = value
   oModified = True
   
End Sub

Public Function getLanguage() As String
   getLanguage = oLanguage
   
End Function

Public Sub setLanguage(value As String)
   oLanguage = value
   oModified = True
   
End Sub

Public Function getURL() As String
   getURL = oURL
   
End Function

Public Sub setURL(value As String)
   oURL = value
   oModified = True
   
End Sub

Public Property Let isPersistent(value As Boolean)
   oPersistent = value
   
End Property
Public Property Get isPersistent() As Boolean
   isPersistent = oPersistent
   
End Property

Public Property Let isModified(value As Boolean)
   oModified = value
   
End Property

Public Property Get isModified() As Boolean
   isModified = oModified
   
End Property

