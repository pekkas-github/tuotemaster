VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "If_Entity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit


Public Function getCode() As String

End Function

Public Function getCurrentStatus() As dom_Status

End Function

Public Function getDecision(decisionId As Long) As dom_Decision

End Function

Public Function getDescription(language As String) As String

End Function

Public Function getLastDecision() As dom_Decision

End Function

Public Function getName(language As String) As String

End Function

Public Function getNewDecision() As dom_Decision

End Function

Public Function getOwner() As dom_Owner

End Function

Public Function getProperty(propertyType As String) As dom_Property

End Function

Public Function getType() As String

End Function

Public Function getVersionNumber() As String

End Function

Public Sub setDecision(newDecision As dom_Decision)

End Sub

Public Sub setDescription(descriptionText As String, language As String)

End Sub

Public Sub setName(newName As String, language As String)

End Sub

Public Sub setOwner(partyId As Long, startDate As Date)

End Sub

Public Sub setProperty(newProperty As dom_Property)

End Sub

Public Sub setStatus(statusId As Integer, startDate As Date)

End Sub

Public Function getDocuments() As repo_Documents

End Function
