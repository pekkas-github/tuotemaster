VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "dom_Decision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'  This class represents single decision object.

Private oId          As Long
Private oStartDate   As Date
Private oEndDate     As Date
Private oText        As String
Private oAuthority   As String
Private oNew         As Boolean

Public Sub init(isNew As Boolean)
'  Constructor

   oNew = isNew
   
'  Defaulf values are set if the object is new
   If isNew Then
      oId = 0
      oStartDate = Date
      oEndDate = "31.12.2050"
      oText = "Decision text..."
      oAuthority = "Admin"
   End If
         
End Sub

Public Sub setStartDate(value As Date)
   oStartDate = value
End Sub

Public Function getDecisionText() As String
   getDecisionText = oText
End Function

Public Sub setDecisionText(value As String)
   oText = value
End Sub

Public Function getAuthority() As String
   getAuthority = oAuthority
End Function

Public Sub setAuthority(value As String)
   oAuthority = value
End Sub

Public Function getStartDate() As Date
   getStartDate = oStartDate
End Function

Public Function getEndDate() As Date
   getEndDate = oEndDate
End Function

Public Sub setEndDate(value As Date)
   oEndDate = value
End Sub

Public Function getId() As Long
   getId = oId
End Function

Public Sub setId(value As Long)
   oId = value
End Sub

Public Property Get isNew() As Boolean
   isNew = oNew
End Property

Public Property Let isNew(value As Boolean)
   oNew = value
End Property



