VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Application_API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Function Documents() As Documents
   
   Set Documents = New Documents
   Call Documents.init
      
End Function
Public Function Entities() As Entities
   
   Set Entities = New Entities
   Call Entities.init
      
End Function


Public Function Services() As Services

   Set Services = New Services
   
End Function

Public Function Descriptions() As Descriptions

   Set Descriptions = New Descriptions
   Call Descriptions.init
   
End Function

Public Function Statuses() As Statuses

   Set Statuses = New Statuses
   Call Statuses.init
   
End Function


Public Function Decisions() As Decisions

   Set Decisions = New Decisions
   Call Decisions.init
   
End Function
