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

Public Function Entities() As Entities

   Dim newEntities   As Entities
   
   Set newEntities = New Entities
   Call newEntities.init
   
   Set Entities = newEntities
   
End Function


Public Function Services() As Services

   Set Services = New Services
   
End Function

Public Function Descriptions() As Descriptions

   Dim newDescriptions  As Descriptions
   
   Set newDescriptions = New Descriptions
   Call newDescriptions.init
   
   Set Descriptions = newDescriptions
   Set newDescriptions = Nothing
   
End Function

Public Function Statuses() As Statuses

   Set Statuses = New Statuses
   Call Statuses.init
   
End Function


Public Function Decisions() As Decisions

   Dim newDecisions  As Decisions
   Set newDecisions = New Decisions
   Call newDecisions.init
   
   Set Decisions = newDecisions
   Set newDecisions = Nothing
   
End Function
