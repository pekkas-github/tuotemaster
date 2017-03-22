VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Application_API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Public Function getDbConn() As ADODB.Connection
   Static dbConn  As ADODB.Connection
   
   If dbConn Is Nothing Then
      Set dbConn = CurrentProject.Connection
   End If
   
   Set getDbConn = dbConn
   
End Function


Public Function getCategories() As repo_Categories
   
   Set getCategories = new_Categories
   
End Function


Public Function getProducts() As repo_Products
   
   Set getProducts = new_Products
   
End Function

Public Function Statuses() As repo_Statuses

   Set Statuses = new_Statuses
   
End Function

Public Function Services() As Services

   Set Services = new_Services
   
End Function
