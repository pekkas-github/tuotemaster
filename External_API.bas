Attribute VB_Name = "External_API"
Option Compare Database
Option Explicit

' API internface for external applications
'-----------------------------------------

Public Function getAPI() As Application_API
' Return API interface for external applications referencing to Tuotemaster

   Static MyAPI   As Application_API
   
   If MyAPI Is Nothing Then
      Set MyAPI = new_Application_API
   End If
   
   Set getAPI = MyAPI
   
End Function
