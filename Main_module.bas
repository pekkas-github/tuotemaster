Attribute VB_Name = "Main_module"
Option Compare Database
Option Explicit

Public Function Main() As Boolean
'  Application starts from here.
'  LIVE version opens the login form automatically
'  Develop version keeps the system in design mode

   On Error GoTo Catch
   
   Globals.lang = "fin"            ' Default data representation language
   
   If Globals.IS_LIVE Then
      StartApplication
   End If

exitproc:
   Exit Function

Catch:
   MsgBox Err.description, , "Error in start-up"
   Resume exitproc
   Resume
   
End Function


Private Sub StartApplication()
'   Start the application by opening Login form.

   On Error GoTo Catch
   
   DoCmd.OpenForm "Login"
   
exitproc:
   Exit Sub
        
Catch:
   MsgBox Err.description, , "Opening the login form failed"
   Resume exitproc
   Resume
   
End Sub
