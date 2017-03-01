Attribute VB_Name = "Main_module"
Option Compare Database
Option Explicit

Public Function Main(defaultLanguage As String) As Boolean
'  Application starts from here.
'  LIVE version opens the login form automatically
'  Develop version keeps the system in design mode

   On Error GoTo catch
   
   Globals.lang = defaultLanguage   ' Default data representation language
   FormUtility.initFormUtility      ' Initialize Form Utility service
   
   If Globals.IS_LIVE Then
      StartApplication
   End If

exitproc:
   Exit Function

catch:
   MsgBox Err.description, , "Error in start-up"
   Resume exitproc
   Resume
   
End Function


Private Sub StartApplication()
'   Start the application by opening Login form.

   On Error GoTo catch
   
   DoCmd.openForm "Login"
   
exitproc:
   Exit Sub
        
catch:
   MsgBox Err.description, , "Opening the login form failed"
   Resume exitproc
   Resume
   
End Sub
