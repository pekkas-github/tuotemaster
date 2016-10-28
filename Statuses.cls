VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Statuses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'CREATED in HOT#005
'  This is an abstraction of status repository.

Private app             As Application_API
Private dbconn          As ADODB.Connection

Public Sub init()
   
   Set app = New Application_API
   Set dbconn = CurrentProject.Connection
   
End Sub


Public Function getStatus(entityCode As String, versionNro As String) As Status
' Return the current status of the entity version as an object or
' return nothing, if there is no statuses yet

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT s.Id, s.StatusName_Id, s.Valid_From, s.Valid_To " & _
         "FROM " & VERSION_STATUS & " AS s " & _
         "INNER JOIN " & ITEM_VERSION & " AS v ON v.Id = s.ItemVersion_Id " & _
         "WHERE v.Item_Code = '" & entityCode & "' AND v.VersionNumber = '" & versionNro & "' " & _
         "AND s.Valid_From <= " & TODAY & " AND s.Valid_To > " & TODAY
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   If rst.EOF Then                     'current status is empty
      Set getStatus = Nothing
   Else
      Set getStatus = new_Status(rst!StatusName_Id, rst!Valid_From)
   End If
   

End Function


Public Sub saveStatus(entityCode As String, versionNro As String, newStatus As Status)
' Insert a new status object (and change end date in the previous one) or
' update the current status

   Dim currentStatus As Status
   Dim versionId     As Long
   
   versionId = DLookup("Id", ITEM_VERSION, "Item_Code = '" & entityCode & "' AND VersionNumber = '" & versionNro & "'")
   Set currentStatus = Me.getStatus(entityCode, versionNro)
   
   If currentStatus Is Nothing Then
      Call insertStatusData(versionId, newStatus)
   Else
      If newStatus.getStartDate = currentStatus.getStartDate Then
         currentStatus.setType (newStatus.getType)
         Call updateStatusData(versionId, currentStatus)
      Else
         currentStatus.setEndDate (newStatus.getStartDate - 1)
         Call updateStatusData(versionId, currentStatus)    ' Current status must be updated before the new status is saved (new becomes current)
         Call insertStatusData(versionId, newStatus)
      End If
   End If
   
End Sub


Private Sub insertStatusData(versionId As Long, newStatus As Status)
' Add a new status into database

   Dim rst  As New ADODB.Recordset
   
   rst.Open VERSION_STATUS, dbconn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !StatusName_Id = Nz(newStatus.getType)
      !ItemVersion_Id = versionId
      !Valid_From = newStatus.getStartDate
      !valid_to = newStatus.getEndDate
      .Update
   End With
   
   Set rst = Nothing

End Sub


Private Sub updateStatusData(versionId As Long, currentStatus As Status)
' Update an existing status in database; change only status type and end date.
' Status type can be changed during the same day when status was created
' End date is changed when a new status replaces an old one.

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
      
   sql = "SELECT * FROM " & VERSION_STATUS & " WHERE ItemVersion_Id = " & versionId & _
         " AND Valid_From <= " & TODAY & " AND " & TODAY & " < Valid_To"
   
   rst.Open sql, dbconn, adOpenDynamic, adLockPessimistic
   
   With rst
      !StatusName_Id = Nz(currentStatus.getType)
      !valid_to = currentStatus.getEndDate
      .Update
   End With
   
   Set rst = Nothing

End Sub


Public Function rsString(entityType As String, language As String, hideTerminated As Boolean) As String
'  Return valid status values and names of this entity type.
'  Includes or excludes terminated status.
   
   Dim sql  As String
   
   sql = "SELECT n.Id, n.Text FROM " & STATUS_NAME & " AS n " & _
         "INNER JOIN Statuses AS s ON n.Id = s.StatusName_Id " & _
         "WHERE s.ItemType_Id = '" & entityType & "' AND n.Lang_Code = '" & language & "' "
   
   If hideTerminated Then
      sql = sql & "AND NOT n.Id = 4 AND NOT n.Id = 9 "
   End If
   
   sql = sql & "ORDER BY n.[Order]"
      
   rsString = sql
   
End Function


Private Sub Class_Terminate()

   Set app = Nothing
   Set dbconn = Nothing
   
End Sub
