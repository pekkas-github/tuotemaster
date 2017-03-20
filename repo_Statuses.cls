VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "repo_Statuses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

'  This is an abstraction of status repository.

Private app             As Application_API
Private util            As Services

Private Sub Class_Initialize()
   
   Set app = new_Application_API
   Set util = new_Services
   
End Sub


Public Function getStatus(entityCode As String, versionNro As String) As dom_Status
' Return the current status of the entity version as an object or
' return nothing, if there is no statuses yet

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
   
   sql = "SELECT s.Id, s.StatusName_Id, s.Valid_From, s.Valid_To " & _
         "FROM " & VERSION_STATUS & " AS s " & _
         "INNER JOIN " & ITEM_VERSION & " AS v ON v.Id = s.ItemVersion_Id " & _
         "WHERE v.Item_Code = '" & entityCode & "' AND v.VersionNumber = '" & versionNro & "' " & _
         "AND s.Valid_From <= " & TODAY & " AND s.Valid_To > " & TODAY
   
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   If rst.EOF Then                     'current status is empty
      Set getStatus = Nothing
   Else
      Set getStatus = new_Status(rst!StatusName_Id, rst!Valid_From)
   End If
   

End Function


Public Sub saveStatus(entityCode As String, versionNro As String, NewStatus As dom_Status)
' Insert a new status object (and change end date in the previous one) or
' update the current status

   Dim currentStatus As dom_Status
   Dim versionID     As Long
   
   versionID = util.getVersionId(entityCode, versionNro)
   Set currentStatus = Me.getStatus(entityCode, versionNro)
   
   If currentStatus Is Nothing Then
      Call insertStatusData(versionID, NewStatus)
   Else
      If NewStatus.getStartDate = currentStatus.getStartDate Then
         currentStatus.setType (NewStatus.getType)
         Call updateStatusData(versionID, currentStatus)
      Else
         currentStatus.setEndDate (NewStatus.getStartDate - 1)
         Call updateStatusData(versionID, currentStatus)    ' Current status must be updated before the new status is saved (new becomes current)
         Call insertStatusData(versionID, NewStatus)
      End If
   End If
   
End Sub


Private Sub insertStatusData(versionID As Long, NewStatus As dom_Status)
' Add a new status into database

   Dim rst  As New ADODB.Recordset
   
   rst.Open VERSION_STATUS, app.getDbConn, adOpenDynamic, adLockPessimistic
   
   With rst
      .AddNew
      !StatusName_Id = Nz(NewStatus.getType)
      !ItemVersion_Id = versionID
      !Valid_From = NewStatus.getStartDate
      !valid_to = NewStatus.getEndDate
      .Update
   End With
   
   Set rst = Nothing

End Sub


Private Sub updateStatusData(versionID As Long, currentStatus As dom_Status)
' Update an existing status in database; change only status type and end date.
' Status type can be changed during the same day when status was created
' End date is changed when a new status replaces an old one.

   Dim rst  As New ADODB.Recordset
   Dim sql  As String
      
   sql = "SELECT * FROM " & VERSION_STATUS & " WHERE ItemVersion_Id = " & versionID & _
         " AND Valid_From <= " & TODAY & " AND " & TODAY & " < Valid_To"
   
   rst.Open sql, app.getDbConn, adOpenDynamic, adLockPessimistic
   
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
   Set util = Nothing
   
End Sub
