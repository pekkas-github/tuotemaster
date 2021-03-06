Attribute VB_Name = "Rowsources"
Option Compare Database
Option Explicit

'	This class carries a collection of rowsource sql statements for UI controls.

Public Function rsCategories(filters As Collection) As String
' This method returns a rowsource for UI.
' It shows a list of (last versions of the) categories
' There is a subquery sqlLastVesions to filter out only the
' latest versions (basically there doesn't exist versions)

   Dim sql     As String
   Dim where   As String
   
   sql = "SELECT i.Code, n.Text AS Name, sn.Text AS Status, p.PersonName AS Owner " & _
         "FROM (((((((" & ITEM_VERSION & " AS v " & _
         "INNER JOIN (" & sqlLastVersions & ") AS v1 ON v.Id = v1.Id) " & _
         "INNER JOIN " & VERSION_STATUS & " AS s ON v.Id = s.ItemVersion_Id) " & _
         "INNER JOIN " & STATUS_NAME & " AS sn ON sn.Id = s.StatusName_Id) " & _
         "INNER JOIN " & CORE_ITEM & " AS i ON i.Code = v.Item_Code) " & _
         "INNER JOIN " & ITEM_NAME & " AS n ON i.Code = n.Item_Code) " & _
         "INNER JOIN " & ITEM_OWNER & " AS o ON i.Code = o.Item_Code) " & _
         "INNER JOIN " & PARTY & " AS p ON p.Id = o.Party_Id) " & _
         "WHERE n.Lang_Code = '" & Globals.lang & "' AND sn.Lang_Code = '" & Globals.lang & "' " & _
              "AND o.Valid_From <= " & TODAY & " AND o.Valid_To > " & TODAY & _
              " AND s.Valid_From <=" & TODAY & " AND s.Valid_To > " & TODAY & _
              " AND i.Type = 'CAT'"
         
'  Check filters
   'Don't show "Deleted" and (optionaly) "Terminated"
   If Not filters("All") Then
       where = where & " AND NOT (s.StatusName_Id = 4 OR s.StatusName_Id = 9) "
   End If

   sql = sql & where & " ORDER BY n.Text"

   rsCategories = sql
   
End Function


Public Function rsProducts(filters As Collection) As String
' This method returns a list of last versions of the products
' filtered by code / name / status /owner
' There is a subquery sqlLastVesions to filter out only the
' latest versions

   Dim sql     As String
   Dim where   As String
   
   sql = "SELECT i.Code, n.Text AS Name, sn.Text AS Status, p.PersonName AS Owner " & _
         "FROM (((((((" & ITEM_VERSION & " AS v " & _
         "INNER JOIN (" & sqlLastVersions & ") AS v1 ON v.Id = v1.Id) " & _
         "INNER JOIN " & VERSION_STATUS & " AS s ON v.Id = s.ItemVersion_Id) " & _
         "INNER JOIN " & STATUS_NAME & " AS sn ON sn.Id = s.StatusName_Id) " & _
         "INNER JOIN " & CORE_ITEM & " AS i ON i.Code = v.Item_Code) " & _
         "INNER JOIN " & ITEM_NAME & " AS n ON i.Code = n.Item_Code) " & _
         "INNER JOIN " & ITEM_OWNER & " AS o ON i.Code = o.Item_Code) " & _
         "INNER JOIN " & PARTY & " AS p ON p.Id = o.Party_Id) " & _
         "WHERE n.Lang_Code = '" & Globals.lang & "' AND sn.Lang_Code = '" & Globals.lang & "' " & _
              "AND o.Valid_From <= " & TODAY & " AND o.Valid_To > " & TODAY & _
              " AND s.Valid_From <=" & TODAY & " AND s.Valid_To > " & TODAY & _
              " AND i.Type = 'MR'"
         
'  Check filters
   'Don't show "Deleted" and (optionaly) "Terminated"
   If Not filters("All") Then
       where = where & " AND NOT (s.StatusName_Id = 4 OR s.StatusName_Id = 9) "
   End If
       
   'Code
   If Not filters("Code") = "" Then
       where = where & " AND i.Code Like '" & filters("Code") & "'"
   End If
   'Name
   If Not filters("Name") = "" Then
       where = where & " AND n.Text Like '" & filters("Name") & "'"
   End If
   'Status
   If filters("Status") > 0 Then
       where = where & " AND s.StatusName_Id = " & filters("Status")
   End If
   'Owner
   If filters("Owner") > 0 Then
       where = where & " AND p.Id = " & filters("Owner")
   End If
   
   sql = sql & where & " ORDER BY n.Text"
    
'  SQL Server uses % as a wildcard character. In UI we want to use * as a wildcard character
   sql = Replace(sql, "*", WILD_CARD)

   rsProducts = sql
         
End Function


Public Function rsServices(filters As Collection) As String
' This method returns a rowsource for UI.
' It shows a list of last versions of the services filtered
' by code / name / status /owner
' There is a subquery sqlLastVesions to filter out only the
' latest versions

   Dim sql     As String
   Dim where   As String
   
   sql = "SELECT i.Code, n.Text AS Name, sn.Text AS Status, p.PersonName AS Owner " & _
         "FROM (((((((" & ITEM_VERSION & " AS v " & _
         "INNER JOIN (" & sqlLastVersions & ") AS v1 ON v.Id = v1.Id) " & _
         "INNER JOIN " & VERSION_STATUS & " AS s ON v.Id = s.ItemVersion_Id) " & _
         "INNER JOIN " & STATUS_NAME & " AS sn ON sn.Id = s.StatusName_Id) " & _
         "INNER JOIN " & CORE_ITEM & " AS i ON i.Code = v.Item_Code) " & _
         "INNER JOIN " & ITEM_NAME & " AS n ON i.Code = n.Item_Code) " & _
         "INNER JOIN " & ITEM_OWNER & " AS o ON i.Code = o.Item_Code) " & _
         "INNER JOIN " & PARTY & " AS p ON p.Id = o.Party_Id) " & _
         "WHERE n.Lang_Code = '" & Globals.lang & "' AND sn.Lang_Code = '" & Globals.lang & "' " & _
              "AND o.Valid_From <= " & TODAY & " AND o.Valid_To > " & TODAY & _
              " AND s.Valid_From <=" & TODAY & " AND s.Valid_To > " & TODAY & _
              " AND i.Type = 'CF'"
         
'  Check filters
   'Don't show "Deleted" and (optionaly) "Terminated"
   If Not filters("All") Then
       where = where & " AND NOT (s.StatusName_Id = 4 OR s.StatusName_Id = 9) "
   End If
       
   'Code
   If Not filters("Code") = "" Then
       where = where & " AND i.Code Like '" & filters("Code") & "'"
   End If
   'Name
   If Not filters("Name") = "" Then
       where = where & " AND n.Text Like '" & filters("Name") & "'"
   End If
   'Status
   If filters("Status") > 0 Then
       where = where & " AND s.StatusName_Id = " & filters("Status")
   End If
   'Owner
   If filters("Owner") > 0 Then
       where = where & " AND p.Id = " & filters("Owner")
   End If
   
   sql = sql & where & " ORDER BY n.Text"
    
'  SQL Server uses % as a wildcard character. In UI we want to use * as a wildcard character
   sql = Replace(sql, "*", WILD_CARD)

   rsServices = sql
         
End Function


Private Function sqlLastVersions() As String
' This subquery is used in rowsources where we show only entity level
' without individua version. This query returns version id of the last
' version of each entity.

    sqlLastVersions = "SELECT v0.Item_Code, Max(v0.Id) AS Id " & _
                "FROM " & ITEM_VERSION & " AS v0 " & _
                "GROUP BY v0.Item_Code"

End Function


Public Function rsOwners() As String
'  Return current owners of all entities.
'  This rowsource is used in drop down lists when
'  entities are filtered by owners.

   Dim sql As String
   
   sql = "SELECT DISTINCT p.Id, p.PersonName " & _
         "FROM " & ITEM_OWNER & " AS o " & _
         "INNER JOIN " & PARTY & " AS p ON o.Party_Id = p.Id " & _
         "ORDER BY p.PersonName"

   rsOwners = sql
   
End Function


Public Function rsOwnerHistory(entityCode As String) As String
' Rowsource that returns owner history of a specific entity
' This is used in entity details forms.
   
   Dim sql As String
   
   sql = "SELECT p.Id, p.PersonName AS Owner, o.Valid_From AS StartDate " & _
         "FROM " & ITEM_OWNER & " AS o " & _
         "INNER JOIN " & PARTY & " AS p ON o.Party_Id = p.Id " & _
         "WHERE o.Item_Code = '" & entityCode & "' " & _
         "ORDER BY o.Id DESC"
               
   rsOwnerHistory = sql

End Function


Public Function rsParties() As String
' Return all parties in the person register
' This is used when party is selected for certain reason
' e.g. when defining the product owner.

   Dim sql  As String

    sql = "SELECT Id, PersonName, Heno " & _
          "FROM " & PARTY & _
          " ORDER BY PersonName"
          
    rsParties = sql
    
End Function



Public Function rsStatuses(entityType As String, showAll As Boolean) As String

'  Return valid status values and their names of this entity type.
'  This rowsource is used in drop down lists when a specific status
'  have to be selected.
   
   Dim sql  As String
   
   sql = "SELECT n.Id, n.Text " & _
         "FROM " & STATUS_NAME & " AS n " & _
         "INNER JOIN " & ITEM_STATUS & " AS s ON n.Id = s.StatusName_Id " & _
         "WHERE s.ItemType_Id = '" & entityType & "' AND n.Lang_Code = '" & Globals.lang & "' "
   
   If Not showAll Then
      sql = sql & "AND n.Id <> 4 AND  n.Id <> 9 "
   End If
   
   sql = sql & "ORDER BY n.[Order]"
      
   rsStatuses = sql

End Function


Public Function rsBillingItems(filters As Collection) As String
' Shows a filtered list of all billing codes (MIPA, Tellus, JPD) imported from Netezza and
' shows the mapping status of each code
' If there are no filters then the list is empty.

   Dim sql As String
   Dim where As String
             
   where = ""
   
   'Check the filters
   If Not filters("Järjestelmä") = "" Then
       where = where & " AND Järjestelmä Like '" & filters("Järjestelmä") & "'"
   End If
   
   If Not filters("Tuotetunnus") = "" Then
       where = where & " AND Tuotetunnus Like '" & filters("Tuotetunnus") & "'"
   End If
   
   If Not filters("Tuotenimi") = "" Then
       where = where & " AND Tuotenimi Like '" & filters("Tuotenimi") & "'"
   End If
   
   If Not filters("Tuotetyyppi") = "" Then
       where = where & " AND Tuotetyyppi Like '" & filters("Tuotetyyppi") & "'"
   End If
   
   If Not filters("Tuoteryhmä") = "" Then
       where = where & " AND Tuoteryhmä Like '" & filters("Tuoteryhmä") & "'"
   End If
   
   If Not filters("TRnimi") = "" Then
       where = where & " AND [Tuoteryhmän nimi] Like '" & filters("TRnimi") & "'"
   End If
   
   If Not filters("Käyttöluokka") = "" Then
       where = where & " AND Käyttöluokka Like '" & filters("Käyttöluokka") & "'"
   End If
   
   If Not filters("SAPkoodi") = "" Then
       where = where & " AND [refSAP-koodi] Like '" & filters("SAPkoodi") & "'"
   End If
   
   If Not filters("Mapping") = 0 Then
       where = where & " AND m.StatusName_Id = " & filters("Mapping")
   End If
   
   If filters("Aktiivinen") = True Then
       where = where & " AND (Lopetuspäivämäärä > " & TODAY & " OR Lopetuspäivämäärä IS NULL) "
   End If
       
'  If no filters then view empty list otherwise show line items
   If where = "" Then
      sql = "SELECT * FROM dbo_Model WHERE 1=2"   ' a pseudo rowsource to generate an epmpty list
   Else
      sql = "SELECT Järjestelmä AS System, Tuotetunnus AS Code, Tuotenimi AS ProductName, Tuotetyyppi AS Type, Tuoteryhmä AS [Group], [Tuoteryhmän nimi] AS GroupName, Luontipäivämäärä AS Created, Lopetuspäivämäärä AS Expired, Käyttöluokka AS Status, [refSAP-koodi] AS SAP, m.StatusName_Id, LegacyStatusName AS Mapping, m.Comment " & _
            "FROM (" & BILLING_ITEMS & " AS b INNER JOIN " & BILLING_STATUS & " AS m ON (b.Järjestelmä = m.System AND b.Tuotetunnus = m.Code)) " & _
            "INNER JOIN " & BILLING_STATUS_NAME & " AS n ON m.StatusName_Id = n.Id "
      
      sql = sql & "WHERE 1=1 " & where
      sql = Replace(sql, "*", WILD_CARD)
   End If
   
   rsBillingItems = sql

End Function

Public Function rsCurrentPropertyValues(entityCode As String, entityType As String, language As String) As String
'  Return rowsource with valid properties and their current values

   Dim sqlDefaultProperties   As String
   Dim sqlItemValues          As String
   Dim sqlRowsource           As String
   
   'default properties for this entity type
   sqlDefaultProperties = "SELECT DISTINCT a.Property_Id, b.LookupText AS Property " & _
                          "FROM " & ITEM_TYPE_PROPERTY_VALUE & " AS a " & _
                          "INNER JOIN " & PROPERTY_TYPE & " AS b ON a.Property_Id = b.Id " & _
                          "WHERE ItemType = '" & entityType & "' AND b.Language = '" & language & "'"
   
   'current property values of this entity
   sqlItemValues = "SELECT a.Property_Id, b.LookupText AS Value " & _
                   "FROM " & ITEM_PROPERTY_VALUE & " AS a " & _
                   "INNER JOIN " & PROPERTY_VALUE & " AS b ON a.Property_Id = b.Property_Id AND a.Value_Id = b.Value_Id " & _
                   "WHERE Item_Code = '" & entityCode & "' AND b.Language = '" & language & "'"
   
   '
   sqlRowsource = "SELECT dp.Property_Id, dp.Property, val.Value " & _
                  "FROM (" & sqlDefaultProperties & ") AS dp " & _
                  "LEFT JOIN (" & sqlItemValues & ") AS val ON dp.Property_Id = val.Property_Id"
                  
   rsCurrentPropertyValues = sqlRowsource
                    
End Function


Public Function rsValidPropertyValues(entityType As String, propertyType As String, language As String) As String
'  Return rowsource with valid property values for this entity-property type combination

   Dim sql As String
   
   sql = "SELECT pro.Id, val.Value_Id, pro.LookupText AS Property, val.LookupText AS [Value] " & _
         "FROM (" & ITEM_TYPE_PROPERTY_VALUE & " AS itpv " & _
         "INNER JOIN " & PROPERTY_TYPE & " AS pro ON itpv.Property_Id = pro.Id) " & _
         "INNER JOIN " & PROPERTY_VALUE & " AS val ON itpv.Value_Id = val.Value_Id " & _
         "WHERE itpv.ItemType = '" & entityType & "' " & _
         "AND pro.Id = '" & propertyType & "' " & _
         "AND pro.Language = '" & language & "' " & _
         "AND val.Language = '" & language & "' " & _
         "ORDER BY pro.LookupText"
         
   rsValidPropertyValues = sql
   
End Function

Public Function rsSalesItems(productCode As String) As String

    Dim sql     As String
    
    sql = "SELECT i.Code, n.Text AS Name, v.VersionNumber AS Version, sn.Text AS Status " & _
          "FROM (((( " & CORE_ITEM & " AS i " & _
          "INNER JOIN " & ITEM_NAME & " AS n ON i.Code = n.Item_Code) " & _
          "INNER JOIN " & ITEM_VERSION & " AS v ON i.Code = v.Item_Code) " & _
          "INNER JOIN " & VERSION_STATUS & " AS s ON v.Id = s.ItemVersion_Id) " & _
          "INNER JOIN " & STATUS_NAME & " AS sn ON sn.Id = s.StatusName_Id) " & _
          "INNER JOIN " & GROUP_HIERARCHY & " AS g ON i.Code = g.Child_Code " & _
          "WHERE n.Lang_Code = '" & lang & "' AND sn.Lang_Code = '" & lang & "' " & _
             "AND s.Valid_From <= " & TODAY & " AND s.Valid_To > " & TODAY & _
             " AND g.Parent_Code = '" & productCode & "' " & _
         "ORDER BY n.Text"

   rsSalesItems = sql
   
End Function


Public Function rsMemberInCategories(childCode As String) As String
' A rowsource that provides all reverse category paths from an entity to the top level category
' The rowsource is a value chain and its format is
'  "CAT000001; cat_1 > group_a > group_b; CAT000002; cat_2 > group_c; " etc.

   Dim db            As ADODB.Connection
   Dim parentCodes   As New ADODB.Recordset
   Dim sql           As String
   Dim rowSource     As String
   
   Set db = CurrentProject.Connection
   rowSource = "Category;Path;"
   
   sql = "SELECT Parent_Code FROM " & GROUP_HIERARCHY & " WHERE Child_Code = '" & childCode & "'"
   
   parentCodes.Open sql, db, adOpenDynamic, adLockPessimistic
   
   While Not parentCodes.EOF
      rowSource = rowSource & getSinglePath(parentCodes!Parent_Code) & ";"
      parentCodes.MoveNext
   Wend
   
   rsMemberInCategories = rowSource
   
   Set parentCodes = Nothing
   Set db = Nothing
   
End Function

Private Function getSinglePath(childCode As String) As String
' Create a reverese group hierarchy path of a child group.
' Assumption is that each subgroup has only one parent.
' Format is "CAT > subGroup > subGroup etc."

   Dim db            As ADODB.Connection
   Dim parentItem    As New ADODB.Recordset
   Dim sqlParent     As String
   Dim thisItemName  As String
   
   Set db = CurrentProject.Connection
   
   sqlParent = "SELECT Parent_Code, Parent_Type FROM " & GROUP_HIERARCHY & " WHERE Child_Code = '" & childCode & "'"
                    
   parentItem.Open sqlParent, db, adOpenDynamic, adLockPessimistic
   
   If parentItem!Parent_Type = "ROOT" Then
      getSinglePath = DLookup("Text", ITEM_NAME, "Item_Code = '" & childCode & "' AND Lang_Code = '" & lang & "'") & ";"
   Else
      thisItemName = DLookup("Text", ITEM_NAME, "Item_Code = '" & childCode & "' AND Lang_Code = '" & lang & "'")
      getSinglePath = getSinglePath(parentItem!Parent_Code) & " --> " & thisItemName
   End If
   
End Function

Public Function rsDecisions(entityCode As String, versionNro As String) As String

   Dim sql  As String
   
   sql = "SELECT d.Id, n.Text AS Status, d.ValidFrom AS [Date], d.DecisionText AS Decision " & _
         "FROM (((" & VERSION_DECISION & " AS d " & _
         "INNER JOIN " & ITEM_VERSION & " AS v ON v.Id = d.ItemVersion_Id) " & _
         "INNER JOIN " & VERSION_STATUS & " AS s " & _
               "ON d.ItemVersion_Id = s.ItemVersion_Id AND " & _
               "d.ValidFrom >= s.Valid_From AND " & _
               "d.ValidFrom <= s.Valid_To) " & _
         "INNER JOIN " & STATUS_NAME & " AS n ON s.StatusName_Id = n.Id) " & _
         "WHERE n.Lang_Code = '" & lang & "' AND v.Item_Code = '" & entityCode & "' AND v.VersionNumber = '" & versionNro & "' " & _
         "ORDER BY d.Id DESC"

   rsDecisions = sql
   
End Function


Public Function rsProperties(entityType As String, entityCode As String) As String
'  Return rowsource with valid properties and their current values

   Dim sqlDefaultProperties   As String
   Dim sqlItemValues          As String
   Dim sqlRowsource           As String
   
   'default properties for this entity type
   sqlDefaultProperties = "SELECT DISTINCT a.Property_Id, b.LookupText AS Property " & _
                          "FROM " & ITEM_TYPE_PROPERTY_VALUE & " AS a " & _
                          "INNER JOIN " & PROPERTY_TYPE & " AS b ON a.Property_Id = b.Id " & _
                          "WHERE ItemType = '" & entityType & "' AND b.Language = '" & lang & "'"
   
   'current property values of this entity
   sqlItemValues = "SELECT a.Property_Id, b.LookupText AS [Value] " & _
                   "FROM " & ITEM_PROPERTY_VALUE & " AS a " & _
                   "INNER JOIN " & PROPERTY_VALUE & " AS b ON a.Property_Id = b.Property_Id AND a.Value_Id = b.Value_Id " & _
                   "WHERE Item_Code = '" & entityCode & "' AND b.Language = '" & lang & "'"
   
   '
   sqlRowsource = "SELECT dp.Property_Id, dp.Property, val.Value " & _
                  "FROM (" & sqlDefaultProperties & ") AS dp " & _
                  "LEFT JOIN (" & sqlItemValues & ") AS val ON dp.Property_Id = val.Property_Id"

   rsProperties = sqlRowsource
   
End Function
