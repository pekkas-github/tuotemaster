Attribute VB_Name = "Globals"
Option Compare Database
Option Explicit
'
Public Args             As Collection   ' Collection is used to pass data between forms
Public lang             As String       ' Selected translation of product data (UI labels are always in English
Public AppVersion       As String       ' Application version in string mode

'Usergroups
   Public Const ADMIN      As Integer = 1
   Public Const READ_ONLY  As Integer = 100
   Public Const UNKNOWN    As Integer = 1000
   
'   T�m�n modulin versio
    Public Const VER_MAJOR As Integer = 0
    Public Const VER_MINOR As Integer = 0
    Public Const VER_PATCH As Integer = 0
    Public Const IS_LIVE   As Boolean = False


'   Lomakepohjien v�rit
    Public Const HeaderBackColor  As Long = &HB4835C
    Public Const FooterBackColor  As Long = &HB4835C
    Public Const HeaderForeColor  As Long = &HFFFFFF
    Public Const ContentBackColor As Long = &HECECEC
    
'   Tietokantataulujen nimet
    Public Const CORE_ITEM             As String = "dbo_Item"
    Public Const ITEM_NAME             As String = "dbo_ItemName"
    Public Const ITEM_OWNER            As String = "dbo_ItemOwner"
    Public Const ITEM_VERSION          As String = "dbo_ItemVersion"
    Public Const VERSION_STATUS        As String = "dbo_ItemVersionStatus"
    Public Const VERSION_DESCRIPTION   As String = "dbo_ItemVersionDescription"
    Public Const VERSION_DECISION      As String = "dbo_ItemVersionStatusDecision"
    Public Const STATUS_NAME           As String = "dbo_ItemVersionStatusName"
    Public Const PARTY                 As String = "dbo_ltbPerson"
    Public Const ITEM_STATUS           As String = "dbo_Statuses"
    Public Const BILLING_ITEMS         As String = "dbo_lnk_BSS_tuotekannat"
    Public Const BILLING_STATUS        As String = "dbo_LegacyStatusMapping"
    Public Const BILLING_STATUS_NAME   As String = "dbo_LegacyStatusName"
    Public Const PRICE_LINE            As String = "dbo_PriceLine"
    Public Const PROPERTY_TYPE         As String = "dbo_Properties"
    Public Const PROPERTY_VALUE        As String = "dbo_PropertyValues"
    Public Const ITEM_PROPERTY_VALUE   As String = "dbo_ItemPropertyValues"
    Public Const ITEM_TYPE_PROPERTY_VALUE As String = "dbo_ItemTypePropertyValues"
    Public Const GROUP_HIERARCHY          As String = "dbo_GroupHierarchy"
    
    
    
    
    
' Differences in characteristics of Access and T-SQL
   Public Const TODAY As String = "DATE()"
   Public Const WILD_CARD As String = "*"


' User defined error codes
   Enum e
      notUniqueName = vbObjectError + 1000
      actionNotAllowed = vbObjectError + 1001
   End Enum
    
'   Filter attributes in queries
    Public Type WhereFilter
        J�rjestelm�  As String
        Tuotetunnus  As String
        Tuotenimi    As String
        Tuotetyyppi  As String
        Tuoteryhm�   As String
        TRnimi       As String
        Luontipvm    As Date
        Lopetuspvm   As Date
        Aktiivinen   As Boolean
        K�ytt�luokka As String
        SAPkoodi     As String
        Mapping      As Integer
    End Type
   
'  Filter attributes in queries
   Public Type entityFilter
      code     As String
      Name     As String
      Status   As Integer
      Owner    As String
   End Type
   
