Attribute VB_Name = "Globals"
Option Compare Database
Option Explicit
'
Public Args             As Collection   ' Collection is used to pass data between forms
Public lang             As String       ' UI language for product data in UI (UI labels are always in English)
Public clipBoard        As Collection   ' temporary storage for copy and paste

'  Usergroups
   Public Const ADMIN      As Integer = 1
   Public Const READ_ONLY  As Integer = 100
   Public Const UNKNOWN    As Integer = 1000
   
'  Tämän modulin versio
   Public Const APP_VERSION   As String = "6.0.0.dev"
   Public Const TARGET_DB     As String = "Local_DB"
   Public Const IS_LIVE       As Boolean = False


'  Lomakepohjien värit
   Public Const HeaderBackColor  As Long = &HB4835C
   Public Const FooterBackColor  As Long = &HB4835C
   Public Const HeaderForeColor  As Long = &HFFFFFF
   Public Const ContentBackColor As Long = &HECECEC
    
'  Tietokantataulujen nimet
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
   Public Const BILLING_MAPPING       As String = "dbo_LegacyMapping"
   Public Const BILLING_STATUS        As String = "dbo_LegacyStatusMapping"
   Public Const BILLING_STATUS_NAME   As String = "dbo_LegacyStatusName"
   Public Const PRICE_LINE            As String = "dbo_PriceLine"
   Public Const PROPERTY_TYPE         As String = "dbo_Properties"
   Public Const PROPERTY_VALUE        As String = "dbo_PropertyValues"
   Public Const ITEM_PROPERTY_VALUE   As String = "dbo_ItemPropertyValues"
   Public Const ITEM_TYPE_PROPERTY_VALUE As String = "dbo_ItemTypePropertyValues"
   Public Const GROUP_HIERARCHY          As String = "dbo_GroupHierarchy"
   Public Const DOC_REFERENCE         As String = "dbo_DocumentReference"
   Public Const DOC_TYPE              As String = "dbo_DocumentType"
    
    
    
    
    
    
    
' Differences in characteristics of Access and T-SQL
   Public Const TODAY As String = "DATE()"
   Public Const WILD_CARD As String = "*"


' User defined error codes
   Enum e
      notUniqueName = vbObjectError + 1000
      actionNotAllowed = vbObjectError + 1001
      recordnotfound = vbObjectError + 1002
   End Enum
    
'   Filter attributes in queries
    Public Type WhereFilter
        Järjestelmä  As String
        Tuotetunnus  As String
        Tuotenimi    As String
        Tuotetyyppi  As String
        Tuoteryhmä   As String
        TRnimi       As String
        Luontipvm    As Date
        Lopetuspvm   As Date
        Aktiivinen   As Boolean
        Käyttöluokka As String
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
   
