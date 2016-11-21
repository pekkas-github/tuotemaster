VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
'
' Services class - all service methods that are not part of any domain class
'
Private db  As Services_Db

Private Sub Class_Initialize()

   Set db = New Services_Db
   
End Sub

Public Function getNextPriceLineCode() As String
' Palauttaa seuraavan vapaan price line -koodin. Prefix on aina "C" ja digittej‰ on 7

    Dim lastPriceLineCode      As String
    Dim nextPriceLineCode      As String
    
    lastPriceLineCode = db.getLastPriceLineCode
    nextPriceLineCode = calculateNextEntityCode(lastPriceLineCode, "C", 7)
    
    getNextPriceLineCode = nextPriceLineCode

End Function
Public Function getNextEntityCode(prefix As String) As String
' MODIFIED 28.8.2015 in ver. 6.2.0 / Number of digits (6) added because of a change in a function call
' Palauttaa seuraavan vapaan entity-koodin kyseiselle entity-tyypille.

    Dim lastEntityCode      As String
    Dim nextEntityCode      As String
    
    lastEntityCode = db.GetLastEntityCode(prefix)
    nextEntityCode = calculateNextEntityCode(lastEntityCode, prefix, 6)
    
    getNextEntityCode = nextEntityCode

End Function

Private Function calculateNextEntityCode(lastCode As String, prefix As String, digits As Integer) As String
' Return the next free code of a given entity type

    Dim numberPart  As String
    Dim number      As Long

    numberPart = Right(lastCode, digits)
    number = CLng(numberPart)
    number = number + 1
    numberPart = CStr(number)
    While Len(numberPart) < digits
        numberPart = "0" & numberPart
    Wend
    calculateNextEntityCode = prefix & numberPart
    
End Function

Public Function getNextVersionNumber(entityCode As String) As String
' Return the next to last major version number

   Dim versionNro    As String
   Dim majorPart     As Integer
   
   versionNro = getLastVersionNumber(entityCode)
   majorPart = Me.ParseVersionMajorPart(versionNro)
   majorPart = majorPart + 1
   
   getNextVersionNumber = Me.ConcatVersionNumber(majorPart, 0)
   
End Function


Public Function getLastVersionNumber(entityCode As String) As String
' Return the last version number of the entity
' If there is no versions then return "0.0"
    
    Dim versionNumbers  As Collection
    Dim Major           As Integer
    Dim Minor           As Integer
    Dim m               As Integer
    Dim n               As Integer
    Dim VersionNumber   As Variant
        
    Major = 0
    Minor = 0
    
    Set versionNumbers = db.GetVersionNumbers(entityCode)
    
    For Each VersionNumber In versionNumbers
        m = ParseVersionMajorPart(CStr(VersionNumber))
        n = ParseVersionMinorPart(CStr(VersionNumber))
        
        If m = Major And n > Minor Then
            Minor = n
        End If
        
        If m > Major Then
            Major = m
            Minor = n
        End If
    Next
    
    getLastVersionNumber = ConcatVersionNumber(Major, Minor)
    
    Set versionNumbers = Nothing
    
End Function
Public Function ConcatVersionNumber(Major As Integer, Minor As Integer) As String
'** Yhdist‰‰ major ja minor lukuarvot merkkijonoksi "m.n"

    ConcatVersionNumber = CStr(Major) & "." & CStr(Minor)

End Function
Public Function ParseVersionMajorPart(VersionNumber As String) As Integer
' Palauttaa version "m.n" major osan lukuna m.
    
    ParseVersionMajorPart = CInt(Left(VersionNumber, (InStr(1, VersionNumber, ".", 1) - 1)))
    
End Function

Public Function ParseVersionMinorPart(VersionNumber As String) As Integer
' Palauttaa version "m.n" minor osan lukuna n.

    ParseVersionMinorPart = CInt(Right(VersionNumber, (Len(VersionNumber) - InStr(1, VersionNumber, ".", 1))))
    
End Function


Public Sub deleteEntityPermanently(entityCode As String)
' Delete the entity and all of its attributes and relations permanently from database
' Basically the only entities that are deleted are categories and subgroups
   Dim versionIdCol  As Collection
   Dim verId         As Variant
   
   Set versionIdCol = db.getVersionIds(entityCode)
   
   For Each verId In versionIdCol
      db.deleteDescriptions (verId)
      db.deleteStatuses (verId)
      db.deleteDecisions (verId)
      db.deleteVersion (verId)
   Next
   
   db.deleteNames (entityCode)
   db.deleteOwners (entityCode)
   db.deleteProperties (entityCode)
   db.deleteEntity (entityCode)
   
End Sub
Private Sub Class_Terminate()
   
   Set db = Nothing
   
End Sub
