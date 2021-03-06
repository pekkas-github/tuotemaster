VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit
'MODIFIED 29.6.2015 in ver 6.1.0 (Data-object removed)
' GroupMapperilla hallitaan ryhmittelyhierarkioita.
' Rakenne on ROOT > Category > Group(s) > Item
' Kaikki hierarkian j�senet ovat Entiteettej�.
' Kokonaisia alihierarkioita voi leikata ja siirt�� paikasta toiseen.
' Orvoiksi j��neet alihierarkiat siivotaan
' Orvoiksi j��neet entiteetit siivotaan TOOLS-ty�kalulla.
'**

Private Data               As GroupMapper_Db
Private thisClipBoard       As Collection     ' Leikep�yt� kopioitujen/leikattujen ryhm�hierarkioiden tilap�iseen tallettamiseeen
Private thisCutPending      As Boolean        ' Ty�p�yd�ll� on kopioituja/leikattuja ryhm�hierarkioita

Public Sub Copy(ChildEntities As Collection)
'   Tallettaa "leikep�yd�lle" ryhm�hierarkioiden juurientiteetit my�hemp�� liitt�mist� varten.
'   Aikaisempi kopioitu tai leikattu joukko h�vi�� "leikep�yd�lt�"

    Set thisClipBoard = ChildEntities
    thisCutPending = False
    
End Sub

Public Sub Cut(ParentEntity As Entity, ChildEntities As Collection)
'   Irrotetaan is�nt�ryhm�st� kaikki kokoelmassa olevat lapset. Is�nt� voi olla ROOT, CAT tai GRP.
'   Lapset voivat olla mit� tahansa tyyppi�. Onnistuneesti irrotetut lapset j��v�t muistiin, kunnes
'       a) tehd��n uusi leikkaus/kopiointi (joukko j�� orvoksi -> garbage collection)
'       b) GroupMapper terminoituu (joukko j�� orvoksi -> garbage collection)

    Dim ChildEntity As Entity
    
    thisCutPending = True
    Set thisClipBoard = New Collection
    
    For Each ChildEntity In ChildEntities
        Call UnMap(ChildEntity, ParentEntity)
        thisClipBoard.Add ChildEntity
    Next
            
End Sub

Public Sub UnMap(ChildEntity As Entity, ParentEntity As Entity)
'   Disconnect this child entity from its parent entity

    Call Data.DeleteFromGroupHierarchy(ChildEntity, ParentEntity)
    
End Sub

Public Sub Paste(ParentEntity As Entity)
'   Liitet��n aikaisemmin kopioitu/leikattu lapsijoukko valittuun is�nt��n.
'   Liitt�misen j�lkeen lapsikokoelmaa ei tyhjennet�, jotta se voidaan liitt��
'   useampaan is�nt��n.

    Dim ChildEntity As Entity
    Dim ParentType  As String
    Dim ChildType   As String
    
    ParentType = ParentEntity.getType
    
    For Each ChildEntity In thisClipBoard
        ChildType = ChildEntity.getType
        If isMappingOk(ParentType, ChildType) Then
            Call Map(ChildEntity, ParentEntity)
        Else
            Err.Raise vbObjectError + 513
        End If
    Next
        
    thisCutPending = False
    
End Sub
Private Function isMappingOk(ParentType As String, ChildType As String) As Boolean
' MODIFIED in ver 6.2.1 (CAT -> MR and MR -> MN added)
'   Lapsett voivat olla tyypilt��n mit� tahansa. Liitt�misess� on huolehdittava, ett�
'       a) ROOT -> CAT
'       b) CAT -> GRP tai MR
'       c) GRP -> GRP, MR, MN, CF
'       d) MR -> MN

    If _
    (ParentType = "ROOT" And ChildType = "CAT") Or _
    (ParentType = "CAT" And (ChildType = "GRP" Or ChildType = "MR")) Or _
    (ParentType = "MR" And ChildType = "MN") Or _
    (ParentType = "GRP" And ChildType <> "CAT") _
    Then
        isMappingOk = True
    Else
        isMappingOk = False
    End If
    
End Function

Public Sub Map(ChildEntity As Entity, ParentEntity As Entity)
'   Connect the child entity under the selected entity
    
    'If item is already mapped then skip
    If Data.IsMappedToParent(ChildEntity, ParentEntity) Then Exit Sub
    
    'Otherwise connect item to the parent
    Call Data.InsertInGroupHierarchy(ChildEntity, ParentEntity)
        
End Sub

Public Function isCutPending() As Boolean
    isCutPending = thisCutPending
End Function

Private Sub Class_Initialize()
    
    Set Data = New GroupMapper_Db
    thisCutPending = False
    
End Sub

Private Sub Class_Terminate()

   Set Data = Nothing
   
End Sub
