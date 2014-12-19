VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PageNavigationClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'
Private WorkingQueryString As String
'
Const PageTypeRootChild = 1
Const PageTypeSibling = 2
Const PageTypeChild = 3
Const PageTypeSiblingChild = 4
'
Const SectionPages = "SectionPages"
Const SisterPages = "SisterPages"
Const ChildPages = "ChildPages"
Const SisterChildPages = "SisterChildPages"
'
Const ContentNamePageContent = "Page Content"
'
Private LoadedPageID As Long
'
Private Main As Object
Private csv As Object
'
' v3.4 Interface
'
Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String)
    '
    Call Init(MainObject)
    Set csv = CsvObject
    Call Main.TestPoint("pageNavigationClass, OptionString=" & OptionString)
    Execute = GetContent(OptionString)
    '
    Exit Function
ErrorTrap:
    Call HandleError("PageExtensionClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
End Function

'
'----------------------------------------------
'   Init()
'----------------------------------------------
'
Public Sub Init(MainObject As Object)
    '
    Set Main = MainObject
    '
    WorkingQueryString = Main.RefreshQueryString
    LoadedPageID = Main.GetStreamInteger("bid")
    '
    Exit Sub
    '
    ' ----- Error Trap
    '
ErrorTrap:
    Call HandleError("PageNavigationClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub
'
'
'
Public Function GetArgumentList()
    GetArgumentList = "PageNavigationType=[SectionPages|SisterPages|ChildPages|SisterChildPages]&Tiers=2"
End Function
'
'   Aggregate Object Interface
'
Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    '
    Dim GivenPageType As String
    Dim PageNavigationType As Long
    Dim OptionArray() As String
    Dim BracketPosition As Long
    Dim TypeName As String
    Dim TierMax As Long
    '
    
    If Not (Main Is Nothing) Then
        '
        '
        '
        GivenPageType = Main.GetAggrOption("PageNavigationType", OptionString)
        TierMax = kmaEncodeInteger(Main.GetAggrOption("Tiers", OptionString))
        If TierMax = 0 Then
            TierMax = 2
        End If
        Call Main.TestPoint("pageNavigationClass, PageNavigationType=" & GivenPageType)
        Call Main.TestPoint("pageNavigationClass, Tiers=" & TierMax)
        '
        'If InStr(1, OptionString, "=") > 1 Then
        '    TypeName = Mid(OptionString, 1, InStr(1, OptionString, "=") - 1)
        '    GivenPageType = GetAggrOption(TypeName, OptionString)
        'End If
        'OptionArray = Split(OptionString, "=", , vbTextCompare)
        'BracketPosition = InStr(1, OptionArray(1), "[", vbTextCompare)
        'GivenPageType = Left(OptionArray(1), BracketPosition - 1)
        '
        If IsNumeric(GivenPageType) Then
            Select Case GivenPageType
                Case SectionPages
                    PageNavigationType = PageTypeRootChild
                Case SisterPages
                    PageNavigationType = PageTypeSibling
                Case ChildPages
                    PageNavigationType = PageTypeChild
                Case SisterChildPages
                    PageNavigationType = PageTypeSiblingChild
            End Select
        Else
            Select Case LCase(GivenPageType)
                Case "section"
                    PageNavigationType = PageTypeRootChild
                Case "sister"
                    PageNavigationType = PageTypeSibling
                Case "child"
                    PageNavigationType = PageTypeChild
                Case "sister child"
                    PageNavigationType = PageTypeSiblingChild
                Case Else
                    PageNavigationType = PageTypeChild
            End Select
        End If
        '
        GetContent = GetPageNavigation(PageNavigationType, TierMax)
        End If
    '
    Exit Function
ErrorTrap:
    Call HandleError("PageNavigationClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetPageNavigation(GivenPageType As Long, TierMax As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim Caption As String
    Dim Stream As String
    Dim ContentPageStructure As String
    '
    Dim ContentPageStructureArray() As String
    Dim RowPointer As Long
    Dim RowCount As Long
    '
    Dim CSPointer As Long
    '
    Dim ColumnDelimiter As String
    Dim ColumnArray() As String
    Dim ColumnCount As Long
    Dim ColumnPointer As Long
    '
    Dim CriteriaString As String
    '
    Dim LinkCaption As String
    Dim Link As String
    Dim CurrentPageID As Long
    Dim RootPageID As Long
    Dim ParentPageID As Long
    Dim ThisPageType As String
    Dim InnerString As String
    Dim RootToCurrentList As String
    Dim SQLNow As String
    '
    Dim SortCriteria As String
    '
    Dim CurrentRecordID As Long
    '
    Dim BakeName As String
    Dim BakeConent As String
    '
    Dim SecondTest As String
    '
    SQLNow = KmaEncodeSQLDate(Now())
    ContentPageStructure = Main.ContentPageStructure
Call Main.TestPoint("ContentPageStructure: " & ContentPageStructure)
    ContentPageStructureArray = Split(ContentPageStructure, vbCrLf)
    RowCount = UBound(ContentPageStructureArray) + 1
Call Main.TestPoint("RowCount: " & RowCount)
    For RowPointer = 0 To RowCount - 1
        ColumnDelimiter = Left(ContentPageStructureArray(RowPointer), 1)
        ColumnArray = Split(ContentPageStructureArray(RowPointer), ColumnDelimiter)
        ColumnCount = UBound(ColumnArray) + 1
        If ColumnCount > 1 Then
            'Call Main.TestPoint("ColumnCount: " & ColumnCount)
            'For ColumnPointer = 0 To ColumnCount - 1
                'Call Main.TestPoint("ColumnArray(" & ColumnPointer & "): " & ColumnArray(ColumnPointer))
            If ColumnArray(1) = "0" Then
                RootPageID = ColumnArray(2)
            End If
            If ColumnArray(1) = "2" Then
                ParentPageID = ColumnArray(3)
                CurrentPageID = ColumnArray(2)
            End If
            If ColumnArray(1) < "3" Then
                RootToCurrentList = RootToCurrentList & "," & ColumnArray(2)
            End If
            'Next
        End If
    Next
    If RootToCurrentList <> "" Then
        RootToCurrentList = Mid(RootToCurrentList, 2)
    End If
    '
    If RootPageID = 0 Then
        RootPageID = CurrentPageID
        End If
    If ParentPageID = 0 Then
        ParentPageID = Main.GetRecordID(ContentNamePageContent, "Landing Page Content")
        End If
    '
    Call Main.TestPoint("RootPageID=" & RootPageID)
    Call Main.TestPoint("ParentPageID=" & ParentPageID)
    Call Main.TestPoint("CurrentPageID=" & CurrentPageID)
    Call Main.TestPoint("RootToCurrentList=" & RootToCurrentList)
    '
    BakeName = "PageNavigation_Type" & GivenPageType & "_Record" & CurrentPageID
    BakeConent = Main.ReadBake(BakeName)
    '
    If BakeConent = "" Then
        Call Main.TestPoint(">>>>> Creating Bake File - " & BakeName)
        Select Case GivenPageType
            Case PageTypeRootChild
                SortCriteria = GetChildPageListSortMethod(ContentNamePageContent, RootPageID)
                CriteriaString = "(ParentID=" & Main.EncodeSQLNumber(RootPageID) & ")"
            Case PageTypeChild
                SortCriteria = GetChildPageListSortMethod(ContentNamePageContent, CurrentPageID)
                CriteriaString = "(ParentID=" & Main.EncodeSQLNumber(CurrentPageID) & ")"
            Case PageTypeSibling
                SortCriteria = GetChildPageListSortMethod(ContentNamePageContent, ParentPageID)
                CriteriaString = "(ParentID=" & Main.EncodeSQLNumber(ParentPageID) & ")"
            Case PageTypeSiblingChild
                SortCriteria = GetChildPageListSortMethod(ContentNamePageContent, ParentPageID)
                CriteriaString = "(ParentID=" & Main.EncodeSQLNumber(RootPageID) & ")"
        End Select
        '
        SecondTest = "" _
            & "(AllowInMenus<>0)" _
            & "And((PubDate is null)or(PubDate<" & SQLNow & "))" _
            & "And((DateArchive is null)or(DateArchive>" & SQLNow & "))" _
            & "And((dateexpires is null)or(dateexpires>" & SQLNow & "))" _
            & ""
        If CriteriaString <> "" Then
            CSPointer = Main.OpenCSContent(ContentNamePageContent, CriteriaString & " AND (" & SecondTest & " )", SortCriteria, , , "ID, Name,MenuHeadline")
            'CSPointer = Main.OpenCSContent(ContentNamePageContent, CriteriaString & " AND (MenuHeadline<>'' and " & SecondTest & " )", SortCriteria, , , "ID, MenuHeadline")
            'CSPointer = Main.OpenCSContent(ContentNamePageContent, CriteriaString & " AND (MenuHeadline<>'')", SortCriteria, , , "ID, MenuHeadline")
            If Main.CSOK(CSPointer) Then
                Stream = Stream & "<DIV Class=NavigationOuterDiv>"
                Stream = Stream & "<TABLE cellpadding=0 cellspacing=0 Class=""NavigationTable"">"
                Do While Main.CSOK(CSPointer)
                    CurrentRecordID = Main.GetCSInteger(CSPointer, "ID")
                    Caption = Main.GetCSText(CSPointer, "menuheadline")
                    If Caption = "" Then
                       Caption = Main.GetCSText(CSPointer, "name")
                       If Caption = "" Then
                            Caption = "Page " & CurrentRecordID
                       End If
                    End If
                    Stream = Stream & vbCrLf & "<TR>"
                    Stream = Stream & "<TD class=""navigation""><a class=""textnavigation"" href=""" & Main.GetPageLink(CurrentRecordID) & """><span class=""textnavigation"">" & Caption & "</span></a></TD>"
                    'Stream = Stream & "<TD class=""navigation""><a class=""textnavigation"" href=""" & Main.GetContentWatchLinkByID(Main.GetContentID(ContentNamePageContent), CurrentRecordID) & """><span class=""textnavigation"">" & Main.GetCSText(CSPointer, "MenuHeadline") & "</span></a></TD>"
                    Stream = Stream & "</TR>"
                    '
                    If (GivenPageType = PageTypeSiblingChild) And IsInDelimitedString(RootToCurrentList, CStr(CurrentRecordID), ",") Then
                    'If (GivenPageType = PageTypeSiblingChild) And (CurrentRecordID = LoadedPageID) Then
                        Stream = Stream & GetChildPages(CurrentRecordID, RootToCurrentList, 1, TierMax)
                        End If
                    Call Main.NextCSRecord(CSPointer)
                Loop
                '
                Stream = Stream & "</TABLE>"
                Stream = Stream & "</DIV>"
                End If
            Call Main.CloseCS(CSPointer)
            End If
        If Stream <> "" Then
            Call Main.SaveBake(BakeName, Stream, ContentNamePageContent)
            End If
    Else
        Call Main.TestPoint(">>>>> Using Bake File - " & BakeName)
        Stream = BakeConent
        End If
    '
    GetPageNavigation = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("PageNavigationClass", "GetPageNavigation", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
'   Returns the child sort method for the given page
'
Private Function GetChildPageListSortMethod(ContentName As String, RecordID As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPage As Integer
    Dim CSSortMethod As Integer
    Dim ChildListSortMethodID As Integer
    '
    GetChildPageListSortMethod = "Name"
    '
    CSPage = Main.OpenCSContent(ContentName, "ID=" & Main.EncodeSQLNumber(RecordID))
    If Main.CSOK(CSPage) Then
        ChildListSortMethodID = Main.GetCSInteger(CSPage, "ChildListSortMethodID")
        CSSortMethod = Main.OpenCSContent("Sort Methods", "ID=" & Main.EncodeSQLNumber(ChildListSortMethodID))
        If Main.CSOK(CSSortMethod) Then
            GetChildPageListSortMethod = Main.GetCSText(CSSortMethod, "OrderByClause")
            End If
        Call Main.CloseCS(CSSortMethod)
        End If
    Call Main.CloseCS(CSPage)
    '
    Exit Function
ErrorTrap:
    Call HandleError("PageNavigationClass", "GetChildPageListSortMethod", Err.Number, Err.Source, Err.Description, True, False)
    End Function
'
Private Function GetChildPages(ParentPageID As Long, RootToCurrentList As String, Tier As Long, MaxTier As Long) As String
    On Error GoTo ErrorTrap
    '
    Dim Stream As String
    Dim CSPointer As String
    Dim CurrentRecordID As Long
    Dim MenuHeadline As String
    Dim SubStyleName As String
    '
    If Tier <= MaxTier Then
        CSPointer = Main.OpenCSContent(ContentNamePageContent, "ParentID=" & Main.EncodeSQLNumber(ParentPageID), GetChildPageListSortMethod(ContentNamePageContent, ParentPageID))
        Do While Main.CSOK(CSPointer)
            CurrentRecordID = Main.GetCSInteger(CSPointer, "ID")
            MenuHeadline = Main.GetCSText(CSPointer, "MenuHeadline")
            SubStyleName = "Tier" & CStr(Tier)
            If MenuHeadline <> "" Then
                '
                ' display this page navigation
                '
                Stream = Stream & vbCrLf & "  <TR>"
                Stream = Stream & "<TD class=""subnavigation " & SubStyleName & """><a class=""subtextnavigation " & SubStyleName & """ href=""" & Main.GetPageLink(CurrentRecordID) & """><span class=""subtextnavigation " & SubStyleName & """ >" & MenuHeadline & "</span></a></TD>"
                'Stream = Stream & "<TD class=""subnavigation""><a class=""subtextnavigation"" href=""" & Main.GetContentWatchLinkByID(Main.GetContentID(ContentNamePageContent), CurrentRecordID) & """><span class=""subtextnavigation"">" & MenuHeadline & "</span></a></TD>"
                Stream = Stream & "</TR>"
                '
                ' if this page is in the Root-To-Current list, get its child pages also
                '
                If IsInDelimitedString(RootToCurrentList, CStr(CurrentRecordID), ",") Then
                    Stream = Stream & GetChildPages(CurrentRecordID, RootToCurrentList, Tier + 1, MaxTier)
                End If
            End If
            '
            Call Main.NextCSRecord(CSPointer)
        Loop
        Call Main.CloseCS(CSPointer)
    End If
    '
    GetChildPages = Stream
    '
    Exit Function
ErrorTrap:
    Call HandleError("PageNavigationClass", "GetChildPages", Err.Number, Err.Source, Err.Description, True, False)
    End Function

