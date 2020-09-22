VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl TreeList 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5685
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   HitBehavior     =   2  'Use Paint
   KeyPreview      =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   5685
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeList.ctx":0000
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeList.ctx":059A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TreeList.ctx":0B34
            Key             =   "note"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5520
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   1092
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   100
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   4800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   8467
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   4800
      Left            =   2280
      TabIndex        =   1
      Top             =   0
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   8467
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   2040
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   100
   End
End
Attribute VB_Name = "TreeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const sglSplitLimit = 500
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private mvarTreeViewWidth As Long
Private mvarOuterBorderStyle As MSComctlLib.BorderStyleConstants
Private mvarOuterAppearance As MSComctlLib.AppearanceConstants
Private mvarListViewViewType As MSComctlLib.ListViewConstants
Private mvarInnerBorderStyle As MSComctlLib.BorderStyleConstants
Private mvarInnerAppearance As MSComctlLib.AppearanceConstants
Private mvarTreeViewStyle As MSComctlLib.TreeStyleConstants
Private mvarTreeViewCheckBoxes As Boolean
Private mvarListViewCausesValidation As Boolean
Private mvarTreeViewFullRowSelect As Boolean
Private mvarTreeViewHotTracking As Boolean
Private mvarTreeViewIndentation As Double
Private mvarListViewAutoResize As Boolean
Private mvarTreeViewWhatsThisHelpID As Long
Private mvarTreeViewToolTipText As String
Private mvarTreeViewScroll As Boolean
Private mvarTreeViewSorted As Boolean
Private mvarTreeViewSingleSel As Boolean
Private mvarTreeViewLineStyle As MSComctlLib.TreeLineStyleConstants
Private mbMoving As Boolean
Private mvarHeadings As String
Private mvarListViewCheckBoxes As Boolean
Private mvarListViewFullRowSelect As Boolean
Private mvarTreeViewPathSeparator As String
Private mvarListViewMultiSelect As Boolean
Private mvarListViewSortOrder As MSComctlLib.ListSortOrderConstants
Private mvarTreeViewLabelEdit As MSComctlLib.LabelEditConstants
Private mvarListViewGridLines As Boolean
Private mvarListViewHideColumnHeaders As Boolean
Private mvarListViewHideSelection As Boolean
Private mvarListViewHotTracking As Boolean
Private mvarListViewHoverSelection As Boolean
Private mvarListViewLabelEdit As MSComctlLib.LabelEditConstants
Private mvarListViewLabelWrap As Boolean
Private mvarListViewAllowColumnReorder As Boolean
Private mvarListViewEnabled As Boolean
Private mvarListViewHelpContextID As Long
Private mvarListViewSorted As Boolean
Private mvarListViewSortKey As Integer
Private mvarListViewTag As String
Private mvarListViewToolTipText As String
Private mvarListViewWhatsThisHelpID As Long
Private mvarTreeViewHideSelection As Boolean
Private mvarListViewFlatScrollBar As Boolean
Private mvarListViewRightAlign As String
Private mvarHeadingsIsDate As String
Private mvarHeadingsToSum As String
Private mvarTreeViewEnabled As Boolean
Private mvarTreeViewHelpContextID As Long
Private ListItems() As MyListItem
Private iListItems As Long
Private Type SubItem
    Text As String
    Bold As Boolean
    ForeColor As ColorConstants
    ReportIcon As String
End Type
Private Type MyListItem
    Relative As String
    Bold As Boolean
    Checked As Boolean
    ForeColor As ColorConstants
    Ghosted As Boolean
    Icon As Variant
    Index As Long
    Key As String
    Selected As Boolean
    SmallIcon As Variant
    Tag As String
    Text As String
    ToolTipText As String
    EnsureVisible As Boolean
    SubItems() As SubItem
    BoldAll As Boolean
    ColorAll As Boolean
End Type
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Event ListViewClick()
Event TreeViewLostFocus()
Event TreeViewDragOver(Source As Control, x As Single, y As Single, State As Integer)
Event TreeViewDragDrop(Source As Control, x As Single, y As Single)
Event TreeViewNodeClick(ByVal Node As MSComctlLib.Node)
Event TreeViewMouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event TreeViewClick()
Event TreeViewCollapse(ByVal Node As MSComctlLib.Node)
Event TreeViewDblClick()
Event TreeViewExpand(ByVal Node As MSComctlLib.Node)
Event TreeViewGotFocus()
Event TreeViewKeyDown(KeyCode As Integer, Shift As Integer)
Event TreeViewKeyPress(KeyAscii As Integer)
Event TreeViewKeyUp(KeyCode As Integer, Shift As Integer)
Event TreeViewMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event TreeViewMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event TreeViewNodeCheck(ByVal Node As MSComctlLib.Node)
Event TreeViewAfterLabelEdit(Cancel As Integer, NewString As String)
Event TreeViewBeforeLabelEdit(Cancel As Integer)
Event ListViewAfterLabelEdit(Cancel As Integer, NewString As String)
Event ListViewBeforeLabelEdit(Cancel As Integer)
Event ListViewColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Event ListViewDblClick()
Event ListViewGotFocus()
Event ListViewItemCheck(ByVal Item As MSComctlLib.ListItem)
Event ListViewItemClick(ByVal Item As MSComctlLib.ListItem)
Event ListViewKeyDown(KeyCode As Integer, Shift As Integer)
Event ListViewKeyPress(KeyAscii As Integer)
Event TreeViewValidate(Cancel As Boolean)
Event ListViewKeyUp(KeyCode As Integer, Shift As Integer)
Event ListViewLostFocus()
Event ListViewMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event ListViewMouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event ListViewMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event TreeViewOLECompleteDrag(Effect As Long)
Event TreeViewOLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Event TreeViewOLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Event TreeViewOLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event TreeViewOLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
Event TreeViewOLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
Event ListViewDragDrop(Source As Control, x As Single, y As Single)
Event ListViewDragOver(Source As Control, x As Single, y As Single, State As Integer)
Event ListViewOLECompleteDrag(Effect As Long)
Event ListViewOLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Event ListViewOLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Event ListViewOLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event ListViewOLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
Event ListViewOLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
Event ListViewValidate(Cancel As Boolean)
Public Enum IconvEnum
    Money = 0
    Date = 1
    Proper = 2
End Enum
Public Function LstViewUpdate(Arrfields() As String, Optional ByVal lstIndex As String = "", Optional ByVal sIcon As String = "note", Optional ByVal sSmallIcon = "note") As Long
    On Error Resume Next
    Dim ItmX As ListItem
    Dim fldCnt As Integer
    Dim wCnt As Integer
    Select Case Val(lstIndex)
    Case 0
        Set ItmX = lvListView.ListItems.Add()
    Case Else
        Set ItmX = lvListView.ListItems(Val(lstIndex))
    End Select
    wCnt = UBound(Arrfields) - 1
    With ItmX
        .Text = Arrfields(1)
        For fldCnt = 1 To wCnt
            .SubItems(fldCnt) = Arrfields(fldCnt + 1)
            Err.Clear
        Next
    End With
    If Len(sIcon) > 0 Then
        ItmX.Icon = sIcon
    End If
    If Len(sSmallIcon) > 0 Then
        ItmX.SmallIcon = sSmallIcon
    End If
    LstViewUpdate = ItmX.Index
    Set ItmX = Nothing
    Err.Clear
End Function
Private Function MvFromCollection(objCollection As Collection, ByVal Delimiter As String) As String
    On Error Resume Next
    Dim xTot As Long
    Dim xCnt As Long
    Dim sRet As String
    sRet = ""
    xTot = objCollection.Count
    For xCnt = 1 To xTot
        If xCnt = xTot Then
            sRet = sRet & objCollection.Item(xCnt)
        Else
            sRet = sRet & objCollection.Item(xCnt) & Delimiter
        End If
        Err.Clear
    Next
    MvFromCollection = sRet
    Err.Clear
End Function
Private Function RemDelim(ByVal Dataobj As String, ByVal Delimiter As String) As String
    On Error Resume Next
    Dim intDataSize As Long
    Dim intDelimSize As Long
    Dim strLast As String
    intDataSize = Len(Dataobj)
    intDelimSize = Len(Delimiter)
    strLast = Right$(Dataobj, intDelimSize)
    Select Case strLast
    Case Delimiter
        RemDelim = Left$(Dataobj, (intDataSize - intDelimSize))
    Case Else
        RemDelim = Dataobj
    End Select
    Err.Clear
End Function
Private Function ProperAmount(ByVal strValue As String, Optional Reverse As Boolean = False) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim rsVal As String
    rsStr = ""
    strValue = Trim$(strValue)
    If Len(strValue) = 0 Then
        strValue = "0.00"
    End If
    strValue = Replace$(strValue, ",", "")
    strValue = Replace$(strValue, "*", "")
    If InStr(1, strValue, "E-") > 0 Then
        strValue = Format$(strValue, "#######################0.00")
        If Reverse = True Then
            strValue = Val(strValue) * (0 - 1)
            strValue = Format$(strValue, "#######################0.00")
        Else
            ProperAmount = strValue
        End If
    Else
        rsTot = Len(strValue)
        For rsCnt = 1 To rsTot
            rsVal = Mid$(strValue, rsCnt, 1)
            If InStr(1, "-.0123456789", rsVal) > 0 Then
                rsStr = rsStr & rsVal
            End If
            Err.Clear
        Next
        rsStr = Trim$(rsStr)
        If Len(rsStr) = 0 Then
            rsStr = "0.00"
        End If
        If InStr(1, rsStr, ".") = 0 Then
            rsStr = rsStr & ".00"
        End If
        If Reverse = True Then
            strValue = Val(strValue) * (0 - 1)
        End If
        'strValue = CDbl(rsStr)
        ProperAmount = Format$(strValue, "#######################0.00")
    End If
    Err.Clear
End Function
Private Function MakeMoney(ByVal strValue As String, Optional Reverse As Boolean = False) As String
    On Error Resume Next
    strValue = ProperAmount(strValue)
    If Reverse = True Then
        strValue = Val(strValue) * (0 - 1)
        strValue = ProperAmount(strValue)
    End If
    MakeMoney = Format$(strValue, "#,##0.00")
    Err.Clear
End Function
Public Sub Clear()
    On Error Resume Next
    tvTreeView.Nodes.Clear
    lvListView.ListItems.Clear
    lvListView.ColumnHeaders.Clear
    lvListView.View = lvwIcon
    HeadingsIsDate = ""
    Headings = ""
    HeadingsToRightAlign = ""
    HeadingsToSum = ""
    iListItems = 0
    Err.Clear
End Sub
Private Function MvCount(ByVal StringMv As String, ByVal Delimiter As String) As Long
    On Error Resume Next
    Dim xNew() As String
    xNew = Split(StringMv, Delimiter)
    MvCount = UBound(xNew) + 1
    Err.Clear
End Function
Public Function ListViewAddItem(ByVal Relative As String, ByVal Key As String, ByVal Text As String, Optional ByVal Tag As String = "", Optional ByVal ToolTipText As String = "", Optional ByVal Icon As String = "note", Optional ByVal SmallIcon As String = "note", Optional Bold As Boolean = False, Optional Checked As Boolean = False, Optional ForeColor As ColorConstants = vbBlack, Optional Ghosted As Boolean = False, Optional BoldAll As Boolean = False, Optional ColorAll As Boolean = False) As Long
    On Error Resume Next
    Dim totHeadings As Long
    totHeadings = MvCount(Headings, ",") - 1
    iListItems = iListItems + 1
    ReDim Preserve ListItems(iListItems)
    With ListItems(iListItems)
        .Relative = Relative
        .Key = Key
        .Text = Text
        .Icon = Icon
        .SmallIcon = SmallIcon
        .Bold = Bold
        .Checked = Checked
        .ForeColor = ForeColor
        .Ghosted = Ghosted
        .Selected = False
        .Tag = Tag
        .ToolTipText = ToolTipText
        ReDim .SubItems(totHeadings)
        .BoldAll = BoldAll
        .ColorAll = ColorAll
    End With
    ListViewAddItem = iListItems
    Err.Clear
End Function
Public Sub ListViewListSubItems(ByVal Relative As Long, ByVal Index As Variant, ByVal Text As String, Optional cForeColor As ColorConstants = vbBlack, Optional boolBold As Boolean = False, Optional ByVal ReportIcon As String = "")
    On Error Resume Next
    Dim colPos As Integer
    If (VarType(Relative) <> vbLong) Then
        Err.Raise 1, "ListViewListSubItems", "ListViewListSubItems: Relative not of required type Long."
        Err.Clear
        Exit Sub
    End If
    If (VarType(Index) <> vbLong) And (VarType(Index) <> vbString) Then
        Err.Raise 1, "ListViewListSubItems", "ListViewListSubItems: Index not of required type (String or Long)."
        Err.Clear
        Exit Sub
    End If
    If (VarType(Index) = vbLong) Then
        colPos = Index
    Else
        colPos = ListViewColumnPosition(Index)
    End If
    ListItems(Relative).SubItems(colPos - 1).Text = Text
    ListItems(Relative).SubItems(colPos - 1).Bold = boolBold
    ListItems(Relative).SubItems(colPos - 1).ForeColor = cForeColor
    ListItems(Relative).SubItems(colPos - 1).ReportIcon = ReportIcon
    Err.Clear
End Sub
Private Sub lvListView_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error Resume Next
    RaiseEvent ListViewAfterLabelEdit(Cancel, NewString)
    Err.Clear
End Sub
Private Sub lvListView_BeforeLabelEdit(Cancel As Integer)
    On Error Resume Next
    RaiseEvent ListViewBeforeLabelEdit(Cancel)
    Err.Clear
End Sub
Private Sub lvListView_Click()
    On Error Resume Next
    RaiseEvent ListViewClick
    Err.Clear
End Sub
Private Sub lvListView_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    Select Case ListViewSortOrder
    Case lvwAscending
        ListViewSortOrder = lvwDescending
    Case Else
        ListViewSortOrder = lvwAscending
    End Select
    ListViewSortKey = ColumnHeader.Index - 1
    ListViewSorted = True
    lvListView.Refresh
    RaiseEvent ListViewColumnClick(ColumnHeader)
    Err.Clear
End Sub
Private Sub lvListView_DblClick()
    On Error Resume Next
    RaiseEvent ListViewDblClick
    Err.Clear
End Sub
Private Sub lvListView_DragDrop(Source As Control, x As Single, y As Single)
    On Error Resume Next
    RaiseEvent ListViewDragDrop(Source, x, y)
    Err.Clear
End Sub
Private Sub lvListView_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    On Error Resume Next
    RaiseEvent ListViewDragOver(Source, x, y, State)
    Err.Clear
End Sub
Private Sub lvListView_GotFocus()
    On Error Resume Next
    RaiseEvent ListViewGotFocus
    Err.Clear
End Sub
Private Sub lvListView_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    RaiseEvent ListViewItemCheck(Item)
    Err.Clear
End Sub
Private Sub lvListView_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    RaiseEvent ListViewItemClick(Item)
    Err.Clear
End Sub
Private Sub lvListView_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    RaiseEvent ListViewKeyDown(KeyCode, Shift)
    Err.Clear
End Sub
Private Sub lvListView_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    RaiseEvent ListViewKeyPress(KeyAscii)
    Err.Clear
End Sub
Private Sub lvListView_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    RaiseEvent ListViewKeyUp(KeyCode, Shift)
    Err.Clear
End Sub
Private Sub lvListView_LostFocus()
    On Error Resume Next
    RaiseEvent ListViewLostFocus
    Err.Clear
End Sub
Private Sub lvListView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    RaiseEvent ListViewMouseDown(Button, Shift, x, y)
    Err.Clear
End Sub
Private Sub lvListView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    RaiseEvent ListViewMouseMove(Button, Shift, x, y)
    Err.Clear
End Sub
Private Sub lvListView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    RaiseEvent ListViewMouseUp(Button, Shift, x, y)
    Err.Clear
End Sub
Private Sub lvListView_OLECompleteDrag(Effect As Long)
    On Error Resume Next
    RaiseEvent ListViewOLECompleteDrag(Effect)
    Err.Clear
End Sub
Private Sub lvListView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    RaiseEvent ListViewOLEDragDrop(Data, Effect, Button, Shift, x, y)
    Err.Clear
End Sub
Private Sub lvListView_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    On Error Resume Next
    RaiseEvent ListViewOLEDragOver(Data, Effect, Button, Shift, x, y, State)
    Err.Clear
End Sub
Private Sub lvListView_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    On Error Resume Next
    RaiseEvent ListViewOLEGiveFeedback(Effect, DefaultCursors)
    Err.Clear
End Sub
Private Sub lvListView_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
    On Error Resume Next
    RaiseEvent ListViewOLESetData(Data, DataFormat)
    Err.Clear
End Sub
Private Sub lvListView_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    On Error Resume Next
    RaiseEvent ListViewOLEStartDrag(Data, AllowedEffects)
    Err.Clear
End Sub
Private Sub lvListView_Validate(Cancel As Boolean)
    On Error Resume Next
    RaiseEvent ListViewValidate(Cancel)
    Err.Clear
End Sub
Private Sub tvTreeView_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error Resume Next
    RaiseEvent TreeViewAfterLabelEdit(Cancel, NewString)
    Err.Clear
End Sub
Private Sub tvTreeView_BeforeLabelEdit(Cancel As Integer)
    On Error Resume Next
    RaiseEvent TreeViewBeforeLabelEdit(Cancel)
    Err.Clear
End Sub
Public Function TreeViewGetVisibleCount() As Long
    On Error Resume Next
    TreeViewGetVisibleCount = tvTreeView.GetVisibleCount
    Err.Clear
End Function
Public Sub TreeViewSetFocus()
    On Error Resume Next
    tvTreeView.SetFocus
    Err.Clear
End Sub
Public Sub TreeViewStartLabelEdit()
    On Error Resume Next
    tvTreeView.StartLabelEdit
    Err.Clear
End Sub
Private Sub tvTreeView_Click()
    On Error Resume Next
    RaiseEvent TreeViewClick
    Err.Clear
End Sub
Private Sub tvTreeView_Collapse(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    RaiseEvent TreeViewCollapse(Node)
    Err.Clear
End Sub
Private Sub tvTreeView_DblClick()
    On Error Resume Next
    RaiseEvent TreeViewDblClick
    Err.Clear
End Sub
Private Sub tvTreeView_DragOver(Source As Control, x As Single, y As Single, State As Integer)
    On Error Resume Next
    RaiseEvent TreeViewDragOver(Source, x, y, State)
    Err.Clear
End Sub
Private Sub tvTreeView_Expand(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    RaiseEvent TreeViewExpand(Node)
    Err.Clear
End Sub
Private Sub tvTreeView_GotFocus()
    On Error Resume Next
    RaiseEvent TreeViewGotFocus
    Err.Clear
End Sub
Private Sub tvTreeView_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    RaiseEvent TreeViewKeyDown(KeyCode, Shift)
    Err.Clear
End Sub
Private Sub tvTreeView_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    RaiseEvent TreeViewKeyPress(KeyAscii)
    Err.Clear
End Sub
Private Sub tvTreeView_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    RaiseEvent TreeViewKeyUp(KeyCode, Shift)
    Err.Clear
End Sub
Private Sub tvTreeView_LostFocus()
    On Error Resume Next
    RaiseEvent TreeViewLostFocus
    Err.Clear
End Sub
Private Sub tvTreeView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    RaiseEvent TreeViewMouseDown(Button, Shift, x, y)
    Err.Clear
End Sub
Private Sub tvTreeView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    RaiseEvent TreeViewMouseMove(Button, Shift, x, y)
    Err.Clear
End Sub
Private Sub tvTreeView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    RaiseEvent TreeViewMouseUp(Button, Shift, x, y)
    Err.Clear
End Sub
Private Sub tvTreeView_NodeCheck(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    RaiseEvent TreeViewNodeCheck(Node)
    Err.Clear
End Sub
Private Function ListViewListItemPosition(ByVal Key As String) As Long
    On Error Resume Next
    Dim rsCnt As Long
    Dim sKey As String
    For rsCnt = 1 To iListItems
        With ListItems(rsCnt)
            sKey = .Key
            If LCase$(sKey) = LCase$(Key) Then
                ListViewListItemPosition = rsCnt
                Exit For
            End If
        End With
        Err.Clear
    Next
    Err.Clear
End Function
Private Sub tvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    Dim newListItem As MSComctlLib.ListItem
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim sKey As String
    Dim sRelative As String
    Dim rsCntt As Integer
    Dim rsTott As Integer
    lvListView.ListItems.Clear
    ListViewView = lvwReport
    For rsCnt = 1 To iListItems
        With ListItems(rsCnt)
            sKey = .Key
            sRelative = .Relative
            If LCase$(sRelative) = LCase$(Node.Key) Then
                If Len(sKey) = 0 Then
                    Set newListItem = lvListView.ListItems.Add(, sKey, .Text)
                Else
                    Set newListItem = lvListView.ListItems.Add(, , .Text)
                End If
                newListItem.Checked = .Checked
                newListItem.Bold = .Bold
                If .EnsureVisible = True Then newListItem.EnsureVisible
                newListItem.ForeColor = .ForeColor
                newListItem.Ghosted = .Ghosted
                newListItem.Selected = .Selected
                newListItem.ToolTipText = .ToolTipText
                If Len(.Icon) > 0 Then newListItem.Icon = .Icon
                If Len(.SmallIcon) > 0 Then newListItem.SmallIcon = .SmallIcon
                newListItem.Tag = .Tag
                rsTott = UBound(.SubItems)
                For rsCntt = 1 To rsTott
                    newListItem.SubItems(rsCntt) = .SubItems(rsCntt).Text
                    newListItem.ListSubItems(rsCntt).ForeColor = .SubItems(rsCntt).ForeColor
                    newListItem.ListSubItems(rsCntt).Bold = .SubItems(rsCntt).Bold
                    If Len(.SubItems(rsCntt).ReportIcon) > 0 Then newListItem.ListSubItems(rsCntt).ReportIcon = .SubItems(rsCntt).ReportIcon
                    If .BoldAll = True Then newListItem.ListSubItems(rsCntt).Bold = .Bold
                    If .ColorAll = True Then newListItem.ListSubItems(rsCntt).ForeColor = .ForeColor
                    Err.Clear
                Next
            End If
        End With
        Err.Clear
    Next
    If Len(HeadingsToSum) > 0 Then
        If lvListView.ListItems.Count > 0 Then LstViewSumColumns True, HeadingsToSum
    End If
    If ListViewAutoResize = True Then LstViewAutoResize
    RaiseEvent TreeViewNodeClick(Node)
    Err.Clear
End Sub
Private Sub tvTreeView_OLECompleteDrag(Effect As Long)
    On Error Resume Next
    RaiseEvent TreeViewOLECompleteDrag(Effect)
    Err.Clear
End Sub
Private Sub tvTreeView_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    RaiseEvent TreeViewOLEDragDrop(Data, Effect, Button, Shift, x, y)
    Err.Clear
End Sub
Private Sub tvTreeView_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    On Error Resume Next
    RaiseEvent TreeViewOLEDragOver(Data, Effect, Button, Shift, x, y, State)
    Err.Clear
End Sub
Private Sub tvTreeView_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    On Error Resume Next
    RaiseEvent TreeViewOLEGiveFeedback(Effect, DefaultCursors)
    Err.Clear
End Sub
Private Sub tvTreeView_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
    On Error Resume Next
    RaiseEvent TreeViewOLESetData(Data, DataFormat)
    Err.Clear
End Sub
Private Sub tvTreeView_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    On Error Resume Next
    RaiseEvent TreeViewOLEStartDrag(Data, AllowedEffects)
    Err.Clear
End Sub
Private Sub tvTreeView_Validate(Cancel As Boolean)
    On Error Resume Next
    RaiseEvent TreeViewValidate(Cancel)
    Err.Clear
End Sub
Private Sub UserControl_Initialize()
    On Error Resume Next
    TreeViewWidth = UserControl.Width / 3
    Set tvTreeView.ImageList = ImageList1
    Err.Clear
End Sub
Public Property Get TreeViewWidth() As Long
    On Error Resume Next
    TreeViewWidth = mvarTreeViewWidth
    Err.Clear
End Property
Public Property Let TreeViewWidth(vNewValue As Long)
    On Error Resume Next
    mvarTreeViewWidth = vNewValue
    tvTreeView.Width = mvarTreeViewWidth
    SizeControls tvTreeView.Width
    PropertyChanged "TreeViewWidth"
    Err.Clear
End Property
Public Property Get TreeViewSorted() As Boolean
    On Error Resume Next
    TreeViewSorted = mvarTreeViewSorted
    Err.Clear
End Property
Public Property Let TreeViewSorted(vNewValue As Boolean)
    On Error Resume Next
    mvarTreeViewSorted = vNewValue
    tvTreeView.Sorted = vNewValue
    PropertyChanged "TreeViewSorted"
    Err.Clear
End Property
Public Property Get TreeViewCheckBoxes() As Boolean
    On Error Resume Next
    TreeViewCheckBoxes = mvarTreeViewCheckBoxes
    Err.Clear
End Property
Public Property Let TreeViewCheckBoxes(vNewValue As Boolean)
    On Error Resume Next
    mvarTreeViewCheckBoxes = vNewValue
    tvTreeView.Checkboxes = vNewValue
    PropertyChanged "TreeViewCheckBoxes"
    Err.Clear
End Property
Public Property Get TreeViewSingleSel() As Boolean
    On Error Resume Next
    TreeViewSingleSel = mvarTreeViewSingleSel
    Err.Clear
End Property
Public Property Let TreeViewSingleSel(vNewValue As Boolean)
    On Error Resume Next
    mvarTreeViewSingleSel = vNewValue
    tvTreeView.SingleSel = vNewValue
    PropertyChanged "TreeViewSingleSel"
    Err.Clear
End Property
Public Property Get TreeViewToolTipText() As String
    On Error Resume Next
    TreeViewToolTipText = mvarTreeViewToolTipText
    Err.Clear
End Property
Public Property Let TreeViewToolTipText(vNewValue As String)
    On Error Resume Next
    mvarTreeViewToolTipText = vNewValue
    tvTreeView.ToolTipText = vNewValue
    PropertyChanged "TreeViewToolTipText"
    Err.Clear
End Property
Public Property Get TreeViewWhatsThisHelpID() As Long
    On Error Resume Next
    TreeViewWhatsThisHelpID = mvarTreeViewWhatsThisHelpID
    Err.Clear
End Property
Public Property Let TreeViewWhatsThisHelpID(vNewValue As Long)
    On Error Resume Next
    mvarTreeViewWhatsThisHelpID = vNewValue
    tvTreeView.WhatsThisHelpID = vNewValue
    PropertyChanged "TreeViewWhatsThisHelpID"
    Err.Clear
End Property
Public Property Get TreeViewIndentation() As Double
    On Error Resume Next
    TreeViewIndentation = mvarTreeViewIndentation
    Err.Clear
End Property
Public Property Let TreeViewIndentation(vNewValue As Double)
    On Error Resume Next
    mvarTreeViewIndentation = vNewValue
    tvTreeView.Indentation = vNewValue
    PropertyChanged "TreeViewIndentation"
    Err.Clear
End Property
Public Property Get TreeViewScroll() As Boolean
    On Error Resume Next
    TreeViewScroll = mvarTreeViewScroll
    Err.Clear
End Property
Public Property Let TreeViewScroll(vNewValue As Boolean)
    On Error Resume Next
    mvarTreeViewScroll = vNewValue
    tvTreeView.Scroll = vNewValue
    PropertyChanged "TreeViewScroll"
    Err.Clear
End Property
Public Property Get TreeViewPathSeparator() As String
    On Error Resume Next
    TreeViewPathSeparator = mvarTreeViewPathSeparator
    Err.Clear
End Property
Public Property Let TreeViewPathSeparator(vNewValue As String)
    On Error Resume Next
    mvarTreeViewPathSeparator = vNewValue
    tvTreeView.PathSeparator = vNewValue
    PropertyChanged "TreeViewPathSeparator"
    Err.Clear
End Property
Public Property Get TreeViewFullRowSelect() As Boolean
    On Error Resume Next
    TreeViewFullRowSelect = mvarTreeViewFullRowSelect
    Err.Clear
End Property
Public Property Let TreeViewFullRowSelect(vNewValue As Boolean)
    On Error Resume Next
    mvarTreeViewFullRowSelect = vNewValue
    tvTreeView.FullRowSelect = vNewValue
    PropertyChanged "TreeViewFullRowSelect"
    Err.Clear
End Property
Public Property Get ListViewView() As MSComctlLib.ListViewConstants
    On Error Resume Next
    ListViewView = mvarListViewViewType
    Err.Clear
End Property
Public Property Let ListViewView(vNewValue As MSComctlLib.ListViewConstants)
    On Error Resume Next
    mvarListViewViewType = vNewValue
    lvListView.View = vNewValue
    PropertyChanged "ListViewView"
    Err.Clear
End Property
Public Property Get InnerBorderStyle() As MSComctlLib.BorderStyleConstants
    On Error Resume Next
    InnerBorderStyle = mvarInnerBorderStyle
    Err.Clear
End Property
Public Property Let InnerBorderStyle(vNewValue As MSComctlLib.BorderStyleConstants)
    On Error Resume Next
    mvarInnerBorderStyle = vNewValue
    tvTreeView.BorderStyle = vNewValue
    lvListView.BorderStyle = vNewValue
    PropertyChanged "InnerBorderStyle"
    Err.Clear
End Property
Public Property Get OuterBorderStyle() As MSComctlLib.BorderStyleConstants
    On Error Resume Next
    OuterBorderStyle = mvarOuterBorderStyle
    Err.Clear
End Property
Public Property Let OuterBorderStyle(vNewValue As MSComctlLib.BorderStyleConstants)
    On Error Resume Next
    mvarOuterBorderStyle = vNewValue
    UserControl.BorderStyle = vNewValue
    PropertyChanged "OuterBorderStyle"
    Err.Clear
End Property
Public Property Get InnerAppearance() As MSComctlLib.AppearanceConstants
    On Error Resume Next
    InnerAppearance = mvarInnerAppearance
    Err.Clear
End Property
Public Property Get OuterAppearance() As MSComctlLib.AppearanceConstants
    On Error Resume Next
    OuterAppearance = mvarOuterAppearance
    Err.Clear
End Property
Public Property Let OuterAppearance(vNewValue As MSComctlLib.AppearanceConstants)
    On Error Resume Next
    mvarOuterAppearance = vNewValue
    UserControl.Appearance = vNewValue
    PropertyChanged "OuterAppearance"
    Err.Clear
End Property
Public Property Let InnerAppearance(vNewValue As MSComctlLib.AppearanceConstants)
    On Error Resume Next
    mvarInnerAppearance = vNewValue
    tvTreeView.Appearance = vNewValue
    lvListView.Appearance = vNewValue
    PropertyChanged "InnerAppearance"
    Err.Clear
End Property
Private Sub UserControl_InitProperties()
    On Error Resume Next
    ListViewCausesValidation = False
    ListViewFullRowSelect = True
    InnerAppearance = cc3D
    InnerBorderStyle = ccNone
    OuterAppearance = cc3D
    OuterBorderStyle = ccNone
    TreeViewWidth = UserControl.Width / 3
    TreeViewCheckBoxes = False
    TreeViewFullRowSelect = False
    TreeViewSorted = False
    TreeViewWhatsThisHelpID = 0
    TreeViewSingleSel = False
    ListViewAutoResize = True
    TreeViewStyle = tvwTreelinesPlusMinusPictureText
    Headings = ""
    HeadingsIsDate = ""
    HeadingsToRightAlign = ""
    HeadingsToSum = ""
    ListViewView = lvwIcon
    TreeViewToolTipText = ""
    TreeViewLineStyle = tvwTreeLines
    ListViewCheckBoxes = False
    ListViewMultiSelect = False
    ListViewSortOrder = lvwAscending
    ListViewGridLines = True
    ListViewHideColumnHeaders = False
    ListViewHideSelection = True
    ListViewHotTracking = False
    ListViewHoverSelection = False
    ListViewLabelEdit = lvwManual
    ListViewLabelWrap = True
    ListViewAllowColumnReorder = False
    ListViewEnabled = True
    ListViewHelpContextID = 0
    ListViewSorted = False
    ListViewSortKey = 0
    ListViewTag = ""
    ListViewToolTipText = ""
    ListViewWhatsThisHelpID = 0
    ListViewFlatScrollBar = False
    TreeViewEnabled = True
    TreeViewHelpContextID = 0
    TreeViewHideSelection = 0
    TreeViewIndentation = 400.92
    TreeViewHotTracking = False
    TreeViewLabelEdit = lvwManual
    TreeViewPathSeparator = "\"
    TreeViewScroll = True
    Err.Clear
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        ListViewCausesValidation = .ReadProperty("ListViewCausesValidation", False)
        ListViewFullRowSelect = .ReadProperty("ListViewFullRowSelect", True)
        InnerAppearance = .ReadProperty("InnerAppearance", cc3D)
        InnerBorderStyle = .ReadProperty("InnerBorderStyle", ccNone)
        OuterAppearance = .ReadProperty("OuterAppearance", cc3D)
        OuterBorderStyle = .ReadProperty("OuterBorderStyle", ccNone)
        TreeViewWidth = .ReadProperty("TreeViewWidth", UserControl.Width / 3)
        TreeViewCheckBoxes = .ReadProperty("TreeViewCheckBoxes", False)
        TreeViewFullRowSelect = .ReadProperty("TreeViewFullRowSelect", False)
        TreeViewSorted = .ReadProperty("TreeViewSorted", False)
        TreeViewWhatsThisHelpID = .ReadProperty("TreeViewWhatsThisHelpID", 0)
        TreeViewSingleSel = .ReadProperty("TreeViewSingleSel", False)
        ListViewAutoResize = .ReadProperty("ListViewAutoResize", True)
        TreeViewStyle = .ReadProperty("TreeViewStyle", tvwTreelinesPlusMinusPictureText)
        Headings = .ReadProperty("Headings", "")
        ListViewView = .ReadProperty("View", lvwIcon)
        TreeViewToolTipText = .ReadProperty("TreeViewToolTipText", "")
        TreeViewLineStyle = .ReadProperty("TreeViewLineStyle", tvwTreeLines)
        ListViewCheckBoxes = .ReadProperty("ListViewCheckBoxes", False)
        ListViewMultiSelect = .ReadProperty("ListViewMultiSelect", False)
        ListViewSortOrder = .ReadProperty("ListViewSortOrder", lvwAscending)
        ListViewGridLines = .ReadProperty("ListViewGridLines", False)
        ListViewHideColumnHeaders = .ReadProperty("ListViewHideColumnHeaders", False)
        ListViewHideSelection = .ReadProperty("ListViewHideSelection", True)
        ListViewHotTracking = .ReadProperty("ListViewHotTracking", False)
        ListViewHoverSelection = .ReadProperty("ListViewHoverSelection", False)
        ListViewLabelEdit = .ReadProperty("ListViewLabelEdit", lvwManual)
        ListViewLabelWrap = .ReadProperty("ListViewLabelWrap", True)
        ListViewAllowColumnReorder = .ReadProperty("ListViewAllowColumnReorder", False)
        ListViewEnabled = .ReadProperty("ListViewEnabled", True)
        ListViewHelpContextID = .ReadProperty("ListViewHelpContextID", 0)
        ListViewSorted = .ReadProperty("ListViewSorted", False)
        ListViewSortKey = .ReadProperty("ListViewSortKey", 0)
        ListViewTag = .ReadProperty("ListViewTag", "")
        ListViewToolTipText = .ReadProperty("ListViewToolTipText", "")
        ListViewWhatsThisHelpID = .ReadProperty("ListViewWhatsThisHelpID", 0)
        ListViewFlatScrollBar = .ReadProperty("ListViewFlatScrollBar", False)
        HeadingsToRightAlign = .ReadProperty("HeadingsToRightAlign", "")
        HeadingsIsDate = .ReadProperty("HeadingsIsDate", "")
        HeadingsToSum = .ReadProperty("HeadingsToSum", "")
        Set Font = .ReadProperty("Font")
        TreeViewEnabled = .ReadProperty("TreeViewEnabled", True)
        TreeViewHelpContextID = .ReadProperty("TreeViewHelpContextID", 0)
        TreeViewHideSelection = .ReadProperty("TreeViewHideSelection", 0)
        TreeViewIndentation = .ReadProperty("TreeViewIndentation", 400.92)
        TreeViewHotTracking = .ReadProperty("TreeViewHotTracking", False)
        TreeViewLabelEdit = .ReadProperty("TreeViewLabelEdit", lvwManual)
        TreeViewPathSeparator = .ReadProperty("TreeViewPathSeparator", "\")
        TreeViewScroll = .ReadProperty("TreeViewScroll", True)
    End With
    Err.Clear
End Sub
Private Sub UserControl_Resize()
    On Error Resume Next
    TreeViewWidth = UserControl.Width / 3
    Err.Clear
End Sub
Sub SizeControls(x As Single)
    On Error Resume Next
    tvTreeView.Top = UserControl.ScaleTop
    tvTreeView.Left = UserControl.ScaleLeft
    tvTreeView.Height = UserControl.ScaleHeight
    tvTreeView.Width = x
    lvListView.Top = tvTreeView.Top
    lvListView.Height = tvTreeView.Height
    lvListView.Left = x + 40
    lvListView.Width = UserControl.ScaleWidth - (tvTreeView.Width + imgSplitter.Width) + 60
    picSplitter.Top = tvTreeView.Top
    picSplitter.Height = tvTreeView.Height
    picSplitter.Left = x
    picSplitter.Width = imgSplitter.Width
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
    imgSplitter.Left = x
    Err.Clear
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        .WriteProperty "ListViewCausesValidation", ListViewCausesValidation, False
        .WriteProperty "ListViewFullRowSelect", ListViewFullRowSelect, True
        .WriteProperty "InnerAppearance", InnerAppearance, cc3D
        .WriteProperty "InnerBorderStyle", InnerBorderStyle, ccNone
        .WriteProperty "OuterAppearance", OuterAppearance, ccFlat
        .WriteProperty "OuterBorderStyle", OuterBorderStyle, ccNone
        .WriteProperty "TreeViewWidth", TreeViewWidth, UserControl.Width / 3
        .WriteProperty "TreeViewCheckBoxes", TreeViewCheckBoxes, False
        .WriteProperty "TreeViewFullRowSelect", TreeViewFullRowSelect, False
        .WriteProperty "TreeViewSorted", TreeViewSorted, False
        .WriteProperty "TreeViewToolTipText", TreeViewToolTipText, ""
        .WriteProperty "TreeViewWhatsThisHelpID", TreeViewWhatsThisHelpID, 0
        .WriteProperty "TreeViewLineStyle", TreeViewLineStyle, tvwTreeLines
        .WriteProperty "ListViewAutoResize", ListViewAutoResize, True
        .WriteProperty "TreeViewSingleSel", TreeViewSingleSel, False
        .WriteProperty "TreeViewStyle", TreeViewStyle, tvwTreelinesPlusMinusPictureText
        .WriteProperty "Headings", Headings, ""
        .WriteProperty "ListViewView", ListViewView, lvwIcon
        .WriteProperty "ListViewCheckBoxes", ListViewCheckBoxes, False
        .WriteProperty "ListViewMultiSelect", ListViewMultiSelect, False
        .WriteProperty "ListViewSortOrder", ListViewSortOrder, lvwAscending
        .WriteProperty "ListViewGridLines", ListViewGridLines, False
        .WriteProperty "ListViewHideSelection", ListViewHideSelection, True
        .WriteProperty "ListViewHotTracking", ListViewHotTracking, False
        .WriteProperty "ListViewHoverSelection", ListViewHoverSelection, False
        .WriteProperty "ListViewLabelEdit", ListViewLabelEdit, lvwManual
        .WriteProperty "ListViewLabelWrap", ListViewLabelWrap, True
        .WriteProperty "ListViewAllowColumnReorder", ListViewAllowColumnReorder, False
        .WriteProperty "ListViewEnabled", ListViewEnabled, True
        .WriteProperty "ListViewHelpContextID", ListViewHelpContextID, 0
        .WriteProperty "ListViewSorted", ListViewSorted, False
        .WriteProperty "ListViewSortKey", ListViewSortKey, False
        .WriteProperty "TreeViewIndentation", TreeViewIndentation, 400.92
        .WriteProperty "ListViewTag", ListViewTag, ""
        .WriteProperty "ListViewToolTipText", ListViewToolTipText, ""
        .WriteProperty "ListViewWhatsThisHelpID", ListViewWhatsThisHelpID, 0
        .WriteProperty "ListViewFlatScrollBar", ListViewFlatScrollBar, False
        .WriteProperty "Font", Font
        .WriteProperty "HeadingsToRightAlign", HeadingsToRightAlign, ""
        .WriteProperty "HeadingsIsDate", HeadingsIsDate, ""
        .WriteProperty "HeadingsToSum", HeadingsToSum, ""
        .WriteProperty "TreeViewEnabled", TreeViewEnabled, True
        .WriteProperty "TreeViewHelpContextID", TreeViewHelpContextID, 0
        .WriteProperty "TreeViewHideSelection", TreeViewHideSelection, True
        .WriteProperty "TreeViewHotTracking", TreeViewHotTracking, False
        .WriteProperty "TreeViewLabelEdit", TreeViewLabelEdit, lvwManual
        .WriteProperty "TreeViewPathSeparator", TreeViewPathSeparator, "\"
        .WriteProperty "TreeViewScroll", TreeViewScroll, True
    End With
    Err.Clear
End Sub
Public Function TreeViewAddNode(Optional Relative As String = "", Optional RelationShip As MSComctlLib.TreeRelationshipConstants = tvwFirst, Optional Key As String = "", Optional ByVal Text As String = "", Optional ByVal Image As String = "close", Optional SelectedImage As String = "open") As MSComctlLib.Node
    On Error Resume Next
    Dim isThere As Node
    Set isThere = tvTreeView.Nodes(Key)
    If TypeName(isThere) = "Nothing" Then
        If Len(Relative) > 0 Then
            Set isThere = tvTreeView.Nodes.Add(Relative, RelationShip)
        Else
            Set isThere = tvTreeView.Nodes.Add(, RelationShip)
        End If
    End If
    isThere.Text = Text
    If Len(Key) > 0 Then isThere.Key = Key
    If Len(Image) > 0 Then isThere.Image = Image
    If Len(SelectedImage) > 0 Then isThere.SelectedImage = SelectedImage
    Set TreeViewAddNode = isThere
    Err.Clear
End Function
Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
    Err.Clear
End Sub
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim sglPos As Single
    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > UserControl.Width - sglSplitLimit Then
            picSplitter.Left = UserControl.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
    Err.Clear
End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
    Err.Clear
End Sub
Private Sub tvTreeView_DragDrop(Source As Control, x As Single, y As Single)
    On Error Resume Next
    If Source = imgSplitter Then
        SizeControls x
    End If
    RaiseEvent TreeViewDragDrop(Source, x, y)
    Err.Clear
End Sub
Public Property Get Headings() As String
    On Error Resume Next
    Headings = mvarHeadings
    Err.Clear
End Property
Public Property Let Headings(ByVal vNewValue As String)
    On Error Resume Next
    Dim fldCnt As Integer
    Dim FldHead() As String
    Dim fldTot As Integer
    Dim colX As MSComctlLib.ColumnHeader
    mvarHeadings = vNewValue
    FldHead = Split(vNewValue, ",")
    fldTot = UBound(FldHead)
    lvListView.ColumnHeaders.Clear
    For fldCnt = 0 To fldTot
        Set colX = lvListView.ColumnHeaders.Add(, , FldHead(fldCnt), 1440)
        Err.Clear
    Next
    lvListView.Refresh
    PropertyChanged "Headings"
    Err.Clear
End Property
Public Property Get ListViewCheckBoxes() As Boolean
    On Error Resume Next
    ListViewCheckBoxes = mvarListViewCheckBoxes
    Err.Clear
End Property
Public Property Let ListViewCheckBoxes(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewCheckBoxes = vNewValue
    lvListView.Checkboxes = vNewValue
    PropertyChanged "ListViewCheckBoxes"
    Err.Clear
End Property
Public Property Get ListViewFullRowSelect() As Boolean
    On Error Resume Next
    ListViewFullRowSelect = mvarListViewFullRowSelect
    Err.Clear
End Property
Public Property Let ListViewFullRowSelect(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewFullRowSelect = vNewValue
    lvListView.FullRowSelect = vNewValue
    PropertyChanged "ListViewFullRowSelect"
    Err.Clear
End Property
Public Property Get ListViewMultiSelect() As Boolean
    On Error Resume Next
    ListViewMultiSelect = mvarListViewMultiSelect
    Err.Clear
End Property
Public Property Let ListViewMultiSelect(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewMultiSelect = vNewValue
    lvListView.MultiSelect = vNewValue
    PropertyChanged "ListViewMultiSelect"
    Err.Clear
End Property
Public Property Get ListViewSortOrder() As MSComctlLib.ListSortOrderConstants
    On Error Resume Next
    ListViewSortOrder = mvarListViewSortOrder
    Err.Clear
End Property
Public Property Let ListViewSortOrder(vNewValue As MSComctlLib.ListSortOrderConstants)
    On Error Resume Next
    mvarListViewSortOrder = vNewValue
    lvListView.SortOrder = vNewValue
    PropertyChanged "ListViewSortOrder"
    Err.Clear
End Property
Public Property Get ListViewGridLines() As Boolean
    On Error Resume Next
    ListViewGridLines = mvarListViewGridLines
    Err.Clear
End Property
Public Property Let ListViewGridLines(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewGridLines = vNewValue
    lvListView.GridLines = vNewValue
    PropertyChanged "ListViewGridLines"
    Err.Clear
End Property
Public Property Get ListViewHideColumnHeaders() As Boolean
    On Error Resume Next
    ListViewHideColumnHeaders = mvarListViewHideColumnHeaders
    Err.Clear
End Property
Public Property Let ListViewHideColumnHeaders(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewHideColumnHeaders = vNewValue
    lvListView.HideColumnHeaders = vNewValue
    PropertyChanged "ListViewHideColumnHeaders"
    Err.Clear
End Property
Public Property Get ListViewHideSelection() As Boolean
    On Error Resume Next
    ListViewHideSelection = mvarListViewHideSelection
    Err.Clear
End Property
Public Property Let ListViewHideSelection(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewHideSelection = vNewValue
    lvListView.HideSelection = vNewValue
    PropertyChanged "ListViewHideSelection"
    Err.Clear
End Property
Public Property Get TreeViewHideSelection() As Boolean
    On Error Resume Next
    TreeViewHideSelection = mvarTreeViewHideSelection
    Err.Clear
End Property
Public Property Let TreeViewHideSelection(vNewValue As Boolean)
    On Error Resume Next
    mvarTreeViewHideSelection = vNewValue
    tvTreeView.HideSelection = vNewValue
    PropertyChanged "TreeViewHideSelection"
    Err.Clear
End Property
Public Property Get ListViewHotTracking() As Boolean
    On Error Resume Next
    ListViewHotTracking = mvarListViewHotTracking
    Err.Clear
End Property
Public Property Let ListViewHotTracking(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewHotTracking = vNewValue
    lvListView.HotTracking = vNewValue
    PropertyChanged "ListViewHotTracking"
    Err.Clear
End Property
Public Property Get TreeViewHotTracking() As Boolean
    On Error Resume Next
    TreeViewHotTracking = mvarTreeViewHotTracking
    Err.Clear
End Property
Public Property Let TreeViewHotTracking(vNewValue As Boolean)
    On Error Resume Next
    mvarTreeViewHotTracking = vNewValue
    tvTreeView.HotTracking = vNewValue
    PropertyChanged "TreeViewHotTracking"
    Err.Clear
End Property
Public Property Get ListViewHoverSelection() As Boolean
    On Error Resume Next
    ListViewHoverSelection = mvarListViewHoverSelection
    Err.Clear
End Property
Public Property Let ListViewHoverSelection(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewHoverSelection = vNewValue
    lvListView.HoverSelection = vNewValue
    PropertyChanged "ListViewHoverSelection"
    Err.Clear
End Property
Public Property Get ListViewLabelEdit() As LabelEditConstants
    On Error Resume Next
    ListViewLabelEdit = mvarListViewLabelEdit
    Err.Clear
End Property
Public Property Let ListViewLabelEdit(vNewValue As MSComctlLib.LabelEditConstants)
    On Error Resume Next
    mvarListViewLabelEdit = vNewValue
    lvListView.LabelEdit = vNewValue
    PropertyChanged "ListViewLabelEdit"
    Err.Clear
End Property
Public Property Get TreeViewLabelEdit() As LabelEditConstants
    On Error Resume Next
    TreeViewLabelEdit = mvarTreeViewLabelEdit
    Err.Clear
End Property
Public Property Let TreeViewLabelEdit(vNewValue As MSComctlLib.LabelEditConstants)
    On Error Resume Next
    mvarTreeViewLabelEdit = vNewValue
    tvTreeView.LabelEdit = vNewValue
    PropertyChanged "TreeViewLabelEdit"
    Err.Clear
End Property
Public Property Get ListViewLabelWrap() As Boolean
    On Error Resume Next
    ListViewLabelWrap = mvarListViewLabelWrap
    Err.Clear
End Property
Public Property Let ListViewLabelWrap(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewLabelWrap = vNewValue
    lvListView.LabelWrap = vNewValue
    PropertyChanged "ListViewLabelWrap"
    Err.Clear
End Property
Public Property Get ListViewAllowColumnReorder() As Boolean
    On Error Resume Next
    ListViewAllowColumnReorder = mvarListViewAllowColumnReorder
    Err.Clear
End Property
Public Property Let ListViewAllowColumnReorder(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewAllowColumnReorder = vNewValue
    lvListView.AllowColumnReorder = vNewValue
    PropertyChanged "ListViewAllowColumnReorder"
    Err.Clear
End Property
Public Property Get ListViewEnabled() As Boolean
    On Error Resume Next
    ListViewEnabled = mvarListViewEnabled
    Err.Clear
End Property
Public Property Let ListViewEnabled(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewEnabled = vNewValue
    lvListView.Enabled = vNewValue
    PropertyChanged "ListViewEnabled"
    Err.Clear
End Property
Public Property Get TreeViewEnabled() As Boolean
    On Error Resume Next
    TreeViewEnabled = mvarTreeViewEnabled
    Err.Clear
End Property
Public Property Let TreeViewEnabled(vNewValue As Boolean)
    On Error Resume Next
    mvarTreeViewEnabled = vNewValue
    tvTreeView.Enabled = vNewValue
    PropertyChanged "TreeViewEnabled"
    Err.Clear
End Property
Public Property Get ListViewHelpContextID() As Long
    On Error Resume Next
    ListViewHelpContextID = mvarListViewHelpContextID
    Err.Clear
End Property
Public Property Let ListViewHelpContextID(vNewValue As Long)
    On Error Resume Next
    mvarListViewHelpContextID = vNewValue
    lvListView.HelpContextID = vNewValue
    PropertyChanged "ListViewHelpContextID"
    Err.Clear
End Property
Public Property Get TreeViewHelpContextID() As Long
    On Error Resume Next
    TreeViewHelpContextID = mvarTreeViewHelpContextID
    Err.Clear
End Property
Public Property Let TreeViewHelpContextID(vNewValue As Long)
    On Error Resume Next
    mvarTreeViewHelpContextID = vNewValue
    tvTreeView.HelpContextID = vNewValue
    PropertyChanged "TreeViewHelpContextID"
    Err.Clear
End Property
Public Property Get ListViewSorted() As Boolean
    On Error Resume Next
    ListViewSorted = mvarListViewSorted
    Err.Clear
End Property
Public Property Let ListViewSorted(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewSorted = vNewValue
    lvListView.Sorted = vNewValue
    PropertyChanged "ListViewSorted"
    Err.Clear
End Property
Public Property Get ListViewSortKey() As Integer
    On Error Resume Next
    ListViewSortKey = mvarListViewSortKey
    Err.Clear
End Property
Public Property Let ListViewSortKey(vNewValue As Integer)
    On Error Resume Next
    mvarListViewSortKey = vNewValue
    lvListView.SortKey = vNewValue
    PropertyChanged "ListViewSortKey"
    Err.Clear
End Property
Public Property Get ListViewTag() As String
    On Error Resume Next
    ListViewTag = mvarListViewTag
    Err.Clear
End Property
Public Property Let ListViewTag(vNewValue As String)
    On Error Resume Next
    mvarListViewTag = vNewValue
    lvListView.Tag = vNewValue
    PropertyChanged "ListViewTag"
    Err.Clear
End Property
Public Property Get ListViewToolTipText() As String
    On Error Resume Next
    ListViewToolTipText = mvarListViewToolTipText
    Err.Clear
End Property
Public Property Let ListViewToolTipText(vNewValue As String)
    On Error Resume Next
    mvarListViewToolTipText = vNewValue
    lvListView.ToolTipText = vNewValue
    PropertyChanged "ListViewToolTipText"
    Err.Clear
End Property
Public Property Get ListViewWhatsThisHelpID() As Long
    On Error Resume Next
    ListViewWhatsThisHelpID = mvarListViewWhatsThisHelpID
    Err.Clear
End Property
Public Property Let ListViewWhatsThisHelpID(vNewValue As Long)
    On Error Resume Next
    mvarListViewWhatsThisHelpID = vNewValue
    lvListView.WhatsThisHelpID = vNewValue
    PropertyChanged "ListViewWhatsThisHelpID"
    Err.Clear
End Property
Public Property Get ListViewFlatScrollBar() As Boolean
    On Error Resume Next
    ListViewFlatScrollBar = mvarListViewFlatScrollBar
    Err.Clear
End Property
Public Property Let ListViewFlatScrollBar(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewFlatScrollBar = vNewValue
    lvListView.FlatScrollBar = vNewValue
    PropertyChanged "ListViewFlatScrollBar"
    Err.Clear
End Property
Public Property Get Font() As StdFont
    On Error Resume Next
    Set Font = UserControl.Font
    Err.Clear
End Property
Public Property Set Font(newFont As StdFont)
    On Error Resume Next
    tvTreeView.Font.Name = newFont.Name
    tvTreeView.Font.Size = newFont.Size
    tvTreeView.Font.Bold = newFont.Bold
    tvTreeView.Font.Italic = newFont.Italic
    tvTreeView.Font.Strikethrough = newFont.Strikethrough
    tvTreeView.Font.Underline = newFont.Underline
    lvListView.Font.Name = newFont.Name
    lvListView.Font.Size = newFont.Size
    lvListView.Font.Bold = newFont.Bold
    lvListView.Font.Italic = newFont.Italic
    lvListView.Font.Strikethrough = newFont.Strikethrough
    lvListView.Font.Underline = newFont.Underline
    UserControl.Font.Name = newFont.Name
    UserControl.Font.Size = newFont.Size
    UserControl.Font.Charset = newFont.Charset
    UserControl.Font.Strikethrough = newFont.Strikethrough
    UserControl.Font.Underline = newFont.Underline
    UserControl.Font.Bold = newFont.Bold
    UserControl.Font.Italic = newFont.Italic
    PropertyChanged "Font"
    Err.Clear
End Property
Private Sub ListViewRightAlignThese(ParamArray vColumnNames())
    On Error Resume Next
    Dim colItem As Variant
    Dim colPos As Long
    For Each colItem In vColumnNames
        colPos = ListViewColumnPosition(CStr(colItem))
        If colPos <> 0 Then
            lvListView.ColumnHeaders(colPos).Alignment = lvwColumnRight
        End If
        Err.Clear
    Next
    lvListView.Refresh
    Err.Clear
End Sub
Public Function ListViewColumnPosition(ByVal StrColName As String) As Long
    On Error Resume Next
    Dim xCols() As String
    Dim yPos As Long
    xCols = Split(Headings, ",")
    yPos = ArraySearch(xCols, StrColName)
    ListViewColumnPosition = IIf((yPos = -1), -1, yPos + 1)
    Err.Clear
End Function
Private Function MvSearch(ByVal StrMv As String, ByVal StrSearch As String, Delimiter As String) As Long
    On Error Resume Next
    Dim xValues() As String
    Dim xPos As Long
    xValues = Split(StrMv, Delimiter)
    xPos = ArraySearch(xValues, StrSearch)
    MvSearch = IIf((xPos = -1), 0, xPos + 1)
    Err.Clear
End Function
Private Function ArraySearch(varArray() As String, ByVal StrSearch As String) As Long
    On Error Resume Next
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim strCur As String
    Dim arrayLow As Long
    ArrayTot = UBound(varArray)
    arrayLow = LBound(varArray)
    StrSearch = LCase$(Trim$(StrSearch))
    ArraySearch = -1
    For arrayCnt = arrayLow To ArrayTot
        strCur = LCase$(varArray(arrayCnt))
        Select Case strCur
        Case StrSearch
            ArraySearch = arrayCnt
            Exit For
        End Select
        Err.Clear
    Next
    Err.Clear
End Function
Public Property Get HeadingsToRightAlign() As String
    On Error Resume Next
    HeadingsToRightAlign = mvarListViewRightAlign
    Err.Clear
End Property
Public Property Let HeadingsToRightAlign(vNewValue As String)
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr() As String
    If Len(vNewValue) = 0 Then Exit Property
    mvarListViewRightAlign = vNewValue
    PropertyChanged "HeadingsToRightAlign"
    rsTot = lvListView.ColumnHeaders.Count - 1
    For rsCnt = 1 To rsTot
        lvListView.ColumnHeaders(rsCnt).Alignment = lvwColumnLeft
        Err.Clear
    Next
    rsStr = Split(vNewValue, ",")
    rsTot = UBound(rsStr)
    For rsCnt = 0 To rsTot
        ListViewRightAlignThese rsStr(rsCnt)
        Err.Clear
    Next
    lvListView.Refresh
    Err.Clear
End Property
Public Property Get HeadingsIsDate() As String
    On Error Resume Next
    HeadingsIsDate = mvarHeadingsIsDate
    Err.Clear
End Property
Public Property Let HeadingsIsDate(vNewValue As String)
    On Error Resume Next
    mvarHeadingsIsDate = vNewValue
    PropertyChanged "HeadingsIsDate"
    Err.Clear
End Property
Public Property Get HeadingsToSum() As String
    On Error Resume Next
    HeadingsToSum = mvarHeadingsToSum
    Err.Clear
End Property
Public Property Let HeadingsToSum(vNewValue As String)
    On Error Resume Next
    mvarHeadingsToSum = vNewValue
    PropertyChanged "HeadingsToSum"
    Err.Clear
End Property
Private Sub LstViewSumColumns(ToMoney As Boolean, ParamArray MyColumns())
    On Error Resume Next
    Dim rsCnt As Long
    Dim myColumn As Variant
    Dim colSum As String
    Dim totPos As Long
    Dim spLine() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spColumns() As String
    totPos = ListViewFindItem("Totals", lvwText, lvwWhole)
    If totPos = 0 Then
        lvListView.ListItems.Add , , "Totals"
    End If
    For Each myColumn In MyColumns
        spColumns = Split(CStr(myColumn), ",")
        spTot = UBound(spColumns)
        For spCnt = 0 To spTot
            rsCnt = ListViewColumnPosition(spColumns(spCnt))
            colSum = LstViewSumColumn(rsCnt, ToMoney)
            totPos = ListViewFindItem("Totals", lvwText, lvwWhole)
            spLine = ListViewGetRow(totPos)
            If ToMoney = True Then
                colSum = MakeMoney(colSum)
            Else
                colSum = Format$(colSum, "#,###")
            End If
            spLine(rsCnt) = colSum
            totPos = LstViewUpdate(spLine, CStr(totPos), "note", "note")
            lvListView.ListItems(totPos).EnsureVisible
            Err.Clear
        Next
        Err.Clear
    Next
    'MySQL.LstViewAutoResize lstView
    Err.Clear
End Sub
Public Function ListViewFindItem(ByVal StrSearch As String, Optional ByVal SearchWhere As MSComctlLib.ListFindItemWhereConstants = lvwText, Optional SearchItemType As MSComctlLib.ListFindItemHowConstants = lvwWhole) As Long
    On Error Resume Next
    Dim itmFound As ListItem
    ListViewFindItem = 0
    Set itmFound = lvListView.FindItem(StrSearch, SearchWhere, , SearchItemType)
    If TypeName(itmFound) = "Nothing" Then
        Err.Clear
        Exit Function
    End If
    ListViewFindItem = CLng(itmFound.Index)
    Set itmFound = Nothing
    Err.Clear
End Function
Public Function LstViewSelectedItem() As MSComctlLib.ListItem
    On Error Resume Next
    Set LstViewSelectedItem = lvListView.SelectedItem
    Err.Clear
End Function
Public Function LstViewSumColumn(colPos As Long, Optional ToMoney As Boolean = False) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim spLine() As String
    Dim strSum As String
    strSum = "0"
    rsTot = lvListView.ListItems.Count
    For rsCnt = 1 To rsTot
        spLine = ListViewGetRow(rsCnt)
        If spLine(1) = "Totals" Then
        Else
            If ToMoney = True Then
                strSum = Val(ProperAmount(strSum)) + Val(ProperAmount(spLine(colPos)))
                strSum = ProperAmount(strSum)
            Else
                strSum = Val(strSum) + Val(Replace$(spLine(colPos), ",", ""))
            End If
        End If
        If ToMoney = True Then
            spLine(colPos) = MakeMoney(spLine(colPos))
        Else
            spLine(colPos) = Format$(spLine(colPos), "#,###")
        End If
        Call LstViewUpdate(spLine, CStr(rsCnt), "note", "note")
        Err.Clear
    Next
    If ToMoney = True Then
        LstViewSumColumn = ProperAmount(strSum)
    Else
        LstViewSumColumn = Format$(strSum, "#,###")
    End If
    lvListView.Refresh
    Err.Clear
End Function
Public Function ListViewGetRow(ByVal idx As Long) As Variant
    On Error Resume Next
    If idx = 0 Then Exit Function
    Dim retarray() As String
    Dim clsColTot As Long
    Dim clsColCnt As Long
    ListViewGetRow = Array()
    clsColTot = lvListView.ColumnHeaders.Count
    If clsColTot = 0 Then Exit Function
    ReDim retarray(clsColTot)
    retarray(1) = lvListView.ListItems(idx).Text
    clsColTot = clsColTot - 1
    For clsColCnt = 1 To clsColTot
        retarray(clsColCnt + 1) = lvListView.ListItems(idx).SubItems(clsColCnt)
        Err.Clear
    Next
    ListViewGetRow = retarray
    Err.Clear
End Function
Public Property Get TreeViewLineStyle() As MSComctlLib.TreeLineStyleConstants
    On Error Resume Next
    TreeViewLineStyle = mvarTreeViewLineStyle
    Err.Clear
End Property
Public Property Let TreeViewLineStyle(vNewValue As MSComctlLib.TreeLineStyleConstants)
    On Error Resume Next
    mvarTreeViewLineStyle = vNewValue
    tvTreeView.LineStyle = vNewValue
    PropertyChanged "TreeViewLineStyle"
    Err.Clear
End Property
Public Property Get ListViewAutoResize() As Boolean
    On Error Resume Next
    ListViewAutoResize = mvarListViewAutoResize
    Err.Clear
End Property
Public Property Let ListViewAutoResize(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewAutoResize = vNewValue
    PropertyChanged "ListViewAutoResize"
    Err.Clear
End Property
Public Property Get ListViewCausesValidation() As Boolean
    On Error Resume Next
    ListViewCausesValidation = mvarListViewCausesValidation
    Err.Clear
End Property
Public Property Let ListViewCausesValidation(vNewValue As Boolean)
    On Error Resume Next
    mvarListViewCausesValidation = vNewValue
    lvListView.CausesValidation = vNewValue
    PropertyChanged "ListViewCausesValidation"
    Err.Clear
End Property
Public Property Get TreeViewStyle() As MSComctlLib.TreeStyleConstants
    On Error Resume Next
    TreeViewStyle = mvarTreeViewStyle
    Err.Clear
End Property
Public Property Get TreeViewNodes() As MSComctlLib.Nodes
    On Error Resume Next
    Set TreeViewNodes = tvTreeView.Nodes
    Err.Clear
End Property
Public Property Get ListViewListItems() As MSComctlLib.ListItems
    On Error Resume Next
    Set ListViewListItems = lvListView.ListItems
    Err.Clear
End Property
Public Property Let TreeViewStyle(vNewValue As MSComctlLib.TreeStyleConstants)
    On Error Resume Next
    mvarTreeViewStyle = vNewValue
    tvTreeView.Style = vNewValue
    PropertyChanged "TreeViewStyle"
    Err.Clear
End Property
Private Sub LstViewAutoResize()
    On Error Resume Next
    Dim col2adjust As Long
    Dim col2adjust_Tot As Long
    If lvListView.ListItems.Count = 0 Then
        Err.Clear
        Exit Sub
    End If
    col2adjust_Tot = lvListView.ColumnHeaders.Count - 1
    For col2adjust = 0 To col2adjust_Tot
        Call SendMessage(lvListView.HWND, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function ListViewColumnValues(ByVal Column As Variant, Optional ByVal Delimiter As String = ";", Optional RemoveDuplicates As Boolean = False) As String
    On Error Resume Next
    Dim strData As String
    Dim strR As String
    Dim clsRowTot As Long
    Dim clsRowCnt As Long
    Dim intPos As Long
    Dim colCollection As Collection
    Set colCollection = New Collection
    strR = ""
    clsRowTot = lvListView.ListItems.Count
    If clsRowTot = 0 Then
        ListViewColumnValues = ""
        Err.Clear
        Exit Function
    End If
    If (VarType(Column) <> vbInteger) And (VarType(Column) <> vbString) Then
        Err.Raise 1, "ListViewColumnValues", "ListViewColumnValues: Column not of required type (String or Integer)."
        Err.Clear
        Exit Function
    End If
    If (VarType(Column) = vbInteger) Then
        intPos = Column
    Else
        intPos = ListViewColumnPosition(Column)
    End If
    Select Case intPos
    Case 1
        For clsRowCnt = 1 To clsRowTot
            strData = lvListView.ListItems(clsRowCnt).Text
            If RemoveDuplicates = False Then
                strR = strR & strData & Delimiter
            Else
                colCollection.Add strData, strData
            End If
            Err.Clear
        Next
        strR = RemDelim(strR, Delimiter)
    Case Else
        For clsRowCnt = 1 To clsRowTot
            strData = lvListView.ListItems(clsRowCnt).SubItems(intPos - 1)
            If RemoveDuplicates = False Then
                strR = strR & strData & Delimiter
            Else
                colCollection.Add strData, strData
            End If
            Err.Clear
        Next
        strR = RemDelim(strR, Delimiter)
    End Select
    If RemoveDuplicates = False Then
        ListViewColumnValues = strR
    Else
        ListViewColumnValues = MvFromCollection(colCollection, Delimiter)
    End If
    Err.Clear
End Function
Public Function ListViewHitText(x As Single, y As Single) As MSComctlLib.ListItem
    On Error Resume Next
    Set ListViewHitText = lvListView.HitTest(x, y)
    Err.Clear
End Function
Public Sub ListViewDrag(Action As Variant)
    On Error Resume Next
    lvListView.Drag Action
    Err.Clear
End Sub
Public Sub ListViewGetFirstVisible()
    On Error Resume Next
    lvListView.GetFirstVisible
    Err.Clear
End Sub
Public Sub ListViewOLEDrag()
    On Error Resume Next
    lvListView.OLEDrag
    Err.Clear
End Sub
Public Sub ListViewRefresh()
    On Error Resume Next
    lvListView.Refresh
    Err.Clear
End Sub
Public Sub ListViewStartLabelEdit()
    On Error Resume Next
    lvListView.StartLabelEdit
    Err.Clear
End Sub
Public Sub ListViewSetFocus()
    On Error Resume Next
    lvListView.SetFocus
    Err.Clear
End Sub
Function TreeViewPathLocation(ByVal SearchPath As String) As Long
    On Error Resume Next
    Dim myNode As MSComctlLib.Node
    TreeViewPathLocation = 0
    For Each myNode In tvTreeView.Nodes
        If LCase$(myNode.FullPath) = LCase$(SearchPath) Then
            TreeViewPathLocation = myNode.Index
            Exit For
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Public Function TreeViewKeys(Delimiter As String, Optional Checked As Boolean = False) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsKey As String
    Dim rsEnd As String
    rsEnd = ""
    rsTot = tvTreeView.Nodes.Count
    For rsCnt = 1 To rsTot
        rsKey = ""
        If tvTreeView.Nodes(rsCnt).Checked = Checked Then rsKey = tvTreeView.Nodes(rsCnt).Key
        rsEnd = rsEnd & rsKey & Delimiter
        Err.Clear
    Next
    TreeViewKeys = RemDelim(rsEnd, Delimiter)
    Err.Clear
End Function
Public Function TreeViewFullPaths(Delimiter As String, Optional Checked As Boolean = False) As String
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsKey As String
    Dim rsEnd As String
    rsEnd = ""
    rsTot = tvTreeView.Nodes.Count
    For rsCnt = 1 To rsTot
        rsKey = ""
        If tvTreeView.Nodes(rsCnt).Checked = Checked Then rsKey = tvTreeView.Nodes(rsCnt).FullPath
        rsEnd = rsEnd & rsKey & Delimiter
        Err.Clear
    Next
    TreeViewFullPaths = RemDelim(rsEnd, Delimiter)
    Err.Clear
End Function
Public Function ListViewGetRowColumn(ByVal rowPos As Long, ByVal colPos As Long) As String
    On Error Resume Next
    Dim retarray() As String
    retarray = ListViewGetRow(rowPos)
    ListViewGetRowColumn = retarray(colPos)
    Err.Clear
End Function
Public Sub ListViewFilterNew(ByVal ColumnName As String, ByVal ColumnValue As String, ByVal ColumnValueDelim As String, Optional Remove As Boolean = False)
    On Error Resume Next
    Dim rsTot As Long
    Dim rsCnt As Long
    Dim xCols As String
    Dim xPos As Long
    Dim spLine() As String
    Dim curValue As String
    Dim rowPos As Long
    xCols = Headings
    xPos = MvSearch(xCols, ColumnName, ",")
    If xPos = 0 Then
        Err.Clear
        Exit Sub
    End If
    ColumnValue = LCase$(ColumnValue)
    ColumnValue = MvReplaceItem(ColumnValue, "(blank)", "(none)", ColumnValueDelim)
    rsTot = lvListView.ListItems.Count
    For rsCnt = rsTot To 1 Step -1
        spLine = ListViewGetRow(rsCnt)
        curValue = LCase$(Trim$(spLine(xPos)))
        If curValue = "" Then
            curValue = "(none)"
        End If
        rowPos = MvSearch(ColumnValue, curValue, ColumnValueDelim)
        If Remove = False Then
            If rowPos = 0 Then
                lvListView.ListItems.Remove rsCnt
            End If
        Else
            If rowPos > 0 Then
                lvListView.ListItems.Remove rsCnt
            End If
        End If
        DoEvents
        Err.Clear
    Next
    If ListViewAutoResize = True Then LstViewAutoResize
    Err.Clear
End Sub
Private Function StrParse(retarray() As String, ByVal strText As String, ByVal Delimiter As String) As Long
    On Error Resume Next
    Dim varArray() As String
    Dim varCnt As Long
    Dim VarS As Long
    Dim VarE As Long
    Dim varA As Long
    varArray = Split(strText, Delimiter)
    VarS = LBound(varArray)
    VarE = UBound(varArray)
    varA = VarE + 1
    ReDim retarray(varA)
    For varCnt = VarS To VarE
        varA = varCnt + 1
        retarray(varA) = varArray(varCnt)
        Err.Clear
    Next
    StrParse = UBound(retarray)
    Err.Clear
End Function
Private Function MvReplaceItem(ByVal strValue As String, ByVal strItem As String, ByVal StrReplaceWith As String, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim spItems() As String
    Dim spTot As Long
    Dim spCnt As Long
    Call StrParse(spItems, strValue, Delim)
    spTot = UBound(spItems)
    For spCnt = 1 To spTot
        If LCase$(spItems(spCnt)) = LCase$(strItem) Then
            spItems(spCnt) = StrReplaceWith
        End If
        Err.Clear
    Next
    MvReplaceItem = MvFromArray(spItems, Delim)
    Err.Clear
End Function
Private Function MvFromArray(vArray() As String, ByVal Delimiter As String, Optional StartingAt As Long = 1, Optional TrimItem As Boolean = True) As String
    On Error Resume Next
    Dim I As Long
    Dim BldStr As String
    Dim strL As String
    Dim totArray As Long
    totArray = UBound(vArray)
    For I = StartingAt To totArray
        strL = vArray(I)
        If TrimItem = True Then
            strL = Trim$(strL)
        End If
        If I = totArray Then
            BldStr = BldStr & strL
        Else
            BldStr = BldStr & strL & Delimiter
        End If
        Err.Clear
    Next
    MvFromArray = BldStr
    Err.Clear
End Function
Private Function File_Token(ByVal strFileName As String, Optional ByVal Sretrieve As String = "F", Optional ByVal Delim As String = "\") As String
    On Error Resume Next
    Dim intNum As Long
    Dim sNew As String
    File_Token = strFileName
    Select Case UCase$(Sretrieve)
    Case "D"
        File_Token = Left$(strFileName, 3)
    Case "F"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            File_Token = Mid$(strFileName, intNum + 1)
        End If
    Case "P"
        intNum = InStrRev(strFileName, Delim)
        If intNum <> 0 Then
            File_Token = Mid$(strFileName, 1, intNum - 1)
        End If
    Case "E"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            File_Token = Mid$(strFileName, intNum + 1)
        End If
    Case "FO"
        sNew = strFileName
        intNum = InStrRev(sNew, Delim)
        If intNum <> 0 Then
            sNew = Mid$(sNew, intNum + 1)
        End If
        intNum = InStrRev(sNew, ".")
        If intNum <> 0 Then
            sNew = Left$(sNew, intNum - 1)
        End If
        File_Token = sNew
    Case "PF"
        intNum = InStrRev(strFileName, ".")
        If intNum <> 0 Then
            File_Token = Left$(strFileName, intNum - 1)
        End If
    End Select
    Err.Clear
End Function
Public Sub TreeViewFromRecordset(Clear As Boolean, rsAdo As ADODB.Recordset, ByVal TreeFldNames As String, ByVal ListFldNames As String, ByVal TreeDelimiter As String, Optional TreeMergeFields As Boolean = False, Optional TreeIncludeTableName As Boolean = False, Optional ByVal ListFldNamesToSum As String = "", Optional ByVal ListFldNamesToRightAlign As String = "")
    On Error Resume Next
    Dim rsCnt As Long
    Dim rsTot As Long
    Dim rsStr As String
    Dim rsData As String
    Dim spFields() As String
    Dim spCnt As Long
    Dim spTot As Long
    Dim spStr As String
    Dim TbName As String
    Dim mNode As MSComctlLib.Node
    Dim fTot As Long
    Dim fCnt As Long
    Dim fFld() As String
    Dim vFld() As String
    Dim fStr As String
    Dim lstPos As Long
    If Clear = True Then
        Me.Clear
        Me.Headings = ListFldNames
        Me.HeadingsToSum = ListFldNamesToSum
        Me.HeadingsToRightAlign = ListFldNamesToRightAlign
        Me.ListViewListItems.Clear
    End If
    TbName = Table_NameFromSelect(rsAdo.Source)
    spTot = StrParse(spFields, TreeFldNames, ",")
    fTot = StrParse(fFld, ListFldNames, ",")
    ReDim vFld(fTot)
    rsAdo.MoveFirst
    rsTot = rsAdo.RecordCount
    For rsCnt = 1 To rsTot
        ' process tree fields
        rsStr = ""
        For spCnt = 1 To spTot
            spStr = spFields(spCnt)
            spStr = rsAdo.Fields(spStr).Value & ""
            If TreeMergeFields = True Then
                rsStr = rsStr & spStr & TreeDelimiter
            Else
                rsStr = rsStr & spStr & TreeViewPathSeparator
            End If
            Err.Clear
        Next
        If TreeMergeFields = True Then
            rsStr = RemDelim(rsStr, TreeDelimiter)
        Else
            rsStr = RemDelim(rsStr, TreeViewPathSeparator)
        End If
        If TreeIncludeTableName = True Then
            rsStr = TbName & TreeViewPathSeparator & rsStr
        End If
        ' process listview fields
        ReDim vFld(fTot)
        For fCnt = 1 To fTot
            fStr = fFld(fCnt)
            vFld(fCnt) = rsAdo.Fields(fStr).Value & ""
            Err.Clear
        Next
        ' add the tree path as is
        Set mNode = TreeViewAddPath(rsStr)
        ' add the list view item
        lstPos = ListViewAddItem(mNode.Key, mNode.Key, vFld(1))
        For fCnt = 2 To fTot
            Call ListViewListSubItems(lstPos, fFld(fCnt), vFld(fCnt))
            Err.Clear
        Next
        rsAdo.MoveNext
        Err.Clear
    Next
    Err.Clear
End Sub
Private Function Table_NameFromSelect(ByVal strQuery As String) As String
    On Error Resume Next
    Dim fromPos As Long
    Dim lenStr As Long
    Dim lenCnt As Long
    Dim restStr As String
    Dim delimPos As Long
    Dim TbName As String
    fromPos = InStr(1, strQuery, " from ", vbTextCompare)
    If fromPos > 0 Then
        restStr = Trim$(Mid$(strQuery, fromPos + 5))
        ' check using the table Quote
        delimPos = InStr(1, restStr, "[", vbTextCompare)
        lenStr = Len(restStr)
        TbName = ""
        If delimPos > 0 Then
            For lenCnt = 2 To lenStr
                If Mid$(restStr, lenCnt, 1) = "]" Then
                    Exit For
                Else
                    TbName = TbName & Mid$(restStr, lenCnt, 1)
                End If
                Err.Clear
            Next
        Else
            For lenCnt = 1 To lenStr
                If Mid$(restStr, lenCnt, 1) = " " Then
                    Exit For
                Else
                    TbName = TbName & Mid$(restStr, lenCnt, 1)
                End If
                Err.Clear
            Next
        End If
        Table_NameFromSelect = Replace$(TbName, ";", "", , , vbTextCompare)
    Else
        Table_NameFromSelect = strQuery
    End If
    Err.Clear
End Function
Public Function TreeViewAddPath(ByVal sPath As String, Optional ByVal Image As String = "close", Optional ByVal SelectedImage As String = "open") As MSComctlLib.Node
    On Error Resume Next
    Dim arrayPath() As String
    Dim ArrayTot As Long
    Dim arrayCnt As Long
    Dim sParent As String
    Dim pParent As String
    Dim nText As String
    ' split the path to be subitems
    ArrayTot = StrParse(arrayPath, sPath, TreeViewPathSeparator)
    For arrayCnt = 1 To ArrayTot
        ' get the current path
        sParent = MvFromMv(sPath, 1, arrayCnt, TreeViewPathSeparator)
        ' get the text
        nText = MvField(sParent, -1, TreeViewPathSeparator)
        ' get the relative
        pParent = MvFromMv(sPath, 1, arrayCnt - 1, TreeViewPathSeparator)
        If StartWithNumber(pParent) = True Then pParent = "K-" & pParent
        If StartWithNumber(sParent) = True Then sParent = "K" & sParent
        Set TreeViewAddPath = tvTreeView.Nodes(sParent)
        If Err.Number = 35601 Then
            ' element not found
            If Len(pParent) = 0 Then
                Set TreeViewAddPath = tvTreeView.Nodes.Add(, tvwChild, sParent, nText, Image, SelectedImage)
            Else
                Set TreeViewAddPath = tvTreeView.Nodes.Add(pParent, tvwChild, sParent, nText, Image, SelectedImage)
            End If
        End If
        Err.Clear
    Next
    Err.Clear
End Function
Private Function MvFromMv(ByVal strOriginalMv As String, ByVal startPos As Long, Optional ByVal NumOfItems As Long = -1, Optional ByVal Delim As String = "") As String
    On Error Resume Next
    Dim spOriginal() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim sLine As String
    Dim endPos As Long
    sLine = ""
    spTot = StrParse(spOriginal, strOriginalMv, Delim)
    If NumOfItems = -1 Then
        endPos = spTot
    ElseIf NumOfItems = -2 Then
        endPos = spTot - 2
    Else
        endPos = (startPos + NumOfItems) - 1
    End If
    For spCnt = startPos To endPos
        If spCnt = endPos Then
            sLine = sLine & spOriginal(spCnt)
        Else
            sLine = sLine & spOriginal(spCnt) & Delim
        End If
        Err.Clear
    Next
    MvFromMv = sLine
    Err.Clear
End Function
Private Function StartWithNumber(ByVal strValue As String) As Boolean
    On Error Resume Next
    Dim strLeft As String
    If Len(strValue) = 0 Then
        StartWithNumber = False
    Else
        strLeft = Left$(strValue, 1)
        If InStr(1, "0123456789", strLeft) > 0 Then
            StartWithNumber = True
        Else
            StartWithNumber = False
        End If
    End If
    Err.Clear
End Function
Private Function MvField(ByVal strData As String, fldPos As Long, ByVal Delim As String) As String
    On Error Resume Next
    ' returns a substring from a delimted string
    Dim spData() As String
    Dim spCnt As Long
    MvField = ""
    If Len(Delim) = 0 Then
        Delim = Chr$(253)
    End If
    If Len(strData) = 0 Then
        Err.Clear
        Exit Function
    End If
    Call StrParse(spData, strData, Delim)
    spCnt = UBound(spData)
    Select Case fldPos
    Case -1
        MvField = Trim$(spData(spCnt))
    Case -2
        MvField = Trim$(spData(spCnt - 1))
    Case Else
        If fldPos <= spCnt Then
            MvField = Trim$(spData(fldPos))
        End If
    End Select
    Err.Clear
End Function
Public Sub ListViewCheckAll(Optional ByVal bOp As Boolean = True)
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    lstTot = lvListView.ListItems.Count
    For lstCnt = 1 To lstTot
        lvListView.ListItems(lstCnt).Checked = bOp
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function ListViewCheckedToMV(ByVal Index As Variant, ByVal Delimiter As String, Optional bRemoveDuplicates As Boolean = False, Optional bRemoveBlanks As Boolean = False, Optional bRemoveStars As Boolean = True, Optional bRemoveTotals As Boolean = True) As String
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    Dim bOp As Boolean
    Dim lstStr() As String
    Dim retStr As String
    Dim colPos As Long
    If (VarType(Index) <> vbInteger) And (VarType(Index) <> vbString) Then
        Err.Raise 1, "ListViewCheckedToMV", "ListViewCheckedToMV: Index not of required type (String or Integer)."
        Err.Clear
        Exit Function
    End If
    If (VarType(Index) = vbInteger) Then
        colPos = Index
    Else
        colPos = ListViewColumnPosition(Index)
    End If
    retStr = ""
    lstTot = lvListView.ListItems.Count
    For lstCnt = 1 To lstTot
        bOp = lvListView.ListItems(lstCnt).Checked
        Select Case bOp
        Case True
            lstStr = ListViewGetRow(lstCnt)
            retStr = retStr & lstStr(colPos) & Delimiter
        End Select
        Err.Clear
    Next
    retStr = RemDelim(retStr, Delimiter)
    If bRemoveTotals = True Then
        retStr = Replace$(retStr, "Totals", "")
    End If
    If bRemoveStars = True Then
        retStr = Replace$(retStr, "*", "")
    End If
    If bRemoveDuplicates = True Then
        retStr = MvRemoveDuplicates(retStr, Delimiter)
    End If
    ListViewCheckedToMV = retStr
    Err.Clear
End Function
Private Function MvRemoveDuplicates(ByVal StrMvString As String, ByVal Delimiter As String) As String
    On Error Resume Next
    ' returns a string from a string after removing all duplicated sub strings of a delimited string
    Dim spData() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim xCol As Collection
    Set xCol = New Collection
    spTot = StrParse(spData, StrMvString, Delimiter)
    For spCnt = 0 To spTot
        spData(spCnt) = Trim$(spData(spCnt))
        If Len(spData(spCnt)) > 0 Then
            xCol.Add spData(spCnt), spData(spCnt)
        End If
        Err.Clear
    Next
    MvRemoveDuplicates = MvFromCollection(xCol, Delimiter)
    Err.Clear
End Function
Public Sub ListViewCheckFromMv(ByVal Index As Variant, ByVal StrUseMv As String, ByVal Delimiter As String, Optional boolCheck As Boolean = True, Optional useColor As Long = vbBlack, Optional bShow As Boolean = False)
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    Dim lstPos As Long
    Dim colPos As Long
    Dim useData() As String
    If (VarType(Index) <> vbInteger) And (VarType(Index) <> vbString) Then
        Err.Raise 1, "ListViewCheckFromMv", "ListViewCheckFromMv: Index not of required type (String or Integer)."
        Err.Clear
        Exit Sub
    End If
    If (VarType(Index) = vbInteger) Then
        colPos = Index
    Else
        colPos = ListViewColumnPosition(Index)
    End If
    ' uncheck or check all items at first
    lstTot = lvListView.ListItems.Count
    For lstCnt = 1 To lstTot
        lvListView.ListItems(lstCnt).Checked = Not boolCheck
        Err.Clear
    Next
    lstTot = StrParse(useData, StrUseMv, Delimiter)
    For lstCnt = 1 To lstTot
        Select Case colPos
        Case 1
            lstPos = ListViewFindItem(useData(lstCnt))
        Case Else
            lstPos = ListViewFindItem(useData(lstCnt), lvwSubItem, lvwWhole)
        End Select
        If lstPos > 0 Then
            lvListView.ListItems(lstPos).Checked = boolCheck
            lvListView.ListItems(lstPos).ForeColor = useColor
        End If
        Err.Clear
    Next
    Err.Clear
End Sub
Public Sub ListViewColIconv(IConvCode As IconvEnum, ParamArray Indexes())
    On Error Resume Next
    Dim strData As String
    Dim clsRowTot As Long
    Dim clsRowCnt As Long
    Dim colPos As Long
    Dim strHeads As String
    Dim Index As Variant
    Dim xLine() As String
    Dim spHeads() As String
    Dim spCnt As Long
    Dim spTot As Long
    Dim strValue As String
    strHeads = ""
    For Each Index In Indexes
        strHeads = strHeads & CStr(Index)
        Err.Clear
    Next
    strHeads = RemDelim(strHeads, ",")
    spTot = StrParse(spHeads, strHeads, ",")
    clsRowTot = lvListView.ListItems.Count
    For clsRowCnt = 1 To clsRowTot
        xLine = ListViewGetRow(clsRowCnt)
        For spCnt = 1 To spTot
            colPos = ListViewColumnPosition(spHeads(spCnt))
            Select Case IConvCode
            Case Money
                strValue = ProperAmount(xLine(colPos))
            Case Date
                strValue = DateIconv(xLine(colPos))
            Case Proper
                strValue = ProperCase(xLine(colPos))
            End Select
            Select Case colPos
            Case 1
                lvListView.ListItems(clsRowCnt).Text = strValue
            Case Else
                lvListView.ListItems(clsRowCnt).SubItems(colPos - 1) = strValue
            End Select
            Err.Clear
        Next
        Err.Clear
    Next
    Err.Clear
End Sub
Public Function DateIconv(ByVal sDate As String) As String
    On Error Resume Next
    DateIconv = sDate
    If Len(sDate) = 0 Then
        Err.Clear
        Exit Function
    End If
    If IsDate(sDate) = True Then DateIconv = DateDiff("d", "31/12/1967", Format$(Now, "dd/mm/yyyy"))
    Err.Clear
End Function
Private Function ProperCase(ByVal StrString As String, Optional Delim As String = "\") As String
    On Error Resume Next
    Dim spItems() As String
    Dim spSubs() As String
    Dim spTot As Long
    Dim spCnt As Long
    Dim spTott As Long
    Dim spCntt As Long
    Dim spSubss() As String
    StrString = Trim$(StrString)
    spTot = StrParse(spItems, StrString, Delim)
    For spCnt = 1 To spTot
        spItems(spCnt) = StrConv(spItems(spCnt), vbProperCase)
        spTott = StrParse(spSubss, spItems(spCnt), "|")
        For spCntt = 1 To spTott
            spSubss(spCntt) = StrConv(spSubss(spCntt), vbProperCase)
            Err.Clear
        Next
        spItems(spCnt) = MvFromArray(spSubss, "|")
        Err.Clear
    Next
    ProperCase = MvFromArray(spItems, Delim)
    Erase spItems
    Erase spSubs
    Err.Clear
End Function
Public Function ListViewColMV(ByVal Index As Variant, ByVal Delimiter As String, Optional bDistinct As Boolean = False) As String
    On Error Resume Next
    ListViewColMV = ListViewColumnValues(Index, Delimiter, bDistinct)
    Err.Clear
End Function
Public Function ListViewColNames(lstView As MSComctlLib.ListView) As String
    On Error Resume Next
    Dim strHead As String
    Dim strName As String
    Dim clsColTot As Long
    Dim clsColCnt As Long
    strHead = ""
    clsColTot = lstView.ColumnHeaders.Count
    For clsColCnt = 1 To clsColTot
        strName = lstView.ColumnHeaders(clsColCnt).Text
        Select Case clsColCnt
        Case clsColTot
            strHead = strHead & strName
        Case Else
            strHead = strHead & strName & ","
        End Select
        Err.Clear
    Next
    ListViewColNames = strHead
    Err.Clear
End Function
Private Function LstViewGetRow(lstView As MSComctlLib.ListView, ByVal idx As Long) As Variant
    On Error Resume Next
    If idx = 0 Then Exit Function
    Dim retarray() As String
    Dim clsColTot As Long
    Dim clsColCnt As Long
    LstViewGetRow = Array()
    clsColTot = lstView.ColumnHeaders.Count
    If clsColTot = 0 Then Exit Function
    ReDim retarray(clsColTot)
    retarray(1) = lstView.ListItems(idx).Text
    clsColTot = clsColTot - 1
    For clsColCnt = 1 To clsColTot
        retarray(clsColCnt + 1) = lstView.ListItems(idx).SubItems(clsColCnt)
        Err.Clear
    Next
    LstViewGetRow = retarray
    Err.Clear
End Function
Public Sub ListViewCopyChechecked(lstSource As MSComctlLib.ListView, lstTarget As MSComctlLib.ListView)
    On Error Resume Next
    Dim spLine() As String
    Dim spCnt As Long
    Dim spTot As Long
    Dim sHeads As String
    sHeads = ListViewColNames(lstSource)
    lstTarget.ListItems.Clear
    LstViewMakeHeadings lstTarget, sHeads
    spTot = lstSource.ListItems.Count
    For spCnt = 1 To spTot
        If lstSource.ListItems(spCnt).Checked = False Then
            GoTo nextLine
        End If
        spLine = LstViewGetRow(lstSource, spCnt)
        Call LstViewUpdate(spLine, lstTarget, "")
nextLine:
        Err.Clear
    Next
    ' align columns
    spTot = StrParse(spLine, sHeads, ",")
    For spCnt = 1 To spTot
        lstTarget.ColumnHeaders(spCnt).Alignment = lstSource.ColumnHeaders(spCnt).Alignment
        Err.Clear
    Next
    lstTarget.Refresh
    DoEvents
    Err.Clear
End Sub
Private Sub LstViewMakeHeadings(lstView As MSComctlLib.ListView, ByVal strHeads As String)
    On Error Resume Next
    ' used to create columns in a listview
    Dim fldCnt As Integer
    Dim FldHead() As String
    Dim fldTot As Integer
    Dim colX As MSComctlLib.ColumnHeader
    FldHead = Split(strHeads, ",")
    fldTot = UBound(FldHead)
    lstView.ColumnHeaders.Clear
    lstView.ListItems.Clear
    lstView.Sorted = False
    ' first column should be left aligned
    Set colX = lstView.ColumnHeaders.Add(, , FldHead(0), 1440)
    For fldCnt = 1 To fldTot
        Set colX = lstView.ColumnHeaders.Add(, , FldHead(fldCnt), 1440)
        Err.Clear
    Next
    Err.Clear
End Sub
Public Sub ListViewFromCollection(varCollection As Collection, ByVal Delimiter As String, Optional ByVal boolClear As Boolean = True)
    On Error Resume Next
    Dim spLine() As String
    Dim varTot As Long
    Dim varCnt As Long
    Dim xTot As Long
    If boolClear = True Then lvListView.ListItems.Clear
    varTot = varCollection.Count
    For varCnt = 1 To varTot
        Call StrParse(spLine, varCollection.Item(varCnt), Delimiter)
        Call LstViewUpdate(spLine, "", "note", "note")
        Err.Clear
    Next
    Err.Clear
End Sub
Public Sub ListViewRemoveChecked(Optional ByVal bCheckedStatus As Boolean = True)
    On Error Resume Next
    Dim bOp As Boolean
    Dim lstTot As Long
    Dim lstCnt As Long
    lstTot = lvListView.ListItems.Count
    For lstCnt = lstTot To 1 Step -1
        bOp = lvListView.ListItems(lstCnt).Checked
        If bOp = bCheckedStatus Then
            lvListView.ListItems.Remove lstCnt
        End If
        Err.Clear
    Next
    Err.Clear
End Sub
Public Sub ListViewRemoveDuplicates()
    On Error Resume Next
    Dim lstTot As Long
    Dim lstCnt As Long
    Dim spLines() As String
    Dim newCol As Collection
    Set newCol = New Collection
    Dim spStr As String
    lstTot = lvListView.ListItems.Count
    For lstCnt = 1 To lstTot
        spLines = ListViewGetRow(lstCnt)
        spStr = MvFromArray(spLines, Chr$(193))
        newCol.Add spStr, spStr
        DoEvents
        Err.Clear
    Next
    lvListView.ListItems.Clear
    lstTot = newCol.Count
    For lstCnt = 1 To lstTot
        Call StrParse(spLines, newCol.Item(lstCnt), Chr$(193))
        LstViewUpdate spLines, "", "note", "note"
        DoEvents
        Err.Clear
    Next
    Set newCol = Nothing
    Err.Clear
End Sub
Public Sub ListViewRowColumnColor(lstView As ListView, rowPos As Long, colPos As Variant, Optional ByVal cForeColor As ColorConstants = vbBlack)
    On Error Resume Next
    Dim colPosNew As Long
    If (VarType(colPos) <> vbInteger) And (VarType(colPos) <> vbString) Then
        Err.Raise 1, "ListViewRowColumnColor", "ListViewRowColumnColor: colPos not of required type (String or Integer)."
        Err.Clear
        Exit Sub
    End If
    If (VarType(colPos) = vbInteger) Then
        colPosNew = colPos
    Else
        colPosNew = ListViewColumnPosition(colPos)
    End If
    Select Case colPos
    Case 1
        lvListView.ListItems(rowPos).ForeColor = cForeColor
    Case Else
        lvListView.ListItems(rowPos).ListSubItems(colPosNew - 1).ForeColor = cForeColor
    End Select
    lstView.Refresh
    Err.Clear
End Sub
Public Sub ListViewRowBold(rowPos As Long, Optional boolBold As Boolean = False)
    On Error Resume Next
    Dim numColumns As Long
    Dim numColumn As Long
    numColumns = lvListView.ColumnHeaders.Count - 1
    lvListView.ListItems(rowPos).Bold = boolBold
    For numColumn = 1 To numColumns
        lvListView.ListItems(rowPos).ListSubItems(numColumn).Bold = boolBold
        Err.Clear
    Next
    lvListView.Refresh
    Err.Clear
End Sub
Public Sub ListViewRowForeColor(rowPos As Long, Optional ByVal cForeColor As ColorConstants = vbBlack)
    On Error Resume Next
    Dim numColumns As Long
    Dim numColumn As Long
    numColumns = lvListView.ColumnHeaders.Count - 1
    lvListView.ListItems(rowPos).ForeColor = cForeColor
    For numColumn = 1 To numColumns
        lvListView.ListItems(rowPos).ListSubItems(numColumn).ForeColor = cForeColor
        Err.Clear
    Next
    lvListView.Refresh
    Err.Clear
End Sub
Public Sub ListViewRowFormat(rowPos As Long, Optional bChecked As Boolean = False, Optional ByVal bBold As Boolean = False, Optional ByVal cForeColor As ColorConstants = vbBlack, Optional rIcon As String = "", Optional rSmallIcon As String = "", Optional ByVal tTooltip As String = "", Optional ByVal tTag As String = "")
    On Error Resume Next
    Dim numColumns As Long
    Dim numColumn As Long
    numColumns = lvListView.ColumnHeaders.Count - 1
    lvListView.ListItems(rowPos).Bold = bBold
    If Len(rIcon) > 0 Then lvListView.ListItems(rowPos).Icon = rIcon
    If Len(rSmallIcon) > 0 Then lvListView.ListItems(rowPos).SmallIcon = rSmallIcon
    lvListView.ListItems(rowPos).Tag = tTag
    lvListView.ListItems(rowPos).ToolTipText = tTooltip
    lvListView.ListItems(rowPos).Checked = bChecked
    lvListView.ListItems(rowPos).ForeColor = cForeColor
    For numColumn = 1 To numColumns
        lvListView.ListItems(rowPos).ListSubItems(numColumn).Bold = bBold
        lvListView.ListItems(rowPos).ListSubItems(numColumn).Tag = tTag
        lvListView.ListItems(rowPos).ListSubItems(numColumn).ForeColor = cForeColor
        lvListView.ListItems(rowPos).ListSubItems(numColumn).ToolTipText = tTooltip
        Err.Clear
    Next
    Err.Clear
End Sub
