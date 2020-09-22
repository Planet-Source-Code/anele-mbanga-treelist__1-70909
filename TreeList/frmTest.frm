VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Tree List Sample"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   14025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoadFromDB 
      Caption         =   "Load From Recordset"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   7560
      Width           =   2055
   End
   Begin Project1.TreeList TreeList 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13695
      _extentx        =   24156
      _extenty        =   11456
      outerappearance =   1
      listviewgridlines=   -1  'True
      font            =   "frmTest.frx":0000
      treeviewhideselection=   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Manually"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   7560
      Width           =   2055
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MyADO As clsADO
Private Sub cmdLoadFromDB_Click()
    On Error Resume Next
    Caption = "TreeList - Database Load"
    Dim adoRs As ADODB.Recordset
    MyADO.OpenConnection App.Path & "\Family Tree.mdb"
    Set adoRs = MyADO.OpenRs("Family Tree")
    TreeList.TreeViewFromRecordset True, adoRs, "First Name,Last Name", _
    "Index,First Name,Last Name,Maiden Name,Date Of Birth,Identity Number,Cell Phone,Fax,Total Children,Father,Mother", " ", True, True, "Total Children", "Identity Number,Cell Phone,Fax,Total Children"
    TreeList.TreeViewFromRecordset False, adoRs, "Father", _
    "Index,First Name,Last Name,Maiden Name,Date Of Birth,Identity Number,Cell Phone,Fax,Total Children,Father,Mother", "", False, True
    TreeList.TreeViewFromRecordset False, adoRs, "Mother", _
    "Index,First Name,Last Name,Maiden Name,Date Of Birth,Identity Number,Cell Phone,Fax,Total Children,Father,Mother", "", False, True
    adoRs.Close
    Err.Clear
End Sub
Private Sub Command1_Click()
    On Error Resume Next
    Caption = "TreeList - Manual Load"
    Dim xPos As Long
    Dim xNode(5) As MSComctlLib.Node
    Dim xListItem(3) As Long
    TreeList.Clear
    TreeList.Headings = "Index,First Name,Last Name,Maiden Name,Date Of Birth,Identity Number,Cell Phone,Fax,Total Children,Father,Mother"
    TreeList.HeadingsToSum = "Total Children"
    TreeList.HeadingsToRightAlign = "Identity Number,Cell Phone,Fax,Total Children"
    Set xNode(1) = TreeList.TreeViewAddPath("Family Tree")
    Set xNode(2) = TreeList.TreeViewAddNode(xNode(1).Key, tvwChild, xNode(1).FullPath & "\Anele Mbanga", "Anele Mbanga")
    Set xNode(3) = TreeList.TreeViewAddPath("Family Tree\Sikelela Mbanga")
    Set xNode(4) = TreeList.TreeViewAddPath("Family Tree\Usibabale Mbanga")
    xListItem(1) = TreeList.ListViewAddItem(xNode(2).FullPath, xNode(2).FullPath, "1", , , , , True)
    TreeList.ListViewListSubItems xListItem(1), "First Name", "Anele", vbGreen
    TreeList.ListViewListSubItems xListItem(1), "Last Name", "Mbanga", vbYellow
    TreeList.ListViewListSubItems xListItem(1), "Date Of Birth", "15/04/1973", vbBlue
    TreeList.ListViewListSubItems xListItem(1), "Identity Number", "730415...", vbMagenta
    TreeList.ListViewListSubItems xListItem(1), "Cell Phone", "083 310 2077", , True
    TreeList.ListViewListSubItems xListItem(1), "Fax", "086 656 9921"
    TreeList.ListViewListSubItems xListItem(1), "Total Children", "1"
    TreeList.ListViewListSubItems xListItem(1), "Father", "Mbulelo Mbanga"
    TreeList.ListViewListSubItems xListItem(1), "Mother", "Sandra Mbanga"
    TreeList.ListViewListSubItems xListItem(1), "Maiden Name", "None", , , "note"
    xListItem(2) = TreeList.ListViewAddItem(xNode(3).FullPath, xNode(3).FullPath, "2", , , , , , , vbGreen, , , True)
    TreeList.ListViewListSubItems xListItem(2), "First Name", "Sikelela"
    TreeList.ListViewListSubItems xListItem(2), "Last Name", "Mbanga"
    TreeList.ListViewListSubItems xListItem(2), "Maiden Name", "Tywala"
    TreeList.ListViewListSubItems xListItem(2), "Date Of Birth", "27/10/1981"
    TreeList.ListViewListSubItems xListItem(2), "Identity Number", "811027..."
    TreeList.ListViewListSubItems xListItem(2), "Cell Phone", "073 530 4869"
    TreeList.ListViewListSubItems xListItem(2), "Total Children", "1"
    xListItem(3) = TreeList.ListViewAddItem(xNode(4).FullPath, xNode(4).FullPath, "3", , , , , True, , , , True)
    TreeList.ListViewListSubItems xListItem(3), "First Name", "Usibabale"
    TreeList.ListViewListSubItems xListItem(3), "Last Name", "Mbanga"
    TreeList.ListViewListSubItems xListItem(3), "Total Children", "0"
    TreeList.ListViewListSubItems xListItem(3), "Father", "Anele Mbanga"
    TreeList.ListViewListSubItems xListItem(3), "Mother", "Sikelela Mbanga"
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Set MyADO = New clsADO
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    TreeList.Width = Me.ScaleWidth - 250
    Err.Clear
End Sub
