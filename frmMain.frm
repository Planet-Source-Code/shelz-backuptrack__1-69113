VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1058
      ButtonWidth     =   1217
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "imlIco32x"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Add"
            Key             =   ""
            Object.ToolTipText     =   "Add a new track to the library"
            Object.Tag             =   "Add"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Open"
            Key             =   ""
            Object.ToolTipText     =   "Open an existing library"
            Object.Tag             =   "Open"
            Style           =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Save As"
            Key             =   ""
            Object.ToolTipText     =   "Save a copy of the currently active library to disk"
            Object.Tag             =   "Save As"
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Options"
            Key             =   ""
            Object.ToolTipText     =   "Configure various aspects of the app. interface"
            Object.Tag             =   "Interface"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Help"
            Key             =   ""
            Object.ToolTipText     =   "Open help in your default web-browser"
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "About"
            Key             =   ""
            Object.ToolTipText     =   "Information about this software"
            Object.Tag             =   "About"
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit"
            Key             =   ""
            Object.ToolTipText     =   "Close BackupTrack"
            Object.Tag             =   "Exit"
         EndProperty
      EndProperty
      Begin VB.PictureBox pSearch 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   5280
         ScaleHeight     =   735
         ScaleWidth      =   3855
         TabIndex        =   2
         Top             =   60
         Width           =   3855
         Begin ComctlLib.Toolbar tbSearch 
            Height          =   390
            Left            =   360
            TabIndex        =   3
            Top             =   0
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   688
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            ImageList       =   "imlIco16x"
            _Version        =   327682
            BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
               NumButtons      =   4
               BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Style           =   1
                  Object.Width           =   2535
                  Value           =   1
               EndProperty
               BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Style           =   3
               EndProperty
               BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Key             =   ""
                  Object.Tag             =   ""
                  Style           =   4
                  Object.Width           =   2535
                  MixedState      =   -1  'True
               EndProperty
               BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                  Enabled         =   0   'False
                  Key             =   "tbSearchShowAll"
                  Object.Tag             =   ""
               EndProperty
            EndProperty
            Begin VB.ComboBox cmbSearchBox 
               Height          =   315
               Left            =   360
               TabIndex        =   4
               Top             =   0
               Width           =   2535
            End
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblSearchResults 
            Alignment       =   2  'Center
            BackColor       =   &H8000000A&
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Width           =   3495
         End
         Begin VB.Image imgSearchIcon 
            Height          =   240
            Left            =   0
            Stretch         =   -1  'True
            Top             =   30
            Width           =   240
         End
      End
   End
   Begin ComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   7680
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin prjBackupTrack.ctlSplitter cSplitterMain 
      Height          =   6255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11033
      SplitterPosition=   3375
      HasCaptions     =   -1  'True
      Caption1        =   "Track List"
      Caption2        =   "Track Browser"
      Begin prjBackupTrack.ctlSplitter cSplitterExplorer 
         Height          =   2535
         Left            =   120
         TabIndex        =   8
         Top             =   3600
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4471
         SplitterOrientation=   1
         SplitterPosition=   3160
         Caption1        =   ""
         Caption2        =   ""
         Begin ComctlLib.ListView LvwBrowser 
            Height          =   2295
            Left            =   3360
            TabIndex        =   10
            Top             =   120
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   4048
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            SmallIcons      =   "imlIco16x"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin ComctlLib.TreeView TvwBrowser 
            Height          =   2295
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   4048
            _Version        =   327682
            Indentation     =   353
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imlIco16x"
            Appearance      =   1
         End
      End
      Begin prjBackupTrack.ctlSplitter cSplitterNavigator 
         Height          =   2535
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4471
         SplitterOrientation=   1
         SplitterPosition=   3160
         Caption1        =   ""
         Caption2        =   ""
         Begin ComctlLib.ListView LvwData 
            Height          =   2295
            Left            =   3360
            TabIndex        =   14
            Top             =   120
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   4048
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            _Version        =   327682
            SmallIcons      =   "imlIco16x"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
         Begin ComctlLib.TreeView TvwData 
            Height          =   2295
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   4048
            _Version        =   327682
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imlIco16x"
            Appearance      =   1
         End
      End
   End
   Begin VB.PictureBox picSmallIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   9000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComDlg.CommonDialog cmDlg 
      Left            =   8880
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrBeginSearch 
      Enabled         =   0   'False
      Interval        =   760
      Left            =   8910
      Top             =   1080
   End
   Begin ComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   7200
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   503
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7938
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   503
            MinWidth        =   503
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList imlIco32x 
      Left            =   8835
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList imlIco16x 
      Left            =   8835
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuTrackFunctions 
      Caption         =   "mnuTrackFunctions"
      Visible         =   0   'False
      Begin VB.Menu mnuTrackFunctionsExplore 
         Caption         =   "Explore"
      End
      Begin VB.Menu mnuTrackFunctionsProperties 
         Caption         =   "View Properties"
      End
      Begin VB.Menu mnuTrackFunctionsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrackFunctionsExport 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuTrackFunctionsDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu mnuOpenOptions 
      Caption         =   "mnuOpenOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenOptionsOpenExisting 
         Caption         =   "Open an existing library..."
      End
      Begin VB.Menu mnuOpenOptionsCreate 
         Caption         =   "Create a new library..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const ROOT_KEY = "*"
Private Const SEARCH_HISTORY_SIZE = 20      '// Max maintained search history items
Private Const PATH_QUALIFIED = "PQC"

Private bIsBusy             As Boolean      '// flag to stop searching if the search string changes
Private vTrackList          As Boolean
Private LvDataSelItem       As String
Private LvDataSelPath       As String
Private LvContent           As Collection

'// Manages the status bar content
Public Sub SetSBStatus(ByVal sbPanel As Integer, ByVal Data As String, Optional ByVal sbStatus As Boolean = True)
    If sbPanel < 5 Then
        SB.Panels(sbPanel).Text = Data
    ElseIf sbPanel = 5 Then
        If sbStatus Then
            '// Set icons
        Else
            '// Set icons
        End If
    End If
End Sub

'// LvContent Handler--------------------------------

'// Search Handler   --------------------------------
Private Property Get BTSearchItem(vntIndexKey As Variant) As clsSearchItem
On Error Resume Next
    Set BTSearchItem = LvContent(vntIndexKey)
End Property

Private Property Get BTSearchCount() As Long
    BTSearchCount = LvContent.Count
End Property

'Private Sub BTSearchRemoveItem(vntIndexKey As String)
'On Error Resume Next
'    LvContent.Remove vntIndexKey
'End Sub

Private Sub BTSearchClear()
    Set LvContent = Nothing
    DoEvents
    Set LvContent = New Collection
End Sub

'// Track Handler   --------------------------------
Private Property Get BTTrackItem(vntIndexKey As Variant) As clsSearchItem
On Error Resume Next
    Set BTTrackItem = LvContent(vntIndexKey)
End Property

Private Property Get BTTrackCount() As Long
    BTTrackCount = LvContent.Count
End Property

'Private Sub BTTrackRemoveItem(vntIndexKey As String)
'On Error Resume Next
'    LvContent.Remove vntIndexKey
'End Sub

Private Sub BTTrackClear()
    Set LvContent = Nothing
    DoEvents
    Set LvContent = New Collection
End Sub
'// LvContent Handler--------------------------------


Private Sub AddSearchItem(cData As clsSearchItem)
Dim tLvI As ComctlLib.ListItem
Dim tmpStr As String
    With cData
        On Error Resume Next
            tmpStr = Replace$(.FileType, ChrW$(32), ChrW$(95))
            Set tLvI = LvwData.ListItems.Add(, , .FileName, , tmpStr)
            If Err Then
                '// Add the icon and retry
                Err.Clear
                Call GetIcon(.FileName, tmpStr, picSmallIcon, imlIco16x)
                Set tLvI = LvwData.ListItems.Add(, , .FileName, , tmpStr)
            End If
        On Error GoTo 0
        tLvI.SubItems(1) = .Size
        If (.Attributes And vbDirectory) = vbDirectory Then
            tLvI.SubItems(2) = "Directory"
        Else
            tLvI.SubItems(2) = .FileType
        End If
        tLvI.SubItems(3) = .TrackName
        tLvI.SubItems(4) = .Path
        tLvI.SubItems(5) = .Attributes
    End With
End Sub

Private Sub AddTrack(cData As clsTrackItem)
Dim tLvI As ComctlLib.ListItem
    With cData
        Set tLvI = LvwData.ListItems.Add(, , .TrackName, , 4)
        tLvI.SubItems(1) = .Category
        tLvI.SubItems(2) = .Created
        tLvI.SubItems(3) = .NumRecords
        tLvI.SubItems(4) = .TrackSize
        tLvI.SubItems(5) = .BasePath
        tLvI.SubItems(6) = .Comment
        tLvI.Tag = .BasePath
    End With
End Sub

'// Displays search results (From StartIndex) in the Listview
Public Sub LoadSearchResults(dispMax As Long)
Dim i As Long, uL As Long, j As Long
Dim File As String, Path As String, Size As String, Attributes As Long, hdn As Boolean, aName As String
Dim tLvI As ComctlLib.ListItem, typeKey As String, tmpStr As String
Dim TrackNode As ComctlLib.Node
Dim cTx As clsScanner, cSrchItem As clsSearchItem

    'cmdCancel.Move lblSearchResults.Left, lblSearchResults.Top, lblSearchResults.Width
    'cmdCancel.Visible = True
    'cmdCancel.ZOrder 0
    With frmWait
        .SetCaption ("Searching...Please stand by...")
        .Show vbModeless, Me
        .Refresh
    End With
    
    DoEvents
    
    LvwData.ListItems.Clear
    With LvwData.ColumnHeaders
        .Clear
        .Add , , "File"
        .Add , , "Size"
        .Add , , "Type"
        .Add , , "Track Name"
        .Add , , "Path"
        .Add , , "Attributes"
        
        '// Set the sorting sequences
        .Item(3).Tag = cscSortFormattedSize
        .Item(6).Tag = cscSortNumber
        DoEvents
    End With
        
        
    Call BTSearchClear
    Set cTx = New clsScanner
    '// Add the root category node
    TvwData.Nodes.Clear
    Call TvwData.Nodes.Add(, , ROOT_KEY, "ALL RESULTS", 5)
    
    bIsBusy = True
    vTrackList = False
    uL = clib.GetNumResults - 1
    
    If uL > 0 Then
        PB.Visible = True
        PB.Max = uL
        
        Call LockWindowUpdate(LvwData.hWnd)
        For i = 0 To uL
            If GetInputState() Then DoEvents
            If bIsBusy Then
                If i Mod REF_LIM = 0 Then
                    PB.Value = i
                    'Call SetSBStatus(3, i & " results found. Please stand by")
                    frmWait.SetCaption ("Searching...Please stand by..." & i & " matches found.")
                    frmWait.Refresh
                End If
    
                File = clib.GetRecordFromSearch(i, aName, Path, Size, Attributes, hdn)
                typeKey = cTx.FileType(File)
                tmpStr = UCase$(Replace$(typeKey, ChrW$(32), ChrW$(95)))
                Set cSrchItem = New clsSearchItem
                With cSrchItem
                    .Attributes = Attributes
                    .FileName = File
                    .FileType = typeKey
                    .Path = Path
                    .Size = Size
                    .TrackName = aName
                End With
                
                On Error Resume Next
                    Set TrackNode = TvwData.Nodes(aName)
                    If Err Then
                        Err.Clear
                        Set TrackNode = TvwData.Nodes.Add(ROOT_KEY, tvwChild, aName, aName, 5)
                    End If
                On Error GoTo 0
                j = CLng(Val(TrackNode.Tag)) + 1
                Call LvContent.Add(cSrchItem, aName & j)
                TrackNode.Tag = j
            
                Call AddSearchItem(cSrchItem)       '// Add one page of data
            Else
                Exit For
            End If
        Next
        Call LockWindowUpdate(0&)
        bIsBusy = False
        Call FixColumnWidths(LvwData.hWnd)
        PB.Visible = False
        
        '// Update the Lvw
        LvwData.Sorted = True
        LvwData.SortKey = 1
        Set cTx = Nothing
        
        With TvwData.Nodes
            For i = 1 To .Count
                Set TrackNode = .Item(i)
                If TrackNode.Tag <> vbNullString Then TrackNode.Text = TrackNode.Text & " (" & TrackNode.Tag & ")"
            Next
        
            With .Item(1)
                .Sorted = True
                .Expanded = True
                .Selected = True
                .EnsureVisible
            End With
        End With
    End If
    Unload frmWait
    cmdCancel.Visible = False
End Sub

'// This sub maintains a cyclic search history in the cmbSearchBox dropdown list
Private Sub AddToCyclicStack(ByVal vStr As String)
    With cmbSearchBox
        '// Is this string already in the list?
        If SendMessage(.hWnd, CB_FINDSTRINGEXACT, -1, ByVal vStr) = CB_ERR Then
            '// We need to add it to the list
            '// Are there SEARCH_HISTORY_SIZE items already in the list?
            If cmbSearchBox.ListCount >= SEARCH_HISTORY_SIZE Then
                Call cmbSearchBox.RemoveItem(SEARCH_HISTORY_SIZE - 1)
            End If
            '// Add this string (At the top of the list)
            Call cmbSearchBox.AddItem(vStr, 0)
            '// Save it
            Call cConfig.WriteSearchHistoryItem(vStr)
        End If
    End With
End Sub

'// Autosizes the column widths
Private Sub FixColumnWidths(ByVal lvHwnd As Long)
Dim i As Long
    For i = 0 To LvwData.ColumnHeaders.Count - 1
        Call SendMessage(lvHwnd, LVM_SETCOLUMNWIDTH, i, LVSCW_AUTOSIZE Or LVSCW_AUTOSIZE_USEHEADER)
    Next
End Sub

Private Function AddCategory(ByVal catName As String) As ComctlLib.Node
Dim currCat As String
Dim cL As Long
    catName = StripSlashes(catName, True)               '// Remove any ending slashes
    '// Does this category exist?
    On Error Resume Next
    Set AddCategory = TvwData.Nodes(catName)
    If Err Then
        On Error GoTo 0
        cL = InStrRev(catName, ChrW$(92))
        If cL > 0 Then
            currCat = Mid$(catName, cL + 1)             '// Get the category name
            '// Add it recursively to its parent
            Set AddCategory = TvwData.Nodes.Add(AddCategory(Left$(catName, cL - 1)), tvwChild, catName, currCat, 5)
        Else
            Set AddCategory = TvwData.Nodes.Add(ROOT_KEY, tvwChild, catName, catName, 5)
        End If
    End If
End Function

'// Loads archives present in the current library
Public Sub LoadTracks()
Dim archiveCnt As Long, i As Long, j As Long
Dim aCategory As String, aCreated As Date, aNumRec As Long, aComment As String, aName As String, aSize As String, aBaseNode As String
Dim tLvI As ComctlLib.ListItem, tTvN As ComctlLib.Node
Dim cTrData As clsTrackItem
    
    TvwData.Nodes.Clear
    '// Add the root category node
    Call TvwData.Nodes.Add(, , ROOT_KEY, "ALL TRACKS", 5)
    Call BTTrackClear
    With LvwData
        .ListItems.Clear
        
        With .ColumnHeaders
            .Clear
            .Add , , "Track Name"
            .Add , , "Category"
            .Add , , "Created"
            .Add , , "Records"
            .Add , , "Size"
            .Add , , "Base Path"
            .Add , , "Comment"
            
            '// Set the sorting sequences
            .Item(3).Tag = cscSortDate
            .Item(4).Tag = cscSortNumber
        End With

        '// Load the Tracks
        archiveCnt = clib.GetTrackCount
        vTrackList = True
        PB.Visible = True
        PB.Max = archiveCnt + 1
        For i = 1 To archiveCnt
            PB.Value = i
            aName = clib.GetTrackDetails(i, aCategory, aCreated, aNumRec, aComment, aSize, aBaseNode)
            
            '// Add the track to the collection
            Set cTrData = New clsTrackItem
            
            cTrData.BasePath = aBaseNode
            cTrData.Category = aCategory
            cTrData.Comment = Replace$(aComment, vbLf, Chr$(149))
            cTrData.Created = aCreated
            cTrData.NumRecords = aNumRec
            cTrData.TrackName = aName
            cTrData.TrackSize = aSize
            
            Set tTvN = AddCategory(aCategory)
            j = CLng(Val(tTvN.Tag) + 1)
            tTvN.Tag = j
            Call LvContent.Add(cTrData, tTvN.Key & j)
            
            '// And then add it to the list
            Call AddTrack(cTrData)
        Next
        
        Call FixColumnWidths(.hWnd)
        PB.Visible = False
        
        Call SetSBStatus(3, clib.CurrentDataBasePath)
        Call SetSBStatus(4, .ListItems.Count & " trx.")
    End With
    
    With TvwData.Nodes(1)
        .Sorted = True
        .Expanded = True
        .Selected = True
        .EnsureVisible
    End With
End Sub

Private Sub ExpandFolders(ByVal Track As String, ByVal Path As String)
Dim aName As String, aPath As String, aSize As String, aType As String, aAttr As Long, aHidden As Boolean
Dim i As Long, cL As Long, ubCache As Long, x As Long
Dim tmpStr As String, pathCache() As String, typeKey As String
Dim cScnr As clsScanner
Dim currDirNode As ComctlLib.Node, tmpNode As ComctlLib.Node

    cL = clib.GetRecordsByPath(Track, Path, grpMatchDirsOnly)
    If cL > 0 Then
        Screen.MousePointer = vbHourglass
        Set cScnr = New clsScanner
        ReDim pathCache(0)
        Set currDirNode = TvwBrowser.Nodes(Path)
        PB.Visible = True
        PB.Max = cL
        
        For i = 0 To cL - 1
            If GetInputState() Then DoEvents
            aName = clib.GetRecordFromSearch(i, tmpStr, aPath, aSize, aAttr, aHidden)
            'fName = StripSlashes(aName, True)
            '// Is it an archive or a directory
            '// This is a record we are interested in so add it and cache it
            ubCache = UBound(pathCache())
            typeKey = cScnr.FileType(Replace$(aName, ChrW$(92), vbNullString))
            typeKey = UCase$(Replace$(typeKey, ChrW$(32), ChrW$(95)))
            tmpStr = UCase$(AddBackSlash(aPath) & aName)

            On Error Resume Next
                TvwBrowser.Nodes.Add currDirNode, tvwChild, tmpStr, aName, typeKey
                If Err Then
                    On Error GoTo 0
                    '// Add the icon and retry
                    Err.Clear
                    Call GetIcon(aName, typeKey, picSmallIcon, TvwBrowser.ImageList)
                    TvwBrowser.Nodes.Add currDirNode, tvwChild, tmpStr, aName, typeKey
                End If
            ReDim Preserve pathCache(ubCache + 1)
            pathCache(ubCache) = tmpStr
            
            If i Mod REF_LIM = 0 Then
                TvwBrowser.Refresh
                PB.Value = i
                Call SetSBStatus(3, "Processing..." & i & " files examined")
            End If
        Next
        TvwBrowser.Refresh
        currDirNode.Sorted = True
        
        '// Now partially expand all cached Nodes
        For i = 0 To ubCache
            PB.Max = ubCache + 1
            PB.Value = i
            Call SetSBStatus(3, i + 1 & " sub-directories/archives examined")
            cL = clib.GetRecordsByPath(Track, pathCache(i), grpMatchDirsOnly)
            If cL > 0 Then
                PB.Max = cL
                For x = 0 To cL - 1
                    If GetInputState() Then DoEvents
                    PB.Value = x
                    aName = clib.GetRecordFromSearch(x, tmpStr, aPath, aSize, aAttr, aHidden)
                    '// It is an archive or a directory so add it and exit this loop
                    Set tmpNode = TvwBrowser.Nodes(pathCache(i))
                    tmpStr = UCase$(aPath & aName & ChrW$(92))
                    TvwBrowser.Nodes.Add tmpNode, tvwChild, tmpStr, aName
                    Exit For
                Next
            End If
        Next
        
        currDirNode.Tag = PATH_QUALIFIED
        Set cScnr = Nothing
        Screen.MousePointer = vbDefault
        PB.Visible = False
        SB.Panels(1).Text = "Ready"
    End If
End Sub

Private Sub ShowFilesInDir(ByVal Track As String, ByVal Path As String)
Dim aName As String, aPath As String, aSize As String, aType As String, aAttr As Long, aHidden As Boolean
Dim tmpStr As String, typeKey As String
Dim cL As Long, i As Long
Dim tmpLvi As ComctlLib.ListItem
Dim cScan As clsScanner
    
    LvwBrowser.ListItems.Clear
                
    cL = clib.GetRecordsByPath(Track, Path, grpMatchFilesOnly)
    If cL > 0 Then
        Screen.MousePointer = vbHourglass
        Set cScan = New clsScanner
        LvwBrowser.Tag = Path
        PB.Visible = True
        PB.Max = cL
        DoEvents
        
        Call SendMessage(LvwBrowser.hWnd, LVM_SETITEMCOUNT, cL, LVSICF_NOINVALIDATEALL Or LVSICF_NOSCROLL)
        LvwBrowser.Sorted = False
            For i = 0 To cL - 1
                If GetInputState() Then DoEvents
                If i Mod REF_LIM = 0 Then
                    PB.Value = i
                    Call SetSBStatus(3, "Displaying..." & i & " files found")
                    LvwBrowser.Refresh
                End If
                aName = clib.GetRecordFromSearch(i, tmpStr, aPath, aSize, aAttr, aHidden)
                aType = cScan.FileType(aName)
                typeKey = UCase$(Replace$(aType, ChrW$(32), ChrW$(95)))
                On Error Resume Next
                    Set tmpLvi = LvwBrowser.ListItems.Add(, , aName, , typeKey)
                    If Err Then
                        On Error GoTo 0
                        '// Add the icon and retry
                        Err.Clear
                        Call GetIcon(aName, typeKey, picSmallIcon, imlIco16x)
                        Set tmpLvi = LvwBrowser.ListItems.Add(, , aName, , typeKey)
                    End If
                tmpLvi.SubItems(1) = aSize
                If (aAttr And vbDirectory) = vbDirectory Then
                    tmpLvi.SubItems(2) = "Directory"
                Else
                    tmpLvi.SubItems(2) = aType
                End If
                tmpLvi.SubItems(3) = aAttr
            Next
        LvwBrowser.Sorted = True
        Call FixColumnWidths(LvwBrowser.hWnd)
        
        Call SetSBStatus(3, cL & " files. Not showing folders/directories and archives.")
        Screen.MousePointer = vbDefault
        Set cScan = Nothing
        PB.Visible = False
    End If
End Sub

Private Sub GetSelInfo(Track As String, ItemData As String, itemIndex As String)
Dim htInfo As HITTESTINFO, lItem As LV_ITEM, pT As POINTAPI
Dim Index As Long
    
    Call GetCursorPos(pT)
    Call ScreenToClient(LvwData.hWnd, pT)
    
    With htInfo
        LSet .pT = pT
        .flags = LVHT_ABOVE Or LVHT_BELOW Or _
                 LVHT_TOLEFT Or LVHT_TORIGHT Or _
                 LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_NOWHERE
    End With
    
    Index = SendMessage(LvwData.hWnd, LVM_SUBITEMHITTEST, 0, htInfo)
    
    If Index = -1 And (htInfo.iSubItem = -1 Or htInfo.iSubItem = 0) Then
        Track = vbNullString
    Else
        '// Get the track
        With lItem
            .mask = LVIF_TEXT
            .iSubItem = 0
            .cchTextMax = MAX_PATH
            .pszText = Space$(MAX_PATH)
        End With
        
        Call SendMessage(LvwData.hWnd, LVM_GETITEMTEXT, Index, lItem)
        Track = TrimNulls(lItem.pszText)
        
        '// Get its path
        With lItem
            .mask = LVIF_TEXT
            .iSubItem = itemIndex
            .cchTextMax = MAX_PATH
            .pszText = Space$(MAX_PATH)
        End With
        
        Call SendMessage(LvwData.hWnd, LVM_GETITEMTEXT, Index, lItem)
        ItemData = TrimNulls(lItem.pszText)
        
        Debug.Print Track, ItemData
    End If
End Sub

Private Sub Explore(ByVal Track As String, ByVal BasePath As String)
Dim newNode As ComctlLib.Node
    
    TvwBrowser.Nodes.Clear
    
    Set newNode = TvwBrowser.Nodes.Add(, , BasePath, Track, 4)
    Call ExpandFolders(Track, BasePath)
    
    newNode.Sorted = True
    newNode.Expanded = True
    cSplitterMain.Caption2 = "Track Browser - " & Track & " - [" & BasePath & "]"
End Sub

Private Sub cmdCancel_Click()
    bIsBusy = False
End Sub

Private Sub cSplitterExplorer_Panel1Resize(ByVal pLeft As Long, ByVal pTop As Long, ByVal pWidth As Long, ByVal pHeight As Long)
    TvwBrowser.Move pLeft, pTop, pWidth, pHeight
End Sub

Private Sub cSplitterExplorer_Panel2Resize(ByVal pLeft As Long, ByVal pTop As Long, ByVal pWidth As Long, ByVal pHeight As Long)
    LvwBrowser.Move pLeft, pTop, pWidth, pHeight
End Sub

Private Sub cSplitterMain_Panel1Resize(ByVal pLeft As Long, ByVal pTop As Long, ByVal pWidth As Long, ByVal pHeight As Long)
On Error Resume Next
    cSplitterNavigator.Move pLeft, pTop, pWidth, pHeight
End Sub

Private Sub cSplitterMain_Panel2Resize(ByVal pLeft As Long, ByVal pTop As Long, ByVal pWidth As Long, ByVal pHeight As Long)
On Error Resume Next
    cSplitterExplorer.Move pLeft, pTop, pWidth, pHeight
End Sub

Private Sub cSplitterNavigator_Panel1Resize(ByVal pLeft As Long, ByVal pTop As Long, ByVal pWidth As Long, ByVal pHeight As Long)
On Error Resume Next
    TvwData.Move pLeft, pTop, pWidth, pHeight
End Sub

Private Sub cSplitterNavigator_Panel2Resize(ByVal pLeft As Long, ByVal pTop As Long, ByVal pWidth As Long, ByVal pHeight As Long)
On Error Resume Next
    LvwData.Move pLeft, pTop, pWidth, pHeight
End Sub

Private Sub Form_Load()
Dim i As Long
Dim prevOpenTrack As String
    
    LvwBrowser.Font.Size = cConfig.FontSize
    LvwBrowser.Font.Name = cConfig.FontName
    LvwData.Font.Size = cConfig.FontSize
    LvwData.Font.Name = cConfig.FontName
    TvwBrowser.Font.Size = cConfig.FontSize
    TvwBrowser.Font.Name = cConfig.FontName
    TvwData.Font.Size = cConfig.FontSize
    TvwData.Font.Name = cConfig.FontName

    Call ObjectGridlines(LvwData.hWnd, cConfig.Gridlines)
    Call ObjectGridlines(LvwBrowser.hWnd, cConfig.Gridlines)
    
    Call SendMessage(TvwBrowser.hWnd, TVM_SETBKCOLOR, 0, ByVal Translate(cConfig.IBackColor))
    Call SendMessage(TvwData.hWnd, TVM_SETBKCOLOR, 0, ByVal Translate(cConfig.IBackColor))
    
    LvwBrowser.BackColor = cConfig.IBackColor
    LvwData.BackColor = cConfig.IBackColor
    picSmallIcon.BackColor = cConfig.IBackColor

    Call SendMessage(TvwBrowser.hWnd, TVM_SETTEXTCOLOR, 0, ByVal Translate(cConfig.IForeColor))
    Call SendMessage(TvwData.hWnd, TVM_SETTEXTCOLOR, 0, ByVal Translate(cConfig.IForeColor))
    LvwBrowser.ForeColor = cConfig.IForeColor
    LvwData.ForeColor = cConfig.IForeColor
    
    '// Load icons
    With imlIco32x
        .ImageHeight = 32
        .ImageWidth = 32
        .BackColor = cConfig.IBackColor
        
        With .ListImages
            For i = 1 To 7
                .Add , , LoadResPicture(32000 + i, vbResIcon)
            Next
        End With
    End With
    
    With imlIco16x
        .ImageHeight = 16
        .ImageWidth = 16
        .BackColor = cConfig.IBackColor
        
        With .ListImages
            For i = 1 To 7
                .Add , , LoadResPicture(16000 + i, vbResIcon)
            Next
        End With
        .ListImages.Item(5).Key = "FILE"
        Call picSmallIcon.Move(0, 0, .ImageWidth * Screen.TwipsPerPixelX, .ImageHeight * Screen.TwipsPerPixelY)
    End With
    
    Call MakeFlatToolbar(TB.hWnd)
    TB.Buttons(1).Image = 1
    TB.Buttons(3).Image = 2
    TB.Buttons(4).Image = 3
    TB.Buttons(6).Image = 4
    TB.Buttons(7).Image = 5
    TB.Buttons(8).Image = 6
    TB.Buttons(10).Image = 7
    
    Call FixTreeview(TvwBrowser.hWnd)
    Call FixTreeview(TvwData.hWnd)
    
    With LvwBrowser.ColumnHeaders
        .Clear
        .Add , , "File"
        .Add , , "Size"
        .Add , , "Type"
        .Add , , "Attributes"
        
        '// Set the sorting sequences
        .Item(2).Tag = cscSortFormattedSize
        .Item(4).Tag = cscSortNumber
    End With
    
    Call SendMessage(LvwData.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, ByVal LVS_EX_FULLROWSELECT, ByVal LVS_EX_FULLROWSELECT)
    
    Set imgSearchIcon.Picture = imlIco16x.ListImages(1).Picture
    
    With tbSearch
        .Buttons(4).Image = 2
        cmbSearchBox.Move .Buttons(3).Left
    End With
    
    Caption = APP_NAME & " - [" & App.Major & "." & App.Minor & "." & App.Revision & "]"
    
    Call LoadTracks
    
    '// Open the last track
    prevOpenTrack = cConfig.LastOpenTrack
    With LvwData.ListItems
        For i = 1 To .Count
            If .Item(i).Text = prevOpenTrack Then
                LvDataSelItem = prevOpenTrack
                LvDataSelPath = cConfig.LastOpenTrackPath
                Call mnuTrackFunctionsExplore_Click
                Exit For
            End If
        Next
    End With
    
    Call tbSearch_ButtonClick(tbSearch.Buttons(1))
End Sub

Private Sub Form_Resize()
On Error GoTo Form_Resize_ERH
    DoEvents
    pSearch.Left = ScaleWidth - pSearch.Width
    
    cSplitterMain.Move 0, TB.Height, ScaleWidth, ScaleHeight - (TB.Height + SB.Height)

    With SB
        .Move 0, ScaleHeight - .Height, ScaleWidth
        
        With .Panels(4)
            PB.Move .Left + 30, SB.Top + 30, .Width - 90
        End With
    End With
Form_Resize_ERH:
    If Err Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim xFrm As Form, i As Long
    '// Setup a splash screen
    With frmWait
        .SetCaption ("Compressing library. Please stand by.")
        .Show vbModeless, Me
        .Refresh
    End With
    
    With cConfig
        .MainWindowState = WindowState
        If WindowState <> vbMaximized Then Call .SetWindowRect(Left, Top, Width, Height)
        .LastActiveLibrary = clib.CurrentDataBasePath
        .HSplitterPosition = cSplitterMain.SplitterPosition
        .VSplitterExplorerPosition = cSplitterExplorer.SplitterPosition
        .VSplitterNavigatorPosition = cSplitterNavigator.SplitterPosition
        
        If TvwBrowser.Nodes.Count > 0 Then
            .LastOpenTrack = TvwBrowser.Nodes(1).Text
            .LastOpenTrackPath = TvwBrowser.Nodes(1).Key
        End If
        .WriteConfiguration
    End With
        
    '// Save the library
    Call clib.CloseLibrary
    
    Set LvContent = Nothing
    
    Call LockWindowUpdate(LvwData.hWnd)
        LvwData.ListItems.Clear
    Call LockWindowUpdate(0&)
    
    Call LockWindowUpdate(LvwBrowser.hWnd)
        LvwBrowser.ListItems.Clear
    Call LockWindowUpdate(0&)
    
    Call LockWindowUpdate(TvwBrowser.hWnd)
        TvwBrowser.Nodes.Clear
    Call LockWindowUpdate(0&)
    
    Call LockWindowUpdate(TvwData.hWnd)
        TvwData.Nodes.Clear
    Call LockWindowUpdate(0&)
    
    '// Clean up
    Set clib = Nothing
    Set cConfig = Nothing
    
    For Each xFrm In Forms
        Unload xFrm
    Next
End Sub

Private Sub LvwBrowser_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    LvwBrowser.Sorted = False
    If ColumnHeader.Index - 1 = LvwBrowser.SortKey Then
        If LvwBrowser.SortOrder = lvwAscending Then
            LvwBrowser.SortOrder = lvwDescending
        Else
            LvwBrowser.SortOrder = lvwAscending
        End If
    Else
        LvwBrowser.SortKey = ColumnHeader.Index - 1
        LvwBrowser.SortOrder = lvwAscending
    End If
    
    If ColumnHeader.Tag <> vbNullString Then
        LvwBrowser.Sorted = False
        Call SortColumn(LvwBrowser.hWnd, ColumnHeader.Index - 1, ColumnHeader.Tag, LvwBrowser.SortOrder)
    Else
        LvwBrowser.Sorted = True
    End If
End Sub

Private Sub LvwData_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    LvwData.Sorted = False
    If ColumnHeader.Index - 1 = LvwData.SortKey Then
        If LvwData.SortOrder = lvwAscending Then
            LvwData.SortOrder = lvwDescending
        Else
            LvwData.SortOrder = lvwAscending
        End If
    Else
        LvwData.SortKey = ColumnHeader.Index - 1
        LvwData.SortOrder = lvwAscending
    End If
    
    If ColumnHeader.Tag <> vbNullString Then
        LvwData.Sorted = False
        Call SortColumn(LvwData.hWnd, ColumnHeader.Index - 1, ColumnHeader.Tag, LvwData.SortOrder)
    Else
        LvwData.Sorted = True
    End If
End Sub

Private Sub LvwData_DblClick()
    Call GetSelInfo(LvDataSelItem, LvDataSelPath, 5)
    Call mnuTrackFunctionsExplore_Click
End Sub

Private Sub LvwData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And (Shift And vbShiftMask) > 0 Then
        Call mnuTrackFunctionsDelete_Click
    End If
End Sub

Private Sub LvwData_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '// Get details about the selected item
    Call GetSelInfo(LvDataSelItem, LvDataSelPath, 5)
        
    If (Button = vbRightButton) And (vTrackList) Then
        If LvDataSelItem <> vbNullString Then
            mnuTrackFunctionsProperties.Caption = "Properties " & LvDataSelItem & "..."
            mnuTrackFunctionsExplore.Caption = "Browse " & LvDataSelItem
            PopupMenu mnuTrackFunctions
        End If
    End If
End Sub

Private Sub mnuOpenOptionsCreate_Click()
    With cmDlg
        .CancelError = False
        .DialogTitle = "Select new " & APP_NAME & " library"
        .Filter = APP_NAME & " library Files (*.blf)|*.blf|Extensibe Markup Language(XML) Files (*.xml)|*.xml"
        .ShowSave

        If .FileName <> vbNullString Then
            With frmWait
                .SetCaption ("Opening Library...")
                .Show
                .Refresh
            End With
            Call clib.CloseLibrary
            DoEvents
            Call clib.CreateNewLibrary(.FileName)
            Call clib.OpenFileAsLibrary(.FileName)
            DoEvents
            Call LoadTracks
            Unload frmWait
        End If
    End With
End Sub

Private Sub mnuOpenOptionsOpenExisting_Click()
    With cmDlg
        .CancelError = False
        .DialogTitle = "Select new " & APP_NAME & " library"
        .Filter = APP_NAME & " library Files (*.blf)|*.blf|Extensibe Markup Language(XML) Files (*.xml)|*.xml"
        .ShowOpen

        If .FileName <> vbNullString Then
            With frmWait
                .SetCaption ("Opening Library...")
                .Show
                .Refresh
            End With
            Call clib.CloseLibrary
            DoEvents
            Call clib.OpenFileAsLibrary(.FileName)
            DoEvents
            Call LoadTracks
            Unload frmWait
        End If
    End With
End Sub

Private Sub mnuTrackFunctionsDelete_Click()
    If (LvDataSelItem <> vbNullString) And vTrackList Then
        If MsgBox("Are you sure you want to delete the track " & ChrW$(39) & LvDataSelItem & ChrW$(39) & ChrW$(63), vbQuestion Or vbYesNoCancel, "Confirm Delete") = vbYes Then
            If clib.DeleteTrack(LvDataSelItem) Then
                Call LoadTracks
                MsgBox "You need to restart " & APP_NAME & " to make the changes permanent.", vbInformation Or vbOKOnly, "Track Deleted"
            Else
                MsgBox "An error occured. The track cannot be deleted.", vbCritical Or vbOKOnly, "Error"
            End If
        End If
    End If
End Sub

Private Sub mnuTrackFunctionsExplore_Click()
    If (LvDataSelItem <> vbNullString) And vTrackList Then
        LvwBrowser.ListItems.Clear
        Call Explore(LvDataSelItem, LvDataSelPath)
    End If
End Sub

Private Sub mnuTrackFunctionsProperties_Click()
    If (LvDataSelItem <> vbNullString) And vTrackList Then
        If frmProperties.Visible = False Then
            Load frmProperties
            frmProperties.Show vbModeless, Me
        End If
        Call frmProperties.SetProperties(LvDataSelItem)
    End If
End Sub

Private Sub TB_ButtonClick(ByVal Button As ComctlLib.Button)
Dim tmpStr As String
    Select Case Button.Index
        Case 1
            'PopupMenu mnuAddSelection, , Button.Left, Button.Top + Button.Height + 15
            'Button.Value = tbrUnpressed
            Load frmAddNew
            frmAddNew.Show vbModal, Me
            Call LoadTracks
        
        Case 3
            PopupMenu mnuOpenOptions, , Button.Left, Button.Top + Button.Height + 15
            Button.Value = tbrUnpressed
            
        
        Case 4
            With cmDlg
                .CancelError = False
                .DialogTitle = "Save current library..."
                .Filter = APP_NAME & " library Files (*.blf)|*.blf|Extensibe Markup Language(XML) Files (*.xml)|*.xml"
                .ShowSave

                If .FileName <> vbNullString Then
                    With frmWait
                        .SetCaption ("Opening Library...")
                        .Show
                        .Refresh
                    End With
                    Call clib.SaveLibraryAsFile(.FileName)
                    DoEvents
                    Unload frmWait
                End If
            End With
            
        Case 6
            frmOptions.Show vbModal, Me
            
        Case 7
            frmAbout.Show vbModal, Me
            
        Case 9
            Unload Me
    End Select
End Sub

Private Sub tbSearch_ButtonClick(ByVal Button As ComctlLib.Button)
    If Button.Index = 1 Then
        With Button
            If TvwBrowser.Nodes.Count > 0 Then
                If .Value = tbrPressed Then
                    .Image = 3
                    .ToolTipText = "Search in all tracks"
                Else
                    .Image = 4
                    .ToolTipText = "Search only in Track - " & TvwBrowser.Nodes(1).Text
                End If
            Else
                .Value = tbrPressed
                .Image = 3
                .ToolTipText = "Search in all tracks"
                MsgBox "No track is currently being explored. " & APP_NAME & " will search in all tracks", vbOKOnly Or vbExclamation, "Cannot perform selected operation"
            End If
            If cmbSearchBox.Text <> vbNullString Then
                tmrBeginSearch.Enabled = True
            End If
        End With
    ElseIf Button.Index = 4 Then
        Call TvwData_NodeClick(TvwData.Nodes(1))
    End If
End Sub

Private Sub tmrBeginSearch_Timer()
Dim numHits As Long, srchStr As String
Dim lvwCpp As Long, lDatamax As Long

    srchStr = cmbSearchBox.Text
    If (srchStr <> vbNullString) And (srchStr <> SEARCH_NOT) And (srchStr <> SEARCH_DIR) And (srchStr <> SEARCH_FIL) Then
        If tbSearch.Buttons(1).Value = tbrPressed Then
            If (Left$(srchStr, 2) = SEARCH_NOT) Then
                numHits = clib.FindInLibrary(scDoesNotContain, Mid$(UCase$(srchStr), 3))
            ElseIf (Left$(srchStr, 2) = SEARCH_DIR) Then
                numHits = clib.FindInLibrary(scContains, Mid$(UCase$(srchStr), 3), False, sfcSearchFilePath)
            ElseIf (Left$(srchStr, 2) = SEARCH_FIL) Then
                numHits = clib.FindInLibrary(scContains, Mid$(UCase$(srchStr), 3), False, sfcSearchFileName)
            Else
                numHits = clib.FindInLibrary(scContains, UCase$(srchStr))
            End If
        Else
            If (Left$(srchStr, 2) = SEARCH_NOT) Then
                numHits = clib.FindInTrack(TvwBrowser.Nodes(1).Text, scDoesNotContain, Mid$(UCase$(srchStr), 3))
            ElseIf (Left$(srchStr, 2) = SEARCH_DIR) Then
                numHits = clib.FindInTrack(TvwBrowser.Nodes(1).Text, scContains, Mid$(UCase$(srchStr), 3), False, sfcSearchFilePath)
            ElseIf (Left$(srchStr, 2) = SEARCH_FIL) Then
                numHits = clib.FindInTrack(TvwBrowser.Nodes(1).Text, scContains, Mid$(UCase$(srchStr), 3), False, sfcSearchFileName)
            Else
                numHits = clib.FindInTrack(TvwBrowser.Nodes(1).Text, scContains, UCase$(srchStr))
            End If
        End If
        Call SetSBStatus(4, ChrW$(60) & numHits)
        lblSearchResults.Caption = numHits & " matches (approx.)"
        If numHits > 0 Then
            If numHits > 1500 Then
                If MsgBox("Atleast " & numHits & " possible matches have been found for your search. Display the results?", vbQuestion Or vbYesNoCancel, "Confirm") <> vbYes Then
                    tmrBeginSearch.Enabled = False
                    Exit Sub
                End If
            End If
            '// Load the results
            With LvwData
                '// Show only one screenful of results (We subtract 1 for the bottom scrollbar)
                lvwCpp = SendMessage(.hWnd, LVM_GETCOUNTPERPAGE, 0, 0)
                
                If numHits > lvwCpp Then
                    lDatamax = lvwCpp
                    With tbSearch.Buttons("tbSearchShowAll")
                        .Enabled = True
                        .ToolTipText = "Click here to show all " & numHits & " results."
                    End With
                    Debug.Print "Showing only " & lDatamax & " items."
                Else
                    lDatamax = numHits
                    With tbSearch.Buttons("tbSearchShowAll")
                        .Enabled = False
                        .ToolTipText = vbNullString
                    End With
                End If
                
                '// Show the results
                Call LoadSearchResults(lDatamax)
                Call SetSBStatus(3, numHits & " matches found. Showing " & lDatamax & " of " & numHits)
            End With
        Else
            LvwData.ListItems.Clear
        End If
    End If
        
    tmrBeginSearch.Enabled = False
End Sub

Private Sub cmbSearchBox_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call tbSearch_ButtonClick(tbSearch.Buttons(4))
End Sub

Private Sub cmbSearchBox_LostFocus()
Dim tmpStr As String
    tmpStr = cmbSearchBox.Text
    If tmpStr <> vbNullString Then
        AddToCyclicStack (tmpStr)
    End If
End Sub

Private Sub cmbSearchBox_Click()
    Call cmbSearchBox_Change
End Sub

Private Sub cmbSearchBox_Change()
Dim tmpStr As String
    If cmbSearchBox.Text = vbNullString Then
        Call SetSBStatus(4, vbNullString)
        lblSearchResults.Caption = vbNullString
        Call LoadTracks
    End If
    bIsBusy = False
    tmrBeginSearch.Enabled = True
End Sub

Private Sub cmbSearchBox_Validate(Cancel As Boolean)
    tmrBeginSearch.Enabled = False
End Sub

Private Sub TvwBrowser_Expand(ByVal Node As ComctlLib.Node)
    If Node.Tag <> PATH_QUALIFIED Then
        Call TvwBrowser.Nodes.Remove(Node.Child.Key)
        Call ExpandFolders(TvwBrowser.Nodes(1).Text, Node.Key)
    End If
End Sub

Private Sub TvwBrowser_NodeClick(ByVal Node As ComctlLib.Node)
    If LvwBrowser.Tag <> Node.Key Then
        LvwBrowser.ListItems.Clear
        Call ShowFilesInDir(TvwBrowser.Nodes(1).Text, Node.Key)
        Node.Expanded = True
    End If
End Sub
                
Private Sub TvwData_NodeClick(ByVal Node As ComctlLib.Node)
Dim i As Long, uL As Long
Dim cTrData As clsTrackItem, cSrData As clsSearchItem

    If vTrackList Then
        LvwData.ListItems.Clear
        If Not (LvContent Is Nothing) Then
            PB.Visible = True
            If Node.Key = ROOT_KEY Then
                '// Display all results
                uL = BTTrackCount
                If uL > 0 Then
                    PB.Max = uL
                    For i = 1 To uL
                        If i Mod REF_LIM = 0 Then
                            PB.Value = i
                            Call SetSBStatus(3, i & " tracks found. Please stand by...")
                        End If
                        Set cTrData = LvContent.Item(i)
                        Call AddTrack(cTrData)
                    Next
                End If
            Else
                '// Display select results
                uL = Val(Node.Tag)
                If uL > 0 Then
                    PB.Max = uL
                    For i = 1 To uL
                        If i Mod REF_LIM = 0 Then
                            PB.Value = i
                            Call SetSBStatus(3, i & " tracks found. Please stand by...")
                        End If
                        Set cTrData = LvContent.Item(Node.Key & i)
                        Call AddTrack(cTrData)
                    Next
                End If
            End If
            PB.Visible = False
            Call SetSBStatus(4, LvwData.ListItems.Count & " trx.")
        End If
    Else
        LvwData.ListItems.Clear
        If Not (LvContent Is Nothing) Then
            PB.Visible = True
            If Node.Key = ROOT_KEY Then
                uL = BTSearchCount
                If uL > 0 Then
                    PB.Max = uL
                    For i = 1 To uL
                        If i Mod REF_LIM = 0 Then
                            PB.Value = i
                            Call SetSBStatus(3, i & " results found. Please wait...")
                        End If
                        Set cSrData = LvContent.Item(i)
                        Call AddSearchItem(cSrData)
                    Next
                End If
            Else
                uL = Val(Node.Tag)
                If uL > 0 Then
                    PB.Max = uL
                    For i = 1 To uL
                        If i Mod REF_LIM = 0 Then
                            PB.Value = i
                            Call SetSBStatus(3, i & " results found. Please wait...")
                        End If
                        Set cSrData = LvContent.Item(Node.Key & i)
                        Call AddSearchItem(cSrData)
                    Next
                End If
            End If
            PB.Visible = False
        End If
    End If
    Node.Expanded = True
    Call FixColumnWidths(LvwData.hWnd)
    Call SetSBStatus(3, "Ready")
End Sub
