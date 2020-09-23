Attribute VB_Name = "mod0DataJunkey"
Option Explicit

Private Const vMSXML6 = "Msxml2.DOMDocument.6.0"
Private Const vMSXML5 = "Msxml2.DOMDocument.5.0"
Private Const vMSXML4 = "Msxml2.DOMDocument.4.0"
Private Const vMSXML3 = "MSXML2.DOMDocument.3.0"

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_STATUSTEXT = &H4

Private Const DLL_UNRAR = "unrar.dll"
Private Const DLL_UNZIP = "Unzip32.dll"
Private Const DLL_ZLIB = "zlib.dll"


Private Type BROWSEINFO
   hOwner As Long
   pidlRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type

Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
Private Declare Function InitCommonControls Lib "comctl32" () As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (ByRef lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByRef pidl As Long, ByVal pszPath As String) As Long

Private shInfo      As SHFILEINFO                           '// Used to get the icon

Public Sub GetIcon(ByVal File As String, ByVal typeKey As String, ByRef PB As VB.PictureBox, ByRef iml As ComctlLib.ImageList)
Dim hIcon  As Long

    shInfo.iIcon = -1
    
    PB.Cls
    'Get a handle to the small icon
    hIcon = SHGetFileInfo(File & vbNullChar, 0, shInfo, Len(shInfo), SHGFI_SMALLICON Or SHGFI_SYSICONINDEX Or SHGFI_USEFILEATTRIBUTES)

    '// Draw Small Icon
    Call ImageList_Draw(hIcon, shInfo.iIcon, PB.hdc, 0, 0, ILD_TRANSPARENT)
    Call iml.ListImages.Add(, typeKey, PB.Image)
End Sub

Public Function MakeFlatToolbar(ByVal tbHwnd As Long) As Long
Const TOOLBARCLASS = "ToolbarWindow32"
Dim bStyle As Long, hCtl As Long

    '// Get the handle of the toolbar
    hCtl = FindWindowEx(tbHwnd, 0&, TOOLBARCLASS, vbNullString)
    
    '// Get its styles
    bStyle = SendMessage(hCtl, TB_GETSTYLE, 0&, ByVal 0&)

    '// Set the new style to the toolbar
    If (bStyle And TBSTYLE_FLAT) <> TBSTYLE_FLAT Then
        bStyle = bStyle Or (TBSTYLE_FLAT)
    End If
    MakeFlatToolbar = SendMessage(hCtl, TB_SETSTYLE, 0, ByVal bStyle)
    DoEvents
End Function

Public Function FixTreeview(ByVal tvhWnd As Long) As Long
Dim bStyle As Long
    bStyle = GetWindowLong(tvhWnd, GWL_STYLE)
    bStyle = bStyle Or TVS_TRACKSELECT
    FixTreeview = SetWindowLong(tvhWnd, GWL_STYLE, bStyle)
    DoEvents
End Function

Public Function TrimNulls(ByVal str As String) As String
Dim nLen As Long
    nLen = InStr(str, vbNullChar)
    If nLen > 1 Then
        TrimNulls = Left$(str, nLen - 1)
    Else
        TrimNulls = str
    End If
End Function

Public Function StripSlashes(ByVal DirStr As String, Optional StripEndSlash As Boolean = False) As String
    If Asc(DirStr) = 92 Then DirStr = Mid$(DirStr, 2)
    If StripEndSlash Then _
        If Asc(Mid$(DirStr, Len(DirStr))) = 92 Then DirStr = Left$(DirStr, Len(DirStr) - 1)
    StripSlashes = DirStr
End Function

Public Function BrowsePath(hWnd As Long, Optional sTitle As String = "Select Folder") As String
Dim bInfo As BROWSEINFO
Dim pidl As Long
Dim rPath As String

    With bInfo
        .hOwner = hWnd
        .pidlRoot = 0&
        .lpszTitle = sTitle
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT
    End With
    
    pidl = SHBrowseForFolder(bInfo)
    rPath = Space$(MAX_PATH)
    
    If SHGetPathFromIDList(ByVal pidl, ByVal rPath) Then _
        BrowsePath = Left$(rPath, InStr(rPath, vbNullChar) - 1)
    
    Call CoTaskMemFree(pidl)
End Function

Sub Main()
Dim bData() As Byte
Dim frf As Integer
Dim wLeft As Long, wTop As Long, wWidth As Long, wHt As Long
Dim objDummy As Object, bMSXMLFound As Boolean, MSXMLver As String
Dim tmpStr As String
    
    Call InitCommonControls
    '// Setup a splash screen
    With frmWait
        .SetCaption ("Loading Files...")
        .Show
        .Refresh
    End With
            
    '// Unpack the DLL's we require
    '// Unrar
    If Dir$(App.Path & Chr$(92) & DLL_UNRAR, vbNormal) = vbNullString Then
        '// Unpack the DLL
        bData = LoadResData(101, "BINARY")
        
        '// Create a local instance of it
        frf = FreeFile()
        Open App.Path & Chr$(92) & DLL_UNRAR For Binary As #frf
            Put #frf, , bData
        Close #frf
    End If
    
    '// Unzip
    If Dir$(App.Path & Chr$(92) & DLL_UNZIP, vbNormal) = vbNullString Then
        '// Unpack the DLL
        bData = LoadResData(102, "BINARY")
        
        '// Create a local instance of it
        frf = FreeFile()
        Open App.Path & Chr$(92) & DLL_UNZIP For Binary As #frf
            Put #frf, , bData
        Close #frf
    End If
    
    '// zlib
    If Dir$(App.Path & Chr$(92) & DLL_ZLIB, vbNormal) = vbNullString Then
        '// Unpack the DLL
        bData = LoadResData(103, "BINARY")
        
        '// Create a local instance of it
        frf = FreeFile()
        Open App.Path & Chr$(92) & DLL_ZLIB For Binary As #frf
            Put #frf, , bData
        Close #frf
    End If
    
    With frmWait
        .SetCaption ("Examining System...")
        .Refresh
    End With
    
    '// Determine if MSXML3.0+ is installed
    On Error Resume Next
        Set objDummy = CreateObject(vMSXML6)
        If Err Then
            bMSXMLFound = False
            Err.Clear
            Set objDummy = CreateObject(vMSXML5)
            If Err Then
                bMSXMLFound = bMSXMLFound Or False
                Err.Clear
                Set objDummy = CreateObject(vMSXML4)
                If Err Then
                    bMSXMLFound = bMSXMLFound Or False
                    Err.Clear
                    Set objDummy = CreateObject(vMSXML3)
                    If Err Then
                        bMSXMLFound = bMSXMLFound Or False
                        Err.Clear
                    Else
                        bMSXMLFound = True
                        MSXMLver = "MSXML v3.0"
                    End If
                Else
                    bMSXMLFound = True
                    MSXMLver = "MSXML v4.0"
                End If
            Else
                bMSXMLFound = True
                MSXMLver = "MSXML v5.0"
            End If
        Else
            bMSXMLFound = True
            MSXMLver = "MSXML v6.0"
        End If
        
        If bMSXMLFound = False Then
            Unload frmWait
            MsgBox "MSXML3.0 or greater is needed to run " & APP_NAME & ". Application will now terminate as MSXML3.0 was not found on this system", vbCritical Or vbOKOnly, "Fatal Error"
        Else
            '// Set the Current Workspace
            ChDir App.Path
            ChDrive App.Path
            
            With frmWait
                .SetCaption ("Reading preferences...")
                .Refresh
            End With
            
            Set cConfig = New clsConfig
            Call cConfig.ReadConfiguration
            
            With frmWait
                .SetCaption ("Opening Library...")
                .Refresh
            End With
            Set clib = New clsTrackLibrary
            
            If Not clib.OpenFileAsLibrary(cConfig.LastActiveLibrary) Then
                Call clib.OpenFileAsLibrary(cConfig.CreateLibrary)
                
                MsgBox "An error occured while opening the library: " & cConfig.LastActiveLibrary & vbCrLf & _
                       "A new library was created.", vbCritical Or vbOKOnly, "Library access error"
            End If
            
            
            With frmWait
                .SetCaption ("Loading main interface...")
                .Refresh
            End With
            
            Load frmMain
            With frmMain
                With frmWait
                    .SetCaption ("Reading preferences...")
                    .Refresh
                End With
                
                .WindowState = cConfig.MainWindowState
                Call cConfig.GetWindowRect(wLeft, wTop, wWidth, wHt)
                .Top = wTop
                .Left = wLeft
                .Width = wWidth
                .Height = wHt
                
                .cSplitterMain.HasCaptions = cConfig.InterfaceCaptions
                
                .cSplitterMain.SplitterPosition = cConfig.HSplitterPosition
                .cSplitterExplorer.SplitterPosition = cConfig.VSplitterExplorerPosition
                .cSplitterNavigator.SplitterPosition = cConfig.VSplitterNavigatorPosition
                
                If cConfig.DisplayDbInTitle Then
                    .Caption = .Caption & "  [" & clib.CurrentDataBasePath & "]"
                End If
                
                For frf = 1 To .TB.Buttons.Count
                    If (.TB.Buttons(frf).Style = tbrDefault) Or (.TB.Buttons(frf).Style = tbrCheck) Then
                        If Not cConfig.ToolbarCaptions Then .TB.Buttons(frf).Caption = vbNullString
                    End If
                Next
                
                For frf = 1 To cConfig.GetSearchHistoryCount
                    Call .cmbSearchBox.AddItem(cConfig.ReadSearchHistoryItem(frf))
                Next
                
                .Show
            End With
            Call frmMain.SetSBStatus(2, MSXMLver)
            
            With frmWait
                .SetCaption ("Done...")
                .Refresh
            End With
            Unload frmWait
        End If
    On Error GoTo 0
End Sub

Public Function FnPtr(ByVal lp As Long) As Long
    FnPtr = lp
End Function

Public Function AddBackSlash(ByVal sVal As String) As String
    If sVal = vbNullString Then
        sVal = ChrW$(92)
    ElseIf Not (Asc(Mid$(sVal, Len(sVal))) = 92) Then
        sVal = sVal & ChrW$(92)
    End If
    AddBackSlash = sVal
End Function

Public Function ObjectGridlines(ByVal hWnd As Long, ByVal showGridlines As Boolean)
Dim lStyle As Long
    lStyle = SendMessage(hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, ByVal 0&)
    
    If showGridlines And ((lStyle And LVS_EX_GRIDLINES) <> LVS_EX_GRIDLINES) Then
        Call SendMessage(hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal lStyle Or LVS_EX_GRIDLINES)
    ElseIf (lStyle And LVS_EX_GRIDLINES) = LVS_EX_GRIDLINES Then
        Call SendMessage(hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, ByVal lStyle Xor LVS_EX_GRIDLINES)
    End If
End Function

Public Function Translate(ByVal color As OLE_COLOR) As Long
    Call OleTranslateColor(color, 0, Translate)
End Function
