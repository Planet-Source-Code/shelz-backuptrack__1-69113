VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const MAX_PATH = 260
Private Const INI_FILE = "BTCFG.INI"
Private Const DEF_WINDOW_WH = 8000
Private Const DEF_SPLITTER_POS = 2480

Private m_ConfigFile                    As String
Private m_MainWindowState               As Integer
Private m_LastActiveLibrary             As String
Private m_LastOpenTrack                 As String
Private m_LastOpenTrackPath             As String
Private m_MainWindowLeft                As Long
Private m_MainWindowTop                 As Long
Private m_MainWindowWidth               As Long
Private m_MainWindowHeight              As Long
Private m_HSplitterPosition             As Long
Private m_VSplitter1Position            As Long
Private m_VSplitter2Position            As Long
Private m_DisplayDbInTitle              As Boolean
Private m_Gridlines                     As Boolean
Private m_ToolbarCaptions               As Boolean
Private m_IntefaceHasCaptions           As Boolean
Private m_FontName                      As String
Private m_FontSize                      As Long
Private m_BackColor                     As Long
Private m_ForeColor                     As Long
Private m_SearchHistory()               As String


Private Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function CreateLibrary() As String
Dim tmpStr As String
Dim cLibCreator As New clsTrackLibrary
    '// Create a library
    tmpStr = App.Path & "\BLFIB" & Format(Now, "yydm-hms") & ".BLF"
    Call cLibCreator.CreateNewLibrary(tmpStr)
    CreateLibrary = tmpStr
    
    Set cLibCreator = Nothing
End Function

Public Function ReadSearchHistoryItem(ByVal Index As Long) As String
    If Index <= UBound(m_SearchHistory()) Then ReadSearchHistoryItem = m_SearchHistory(Index)
End Function

Public Sub WriteSearchHistoryItem(ByVal ItemData As String)
    ReDim Preserve m_SearchHistory(UBound(m_SearchHistory) + 1)
    m_SearchHistory(UBound(m_SearchHistory)) = ItemData
End Sub

Public Function GetSearchHistoryCount() As Long
    GetSearchHistoryCount = UBound(m_SearchHistory())
End Function

Public Sub GetWindowRect(wLeft As Long, wTop As Long, wWidth As Long, wHeight As Long)
    wLeft = m_MainWindowLeft
    wTop = m_MainWindowTop
    wWidth = m_MainWindowWidth
    wHeight = m_MainWindowHeight
End Sub

Public Sub SetWindowRect(ByVal wLeft As Long, ByVal wTop As Long, ByVal wWidth As Long, ByVal wHeight As Long)
    m_MainWindowLeft = wLeft
    m_MainWindowTop = wTop
    m_MainWindowWidth = wWidth
    m_MainWindowHeight = wHeight
End Sub

Public Sub WriteConfiguration()
Dim frf As Long

    If Dir$(m_ConfigFile) = vbNullString Then
        '// Create the INI file
        frf = FreeFile()
        Open m_ConfigFile For Output As #frf
        Close #frf
        DoEvents
    End If
    
    '// Save the WindowState
    Call WritePrivateProfileString("WINDOW", "STATE", CStr(m_MainWindowState), m_ConfigFile)
    '// Save the window orientation
    Call WritePrivateProfileString("WINDOW", "LEFT", CStr(m_MainWindowLeft), m_ConfigFile)
    Call WritePrivateProfileString("WINDOW", "TOP", CStr(m_MainWindowTop), m_ConfigFile)
    Call WritePrivateProfileString("WINDOW", "WIDTH", CStr(m_MainWindowWidth), m_ConfigFile)
    Call WritePrivateProfileString("WINDOW", "HEIGHT", CStr(m_MainWindowHeight), m_ConfigFile)
    '// Splitter
    Call WritePrivateProfileString("WINDOW", "SPLIT-H", CStr(m_HSplitterPosition), m_ConfigFile)
    Call WritePrivateProfileString("WINDOW", "SPLIT-V1", CStr(m_VSplitter1Position), m_ConfigFile)
    Call WritePrivateProfileString("WINDOW", "SPLIT-V2", CStr(m_VSplitter2Position), m_ConfigFile)
    '// Save the Last Active Database
    Call WritePrivateProfileString("DATABASE", "LASTACTIVE", m_LastActiveLibrary, m_ConfigFile)
    '// Last open track
    Call WritePrivateProfileString("DATABASE", "LASTACTIVETRACK", m_LastOpenTrack, m_ConfigFile)
    '// Last open track path
    Call WritePrivateProfileString("DATABASE", "LASTACTIVETRACKPATH", m_LastOpenTrackPath, m_ConfigFile)
    '// Interface
    Call WritePrivateProfileString("INTERFACE", "DBINCAPTION", CStr(CLng(m_DisplayDbInTitle)), m_ConfigFile)
    Call WritePrivateProfileString("INTERFACE", "GRIDLINES", CStr(CLng(m_Gridlines)), m_ConfigFile)
    Call WritePrivateProfileString("INTERFACE", "TOOLBARCAPTIONS", CStr(CLng(m_ToolbarCaptions)), m_ConfigFile)
    Call WritePrivateProfileString("INTERFACE", "INTERFACECAPTIONS", CStr(CLng(m_IntefaceHasCaptions)), m_ConfigFile)
    Call WritePrivateProfileString("INTERFACE", "FONTSIZE", CStr(m_FontSize), m_ConfigFile)
    Call WritePrivateProfileString("INTERFACE", "BACKCOLOR", CStr(m_BackColor), m_ConfigFile)
    Call WritePrivateProfileString("INTERFACE", "FORECOLOR", CStr(m_ForeColor), m_ConfigFile)
    Call WritePrivateProfileString("INTERFACE", "FONTNAME", m_FontName, m_ConfigFile)
    
    '// Save the Search History
    Call WritePrivateProfileString("SEARCH", "COUNT", CStr(UBound(m_SearchHistory)), m_ConfigFile)
    For frf = 1 To UBound(m_SearchHistory)
        Call WritePrivateProfileString("SEARCH", "ITEM" & frf, m_SearchHistory(frf), m_ConfigFile)
    Next
End Sub

Public Sub ReadConfiguration()
Dim buff As String, i As Long, sMax As Long, lRet As Long

    If Dir$(m_ConfigFile) <> vbNullString Then
        '// Read the WindowState
        m_MainWindowState = GetPrivateProfileInt("WINDOW", "STATE", 0, m_ConfigFile)
        '// Read the window orientation
        m_MainWindowLeft = GetPrivateProfileInt("WINDOW", "LEFT", 0, m_ConfigFile)
        m_MainWindowTop = GetPrivateProfileInt("WINDOW", "TOP", 0, m_ConfigFile)
        m_MainWindowWidth = GetPrivateProfileInt("WINDOW", "WIDTH", DEF_WINDOW_WH, m_ConfigFile)
        m_MainWindowHeight = GetPrivateProfileInt("WINDOW", "HEIGHT", DEF_WINDOW_WH, m_ConfigFile)
        '// Split
        m_HSplitterPosition = GetPrivateProfileInt("WINDOW", "SPLIT-H", DEF_SPLITTER_POS, m_ConfigFile)
        m_VSplitter1Position = GetPrivateProfileInt("WINDOW", "SPLIT-V1", DEF_SPLITTER_POS, m_ConfigFile)
        m_VSplitter2Position = GetPrivateProfileInt("WINDOW", "SPLIT-V2", DEF_SPLITTER_POS, m_ConfigFile)
        '// Read the Last Active Database
        buff = Space$(MAX_PATH)
        lRet = GetPrivateProfileString("DATABASE", "LASTACTIVE", vbNullString, buff, Len(buff), m_ConfigFile)
        m_LastActiveLibrary = Left$(buff, lRet)
        
        '// Read the last open track
        buff = Space$(MAX_PATH)
        lRet = GetPrivateProfileString("DATABASE", "LASTACTIVETRACK", vbNullString, buff, Len(buff), m_ConfigFile)
        m_LastOpenTrack = Left$(buff, lRet)
        
        '// Read the last open track path
        buff = Space$(MAX_PATH)
        lRet = GetPrivateProfileString("DATABASE", "LASTACTIVETRACKPATH", vbNullString, buff, Len(buff), m_ConfigFile)
        m_LastOpenTrackPath = Left$(buff, lRet)
        
        
        
        If m_LastActiveLibrary = vbNullString Then
            '// Create a library
            m_LastActiveLibrary = CreateLibrary()
        End If
            
        '// Read the Search History
        sMax = GetPrivateProfileInt("SEARCH", "COUNT", 0, m_ConfigFile)
        ReDim m_SearchHistory(sMax)
        For i = sMax To 1 Step -1
            buff = Space$(MAX_PATH)
            lRet = GetPrivateProfileString("SEARCH", "ITEM" & i, vbNullString, buff, Len(buff), m_ConfigFile)
            m_SearchHistory(i) = Left$(buff, lRet)
        Next
        
        '// Get the interface information
        m_DisplayDbInTitle = CBool(GetPrivateProfileInt("INTERFACE", "DBINCAPTION", True, m_ConfigFile))
        m_Gridlines = CBool(GetPrivateProfileInt("INTERFACE", "GRIDLINES", False, m_ConfigFile))
        m_ToolbarCaptions = CBool(GetPrivateProfileInt("INTERFACE", "TOOLBARCAPTIONS", True, m_ConfigFile))
        m_IntefaceHasCaptions = CBool(GetPrivateProfileInt("INTERFACE", "INTERFACECAPTIONS", True, m_ConfigFile))
        m_FontSize = GetPrivateProfileInt("INTERFACE", "FONTSIZE", 8, m_ConfigFile)
        m_BackColor = GetPrivateProfileInt("INTERFACE", "BACKCOLOR", vbWindowBackground, m_ConfigFile)
        m_ForeColor = GetPrivateProfileInt("INTERFACE", "FORECOLOR", vbWindowText, m_ConfigFile)
        
        buff = Space$(MAX_PATH)
        lRet = GetPrivateProfileString("INTERFACE", "FONTNAME", "Tahoma", buff, Len(buff), m_ConfigFile)
        m_FontName = Left$(buff, lRet)
        
        If m_FontName = vbNullString Then
            '// Create a library
            m_FontName = "Tahoma"
        End If
    Else
        '// Create a library
        m_LastActiveLibrary = CreateLibrary()
        m_FontName = "Tahoma"
    End If
End Sub


Public Property Let LastActiveLibrary(ByVal vData As String)
    m_LastActiveLibrary = vData
End Property

Public Property Get LastActiveLibrary() As String
    LastActiveLibrary = m_LastActiveLibrary
End Property



Public Property Let LastOpenTrack(ByVal vData As String)
    m_LastOpenTrack = vData
End Property

Public Property Get LastOpenTrack() As String
    LastOpenTrack = m_LastOpenTrack
End Property



Public Property Let LastOpenTrackPath(ByVal vData As String)
    m_LastOpenTrackPath = vData
End Property

Public Property Get LastOpenTrackPath() As String
    LastOpenTrackPath = m_LastOpenTrackPath
End Property



Public Property Let MainWindowState(ByVal vData As Integer)
    m_MainWindowState = vData
End Property

Public Property Get MainWindowState() As Integer
    MainWindowState = m_MainWindowState
End Property



Public Property Let HSplitterPosition(ByVal vData As Long)
    m_HSplitterPosition = vData
End Property

Public Property Get HSplitterPosition() As Long
    HSplitterPosition = m_HSplitterPosition
End Property



Public Property Let VSplitterExplorerPosition(ByVal vData As Long)
    m_VSplitter1Position = vData
End Property

Public Property Get VSplitterExplorerPosition() As Long
    VSplitterExplorerPosition = m_VSplitter1Position
End Property



Public Property Let VSplitterNavigatorPosition(ByVal vData As Long)
    m_VSplitter2Position = vData
End Property

Public Property Get VSplitterNavigatorPosition() As Long
    VSplitterNavigatorPosition = m_VSplitter2Position
End Property



Public Property Let IForeColor(ByVal vData As Long)
    m_ForeColor = vData
End Property

Public Property Get IForeColor() As Long
    IForeColor = m_ForeColor
End Property



Public Property Let IBackColor(ByVal vData As Long)
    m_BackColor = vData
End Property

Public Property Get IBackColor() As Long
    IBackColor = m_BackColor
End Property



Public Property Let FontSize(ByVal vData As Long)
    m_FontSize = vData
End Property

Public Property Get FontSize() As Long
    FontSize = m_FontSize
End Property



Public Property Let FontName(ByVal vData As String)
    m_FontName = vData
End Property

Public Property Get FontName() As String
    FontName = m_FontName
End Property



Public Property Let InterfaceCaptions(ByVal vData As Boolean)
    m_IntefaceHasCaptions = vData
End Property

Public Property Get InterfaceCaptions() As Boolean
    InterfaceCaptions = m_IntefaceHasCaptions
End Property



Public Property Let ToolbarCaptions(ByVal vData As Boolean)
    m_ToolbarCaptions = vData
End Property

Public Property Get ToolbarCaptions() As Boolean
    ToolbarCaptions = m_ToolbarCaptions
End Property



Public Property Let Gridlines(ByVal vData As Boolean)
    m_Gridlines = vData
End Property

Public Property Get Gridlines() As Boolean
    Gridlines = m_Gridlines
End Property



Public Property Let DisplayDbInTitle(ByVal vData As Boolean)
    m_DisplayDbInTitle = vData
End Property

Public Property Get DisplayDbInTitle() As Boolean
    DisplayDbInTitle = m_DisplayDbInTitle
End Property


Private Sub Class_Initialize()
    m_MainWindowWidth = DEF_WINDOW_WH
    m_MainWindowHeight = DEF_WINDOW_WH
    m_HSplitterPosition = DEF_SPLITTER_POS
    m_VSplitter1Position = DEF_SPLITTER_POS
    m_VSplitter2Position = DEF_SPLITTER_POS
    
    m_DisplayDbInTitle = True
    m_ToolbarCaptions = True
    m_IntefaceHasCaptions = True
    m_FontSize = 8
    m_FontName = "Tahoma"
    m_BackColor = vbWindowBackground
    m_ForeColor = vbWindowText
        
    ReDim m_SearchHistory(0)
    m_ConfigFile = App.Path & Chr$(92) & INI_FILE
End Sub
