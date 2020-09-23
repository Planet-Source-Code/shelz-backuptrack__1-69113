Attribute VB_Name = "modDeclarations"
Option Explicit
Public Const APP_NAME = "BackupTrack"

Public Const REF_LIM = 17                               '// Use this as a refresh trigger in iterations

Public Const GWL_STYLE = -16
Public Const PBS_MARQUEE = 8

Public Const SEARCH_NOT = "!="
Public Const SEARCH_DIR = "\="
Public Const SEARCH_FIL = ".="

Public Enum SearchConstants
    scContains
    scDoesNotContain
End Enum

Public Enum SearchFieldConstants
    sfcSearchBoth
    sfcSearchFileName
    sfcSearchFilePath
End Enum

Public Enum ColumnSortConstants
    cscSortNumber
    cscSortDate
    cscSortFormattedSize
    cscSortAlphaNumeric
End Enum

Public Const WM_USER = &H400
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TBSTYLE_FLAT = &H800

Public Const LB_SETCURSEL = &H186

Private Const TV_FIRST = &H1100
Public Const TVS_TRACKSELECT = &H200
Public Const TVM_SETBKCOLOR = (TV_FIRST + 29)
Public Const TVM_SETTEXTCOLOR = (TV_FIRST + 30)

Public Const ILD_TRANSPARENT = &H1

Public Const SHGFI_SMALLICON = &H1
Public Const SHGFI_SYSICONINDEX = &H4000
Public Const SHGFI_USEFILEATTRIBUTES = &H10


Public Const MAX_PATH As Long = 260

Public Const CB_ERR = (-1)                             '// To maintain the cyclic search stack
Public Const CB_FINDSTRINGEXACT = &H158

Public Const LVM_FIRST = &H1000
Public Const LVM_FINDITEM = LVM_FIRST + 13
Public Const LVM_GETCOUNTPERPAGE = (LVM_FIRST + 40)
Public Const LVM_GETITEM = LVM_FIRST + 5
Public Const LVM_GETITEMCOUNT = (LVM_FIRST + 4)
Public Const LVM_GETITEMTEXT = LVM_FIRST + 45
Public Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Public Const LVM_SETBKCOLOR = (LVM_FIRST + 1)
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVM_SETITEMCOUNT = (LVM_FIRST + 47)
Public Const LVM_SETTEXTCOLOR = (LVM_FIRST + 36)
Public Const LVM_SORTITEMS = LVM_FIRST + 48
Public Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)


Public Const LVIF_TEXT As Long = 1
Public Const LVIF_PARAM As Long = 4


Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1

Public Const LVSICF_NOINVALIDATEALL = &H1
Public Const LVSICF_NOSCROLL = &H2
 
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Const LVHT_ABOVE = &H8
Public Const LVHT_BELOW = &H10
Public Const LVHT_TORIGHT = &H20
Public Const LVHT_TOLEFT = &H40
Public Const LVHT_NOWHERE = &H1
Public Const LVHT_ONITEMICON = &H2
Public Const LVHT_ONITEMLABEL = &H4
Public Const LVHT_ONITEMSTATEICON = &H8
Public Const LVHT_ONITEM = (LVHT_ONITEMICON Or LVHT_ONITEMLABEL Or LVHT_ONITEMSTATEICON)

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type HITTESTINFO
    pT As POINTAPI
    flags As Long
    iItem As Long
    iSubItem  As Long
End Type

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Public Type LV_FINDINFO
  flags As Long
  psz As String
  lParam As Long
  pT As POINTAPI
  vkDirection As Long
End Type

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type


Public Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Public Declare Function GetInputState Lib "user32.dll" () As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Public Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
Public Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal pszPath As String) As Long
Public Declare Function PathMatchSpec Lib "shlwapi.dll" Alias "PathMatchSpecA" (ByVal pszFile As String, ByVal pszSpec As String) As Long


Public clib         As clsTrackLibrary                      '// The CDL File Interface
Public cConfig      As clsConfig                            '// The Config File Interface

