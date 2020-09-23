VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Option Compare Text

Private Const PR_EV = 29

Private Const INVALID_HANDLE_VALUE = -1
Private Const BUFF_SZ = 32
Private Const MAXDWORD = &HFFFFFFFF
Private Const SHGFI_TYPENAME = &H400
Private Const SHGFI_USEFILEATTRIBUTES = &H10

Private Type FILETIME
    dwLowDateTime                       As Long
    dwHighDateTime                      As Long
End Type

Private Type SYSTEMTIME
    wYear                               As Integer
    wMonth                              As Integer
    wDayOfWeek                          As Integer
    wDay                                As Integer
    wHour                               As Integer
    wMinute                             As Integer
    wSecond                             As Integer
    wMilliseconds                       As Integer
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes                    As Long
    ftCreationTime                      As FILETIME
    ftLastAccessTime                    As FILETIME
    ftLastWriteTime                     As FILETIME
    nFileSizeHigh                       As Long
    nFileSizeLow                        As Long
    dwReserved0                         As Long
    dwReserved1                         As Long
    cFileName                           As String * MAX_PATH
    cAlternate                          As String * 14
End Type

Private Declare Sub GetSystemTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME)

Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'Private Declare Function GetInputState Lib "user32.dll" () As Long
'Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32.dll" (ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Long
Private Declare Function PathMatchSpec Lib "shlwapi.dll" Alias "PathMatchSpecA" (ByVal pszFile As String, ByVal pszSpec As String) As Long
Private Declare Function PathRemoveBackslash Lib "shlwapi.dll" Alias "PathRemoveBackslashA" (ByVal pszPath As String) As Long
Private Declare Function StrFormatByteSizeW Lib "shlwapi" (ByVal qdwLow As Long, ByVal qdwHigh As Long, pwszBuf As Any, ByVal cchBuf As Long) As Long


'// Private Variables
Private m_wfd                           As WIN32_FIND_DATA
Private m_RecurseSubdirectories         As Boolean
Private m_ExpandArchives                As Boolean
Private m_AddHiddenFiles                As Boolean
Private m_CancelScan                    As Boolean
Private m_PathToScan                    As String
Private m_CurrentDirectory              As String
Private m_CurrentFile                   As String
Private m_Filter                        As String
Private m_ArchiveName                   As String
Private m_ByteSizeFormat(4)             As String
Private m_TotalScanSize                 As Double
Private m_TotalFiles                    As Long
Private m_TotalFolders                  As Long


Public Event FileFound(ByVal FileName As String, isDirectory As Boolean)
Public Event ArchiveFileFound(ByVal ArchiveName As String, ByVal ArchivePath As String, ByVal File As String, Size As String, PackSize As Long, ByVal Compression As Single, FILETIME As Date, ByVal Attributes As Long)
Public Event UpdateCaption(ByVal Text As String)
Public Event UpdateProgress(ByVal currentFile As Long, ByVal currentFolder As Long)
Public Event ScanComplete()


Public WithEvents cArchiveLstr As clsArchiveLister
Attribute cArchiveLstr.VB_VarHelpID = -1

Private Function StripLeft(ByVal tmpStr As String, ByVal chRef As String) As String
Dim cL As Long
    
    cL = InStr(tmpStr, chRef)
    If cL > 0 Then
        StripLeft = Left$(tmpStr, cL - (Len(chRef)))
    Else
        StripLeft = tmpStr
    End If
End Function

Private Sub ScanFiles(ByVal currentPath As String)
Dim dirbuff() As String, dbBound As Integer, i As Integer
Dim cleanFname As String, writeFileInfo As Boolean
Dim localfileSize As Double
Dim hFile As Long
Dim bitNumRecs As Long
    
    If Not m_CancelScan Then
        If GetInputState Then DoEvents
        '// Redim dirbuff to prevent a ubound(NULL) in the while loop
        ReDim dirbuff(0)
        m_TotalFolders = m_TotalFolders + 1
        
        m_CurrentDirectory = currentPath
        Call PathRemoveBackslash(m_CurrentDirectory)
    
        RaiseEvent UpdateCaption("Scanning..." & currentPath)
        
        '// Actual Process
        hFile = FindFirstFile(currentPath & "\*", m_wfd)
        If hFile <> INVALID_HANDLE_VALUE Then
            Do
                If GetInputState Then DoEvents
                With m_wfd
                    If AscW(.cFileName) <> 46 Then
                        If (Not (.dwFileAttributes And vbSystem) = vbSystem) Then
                            If ((.dwFileAttributes And vbDirectory) = vbDirectory) Then
                                
                                dbBound = UBound(dirbuff) + 1
                                ReDim Preserve dirbuff(dbBound)
                                dirbuff(dbBound) = Left$(.cFileName, InStr(.cFileName, vbNullChar) - 1)
                                
                                writeFileInfo = True
                            ElseIf ((.dwFileAttributes And vbHidden) <> vbHidden) Or (((.dwFileAttributes And vbHidden) = vbHidden) And m_AddHiddenFiles) Then
                                
                                writeFileInfo = True
                            End If
                        End If
                    End If
                    '// Nothin' writable found :(
                End With
                
                If writeFileInfo Then
                    
                    m_CurrentFile = Trim$(StripLeft$(m_wfd.cFileName, vbNullChar))
                    bitNumRecs = bitNumRecs + 1
                    '// Is it a directory?
                    If (m_wfd.dwFileAttributes And vbDirectory) = vbDirectory Then
                        RaiseEvent FileFound(m_CurrentFile, True)
                        
                    ElseIf (m_ExpandArchives And cArchiveLstr.IsValidArchive(m_CurrentFile)) Then
                        '// Is it an Archive?
                        With cArchiveLstr
                            '// List all the files in this archive
                            RaiseEvent FileFound(m_CurrentFile, True)
                            RaiseEvent UpdateCaption("Extracting ... " & currentPath & m_CurrentFile)
                            m_ArchiveName = m_CurrentFile
                            Call .ListArchive(currentPath & m_CurrentFile, "0")
                            m_ArchiveName = vbNullString
                        End With
                        
                    ElseIf PathMatchSpec(m_wfd.cFileName, m_Filter) = 1 Then
                        '// Does it match the filter?
                        m_TotalScanSize = m_TotalScanSize + Size()
                        m_TotalFiles = m_TotalFiles + 1
                        RaiseEvent FileFound(m_CurrentFile, False)
                    End If
                End If
                
                '// Update progressbar for every 29th record found. This speeds up
                '// Operation as well as prevents the user from analyzing the update
                '// pattern
                
                If bitNumRecs > PR_EV Then
                    RaiseEvent UpdateProgress(m_TotalFiles, m_TotalFolders)
                    bitNumRecs = 0
                End If
                
            Loop While FindNextFile(hFile, m_wfd)
            Call FindClose(hFile)
        Else
            Debug.Assert False
        End If
        
        If Not m_CancelScan Then
            '// Rcursively scan all cached Subdirs
            If m_RecurseSubdirectories Then
                For i = 1 To dbBound
                    ScanFiles currentPath & dirbuff(i) & "\"
                Next
            End If
        End If
    End If
End Sub



Public Sub Scan()
    m_CurrentDirectory = m_PathToScan
    ScanFiles m_PathToScan
    RaiseEvent UpdateProgress(m_TotalFiles, m_TotalFolders)
    RaiseEvent ScanComplete
    m_CancelScan = False
End Sub



Public Property Let CurrentDirectory(ByVal vData As String)
    m_CurrentDirectory = vData
End Property

Public Property Get CurrentDirectory() As String
    CurrentDirectory = m_CurrentDirectory
End Property

Public Property Let Filter(ByVal vData As String)
    m_Filter = UCase$(vData)
End Property

Public Property Get Filter() As String
    Filter = m_Filter
End Property


Public Property Let RecurseSubdirectories(ByVal vData As Boolean)
    m_RecurseSubdirectories = vData
End Property

Public Property Get RecurseSubdirectories() As Boolean
    RecurseSubdirectories = m_RecurseSubdirectories
End Property


Public Property Let PathToScan(ByVal vData As String)
    m_PathToScan = vData
    '// The Scanpath should end in a \
    If Mid$(m_PathToScan, Len(m_PathToScan)) <> ChrW$(92) Then m_PathToScan = m_PathToScan & ChrW$(92)
End Property

Public Property Get PathToScan() As String
    PathToScan = m_PathToScan
End Property

Public Property Let AddHiddenDirectories(ByVal vData As Boolean)
    m_AddHiddenFiles = vData
End Property

Public Property Get AddHiddenDirectories() As Boolean
    AddHiddenDirectories = m_AddHiddenFiles
End Property

Public Property Let ExpandArchives(ByVal vData As Boolean)
    m_ExpandArchives = vData
End Property

Public Property Get ExpandArchives() As Boolean
    AddHiddenDirectories = m_ExpandArchives
End Property

'// Read Only Properties
'// Total size of all files in the direcory(including subdirs if specified)
Public Property Get TotalScanSize() As Double
    TotalScanSize = m_TotalScanSize
End Property

'// Size of current file...use with FileFound event
Public Property Get Size() As Double
    With m_wfd
        Size = (.nFileSizeHigh * (MAXDWORD + 1)) + .nFileSizeLow
    End With
End Property

'// Converts a file size to a string expressed as a size value in bytes, KB, MB, or GB, based on size.
Public Property Get FormatSize(Optional Size As Double = -1) As String
Dim sBuff As String, i As Long, Sizer As Double
    If Size > 0 Then
        For i = 4 To 0 Step -1
            Sizer = 1024 ^ i
            If Size >= Sizer Then
                FormatSize = FormatNumber$(Size / Sizer, 2) & m_ByteSizeFormat(i)
                Exit For
            End If
        Next
    ElseIf Size = 0 Then
        FormatSize = "N.A."
    Else
        sBuff = Space$(BUFF_SZ)
        With m_wfd
            Call StrFormatByteSizeW(.nFileSizeLow, .nFileSizeHigh, ByVal StrPtr(sBuff), BUFF_SZ)
        End With
        FormatSize = StripLeft(sBuff, vbNullChar)
    End If
End Property
    

'// DateModified of current file...use with FileFound event
Public Property Get DateModifiedHi() As Long
    DateModifiedHi = m_wfd.ftLastWriteTime.dwHighDateTime
End Property

Public Property Get DateModifiedLo() As Long
    DateModifiedLo = m_wfd.ftLastWriteTime.dwLowDateTime
End Property

'// DateCreated of current file...use with FileFound event
Public Property Get DateCreatedHi() As Long
    DateCreatedHi = m_wfd.ftCreationTime.dwHighDateTime
End Property

Public Property Get DateCreatedLo() As Long
    DateCreatedLo = m_wfd.ftCreationTime.dwLowDateTime
End Property

'// Attributes of current file...use with FileFound event
Public Property Get Attributes() As Long
    Attributes = m_wfd.dwFileAttributes
End Property

'// FileType of current file...use with FileFound event
Public Property Get FileType(Optional File As String = vbNullString) As String
Dim sTypeData As SHFILEINFO
    If File = vbNullString Then
        If SHGetFileInfo(m_CurrentFile, 0, sTypeData, Len(sTypeData), SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
            FileType = StripLeft(sTypeData.szTypeName, vbNullChar)
        End If
    Else
        If SHGetFileInfo(File, 0, sTypeData, Len(sTypeData), SHGFI_TYPENAME Or SHGFI_USEFILEATTRIBUTES) Then
            FileType = StripLeft(sTypeData.szTypeName, vbNullChar)
        End If
    End If
End Property

Public Sub CancelScan()
    m_CancelScan = True
End Sub

Private Sub cArchiveLstr_FileFound(ByVal File As String, Size As Long, PackSize As Long, ByVal Compression As Single, FILETIME As Date, ByVal afAttributes As Long)
    If PathMatchSpec(File, m_Filter) = 1 Then
        m_TotalScanSize = m_TotalScanSize + CDbl(Size)
        RaiseEvent ArchiveFileFound(m_ArchiveName, StripLeft(m_CurrentDirectory, vbNullChar) & ChrW$(92) & m_ArchiveName, File, FormatSize(CDbl(Size)), PackSize, Compression, FILETIME, vbArchive)
    End If
End Sub

Private Sub Class_Initialize()
    m_Filter = "*.*"
    Set cArchiveLstr = New clsArchiveLister
    
    m_ByteSizeFormat(0) = " bytes"
    m_ByteSizeFormat(1) = " KB"
    m_ByteSizeFormat(2) = " MB"
    m_ByteSizeFormat(3) = " GB"
    m_ByteSizeFormat(4) = " TB"
End Sub

Private Sub Class_Terminate()
    Set cArchiveLstr = Nothing
End Sub
