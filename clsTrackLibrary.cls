VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrackLibrary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'D:\My Documents\My Projects\dataJunkie\CB\BTLIB.XML
Private Const MAX_PATH = 260

'// BTL file format constants and types
Private Const BTLIB_PATH_EXT = "BLF"
Private Const BTLIB_EXT__VER = BTLIB_PATH_EXT & "1.3"

Public Enum GetRecordsByPathConstants
    grpMatchFilesOnly
    grpMatchDirsOnly
    grpMatchAllRecords
End Enum

Private Type BTL_HEADER
    ExtVer                              As String * 6
    PackSize                            As Long
    UnPackSize                          As Long
End Type

'// Tag Names
Private Const BT_TRACK = "BTTRACK"
Private Const BT_TRACK_VER = "VERSION"

Private Const BT_TRACKLIST = "BTAL"

Private Const BT_TRACKLIST_REC = "BTAL_REC"
Private Const BT_TRACKLIST_REC_NAME = "NAME"
Private Const BT_TRACKLIST_REC_CATEGORY = "CAT"
Private Const BT_TRACKLIST_REC_CREATED = "DATE"
Private Const BT_TRACKLIST_REC_SIZE = "SIZE"
Private Const BT_TRACKLIST_REC_COMMENT = "COMM"
Private Const BT_TRACKLIST_REC_NUMRECS = "NREC"
Private Const BT_TRACKLIST_REC_BASEPATH = "BASEP"

Private Const BT_TRACK_NODE = "TRACK"
Private Const BT_TRACK_NAME = "NAME"
Private Const BT_TRACK_CATEGORY = "CAT"

Private Const BT_RECORD = "REC"
Private Const BT_RECORD_ATTDIR = "DIR"
Private Const BT_RECORD_ATTCONTAINSFILES = "FS"
Private Const BT_RECORD_VIS = "REC_VIS"
Private Const BT_RECORD_FILE = "REC_FILE"
Private Const BT_RECORD_DIR = "REC_DIR"
Private Const BT_RECORD_SIZE = "REC_SIZE"
Private Const BT_RECORD_ATTR = "REC_ATTR"


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long


Private m_XML_TmpFileName               As String
Private m_XML_CurrentLibPath            As String
Private m_XML_LibraryName               As String
Private m_XML_CurrentLibIsBLF           As Boolean
Private m_XML_QueryIterator             As IXMLDOMNodeList
Private m_XML_TrackBase                 As IXMLDOMNode
Private m_XML_Track                     As DOMDocument


Private Function CreateTempFile() As String
Dim strTmpPath As String
Dim strTmpName As String
Dim lRet As Long
    
    strTmpPath = Space$(MAX_PATH)
    lRet = GetTempPath(Len(strTmpPath), strTmpPath)
    If (lRet > 0 And lRet < MAX_PATH) Then
        strTmpName = Space$(MAX_PATH)
        lRet = GetTempFileName(strTmpPath, APP_NAME, ByVal 0, strTmpName)
        If lRet <> 0 Then CreateTempFile = Left$(strTmpName, InStr(strTmpName, vbNullChar) - 1)
    End If
End Function

Private Function ZLIB_Compress(ByRef bData() As Byte) As Long
Dim ActualSize As Long
Dim tBuffer() As Byte, tBufferSize As Long
Dim lRet As Long

    ActualSize = UBound(bData) + 1
    
    tBufferSize = ActualSize + (ActualSize * 0.01) + 12
    ReDim tBuffer(tBufferSize)
    
    '// Compress the data
    lRet = compress(tBuffer(0), tBufferSize, bData(0), ActualSize)
    
    '// Reload the input buffer with compressed data
    ReDim Preserve bData(tBufferSize - 1)
    CopyMemory bData(0), tBuffer(0), tBufferSize
    
    Erase tBuffer
    
    If lRet = 0 Then
        '// Return the compressed size
        ZLIB_Compress = UBound(bData) + 1
    Else
        ZLIB_Compress = -1
    End If
End Function

Private Function ZLIB_DeCompress(bData() As Byte, ByVal OrigSize As Long) As Long
Dim ActualSize As Long
Dim tBuffer() As Byte, tBufferSize As Long
Dim lRet As Long
    
    tBufferSize = OrigSize + (OrigSize * 0.01) + 12
    ReDim tBuffer(tBufferSize)
    
    '// Decompress the data
    lRet = uncompress(tBuffer(0), tBufferSize, bData(0), UBound(bData) + 1)
    
    '// Inflate the Return Buffer and load it
    ReDim bData(tBufferSize - 1)
    CopyMemory bData(0), tBuffer(0), tBufferSize
    
    Erase tBuffer
    'Reset properties
    If lRet = 0 Then
        ZLIB_DeCompress = tBufferSize
    Else
        ZLIB_DeCompress = -1
    End If
End Function

Public Function DecompressLibrary(ByVal SourcePath As String) As Boolean
Dim fileHeader As BTL_HEADER
Dim frf As Integer
Dim bLibData() As Byte
    
    '// Open the file
    frf = FreeFile()
    Open SourcePath For Binary Access Read Lock Write As #frf
        Get #frf, , fileHeader
        
        '// Validate the file
        If fileHeader.ExtVer <> BTLIB_EXT__VER Then
            DecompressLibrary = False
            Close #frf
            Exit Function
        End If
        
        '// Load the Compressed data
        ReDim bLibData(fileHeader.PackSize - 1)
        'Debug.Print "UBH", fileHeader.PackSize - 1
        'Debug.Print "UBP", fileHeader.UnPackSize
        Get #frf, , bLibData()
    Close #frf
        
    Call ZLIB_DeCompress(bLibData(), fileHeader.UnPackSize)
    
    '// Kill the dest file if it exists
    m_XML_TmpFileName = CreateTempFile
    If Dir$(m_XML_TmpFileName, vbNormal) <> vbNullString Then Kill m_XML_TmpFileName
    
    '// Write it
    frf = FreeFile()
    Open m_XML_TmpFileName For Binary Access Write As #frf
        '// Write decompressed XML
        Put #frf, , bLibData()
    Close #frf
    
    Erase bLibData
    DecompressLibrary = True
End Function

Public Function CompressLibrary(Optional ByVal DestLoc As String) As Boolean
Dim bLibData() As Byte
Dim frf As Integer, fLen As Long
Dim fileHeader As BTL_HEADER
    
    If Dir$(m_XML_TmpFileName) <> vbNullString Then Kill m_XML_TmpFileName
    
    m_XML_TmpFileName = CreateTempFile
    
    '// First save the file if not already done so
    Call SaveLibrary(m_XML_TmpFileName)
    
    '// Init the data array
    fileHeader.UnPackSize = FileLen(m_XML_TmpFileName) - 1
    ReDim bLibData(fileHeader.UnPackSize)
    
    '// Open the file and fetch it completely
    frf = FreeFile()
    Open m_XML_TmpFileName For Binary Access Read Lock Write As #frf
        Get #frf, , bLibData()
    Close #frf
    DoEvents
        
    '// Zip the data
    fileHeader.PackSize = ZLIB_Compress(bLibData())
    If fileHeader.PackSize > 0 Then
        '// Kill the dest file if it exists
        If Dir$(DestLoc, vbNormal) <> vbNullString Then Kill DestLoc
        
        '// Write it
        frf = FreeFile()
        Open DestLoc For Binary Access Write As #frf
            fileHeader.ExtVer = BTLIB_EXT__VER
            '// Write the header
            Put #frf, , fileHeader
            
            '// Write the compressed data
            Put #frf, , bLibData()
        Close #frf
        CompressLibrary = True
    Else
        CompressLibrary = False
    End If
    Erase bLibData
    Kill m_XML_TmpFileName
End Function

Public Function OpenFileAsLibrary(ByVal FileName As String) As Boolean
    FileName = UCase$(FileName)
    If FileName <> vbNullString Then
        If Mid$(FileName, Len(FileName) - 3) = ".BLF" Then
            '// Load A BLF file
            m_XML_CurrentLibIsBLF = True
            '// Decompress it
            Call DecompressLibrary(FileName)
            '// Open It
            OpenFileAsLibrary = OpenLibrary(m_XML_TmpFileName)
        ElseIf Mid$(FileName, Len(FileName) - 3) = ".XML" Then
            '// Load A XML file
            '// Open It
            m_XML_CurrentLibIsBLF = False
            OpenFileAsLibrary = OpenLibrary(FileName)
        End If
        m_XML_CurrentLibPath = FileName
        '// Get the library name
        m_XML_LibraryName = Dir$(FileName)
        m_XML_LibraryName = Left$(m_XML_LibraryName, InStrRev(m_XML_LibraryName, ChrW$(46)) - 1)
    End If
End Function

Public Sub SaveLibraryAsFile(ByVal FileName As String)
    FileName = UCase$(FileName)
    If FileName <> vbNullString Then
        If Mid$(FileName, Len(FileName) - 3) = ".BLF" Then
            Call CompressLibrary(FileName)
        ElseIf Mid$(FileName, Len(FileName) - 3) = ".XML" Then
            Call SaveLibrary(FileName)
        End If
    End If
End Sub

Public Function OpenLibrary(ByVal SourcePath As String) As Boolean
Dim vAttr As IXMLDOMNode
    Set m_XML_Track = New DOMDocument
    m_XML_Track.setProperty "SelectionLanguage", "XPath"
    If m_XML_Track.Load(SourcePath) Then
        '// Validate it
        m_XML_TmpFileName = vbNullString
        If m_XML_Track.childNodes(1).Attributes.Length > 0 Then
            If m_XML_Track.childNodes(1).Attributes(0).Text = BTLIB_EXT__VER Then
                '// It is a valid archive :)
                OpenLibrary = True
                m_XML_TmpFileName = SourcePath
            Else
                Debug.Print m_XML_Track.childNodes(1).nodeName & " is not valid"
            End If
        End If
    End If
End Function

Public Sub CloseLibrary()
    If m_XML_CurrentLibIsBLF Then
        Call CompressLibrary(m_XML_CurrentLibPath)
    Else
        Call SaveLibrary(m_XML_CurrentLibPath)
    End If
End Sub

'// Will create a new database and return its path
Public Function CreateNewLibrary(ByVal DestLoc As String) As String
Const XML_HEADER = "<?xml version=""1.0"" encoding=""UTF-8""?>"
Const BTLIB_EXT__VER_TMPLATE = "<BTTRACK " & _
                               BT_TRACK_VER & "=" & """" & BTLIB_EXT__VER & """" & _
                               "><BTAL></BTAL></BTTRACK>"

Dim bLibData() As Byte
Dim frf As Integer, fLen As Long
Dim fileHeader As BTL_HEADER

    m_XML_TmpFileName = CreateTempFile
    frf = FreeFile()
    Open m_XML_TmpFileName For Binary Access Write As #frf
        Put #frf, , XML_HEADER
        Put #frf, , BTLIB_EXT__VER_TMPLATE
    Close #frf
    DoEvents
        
    '// Init the data array
    fileHeader.UnPackSize = FileLen(m_XML_TmpFileName) - 1
    ReDim bLibData(fileHeader.UnPackSize)
    
    '// Open the file and fetch it completely
    frf = FreeFile()
    Open m_XML_TmpFileName For Binary Access Read Lock Write As #frf
        Get #frf, , bLibData()
    Close #frf
    DoEvents
        
    '// Zip the data
    fileHeader.PackSize = ZLIB_Compress(bLibData())
    If fileHeader.PackSize > 0 Then
        '// Kill the dest file if it exists
        If Dir$(DestLoc, vbNormal) <> vbNullString Then Kill DestLoc
        
        '// Write it
        frf = FreeFile()
        Open DestLoc For Binary Access Write As #frf
            fileHeader.ExtVer = BTLIB_EXT__VER
            '// Write the header
            Put #frf, , fileHeader
            
            '// Write the compressed data
            Put #frf, , bLibData()
        Close #frf
        CreateNewLibrary = DestLoc
    Else
        CreateNewLibrary = vbNullString
    End If
    Erase bLibData
    Kill m_XML_TmpFileName
End Function

Public Function ValidateName(ByVal rName As String) As Boolean
On Error GoTo ValidateName_ERH
Dim tmpAttr As IXMLDOMAttribute

    If Not (m_XML_Track Is Nothing) Then
        Set tmpAttr = m_XML_Track.createAttribute(BT_TRACK_NAME)
        tmpAttr.Value = rName
    End If
    
    ValidateName = True
    Exit Function
    
ValidateName_ERH:
    Err.Clear
    ValidateName = False
End Function

Public Function SaveLibrary(Optional ByVal DestLoc As String) As Boolean
    If Not (m_XML_Track Is Nothing) Then
        If Dir$(DestLoc, vbNormal) <> vbNullString Then Kill DestLoc
        DoEvents
        m_XML_Track.save DestLoc
    End If
End Function

'// Creates an track and sets up a pointer to the base node for that archive
Public Function CreateTrack(ByVal TrackName As String) As Boolean
On Error GoTo CreateTrack_ERH:
Dim tmpNode As IXMLDOMNode, xAtt As IXMLDOMAttribute

    If TrackName <> vbNullString Then
        Set tmpNode = m_XML_Track.selectSingleNode("//" & BT_TRACK & "/" & BT_TRACKLIST & "/" & BT_TRACKLIST_REC & "[@" & BT_TRACK_NAME & "=" & ChrW$(34) & TrackName & ChrW$(34) & "]")
        If tmpNode Is Nothing Then
            Set m_XML_TrackBase = m_XML_Track.createElement(BT_TRACK_NODE)
                '// Add the Track Name
                Set xAtt = m_XML_Track.createAttribute(BT_TRACK_NAME)
                xAtt.Value = TrackName
                Call m_XML_TrackBase.Attributes.setNamedItem(xAtt)
            CreateTrack = True
        Else
            '// An track with this name exitsts
            CreateTrack = False
        End If
    Else
        CreateTrack = False
    End If
    Exit Function
    
CreateTrack_ERH:
    Debug.Print Err.Description
End Function

Public Function DeleteTrack(ByVal TrackName As String) As Boolean
Dim dNode As IXMLDOMNode
    
    If Not (m_XML_Track Is Nothing) Then
        '// First delete the Track
        Set dNode = m_XML_Track.selectSingleNode(BT_TRACK & "/" & BT_TRACK_NODE & "[@" & BT_TRACK_NAME & "=" & ChrW$(34) & TrackName & Chr$(34) & "]")
        If Not (dNode Is Nothing) Then
            Call m_XML_Track.childNodes(1).removeChild(dNode)
            
            '// ...then remove the track listing
            Set dNode = m_XML_Track.selectSingleNode(BT_TRACK & "/" & BT_TRACKLIST & "/" & BT_TRACKLIST_REC & "[@" & BT_TRACK_NAME & "=" & ChrW$(34) & TrackName & ChrW$(34) & "]")
            If Not (dNode Is Nothing) Then
                Call m_XML_Track.childNodes(1).firstChild.removeChild(dNode)
                
                DeleteTrack = True
            End If
        End If
    End If
End Function

'// Saves the curently built archive to the library
Public Function SaveTrack(ByVal aComment As String, ByVal aCategory As String, ByVal aSize As String, ByVal aBasePath As String) As Long
Dim arcRecord   As IXMLDOMNode
Dim tmpNode     As IXMLDOMNode
Dim xAtt        As IXMLDOMAttribute
    
    '// Create the Track header
    Set arcRecord = m_XML_Track.createElement(BT_TRACKLIST_REC)
    
    '// Add the Track Name
    Set xAtt = m_XML_Track.createAttribute(BT_TRACK_NAME)
    xAtt.Value = m_XML_TrackBase.Attributes.getNamedItem(BT_TRACK_NAME).Text
    Call arcRecord.Attributes.setNamedItem(xAtt)
    
    '// Category
    Set tmpNode = m_XML_Track.createElement(BT_TRACKLIST_REC_CATEGORY)
    tmpNode.Text = aCategory
    Call arcRecord.appendChild(tmpNode)
    
    '// Date Created
    Set tmpNode = m_XML_Track.createElement(BT_TRACKLIST_REC_CREATED)
    tmpNode.Text = Now()
    Call arcRecord.appendChild(tmpNode)
    
    '// Comment
    Set tmpNode = m_XML_Track.createElement(BT_TRACKLIST_REC_COMMENT)
    tmpNode.Text = aComment
    Call arcRecord.appendChild(tmpNode)
    
    '// Update the record count
    Set tmpNode = m_XML_Track.createElement(BT_TRACKLIST_REC_NUMRECS)
    tmpNode.Text = m_XML_TrackBase.childNodes.Length
    Call arcRecord.appendChild(tmpNode)
    
    '// Size
    Set tmpNode = m_XML_Track.createElement(BT_TRACKLIST_REC_SIZE)
    tmpNode.Text = aSize
    Call arcRecord.appendChild(tmpNode)
    
    '// Base Path
    Set tmpNode = m_XML_Track.createElement(BT_TRACKLIST_REC_BASEPATH)
    tmpNode.Text = UCase$(aBasePath)
    Call arcRecord.appendChild(tmpNode)
    
    '// Add the Track header to the archive list
    'Debug.Print m_XML_Track.childNodes(1).firstChild.nodeName
    m_XML_Track.childNodes(1).firstChild.appendChild arcRecord
    
    '// Add the archive
    'Debug.Print m_XML_Track.childNodes(1).nodeName
    m_XML_Track.childNodes(1).appendChild m_XML_TrackBase
    
    SaveTrack = m_XML_TrackBase.childNodes.Length
    Set m_XML_TrackBase = Nothing
End Function

Public Function AddRecord(ByVal File As String, ByVal Path As String, ByVal Size As String, ByVal Attributes As Long, ContainsFiles As Boolean, Optional rHidden As Boolean = True) As Long
Dim tmpNode As IXMLDOMElement
Dim tmpRec  As IXMLDOMNode
Dim xAtt    As IXMLDOMAttribute
    
    If Not (m_XML_Track Is Nothing) Then
        Set tmpRec = m_XML_Track.createElement(BT_RECORD)
        
        '// Create an attribute for this node containing the dir. info. This greatly speeds up browsing
        Set xAtt = m_XML_Track.createAttribute(BT_RECORD_ATTDIR)
        xAtt.Value = UCase$(Path)
        Call tmpRec.Attributes.setNamedItem(xAtt)
        
        Set xAtt = m_XML_Track.createAttribute(BT_RECORD_ATTCONTAINSFILES)
        xAtt.Value = CStr(Abs(ContainsFiles))
        Call tmpRec.Attributes.setNamedItem(xAtt)
        
        '// Hidden
        Set tmpNode = m_XML_Track.createElement(BT_RECORD_VIS)
        tmpNode.Text = CStr(UCase$(rHidden))
        Call tmpRec.appendChild(tmpNode)
        
        '// Name
        Set tmpNode = m_XML_Track.createElement(BT_RECORD_FILE)
        tmpNode.Text = UCase$(File)
        Call tmpRec.appendChild(tmpNode)
        
        '// Dir
        Set tmpNode = m_XML_Track.createElement(BT_RECORD_DIR)
        tmpNode.Text = UCase$(Path)
        Call tmpRec.appendChild(tmpNode)
        
        '// Size
        Set tmpNode = m_XML_Track.createElement(BT_RECORD_SIZE)
        tmpNode.Text = Size
        Call tmpRec.appendChild(tmpNode)
        
        '// Attributes
        Set tmpNode = m_XML_Track.createElement(BT_RECORD_ATTR)
        tmpNode.Text = Attributes
        Call tmpRec.appendChild(tmpNode)
        
        '// Add it to the archive
        Call m_XML_TrackBase.appendChild(tmpRec)
        
        AddRecord = m_XML_TrackBase.childNodes.Length
    End If
End Function

Public Function OpenTrack(ByVal Index As Variant) As Long
    If IsNumeric(Index) Then
        If Index > 0 And Index < m_XML_Track.childNodes(1).childNodes.Length Then
            Set m_XML_TrackBase = m_XML_Track.childNodes(1).childNodes(Index)
        End If
    Else
        Set m_XML_TrackBase = m_XML_Track.selectSingleNode(BT_TRACK & "/" & BT_TRACK_NODE & "[@" & BT_TRACK_NAME & "=" & ChrW$(34) & Index & ChrW$(34) & "]")
    End If
    
    If Not (m_XML_TrackBase Is Nothing) Then OpenTrack = m_XML_TrackBase.childNodes.Length
End Function

Public Function GetTrackDetails(ByVal Index As Variant, aCategory As String, aCreated As Date, aNumRecords As Long, aComment As String, aSize As String, aBaseNode As String) As String
Dim tmpNode As IXMLDOMNode
    If Not (m_XML_Track Is Nothing) Then
        If IsNumeric(Index) Then
            Set tmpNode = m_XML_Track.childNodes(1).firstChild.childNodes(Index - 1)
        Else
            Set tmpNode = m_XML_Track.selectSingleNode(BT_TRACK & "/" & BT_TRACKLIST & "/" & BT_TRACKLIST_REC & "[@" & BT_TRACK_NAME & "=" & ChrW$(34) & CStr(Index) & ChrW$(34) & "]")
        End If
        If Not (tmpNode Is Nothing) Then
            With tmpNode
                GetTrackDetails = tmpNode.Attributes(0).Text ' .childNodes(0).Text
                aCategory = .childNodes(0).Text
                aCreated = CDate(.childNodes(1).Text)
                aComment = .childNodes(2).Text
                aNumRecords = CLng(.childNodes(3).Text)
                aSize = .childNodes(4).Text
                aBaseNode = .childNodes(5).Text
            End With
        End If
    End If
End Function

Public Function SetTrackDetails(ByVal TrackName As String, ByVal aName As String, ByVal aCategory As String, ByVal aComment As String) As Boolean
Dim nodRef As IXMLDOMNode, nodCheckName As IXMLDOMNode
    '// aName must not be an empty string
    '// Another archive must not exist with the same name
    If (TrackName <> vbNullString) And (aName <> vbNullString) Then
        Set nodRef = m_XML_Track.selectSingleNode(BT_TRACK & "/" & BT_TRACKLIST & "/" & BT_TRACKLIST_REC & "[@" & BT_TRACK_NAME & "=" & ChrW$(34) & TrackName & ChrW$(34) & "]")
        If Not (nodRef Is Nothing) Then
            '// Look for a collision
            If TrackName <> aName Then
                Set nodCheckName = m_XML_Track.selectSingleNode(BT_TRACK & "/" & BT_TRACKLIST & "/" & BT_TRACKLIST_REC & "[@" & BT_TRACK_NAME & "=" & ChrW$(34) & aName & ChrW$(34) & "]")
                If Not (nodCheckName Is Nothing) Then Exit Function
            End If
            '// Change the data
            With nodRef
                '// New name
                .Attributes(0).Text = aName
                '// Also change the Track attributes
                Set nodCheckName = m_XML_Track.selectSingleNode(BT_TRACK & "/" & BT_TRACK_NODE & "[@" & BT_TRACK_NAME & "=" & ChrW$(34) & TrackName & Chr$(34) & "]")
                nodCheckName.Attributes(0).Text = aName
                '// New category
                .childNodes(0).Text = aCategory
                '// New Comment
                .childNodes(2).Text = aComment
            End With
            SetTrackDetails = True
        End If
    End If
End Function

Public Function GetTrackCount() As Long
    If Not (m_XML_Track Is Nothing) Then
        GetTrackCount = m_XML_Track.childNodes(1).childNodes.Length - 1
    Else
        GetTrackCount = -1
    End If
End Function

Public Function GetRecord(ByVal Index As Long, ByRef Path As String, ByRef Size As String, ByRef Attributes As Long, ByRef Hidden As Boolean) As String
    If Not (m_XML_Track Is Nothing) Then
        If Not (m_XML_TrackBase Is Nothing) Then
            If Index >= 0 And Index < m_XML_TrackBase.childNodes.Length Then
                With m_XML_TrackBase.childNodes(Index)
                    'TrackName = .parentNode.Attributes(0).Text
                    Hidden = CBool(.childNodes(0).Text)
                    GetRecord = LCase$(.childNodes(1).Text)
                    Path = LCase$(.childNodes(2).Text)
                    Size = .childNodes(3).Text
                    Attributes = CLng(.childNodes(4).Text)
                End With
            End If
        End If
    End If
End Function

Public Function GetRecordFromSearch(ByVal Index As Long, ByRef TrackName As String, ByRef Path As String, ByRef Size As String, ByRef Attributes As Long, ByRef Hidden As Boolean) As String
    If Not (m_XML_Track Is Nothing) Then
        If Not (m_XML_QueryIterator Is Nothing) Then
            If Index < m_XML_QueryIterator.Length Then
                With m_XML_QueryIterator.Item(Index)
                    TrackName = .parentNode.Attributes(0).Text
                    Hidden = CBool(.childNodes(0).Text)
                    GetRecordFromSearch = LCase$(.childNodes(1).Text)
                    Path = LCase$(.childNodes(2).Text)
                    Size = .childNodes(3).Text
                    Attributes = CLng(.childNodes(4).Text)
                End With
            End If
        End If
    End If
End Function

Public Function GetRecordsByPath(ByVal TrackName As String, ByVal Path As String, matchSpec As GetRecordsByPathConstants) As Long
    If Not (m_XML_Track Is Nothing) Then
        Set m_XML_QueryIterator = m_XML_Track.selectNodes(BT_TRACK & "/" & BT_TRACK_NODE & "[@" & BT_TRACK_NAME & "=" & ChrW$(34) & TrackName & Chr$(34) & "]" & _
                                                          "/" & BT_RECORD & "[@" & BT_RECORD_ATTDIR & "=" & ChrW$(34) & Path & Chr$(34) & _
                                                          " and " & "@" & BT_RECORD_ATTCONTAINSFILES & "=" & ChrW$(34) & matchSpec & Chr$(34) & "]")
        If Not (m_XML_QueryIterator Is Nothing) Then
            GetRecordsByPath = m_XML_QueryIterator.Length
        Else
            GetRecordsByPath = -1
        End If
    End If
End Function

'// Sample XPATH:BTTRACK/TestA2/REC[not(contains(REC_FILE,"cb"))]
Public Function FindInTrack(ByVal TrackName As String, Sc As SearchConstants, Arg As String, Optional includeHidden As Boolean = False, Optional Fields As SearchFieldConstants = sfcSearchBoth) As Long
Dim Query           As String
Dim strCondition    As String
Dim strSRange       As String


    If Not (m_XML_Track Is Nothing) Then
        '// Start building the query
        Query = BT_TRACK & "/" & BT_TRACK_NODE & "[@" & BT_TRACK_NAME & "=" & ChrW$(34) & TrackName & Chr$(34) & "]" & "/" & BT_RECORD   '// Get all records in 'TrackName'
        
        '// Create the query condition
        '// Create the query condition
        If Fields = sfcSearchBoth Then
            strCondition = "contains(" & BT_RECORD_FILE & "," & Chr$(34) & Arg & Chr$(34) & ")" & _
                           " or " & "contains(" & BT_RECORD_DIR & "," & Chr$(34) & Arg & Chr$(34) & ")"
        ElseIf Fields = sfcSearchFileName Then
            strCondition = "contains(" & BT_RECORD_FILE & "," & Chr$(34) & Arg & Chr$(34) & ")"
        Else
            strCondition = "contains(" & BT_RECORD_DIR & "," & Chr$(34) & Arg & Chr$(34) & ")"
        End If
        
        If Sc = scDoesNotContain Then strCondition = "not(" & strCondition & ")"

        Query = Query & "[" & strCondition & "]"

        Debug.Print Query

        If Query <> vbNullString Then
            Set m_XML_QueryIterator = m_XML_Track.selectNodes(Query)

            If Not (m_XML_QueryIterator Is Nothing) Then
                FindInTrack = m_XML_QueryIterator.Length
            Else
                FindInTrack = -1
            End If
        End If
    End If
End Function

Public Function FindInLibrary(Sc As SearchConstants, Arg As String, Optional includeHidden As Boolean = False, Optional Fields As SearchFieldConstants = sfcSearchBoth) As Long
Dim Query           As String
Dim strCondition    As String
Dim strSRange       As String


    If Not (m_XML_Track Is Nothing) Then
        '// Start building the query
        Query = BT_TRACK & "//" & BT_RECORD    '// Get all records in 'TrackName'
        
        '// Create the query condition
        If Fields = sfcSearchBoth Then
            strCondition = "contains(" & BT_RECORD_FILE & "," & Chr$(34) & Arg & Chr$(34) & ")" & _
                           " or " & "contains(" & BT_RECORD_DIR & "," & Chr$(34) & Arg & Chr$(34) & ")"
        ElseIf Fields = sfcSearchFileName Then
            strCondition = "contains(" & BT_RECORD_FILE & "," & Chr$(34) & Arg & Chr$(34) & ")"
        Else
            strCondition = "contains(" & BT_RECORD_DIR & "," & Chr$(34) & Arg & Chr$(34) & ")"
        End If
        
        If Sc = scDoesNotContain Then strCondition = "not(" & strCondition & ")"

        Query = Query & "[" & strCondition & "]"

        Debug.Print Query

        If Query <> vbNullString Then
            Set m_XML_QueryIterator = m_XML_Track.selectNodes(Query)

            If Not (m_XML_QueryIterator Is Nothing) Then
                FindInLibrary = m_XML_QueryIterator.Length
            Else
                FindInLibrary = -1
            End If
        End If
    End If
End Function

Public Function GetNumResults() As Long
    GetNumResults = m_XML_QueryIterator.Length
End Function

'// You can run customized searches through this function
Public Function ExecQuery(ByVal Query As String) As Long
On Error GoTo ExecQuery_ERH
Dim xNode As IXMLDOMNode
    If Not (m_XML_Track Is Nothing) Then
        Set m_XML_QueryIterator = m_XML_Track.selectNodes(Query)
        ExecQuery = m_XML_QueryIterator.Length
    End If
    Exit Function
    
ExecQuery_ERH:
    MsgBox "An Xpath error occured. " & vbCrLf & "DETAILS (" & Err.Number & ")" & vbCrLf & _
           Err.Description, vbExclamation Or vbOKOnly, "Xpath Error"
    Err.Clear
End Function

Public Property Get CurrentDataBasePath() As String
    CurrentDataBasePath = m_XML_CurrentLibPath
End Property

Public Property Get DataBaseWorkFilePath() As String
    DataBaseWorkFilePath = m_XML_TmpFileName
End Property
    
Public Property Get LibraryName() As String
    LibraryName = m_XML_LibraryName
End Property

Private Sub Class_Terminate()
    Set m_XML_Track = Nothing
    Set m_XML_TrackBase = Nothing
    Set m_XML_QueryIterator = Nothing
End Sub
