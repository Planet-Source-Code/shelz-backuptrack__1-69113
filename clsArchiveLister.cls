VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArchiveLister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'// Common Constants
Private Const STR_VALID_ARCHIVE_MASK = ".ZIP.RAR"

Private Enum ArchiveFileConstants
    FILE_INDEX_UNKNOWN = 0
    FILE_INDEX_RAR = 1
    FILE_INDEX_ZIP = 2
End Enum

'// RAR-Specific constants
Private Const RAR_OM_LIST = 0
Private Const RAR_OM_EXTRACT = 1

Private Const RAR_NO_PASSWORD = "0"
 
Private Const RAR_SKIP = 0
Private Const RAR_TEST = 1
Private Const RAR_EXTRACT = 2

Public Enum ArchiveErrorConstants
    ERAR_END_ARCHIVE = 10
    ERAR_NO_MEMORY = 11
    ERAR_BAD_DATA = 12
    ERAR_BAD_ARCHIVE = 13
    ERAR_UNKNOWN_FORMAT = 14
    ERAR_EOPEN = 15
    ERAR_ECREATE = 16
    ERAR_ECLOSE = 17
    ERAR_EREAD = 18
    ERAR_EWRITE = 19
    ERAR_SMALL_BUF = 20
End Enum

Enum RarOperations
    OP_EXTRACT = 0
    OP_TEST = 1
    OP_LIST = 2
End Enum

Private Type RARHeaderData
    ArcName                             As String * 260
    FileName                            As String * 260
    flags                               As Long
    PackSize                            As Long
    UnpSize                             As Long
    HostOS                              As Long
    FileCRC                             As Long
    FILETIME                            As Long
    UnpVer                              As Long
    Method                              As Long
    FileAttr                            As Long
    CmtBuf                              As String
    CmtBufSize                          As Long
    CmtSize                             As Long
    CmtState                            As Long
End Type
 
Private Type RAROpenArchiveData
    ArcName                             As String
    OpenMode                            As Long
    OpenResult                          As Long
    CmtBuf                              As String
    CmtBufSize                          As Long
    CmtSize                             As Long
    CmtState                            As Long
End Type


'// RAR-Specific declarations
Private Declare Function RAROpenArchive Lib "unrar.dll" (ByRef ArchiveData As RAROpenArchiveData) As Long
Private Declare Function RARCloseArchive Lib "unrar.dll" (ByVal hArcData As Long) As Long
Private Declare Function RARReadHeader Lib "unrar.dll" (ByVal hArcData As Long, ByRef HeaderData As RARHeaderData) As Long
Private Declare Function RARProcessFile Lib "unrar.dll" (ByVal hArcData As Long, ByVal Operation As Long, ByVal DestPath As String, ByVal DestName As String) As Long
Private Declare Sub RARSetPassword Lib "unrar.dll" (ByVal hArcData As Long, ByVal Password As String)


'// Property Vars
Private m_RecurseArchives               As Boolean
Private m_bHandleZIP                    As Boolean

Event FileFound(ByVal File As String, Size As Long, PackSize As Long, ByVal Compression As Single, FILETIME As Date, ByVal Attributes As Long)

'// Helper subroutines and functions
Private Function StripLeft(ByVal tmpStr As String, ByVal chRef As String) As String
Dim cL As Long
    
    cL = InStr(tmpStr, chRef)
    If cL > 0 Then
        StripLeft = Left$(tmpStr, cL - (Len(chRef)))
    Else
        StripLeft = tmpStr
    End If
End Function

Private Function GetExtentionByIndex(ByVal FileName As String) As ArchiveFileConstants
Dim tmpStr As String
    
    If Len(FileName) > 4 Then
        tmpStr = UCase$(Right$(FileName, 4))
        
        Select Case tmpStr
            Case ".RAR"
                GetExtentionByIndex = FILE_INDEX_RAR
                
            Case ".ZIP"
                GetExtentionByIndex = FILE_INDEX_ZIP
        End Select
    End If
End Function

'// ZIP-Specific routines
Private Function ListZipFiles(ZipFile As String) As ArchiveErrorConstants
Dim s(0) As String
Dim tDCL As DCLIST

    tDCL.lpszZipFN = ZipFile
    tDCL.lpszExtractDir = vbNullChar
    tDCL.nvflag = 1
    tDCL.fQuiet = 2
   ' tDCL.nzflag = 1
    
    m_bHandleZIP = True
        Call modZipFunctions.VBUnzip(Me, tDCL, 0, s(), 0, s())
    m_bHandleZIP = False
End Function

Private Function ListRarFiles(RarFile As String, Optional Password As String = RAR_NO_PASSWORD) As ArchiveErrorConstants
Dim hArchive            As Long
Dim Status              As Integer
Dim urarArchiveData     As RAROpenArchiveData
Dim urarHeader          As RARHeaderData
Dim tmpStr              As String
Dim tmpDir              As String
Dim Compression         As Single
Dim lRet                As Long
Dim RarFileBuff()       As String
    
    ReDim RarFileBuff(0) As String
    
    '// Init our header
    With urarArchiveData
        .ArcName = RarFile
        .CmtBuf = Space$(16384)
        .CmtBufSize = Len(.CmtBuf) '16384
        .OpenMode = RAR_OM_EXTRACT
    End With
    
    '// Open the archive
    hArchive = RAROpenArchive(urarArchiveData)
    
    '// A problem occured
    If urarArchiveData.OpenResult <> 0 Then
        ListRarFiles = urarArchiveData.OpenResult
        
    Else
        '// All OK
        
        '// Set the password if any
        If Password <> "" Then RARSetPassword hArchive, Password
            
        '// Now list the files
        Status = RARReadHeader(hArchive, urarHeader)
        
        Do Until Status <> 0
            tmpStr = StripLeft(urarHeader.FileName, vbNullChar)
            If urarHeader.UnpSize > 0 Then
                Compression = (urarHeader.UnpSize - urarHeader.PackSize) / urarHeader.UnpSize * 100
            Else
                Compression = 0
            End If
            RaiseEvent FileFound(tmpStr, urarHeader.UnpSize, urarHeader.PackSize, Compression, Now(), urarHeader.FileAttr)

            '// We do not need to process this file
            lRet = RARProcessFile(hArchive, RAR_SKIP, vbNullChar, vbNullChar)
            
            If lRet <> 0 Then
                ListRarFiles = lRet
                Exit Do
            End If
            
            Status = RARReadHeader(hArchive, urarHeader)
        Loop
        
        If Status = ERAR_BAD_DATA Then ListRarFiles = ERAR_BAD_DATA
        
        RARCloseArchive hArchive
    End If
    
    Erase RarFileBuff
End Function

Public Function ListArchive(ByVal ArchiveFile As String, Optional Password As String) As Long
Dim fileIndex As ArchiveFileConstants
    
    fileIndex = GetExtentionByIndex(ArchiveFile)
    
    Select Case fileIndex
        Case FILE_INDEX_RAR
            Call ListRarFiles(ArchiveFile)
            
        Case FILE_INDEX_ZIP
            Call ListZipFiles(ArchiveFile)
    End Select
End Function

Public Function IsValidArchive(ByVal File As String) As Boolean
    If InStr(1, STR_VALID_ARCHIVE_MASK, UCase$(Right$(File, 4)), vbBinaryCompare) > 0 Then
        IsValidArchive = True
    End If
End Function

'// ZIP Module Friend functions
Friend Sub ZIP_ListZipFileContents(ByVal File As String, ByVal ucsize As Long, ByVal csiz As Long, ByVal cfactor As Integer, FileDate As Date, ByVal crc As Long, ByVal fCrypt As Byte)
'// This function can be only accessed while a zip is being read
    If m_bHandleZIP Then
        
        '// Clean the File Name
        File = StripLeft(File, vbNullChar)
        File = Replace$(File, ChrW$(47), ChrW$(92))
        
        RaiseEvent FileFound(File, ucsize, csiz, cfactor, FileDate, fCrypt)
    End If
End Sub

Friend Sub ZIP_ShowMesg(ByVal Mesg As String)
    Debug.Print Mesg
End Sub
        


Public Property Let RecurseArchives(newVal As Boolean)
    m_RecurseArchives = newVal
End Property
Public Property Get RecurseArchives() As Boolean
    RecurseArchives = m_RecurseArchives
End Property


Private Sub Class_Initialize()
    ChDrive App.Path
    ChDir App.Path
End Sub
