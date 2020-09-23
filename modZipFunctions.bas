Attribute VB_Name = "modZipFunctions"
Option Explicit
Private Enum EUZOverWriteResponse
   euzDoNotOverwrite = 100
   euzOverwriteThisFile = 102
   euzOverwriteAllFiles = 103
   euzOverwriteNone = 104
End Enum

Private Type UNZIPnames
    s(0 To 1023) As String
End Type

'// Callback large "string" (sic)
Private Type CBChar
    ch(0 To 32800) As Byte
End Type

'// Callback small "string" (sic)
Private Type CBCh
    ch(0 To 255) As Byte
End Type

Public Type DCLIST
   ExtractOnlyNewer As Long      ' 1 to extract only newer
   SpaceToUnderScore As Long     ' 1 to convert spaces to underscore
   PromptToOverwrite As Long     ' 1 if overwriting prompts required
   fQuiet As Long                ' 0 = all messages, 1 = few messages, 2 = no messages
   ncflag As Long                ' write to stdout if 1
   ntflag As Long                ' test zip file
   nvflag As Long                ' verbose listing
   nUflag As Long                ' "update" (extract only newer/new files)
   nzflag As Long                ' display zip file comment
   ndflag As Long                ' all args are files/dir to be extracted
   noflag As Long                ' 1 if always overwrite files
   naflag As Long                ' 1 to do end-of-line translation
   nZIflag As Long               ' 1 to get zip info
   C_flag As Long                ' 1 to be case insensitive
   fPrivilege As Long            ' zip file name
   lpszZipFN As String           ' directory to extract to.
   lpszExtractDir As String
End Type

Private Type USERFUNCTION
   ' Callbacks:
   lptrPrnt As Long           ' Pointer to application's print routine
   lptrSound As Long          ' Pointer to application's sound routine.  NULL if app doesn't use sound
   lptrReplace As Long        ' Pointer to application's replace routine.
   lptrPassword As Long       ' Pointer to application's password routine.
   lptrMessage As Long        ' Pointer to application's routine for
                              ' displaying information about specific files in the archive
                              ' used for listing the contents of the archive.
   lptrService As Long        ' callback function designed to be used for allowing the
                              ' app to process Windows messages, or cancelling the operation
                              ' as well as giving option of progress.  If this function returns
                              ' non-zero, it will terminate what it is doing.  It provides the app
                              ' with the name of the archive member it has just processed, as well
                              ' as the original size.
                              
   ' Values filled in after processing:
   lTotalSizeComp As Long     ' Value to be filled in for the compressed total size, excluding
                              ' the archive header and central directory list.
   lTotalSize As Long         ' Total size of all files in the archive
   lCompFactor As Long        ' Overall archive compression factor
   lNumMembers As Long        ' Total number of files in the archive
   cchComment As Integer      ' Flag indicating whether comment in archive.
End Type

Private Declare Function Wiz_SingleEntryUnzip Lib "unzip32.dll" (ByVal ifnc As Long, ByRef ifnv As UNZIPnames, ByVal xfnc As Long, ByRef xfnv As UNZIPnames, dcll As DCLIST, Userf As USERFUNCTION) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private clsAL As clsArchiveLister

'// Info-Zip Callbacks
Private Sub UnzipMessageCallBack(ByVal ucsize As Long, ByVal csiz As Long, ByVal cfactor As Integer, ByVal mo As Integer, ByVal dy As Integer, ByVal yr As Integer, ByVal hh As Integer, ByVal mm As Integer, ByVal c As Byte, ByRef fName As CBCh, ByRef meth As CBCh, ByVal crc As Long, ByVal fCrypt As Byte)
On Error Resume Next

Dim File As String
Dim sFolder As String
Dim dDate As Date
Dim sMethod As String
Dim iPos As Long
    
    '// Fix the file
    File = StrConv(fName.ch, vbUnicode)
    
    '// Fix the date
    dDate = DateSerial(yr, mo, hh) + TimeSerial(hh, mm, 0)
    
    '// Call back the Archive Lister class
    Call clsAL.ZIP_ListZipFileContents(File, ucsize, csiz, cfactor, dDate, crc, fCrypt)
End Sub

Private Function UnzipPrintCallback(ByRef fName As CBChar, ByVal x As Long) As Long
On Error Resume Next
Dim bArray()    As Byte
Dim vbMesg      As String

    If x > 1 And x < 1024 Then
        ReDim bArray(x) As Byte
        CopyMemory bArray(0), fName, x
    End If
    
    vbMesg = StrConv(bArray, vbUnicode)
    
    '// Adjust Backslashes
    If InStr(vbMesg, ChrW$(47)) > 0 Then
        vbMesg = Replace$(vbMesg, ChrW$(47), ChrW$(92))
    End If
    
    '// Call back the Archive Lister class
    Call clsAL.ZIP_ShowMesg(vbMesg)
    
    Erase bArray()
    
    UnzipPrintCallback = 0
End Function

Private Function UnzipPasswordCallBack(ByRef pwd As CBCh, ByVal x As Long, ByRef s2 As CBCh, ByRef Name As CBCh) As Long
On Error Resume Next
    '// We do not handle Password protected files
    UnzipPasswordCallBack = 1
    Debug.Print "File skipped as password required"
End Function

Private Function UnzipReplaceCallback(ByRef fName As CBChar) As Long
On Error Resume Next
    UnzipReplaceCallback = euzDoNotOverwrite
End Function

Private Function UnZipServiceCallback(ByRef mname As CBChar, ByVal x As Long) As Long
On Error Resume Next
    Debug.Print "UnZipServiceCallback"
    UnZipServiceCallback = 1
End Function

Public Function VBUnzip(cAL As clsArchiveLister, tDCL As DCLIST, iIncCount As Long, sInc() As String, iExCount As Long, sExc() As String) As Long
Dim tInc As UNZIPnames
Dim tExc As UNZIPnames
Dim tUser As USERFUNCTION
Dim lR As Long
Dim i As Long

    Set clsAL = cAL

    '// Init the tUser structure
    tUser.lptrPrnt = FnPtr(AddressOf UnzipPrintCallback)
    tUser.lptrSound = 0& ' not supported
    tUser.lptrReplace = FnPtr(AddressOf UnzipReplaceCallback)
    tUser.lptrPassword = FnPtr(AddressOf UnzipPasswordCallBack)
    tUser.lptrMessage = FnPtr(AddressOf UnzipMessageCallBack)
    tUser.lptrService = FnPtr(AddressOf UnZipServiceCallback)
    
    '// Include all files
    tInc.s(0) = vbNullChar
    tExc.s(0) = vbNullChar
    
    '// Unzip
    VBUnzip = Wiz_SingleEntryUnzip(iIncCount, tInc, iExCount, tExc, tDCL, tUser)
End Function

