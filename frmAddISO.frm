VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAddISO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add new disc image"
   ClientHeight    =   4245
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6000
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pProgress 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3975
      ScaleWidth      =   5775
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3840
         TabIndex        =   24
         Top             =   3600
         Width           =   1935
      End
      Begin VB.ListBox lstStatusMessages 
         Height          =   1455
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   5535
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   661
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         Height          =   555
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   5535
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Folders Examined :"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblFolderCount 
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblPleaseWait 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   120
         Width           =   5295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "More Info..."
      Height          =   285
      Left            =   4560
      TabIndex        =   23
      Top             =   825
      Width           =   1335
   End
   Begin VB.Timer tmrMarquee 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   2640
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Track"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtComment 
      Height          =   1095
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   2400
      Width           =   4935
   End
   Begin VB.TextBox txtScanFilter 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   5010
      TabIndex        =   8
      Text            =   "*.*"
      Top             =   1995
      Width           =   855
   End
   Begin VB.CheckBox chkExpandArchives 
      Caption         =   "Expand archives [zip and rar]"
      Height          =   195
      Left            =   960
      TabIndex        =   7
      Top             =   2040
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtRefName 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   825
      Width           =   3135
   End
   Begin VB.TextBox txtScanPath 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   450
      Width           =   4185
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Top             =   450
      Width           =   375
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1230
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   17
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Track Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Path Mask : "
      Height          =   255
      Left            =   3930
      TabIndex        =   9
      Top             =   2010
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Track Name"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CD Image Path"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   465
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Track Type"
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   4
      Top             =   1260
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cISO As FL_ISO9660Reader
Attribute cISO.VB_VarHelpID = -1
Private cS As clsScanner

Private bCancel As Boolean

Private Function GetLastComponent(ByVal Path As String) As String
Dim cL As Long
    If Asc(Mid$(Path, Len(Path))) = 92 Then Path = Left$(Path, Len(Path) - 1)
    
    cL = InStrRev(Path, ChrW$(92))
    If cL > 0 Then
        GetLastComponent = Mid$(Path, cL + 1)
    Else
        GetLastComponent = Path
    End If
End Function

Private Sub ListDir(ByVal sPath As String)
Dim dirs() As String
Dim files() As String
Dim fPath As String, currDir As String, mFilter As String
Dim i As Integer, j As Integer
Static folderCount As Long, fileCount As Long
    
    If Not bCancel Then
        fPath = txtScanPath.Text
        mFilter = txtScanFilter.Text
        
        lstStatusMessages.AddItem ChrW$(91) & Format$(Now, "hh:mm:ss") & ChrW$(93) & Space$(2) & "Scanning " & fPath & sPath
        Call SendMessage(lstStatusMessages.hWnd, LB_SETCURSEL, lstStatusMessages.NewIndex, ByVal 0)
        DoEvents
        
        '// Get all sub dirs of sPath
        If cISO.HasSubDirs(sPath) Then
            '// Get sub dirs for sPath
            dirs = cISO.GetSubDirs(sPath)
            folderCount = folderCount + UBound(dirs)
            lblFolderCount.Caption = folderCount
            PB.Max = PB.Max + UBound(dirs) + 1
            For i = 0 To UBound(dirs)
                If GetInputState Then DoEvents
                PB.Value = i
                '// Add each directory in this path
                currDir = dirs(i)
                currDir = GetLastComponent(StripSlashes(currDir, True))
                Call clib.AddRecord(currDir, StripSlashes(fPath & sPath, True), "N.A.", 0, True, False)
                lblInfo.Caption = "Current Folder : " & fPath & sPath
                '// Get its files
                If cISO.HasSubFiles(dirs(i)) Then
                    '// Get array with files for this dir
                    files = cISO.GetSubFiles(dirs(i))
                    '// Add the files for this directory to the library
                    For j = 0 To UBound(files)
                        '// Match the files to the filemask
                        If PathMatchSpec(files(j) & vbNullChar, mFilter) = 1 Then
                            Call clib.AddRecord(files(j), StripSlashes(fPath & dirs(i), True), cS.FormatSize(cISO.GetFilesize(files(j))), 0, False, False)
                        End If
                    Next
                End If
                '// ... and Iterate through all sub dirs
                ListDir dirs(i)
            Next
        End If
    End If
End Sub


Private Sub cmdBrowse_Click()
    With frmMain.cmDlg
        .CancelError = False
        .DialogTitle = "Select Disc Image to scan"
        .Filter = "ISO 9660 Image File (*.iso)|*.iso"
        .ShowOpen
    
        If .FileName <> vbNullString Then
            Me.MousePointer = vbHourglass
                Set cISO = New FL_ISO9660Reader
                Set cS = New clsScanner
                If Not cISO.ReadISO(.FileName) Then
                    MsgBox "An error occured. " & .FileName & " Could not be read.", vbCritical Or vbOKOnly, "ISO Read Error"
                    Set cISO = Nothing
                    Me.MousePointer = vbDefault
                    Exit Sub
                Else
                    txtScanPath.Text = .FileName
                    txtRefName.Text = cISO.VolumeID
                    txtComment.Text = cISO.ApplicationID
                    cmdCreate.Enabled = True
                End If
            Me.MousePointer = vbDefault
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    bCancel = True
    lstStatusMessages.AddItem ChrW$(91) & Format$(Now, "hh:mm:ss") & ChrW$(93) & Space$(2) & "Cancelling scan please stand by... "
    Call SendMessage(lstStatusMessages.hWnd, LB_SETCURSEL, lstStatusMessages.NewIndex, ByVal 0)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
Dim TrackName As String, fileCount As Long
    
    With pProgress
        .Move 120, 120
        .Visible = True
        .Refresh
    End With
    
    If (txtRefName.Text <> vbNullString) And (txtScanPath.Text <> vbNullString) Then
        '// Validate the name
        If clib.ValidateName(txtRefName.Text) Then
            lblInfo.Caption = "Adding ... " & txtRefName.Text
    
            TrackName = Trim$(txtRefName.Text)
            
            If clib.CreateTrack(TrackName) Then
                lblPleaseWait.Caption = "Please wait while " & APP_NAME & " is scanning. This may take a while."
                Call ListDir("\")
                
                If Not bCancel Then
                    lblInfo.Caption = "Comitting track..."
                    fileCount = clib.SaveTrack(txtComment.Text, cmbType.Text, cS.FormatSize(cISO.VolumeSize * 2048), txtScanPath.Text)
                    
                    lblInfo.Caption = "Scan Complete. Track created"
                    
                    If MsgBox("Scan Complete." & vbCrLf & vbCrLf & _
                              "Number of files added: " & fileCount & _
                              ". Would you like to add another track?", vbQuestion Or vbYesNo, "Scan Complete. Add Another?") = vbNo Then
                        Unload Me
                    Else
                        txtComment.Text = vbNullString
                        txtScanPath.Text = vbNullString
                        Set cISO = Nothing
                        Set cS = Nothing
                    End If
                Else
                    MsgBox "Scan Cancelled. " & txtScanPath.Text & " was not scanned into a track.", vbOKOnly Or vbInformation, "Cancelled!"
                    bCancel = False
                End If
            Else
                MsgBox "An error occured. " & vbCrLf & _
                       "A possible reason for this is that an track with the name '" & _
                       txtRefName.Text & "' already exists in the library. " & vbCrLf & _
                       "The track was not created", vbCritical Or vbOKOnly, "Error"
            End If
        Else
            MsgBox "An error occured. " & vbCrLf & _
                   "An invalid name was given to the track. " & vbCrLf & _
                   "Only alphabets, numbers and the special characters '.', '-' and '_' are permitted" & vbCrLf & _
                   "The track was not created", vbCritical Or vbOKOnly, "Error"

        End If
    Else
        MsgBox "Please Enter a name for the Track" & vbCrLf & _
               "The track was not created", vbCritical Or vbOKOnly, "Error"
    End If
    
    txtRefName.Text = vbNullString
    pProgress.Visible = False
End Sub

Private Sub Command1_Click()
Dim messageStr As String
        messageStr = "Volume ID: " & cISO.VolumeID & vbCrLf & _
                     "System ID: " & cISO.SystemID & vbCrLf & _
                     "Volume Size: " & cISO.VolumeSize * 2048 & vbCrLf & _
                     "Application: " & cISO.ApplicationID & vbCrLf & _
                     "Data Preparer: " & cISO.DataPreparerID & vbCrLf & _
                     "Publisher: " & cISO.PublisherID & vbCrLf & _
                     "Abstract file: " & cISO.AbstractFile & vbCrLf & _
                     "Bibliographic file: " & cISO.BibliographicFile & vbCrLf & _
                     "Copyright file: " & cISO.CopyrightFile & vbCrLf & _
                     "Creation Date: " & cISO.VolumeCreationDate & vbCrLf & _
                     "Effective Date: " & cISO.VolumeEffectiveDate & vbCrLf & _
                     "Expiration Date: " & cISO.VolumeExpirationDate & vbCrLf & _
                     "Modification Date: " & cISO.VolumeModificationDate
        MsgBox messageStr, vbInformation Or vbOKOnly, "Info: " & cISO.VolumeID
End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_ERH
Dim nStyle As Long

    Me.Caption = APP_NAME & " - Create New Track"
    
    nStyle = GetWindowLong(PB.hWnd, GWL_STYLE)
    SetWindowLong PB.hWnd, GWL_STYLE, nStyle Or PBS_MARQUEE
    
    For nStyle = 0 To 9
        cmbType.AddItem LoadResString(nStyle)
    Next
    cmbType.ListIndex = 9
    
    Exit Sub
    
Form_Load_ERH:
    MsgBox "An internal error occured." & vbCrLf & _
           "Error Details..." & vbCrLf & _
           Space$(5) & "Error Number : " & Err.Number & vbCrLf & _
           Space$(5) & "Description  : " & Err.Description, vbOKOnly Or vbCritical, "Error!"
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cISO = Nothing
    Set cS = Nothing
End Sub

