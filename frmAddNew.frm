VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAddNew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New..."
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      Sorted          =   -1  'True
      TabIndex        =   24
      Text            =   "cmbType"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Timer tmrMarquee 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   120
      Top             =   4200
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4350
      TabIndex        =   17
      Top             =   420
      Width           =   375
   End
   Begin VB.TextBox txtScanPath 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   420
      Width           =   3135
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4695
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      ShowTips        =   0   'False
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txtRefName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   795
      Width           =   3135
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2040
      Width           =   3135
   End
   Begin VB.PictureBox picScanOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   165
      ScaleHeight     =   1980
      ScaleWidth      =   5505
      TabIndex        =   7
      Top             =   2040
      Width           =   5505
      Begin VB.CheckBox chkExpandArchives 
         BackColor       =   &H80000005&
         Caption         =   "Expand archives [zip and rar]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox chkRecurse 
         BackColor       =   &H80000005&
         Caption         =   "Recurse subdirectories"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.TextBox txtScanFilter 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4650
         TabIndex        =   10
         Text            =   "*.*"
         Top             =   120
         Width           =   855
      End
      Begin VB.CheckBox chkCtlHiddenDirs 
         BackColor       =   &H80000005&
         Caption         =   "Add hidden directories"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   440
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox chkCtlHidden 
         BackColor       =   &H80000005&
         Caption         =   "Mark hidden files as hidden"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   760
         Width           =   2895
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Path Mask : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3575
         TabIndex        =   13
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Hidden files are cataloged but not searched unless specified in  preferences."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   1560
         Width           =   5415
      End
   End
   Begin VB.PictureBox pScanFeedback 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1980
      Left            =   240
      ScaleHeight     =   1980
      ScaleWidth      =   5505
      TabIndex        =   18
      Top             =   2040
      Width           =   5505
      Begin VB.ListBox lstStatusMessages 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         IntegralHeight  =   0   'False
         Left            =   0
         TabIndex        =   22
         Top             =   840
         Width           =   5415
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   0
         TabIndex        =   20
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label lblFileCount 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblPleaseWait 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   5295
      End
   End
   Begin ComctlLib.TabStrip tabStrip 
      Height          =   2415
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4260
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Options"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Comment"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Status"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   26
      Top             =   4200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Track"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   4200
      Width           =   1935
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
      TabIndex        =   25
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Track Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   -120
      TabIndex        =   23
      Top             =   1230
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4920
      Picture         =   "frmAddNew.frx":0000
      Stretch         =   -1  'True
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Scan Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   -120
      TabIndex        =   3
      Top             =   435
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Track Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   -120
      TabIndex        =   2
      Top             =   810
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cS           As clsScanner
Attribute cS.VB_VarHelpID = -1

Private markHidden              As Boolean
Private bCancel                 As Boolean

Private Sub cmdBrowse_Click()
Dim tmpStr As String

    tmpStr = BrowsePath(hWnd, "Select Scan Path")
    If Len(tmpStr) = 3 Then
        If Mid$(tmpStr, 2) = ":\" Then txtRefName.Text = Dir$(tmpStr, vbVolume)
    Else
        txtRefName.Text = Dir$(tmpStr, vbDirectory)
    End If
    txtRefName.Text = txtRefName.Text
    txtScanPath.Text = tmpStr
    txtScanPath.Tag = tmpStr
    
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
Dim TrackName As String

    txtScanPath.Enabled = False
    txtRefName.Enabled = False
    picScanOptions.Enabled = False
    cmdClose.Enabled = False
    cmdCreate.Enabled = False
    cmdBrowse.Enabled = False
    cmbType.Enabled = False
    tmrMarquee.Enabled = True
    cmdCancel.Visible = True
    DoEvents

    TabStrip.Tabs(3).Selected = True
    
    If (txtRefName.Text <> vbNullString) And (txtScanPath.Text <> vbNullString) Then
        '// Validate the name
        If clib.ValidateName(txtRefName.Text) Then
            StatusBar.SimpleText = "Adding ... " & txtRefName.Text
    
            Set cS = New clsScanner
            TrackName = Trim$(txtRefName.Text)
            
            If clib.CreateTrack(TrackName) Then
                markHidden = -chkCtlHidden.Value
                With cS
                    .Filter = Trim$(txtScanFilter.Text)
                    .PathToScan = txtScanPath.Text ' Left$(txtScanPath.Text, InStr(cmbScanDrive.Drive, Chr$(58))) & Chr$(92)
                    .RecurseSubdirectories = -chkRecurse.Value
                    .AddHiddenDirectories = -chkCtlHiddenDirs.Value
                    .ExpandArchives = -chkExpandArchives.Value
                    
                    .Scan
                End With
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
        MsgBox "Please Enter a name for the track" & vbCrLf & _
               "The track was not created", vbCritical Or vbOKOnly, "Error"
    End If
    
    txtScanPath.Enabled = True
    txtRefName.Enabled = True
    picScanOptions.Enabled = True
    cmdClose.Enabled = True
    cmdCreate.Enabled = True
    txtRefName.Text = vbNullString
    cmdBrowse.Enabled = True
    cmbType.Enabled = True
    tmrMarquee.Enabled = False
    cmdCancel.Visible = False
    DoEvents
    
    Set cS = Nothing
End Sub

Private Sub cS_ArchiveFileFound(ByVal ArchiveName As String, ByVal ArchivePath As String, ByVal File As String, Size As String, PackSize As Long, ByVal Compression As Single, FILETIME As Date, ByVal Attributes As Long)
Dim sHidden As Boolean
    If bCancel Then
        cS.CancelScan
    Else
        sHidden = ((Attributes And vbHidden) = vbHidden) And markHidden
        Call clib.AddRecord(ArchiveName & "//" & File, ArchivePath, Size, Attributes, sHidden)
    End If
End Sub

Private Sub cS_FileFound(ByVal FileName As String, isDirectory As Boolean)
Dim sHidden As Boolean
    With cS
        If bCancel Then
            .CancelScan
        Else
            sHidden = ((.Attributes And vbHidden) = vbHidden) And markHidden
            If Not isDirectory Then
                Call clib.AddRecord(FileName, .CurrentDirectory, .FormatSize, .Attributes, isDirectory, sHidden)
            Else
                Call clib.AddRecord(FileName, .CurrentDirectory, "N.A.", .Attributes, isDirectory, sHidden)
            End If
        End If
    End With
End Sub

Private Sub cS_ScanComplete()
Dim fileCount As Long
    
    If bCancel Then
        MsgBox "Scan Cancelled. " & txtScanPath.Text & " was not scanned into a track.", vbOKOnly Or vbInformation, "Cancelled!"
        bCancel = False
    
    Else
        Call cS_UpdateCaption("Comitting track...")
        fileCount = clib.SaveTrack(txtComment.Text, cmbType.Text, cS.FormatSize(cS.TotalScanSize), txtScanPath.Tag)
        
        StatusBar.SimpleText = "Scan Complete. Track created"
        
        Call cS_UpdateCaption("Scan complete!")
        If MsgBox("Scan Complete." & vbCrLf & vbCrLf & _
                  "Number of files added: " & fileCount & _
                  ". Would you like to add another track?", vbQuestion Or vbYesNo, "Scan Complete. Add Another?") = vbNo Then
            Unload Me
        Else
            lstStatusMessages.Clear
            lblFileCount.Caption = vbNullString
            TabStrip.Tabs(1).Selected = True
        End If
    End If

    Exit Sub
    
cS_ScanComplete_ERH:
    MsgBox "An internal error occured." & vbCrLf & _
           "Error Details..." & vbCrLf & _
           Space$(5) & "Error Number : " & Err.Number & vbCrLf & _
           Space$(5) & "Description  : " & Err.Description, vbOKOnly Or vbCritical, "Error!"
    Err.Clear
End Sub

Private Sub cS_UpdateCaption(ByVal Text As String)
    StatusBar.SimpleText = Text
    Text = ChrW$(91) & Format$(Now, "hh:mm:ss") & ChrW$(93) & Space$(2) & Text
    lstStatusMessages.AddItem Text
    Call SendMessage(lstStatusMessages.hWnd, LB_SETCURSEL, lstStatusMessages.NewIndex, ByVal 0)
    DoEvents
End Sub

Private Sub cS_UpdateProgress(ByVal currentFile As Long, ByVal currentFolder As Long)
    lblFileCount.Caption = (currentFile + currentFolder) & " items scanned."
    DoEvents
End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_ERH
Dim nStyle As Long

    Me.Caption = APP_NAME & " - Create New Track"
    tabStrip_Click
    
    nStyle = GetWindowLong(ProgressBar1.hWnd, GWL_STYLE)
    SetWindowLong ProgressBar1.hWnd, GWL_STYLE, nStyle Or PBS_MARQUEE
    
    Debug.Print cmbType.Style
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

Private Sub tabStrip_Click()
    With TabStrip
        If .SelectedItem.Index = 1 Then
            picScanOptions.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
            picScanOptions.ZOrder 0
        ElseIf .SelectedItem.Index = 2 Then
            txtComment.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
            txtComment.ZOrder 0
        Else
            pScanFeedback.Move .ClientLeft, .ClientTop, .ClientWidth, .ClientHeight
            pScanFeedback.ZOrder 0
            lblPleaseWait.Caption = "Please wait while " & APP_NAME & " is scanning. This may take a while."
            ProgressBar1.Value = 0
        End If
    End With
End Sub

Private Sub tmrMarquee_Timer()
    With ProgressBar1
        .Value = .Value + 1
        If .Value >= .Max Then .Value = .Min
    End With
End Sub
