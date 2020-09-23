VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   1680
      TabIndex        =   14
      Top             =   1440
      Width           =   2895
   End
   Begin VB.TextBox txtTrackName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   13
      Text            =   "Track Name"
      Top             =   195
      Width           =   3855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdShowResults 
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
      Height          =   380
      Left            =   4200
      TabIndex        =   10
      Top             =   3075
      Width           =   375
   End
   Begin VB.TextBox txtSearch 
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
      Left            =   1680
      TabIndex        =   7
      Top             =   3120
      Width           =   2505
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
      Height          =   975
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Category :"
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
      Index           =   3
      Left            =   0
      TabIndex        =   15
      Top             =   1470
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   2
      X1              =   120
      X2              =   4560
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   120
      X2              =   4560
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   4575
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   120
      X2              =   4575
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblSearchCue 
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
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
      Left            =   1680
      TabIndex        =   9
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Find in track :"
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
      Index           =   4
      Left            =   600
      TabIndex        =   8
      Top             =   3165
      Width           =   975
   End
   Begin VB.Image imgSearchIcon 
      Height          =   240
      Left            =   240
      Picture         =   "frmProperties.frx":0000
      Stretch         =   -1  'True
      Top             =   3135
      Width           =   240
   End
   Begin VB.Label lblNumRecs 
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
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
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lblCreateDate 
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
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
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Comment :"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of records :"
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
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Created on :"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblArchiveName1 
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
      Caption         =   "Archive Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgIco 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SetProperties(ByVal aTrack As String)
Dim aCategory As String, aCreated As Date, aNumRec As Long, aComment As String, aName As String, aSize As String, aBaseNode As String
    
    aName = clib.GetTrackDetails(aTrack, aCategory, aCreated, aNumRec, aComment, aSize, aBaseNode)
    
    txtTrackName.Text = aName
    txtTrackName.Tag = aName
    lblCreateDate.Caption = aCreated
    lblNumRecs.Caption = aNumRec
    txtComment.Text = Replace$(aComment, vbLf, vbCrLf)
    cmbType.Text = aCategory
    cmdOK.Enabled = False
    
    Caption = aName & " - Track Properties"
End Sub

Private Sub cmbType_Change()
    If Not cmdOK.Enabled Then cmdOK.Enabled = True
End Sub

Private Sub cmbType_Click()
    If Not cmdOK.Enabled Then cmdOK.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If clib.SetTrackDetails(txtTrackName.Tag, txtTrackName.Text, cmbType.Text, txtComment.Text) Then
        cmdOK.Enabled = False
        Call frmMain.LoadTracks
    Else
        MsgBox "An error occured. " & vbCrLf & _
               "A possible reason for this is that an archive with the name '" & _
               txtTrackName.Text & "' already exists in the library. " & vbCrLf & _
               "The changes were not made", vbCritical Or vbOKOnly, "Error"
    End If
End Sub

Private Sub cmdShowResults_Click()
    Call frmMain.LoadSearchResults(CLng(Val(lblSearchCue.Caption)))
    frmMain.cmbSearchBox.Text = SEARCH_NOT
    DoEvents
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Long
    Caption = "Track Properties"
    
    Set imgIco.Picture = LoadResPicture(32001, vbResIcon)
    Set imgSearchIcon.Picture = LoadResPicture(16001, vbResIcon)
    
    For i = 0 To 9
        cmbType.AddItem LoadResString(i)
    Next
    cmbType.ListIndex = 9
End Sub

Private Sub txtTrackName_Change()
    If Not cmdOK.Enabled Then cmdOK.Enabled = True
End Sub

Private Sub txtComment_Change()
    If Not cmdOK.Enabled Then cmdOK.Enabled = True
End Sub

Private Sub txtSearch_Change()
Dim srchStr As String
Dim numHits As Long
    
    srchStr = txtSearch.Text
    If (srchStr <> vbNullString) And (srchStr <> SEARCH_NOT) Then
        If (Left$(srchStr, 2) <> SEARCH_NOT) Then
            numHits = clib.FindInTrack(txtTrackName.Tag, scContains, UCase$(srchStr))
        Else
            Debug.Print "Searching NOT"
            numHits = clib.FindInTrack(txtTrackName.Tag, scDoesNotContain, Mid$(UCase$(srchStr), 3))
        End If
    End If
    lblSearchCue.Caption = numHits & " possible matches found."
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdShowResults_Click
End Sub
