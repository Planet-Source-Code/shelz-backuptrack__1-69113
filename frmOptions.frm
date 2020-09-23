VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Interface Options"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5880
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
   ScaleHeight     =   4455
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFont 
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset Colors"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   3180
      Width           =   2535
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4560
      TabIndex        =   12
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdFgCol 
      Caption         =   "Foreground Color"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdBgCol 
      Caption         =   "Background Color"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2940
      Width           =   1935
   End
   Begin ComctlLib.Slider sldFontSize 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   2400
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   327682
      Min             =   6
      Max             =   72
      SelStart        =   6
      Value           =   6
   End
   Begin VB.CheckBox chkDisplayDbInTitlebar 
      Caption         =   "Display database filename in title bar"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   4815
   End
   Begin VB.CheckBox chkHasCaptions 
      Caption         =   "Interface panel captions"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   4815
   End
   Begin VB.CheckBox chkToolbarCaptions 
      Caption         =   "Toolbar captions"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1380
      Width           =   4815
   End
   Begin VB.CheckBox chkGridlines 
      Caption         =   "Gridlines in search, track and file lists"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   4815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   5760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Shape shFgCol 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2520
      Top             =   3420
      Width           =   255
   End
   Begin VB.Shape shBgCol 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2520
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblTestFontSize 
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Colors and Fonts"
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
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Interface Appearance"
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
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "General"
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
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   135
      X2              =   5684
      Y1              =   3860
      Y2              =   3860
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBgCol_Click()
    With frmMain.cmDlg
        .CancelError = False
        .color = shBgCol.FillColor
        .ShowColor
        
        shBgCol.FillColor = .color
    End With
End Sub

Private Sub cmdFgCol_Click()
    With frmMain.cmDlg
        .CancelError = False
        .color = shFgCol.FillColor
        .ShowColor
        
        shFgCol.FillColor = .color
    End With
End Sub

Private Sub cmdFont_Click()
    With frmMain.cmDlg
        .CancelError = False
        .flags = cdlCFBoth
        .FontName = cmdFont.FontName
        .ShowFont
        
        cmdFont.Caption = .FontName & "..."
        cmdFont.FontName = .FontName
        
        frmMain.LvwBrowser.Font.Name = .FontName
        frmMain.LvwData.Font.Name = .FontName
        frmMain.TvwBrowser.Font.Name = .FontName
        frmMain.TvwData.Font.Name = .FontName
    End With
End Sub

Private Sub cmdOK_Click()
    With cConfig
        .DisplayDbInTitle = CBool(chkDisplayDbInTitlebar.Value)
        
        .Gridlines = CBool(chkGridlines.Value)
        Call ObjectGridlines(frmMain.LvwData.hWnd, .Gridlines)
        Call ObjectGridlines(frmMain.LvwBrowser.hWnd, .Gridlines)
        
        .ToolbarCaptions = CBool(chkToolbarCaptions)
        .InterfaceCaptions = CBool(chkHasCaptions.Value)
        frmMain.cSplitterMain.HasCaptions = .InterfaceCaptions
        frmMain.cSplitterMain.SplitterOrientation = socOrientEastWest
        
        .FontSize = sldFontSize.Value
        .FontName = cmdFont.FontName
        frmMain.LvwBrowser.Font.Name = .FontName
        frmMain.LvwData.Font.Name = .FontName
        frmMain.TvwBrowser.Font.Name = .FontName
        frmMain.TvwData.Font.Name = .FontName
        
        .IBackColor = shBgCol.FillColor
        Call SendMessage(frmMain.TvwBrowser.hWnd, TVM_SETBKCOLOR, 0, ByVal Translate(.IBackColor))
        Call SendMessage(frmMain.TvwData.hWnd, TVM_SETBKCOLOR, 0, ByVal Translate(.IBackColor))
        
        frmMain.LvwBrowser.BackColor = .IBackColor
        frmMain.LvwData.BackColor = .IBackColor
        frmMain.picSmallIcon.BackColor = .IBackColor
    
        .IForeColor = shFgCol.FillColor
        Call SendMessage(frmMain.TvwBrowser.hWnd, TVM_SETTEXTCOLOR, 0, ByVal Translate(.IForeColor))
        Call SendMessage(frmMain.TvwData.hWnd, TVM_SETTEXTCOLOR, 0, ByVal Translate(.IForeColor))
        frmMain.LvwBrowser.ForeColor = .IForeColor
        frmMain.LvwData.ForeColor = .IForeColor
    End With
    
    MsgBox "Certain changes will take place after an application restart.", vbInformation Or vbOKOnly, "Application Restart Needed"
    
    Unload Me
End Sub

Private Sub cmdSysInfo_Click()
    With frmMain
        .LvwBrowser.Font.Size = cConfig.FontSize
        .LvwData.Font.Size = cConfig.FontSize
        .TvwBrowser.Font.Size = cConfig.FontSize
        .TvwData.Font.Size = cConfig.FontSize
        
        .LvwBrowser.Font.Name = cConfig.FontName
        .LvwData.Font.Name = cConfig.FontName
        .TvwBrowser.Font.Name = cConfig.FontName
        .TvwData.Font.Name = cConfig.FontName
    End With
    
    Unload Me
End Sub

Private Sub Command1_Click()
    shBgCol.FillColor = vbWhite
    shFgCol.FillColor = vbBlack
End Sub

Private Sub Form_Load()
    '// Load existing configuration
    chkDisplayDbInTitlebar.Value = Abs(cConfig.DisplayDbInTitle)
    chkGridlines.Value = Abs(cConfig.Gridlines)
    chkToolbarCaptions.Value = Abs(cConfig.ToolbarCaptions)
    chkHasCaptions.Value = Abs(cConfig.InterfaceCaptions)
    
    sldFontSize.Value = cConfig.FontSize
    lblTestFontSize.Caption = cConfig.FontSize
    
    shBgCol.FillColor = cConfig.IBackColor
    shFgCol.FillColor = cConfig.IForeColor
    
    cmdFont.Caption = cConfig.FontName & "..."
    cmdFont.FontName = cConfig.FontName
End Sub

Private Sub sldFontSize_Scroll()
    With frmMain
        .LvwBrowser.Font.Size = sldFontSize.Value
        .LvwData.Font.Size = sldFontSize.Value
        .TvwBrowser.Font.Size = sldFontSize.Value
        .TvwData.Font.Size = sldFontSize.Value
    End With
    lblTestFontSize.Caption = sldFontSize.Value
End Sub
