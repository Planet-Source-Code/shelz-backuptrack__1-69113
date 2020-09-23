VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmWait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please Wait"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMarquee 
      Interval        =   50
      Left            =   0
      Top             =   720
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Label lblCaption 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Image imgIco 
      Height          =   480
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function SetCaption(ByVal xCaption As String)
    lblCaption.Caption = xCaption
    Call tmrMarquee_Timer
    'DoEvents
End Function

Private Sub Form_Load()
Dim nStyle As Long
    Caption = APP_NAME
    Set imgIco.Picture = LoadResPicture(32001, vbResIcon)

    nStyle = GetWindowLong(ProgressBar1.hWnd, GWL_STYLE)
    SetWindowLong ProgressBar1.hWnd, GWL_STYLE, nStyle Or PBS_MARQUEE
    DoEvents
End Sub

Private Sub tmrMarquee_Timer()
    With ProgressBar1
        .Value = .Value + 1
        If .Value >= .Max Then .Value = .Min
    End With
End Sub
