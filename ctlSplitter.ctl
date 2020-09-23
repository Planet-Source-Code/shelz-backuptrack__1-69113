VERSION 5.00
Begin VB.UserControl ctlSplitter 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2775
   ScaleWidth      =   720
   ToolboxBitmap   =   "ctlSplitter.ctx":0000
   Begin VB.PictureBox pCaption 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   1
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pCaption 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Index           =   0
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox lblResize 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   315
      MousePointer    =   99  'Custom
      ScaleHeight     =   255
      ScaleWidth      =   90
      TabIndex        =   0
      Top             =   1260
      Width           =   90
   End
   Begin VB.Image imgCur 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "ctlSplitter.ctx":0312
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCur 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "ctlSplitter.ctx":0464
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "ctlSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CTL_STP = 75
Private Const CAPTION_HT = 285 '255
Private Const SPLIT_LINE = 30 '(CTL_STP - 15) \ 2
Private Const SPLIT_SPAN = 20 ' 30
Private Const SPLIT_MINMAX = 500
Private Const SPLIT_COLOR_HILITE = &H3CC8FF
Private Const SPLIT_CHRS_HRZ = "—" '"• • •" '"—" '
Private Const SPLIT_CHRS_VRT = "|" '"· · ·"

Private Const BS_HATCHED = 2
Private Const HS_HORIZONTAL = 0
Private Const HS_VERTICAL = 1
Private Const SPLITBAR_SPAN = 45
'Private Const HS_CROSS = 4
'Private Const HS_DIAGCROSS = 5
'Private Const HS_FDIAGONAL = 2
'Private Const HS_BDIAGONAL = 3



Private Const PATCOPY = &HF00021


Private Const CONTAINER_1 = 0
Private Const CONTAINER_2 = 1

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)

Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BF_FLAT = &H4000
Private Const BF_ADJUST = &H2000
Private Const BF_MONO = &H8000



Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10

Private Const DC_GRADIENT = &H20
Private Const DC_INBUTTON = &H10
Private Const DC_SMALLCAP = &H2

Private Const COLOR_3DDKSHADOW = 21
Private Const COLOR_3DLIGHT = 22
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNSHADOW = 16
Private Const COLOR_GRADIENTACTIVECAPTION = 27
Private Const COLOR_GRADIENTINACTIVECAPTION = 28
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_WINDOW = 5





Private Const def_SplitterOrientation = 0
Private Const def_SplitterPosition = 0
Private Const def_ContainerBorderStyle = 0
Private Const def_PanelEdgeStyle = 0
Private Const def_Container1Enabled = True
Private Const def_Container2Enabled = True
Private Const def_Enabled = True
Private Const def_HasCaptions = False
Private Const def_Caption1 = "Caption 1"
Private Const def_Caption2 = "Caption 2"

Private Const def_Client1Width = 0
Private Const def_Client1Height = 0
Private Const def_Client1Left = 0
Private Const def_Client1Top = 0
Private Const def_Client2Width = 0
Private Const def_Client2Height = 0
Private Const def_Client2Left = 0
Private Const def_Client2Top = 0

Public Enum PanelEdgeStyleConstants
    pescNone = 0
    pescRaisedInner = BDR_RAISEDINNER
    pescRaisedOuter = BDR_RAISEDOUTER
    pescSunkenInner = BDR_SUNKENINNER
    pescSunkenOuter = BDR_SUNKENOUTER
    pescBump = EDGE_BUMP
    pescEtched = EDGE_ETCHED
    pescRaised = EDGE_RAISED
End Enum

Public Enum SplitterOrientationConstants
    socOrientEastWest = 0
    socOrientNorthSouth = 1
End Enum

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long
Private Declare Function SetCapture Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Event Split()
Event Panel1Resize(ByVal pLeft As Long, ByVal pTop As Long, ByVal pWidth As Long, ByVal pHeight As Long)
Event Panel2Resize(ByVal pLeft As Long, ByVal pTop As Long, ByVal pWidth As Long, ByVal pHeight As Long)

Private m_SplitterOrientation   As SplitterOrientationConstants
Private m_PanelEdgeStyle        As PanelEdgeStyleConstants
Private m_SplitterMin           As Single
Private m_SplitterPosition      As Single
Private m_MouseIn               As Boolean
Private m_HasCaptions           As Boolean
Private m_CaptionOffset         As Long
Private m_Captions(1)           As String
Private m_pContainer(1)         As RECT
Private m_splitTxt              As RECT


'// Resize Options
Private vResize As Boolean, pShift As Long, selRval As Long

'Property Variables:
Dim m_Client1Width As Long
Dim m_Client1Height As Long
Dim m_Client1Left As Long
Dim m_Client1Top As Long
Dim m_Client2Width As Long
Dim m_Client2Height As Long
Dim m_Client2Left As Long
Dim m_Client2Top As Long


Private Sub AdjustPanels()
'On Error GoTo AdjustPanels_ERH
    If m_SplitterOrientation = socOrientEastWest Then
        Call SetRect(m_pContainer(CONTAINER_1), 0, 0, ScaleWidth, lblResize.Top)
        Call SetRect(m_pContainer(CONTAINER_2), 0, lblResize.Top + CTL_STP, ScaleWidth, ScaleHeight - (lblResize.Top + CTL_STP))
    Else
        Call SetRect(m_pContainer(CONTAINER_1), 0, 0, lblResize.Left, ScaleHeight)
        Call SetRect(m_pContainer(CONTAINER_2), lblResize.Left + CTL_STP, 0, ScaleWidth - (lblResize.Left + CTL_STP), ScaleHeight)
    End If

    With m_pContainer(CONTAINER_1)
        If m_HasCaptions And (m_SplitterOrientation = socOrientEastWest) Then
            RaiseEvent Panel1Resize(.Left, .Top + CAPTION_HT, .Right, .Bottom - CAPTION_HT)
            pCaption(CONTAINER_1).Move .Left, .Top, .Right, .Top + CAPTION_HT
        Else
            RaiseEvent Panel1Resize(.Left, .Top, .Right, .Bottom)
        End If
        'DoEvents
    End With
    
    With m_pContainer(CONTAINER_2)
        If m_HasCaptions And (m_SplitterOrientation = socOrientEastWest) Then
            RaiseEvent Panel2Resize(.Left, .Top + CAPTION_HT, .Right, .Bottom - CAPTION_HT)
            pCaption(CONTAINER_2).Move .Left, .Top, .Right, CAPTION_HT
        Else
            RaiseEvent Panel2Resize(.Left, .Top, .Right, .Bottom)
        End If
        'DoEvents
    End With
    
AdjustPanels_ERH:
    Err.Clear
End Sub

Private Sub OrientSplitter()
Dim pT As POINTAPI
    '// Orient the Panels
    If m_SplitterOrientation = socOrientEastWest Then
        lblResize.Move 0, m_SplitterPosition, ScaleWidth, CTL_STP
    Else
        lblResize.Move m_SplitterPosition, 0, CTL_STP, ScaleHeight
    End If
    Call PropertyChanged("SplitterPosition")
End Sub


Public Property Let SplitterOrientation(newOrientation As SplitterOrientationConstants)

    m_SplitterOrientation = newOrientation
    
    If m_SplitterOrientation = socOrientEastWest Then
        m_SplitterPosition = (ScaleHeight - CTL_STP) \ 2
    Else
        m_SplitterPosition = (ScaleWidth - CTL_STP) \ 2
    End If
    
    Call UserControl_Resize
    Call PropertyChanged("SplitterOrientation")
End Property

Public Property Get SplitterOrientation() As SplitterOrientationConstants
    SplitterOrientation = m_SplitterOrientation
End Property



Public Property Let SplitterPosition(newSplitterPosition As Single)

    If newSplitterPosition >= SPLIT_MINMAX Then
        m_SplitterPosition = newSplitterPosition
        Call UserControl_Resize
    Else
        MsgBox "This value must always be greater than " & SPLIT_MINMAX, vbCritical Or vbOKOnly, "ActiveX control error!"
    End If
End Property

Public Property Get SplitterPosition() As Single
    SplitterPosition = m_SplitterPosition
End Property


Public Property Let SplitterMin(newSplitterMin As Single)
    If newSplitterMin >= SPLIT_MINMAX Then
        m_SplitterMin = newSplitterMin
        PropertyChanged "SplitterMin"
        
        Call UserControl_Resize
    Else
        MsgBox "This value must always be greater than " & SPLIT_MINMAX, vbCritical Or vbOKOnly, "ActiveX control error!"
    End If
End Property

Public Property Get SplitterMin() As Single
    SplitterMin = m_SplitterMin
End Property


Public Property Let ContainerBorderStyle(newContainerBorderStyle As Integer)
    UserControl.BorderStyle = newContainerBorderStyle
    PropertyChanged "ContainerBorderStyle"
End Property

Public Property Get ContainerBorderStyle() As Integer
    ContainerBorderStyle = UserControl.BorderStyle
End Property


Public Property Get Client1Width() As Long
    Client1Width = m_pContainer(CONTAINER_1).Right
End Property


Public Property Get Client1Height() As Long
    Client1Height = m_pContainer(CONTAINER_1).Bottom
End Property


Public Property Get Client1Left() As Long
    Client1Left = m_pContainer(CONTAINER_1).Left
End Property


Public Property Get Client1Top() As Long
    Client1Top = m_pContainer(CONTAINER_1).Top
End Property


Public Property Get Client2Width() As Long
    Client2Width = m_pContainer(CONTAINER_2).Right
End Property


Public Property Get Client2Height() As Long
    Client2Height = m_pContainer(CONTAINER_2).Bottom
End Property


Public Property Get Client2Left() As Long
    Client2Left = m_pContainer(CONTAINER_2).Left
End Property


Public Property Get Client2Top() As Long
    Client2Top = m_pContainer(CONTAINER_2).Top
End Property


Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property


Public Property Get PanelEdgeStyle() As PanelEdgeStyleConstants
    PanelEdgeStyle = m_PanelEdgeStyle
End Property

Public Property Let PanelEdgeStyle(ByVal New_PanelEdgeStyle As PanelEdgeStyleConstants)
    m_PanelEdgeStyle = New_PanelEdgeStyle
    Call UserControl_Resize
    PropertyChanged "PanelEdgeStyle"
End Property


Public Property Let HasCaptions(ByVal newHasCaptions As Boolean)
    m_HasCaptions = newHasCaptions
    If m_HasCaptions And (m_SplitterOrientation = socOrientEastWest) Then
        m_CaptionOffset = CAPTION_HT
        pCaption(CONTAINER_1).Visible = True
        pCaption(CONTAINER_2).Visible = True
    Else
        m_CaptionOffset = 0
        pCaption(CONTAINER_1).Visible = False
        pCaption(CONTAINER_2).Visible = False
    End If
    
    PropertyChanged "HasCaptions"
End Property

Public Property Get HasCaptions() As Boolean
    HasCaptions = m_HasCaptions
End Property


Public Property Let Caption1(ByVal newCaption1 As String)
    m_Captions(CONTAINER_1) = newCaption1
    Call pCaption_Paint(CONTAINER_1)
    PropertyChanged "Caption1"
End Property

Public Property Get Caption1() As String
    Caption1 = m_Captions(CONTAINER_1)
End Property


Public Property Let Caption2(ByVal newCaption2 As String)
    m_Captions(CONTAINER_2) = newCaption2
    Call pCaption_Paint(CONTAINER_2)
    PropertyChanged "Caption2"
End Property

Public Property Get Caption2() As String
    Caption2 = m_Captions(CONTAINER_2)
End Property



Private Sub lblResize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error GoTo lblResize_MouseDown_ERH
    If Button = vbLeftButton Then
        vResize = True
        If m_SplitterOrientation = socOrientEastWest Then
            pShift = y
        Else
            pShift = x
        End If
        lblResize.ZOrder 0
    End If
lblResize_MouseDown_ERH:
End Sub

Private Sub lblResize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'On Error GoTo lblResize_MouseMove_ERH
Dim tmpHt As Long
    
    If vResize Then
        If m_SplitterOrientation = socOrientEastWest Then
            selRval = lblResize.Top + y - pShift
            Select Case selRval
                Case Is < m_SplitterMin + m_CaptionOffset
                    selRval = m_SplitterMin + m_CaptionOffset
    
                Case Is > ScaleHeight - SPLIT_MINMAX
                    selRval = ScaleHeight - SPLIT_MINMAX
            End Select
        Else
            selRval = lblResize.Left + x - pShift
            Select Case selRval
                Case Is < m_SplitterMin + m_CaptionOffset
                    selRval = m_SplitterMin + m_CaptionOffset
    
                Case Is > ScaleWidth - SPLIT_MINMAX
                    selRval = ScaleWidth - SPLIT_MINMAX
            End Select
        End If
        m_SplitterPosition = selRval
        Call UserControl_Resize
    Else
        If ((x < 0) Or (x > lblResize.ScaleWidth) Or (y < 0) Or (y > lblResize.ScaleHeight)) Then
            lblResize.BackColor = vbButtonFace
            Call ReleaseCapture
            m_MouseIn = False
            Call lblResize_Paint
        Else
            If Not m_MouseIn Then
                Call SetCapture(lblResize.hWnd)
                lblResize.BackColor = vb3DLight
                m_MouseIn = True
                Call lblResize_Paint
            End If
        End If
    End If

    Exit Sub
lblResize_MouseMove_ERH:
End Sub

Private Sub lblResize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If vResize Then
        vResize = False
        lblResize.BackColor = vbButtonFace
        m_MouseIn = False
    End If
End Sub

Private Sub lblResize_Paint()
Dim mPoint As Long 'SPLIT_SPAN
    '// Draw the edges if necessary
    With lblResize
        .Cls
        Call GetClientRect(.hWnd, m_splitTxt)
        If m_SplitterOrientation = socOrientEastWest Then
            mPoint = m_splitTxt.Right \ 2
            SetRect m_splitTxt, mPoint - SPLIT_SPAN, 0, mPoint + SPLIT_SPAN, m_splitTxt.Bottom
        Else
            mPoint = m_splitTxt.Bottom \ 2
            SetRect m_splitTxt, 0, mPoint - SPLIT_SPAN, m_splitTxt.Right, mPoint + SPLIT_SPAN
        End If
        Call FillRect(.hdc, m_splitTxt, GetSysColorBrush(COLOR_BTNSHADOW)) 'COLOR_BTNSHADOW COLOR_3DDKSHADOW
        Call FrameRect(.hdc, m_splitTxt, GetSysColorBrush(COLOR_BTNFACE))
    End With
End Sub

Private Sub pCaption_Paint(Index As Integer)
Dim xRect As RECT
    With pCaption(Index)
        .Cls
        Call GetClientRect(.hWnd, xRect)
        Call FillRect(.hdc, xRect, GetSysColorBrush(COLOR_ACTIVECAPTION))
        Call DrawEdge(.hdc, xRect, EDGE_BUMP, BF_RECT Or BF_FLAT)
        If Index = CONTAINER_1 Then
            Call DrawText(.hdc, Space$(3) & m_Captions(CONTAINER_1) & vbNullChar, -1, xRect, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER)
        Else
            Call DrawText(.hdc, Space$(3) & m_Captions(CONTAINER_2) & vbNullChar, -1, xRect, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER)
        End If
    End With
End Sub

Private Sub UserControl_InitProperties()
    SplitterOrientation = def_SplitterOrientation
    m_Client1Width = def_Client1Width
    m_Client1Height = def_Client1Height
    m_Client1Left = def_Client1Left
    m_Client1Top = def_Client1Top
    m_Client2Width = def_Client2Width
    m_Client2Height = def_Client2Height
    m_Client2Left = def_Client2Left
    m_Client2Top = def_Client2Top
    'm_ContainerBorderStyle = def_ContainerBorderStyle
    m_PanelEdgeStyle = def_PanelEdgeStyle
End Sub

Private Sub UserControl_Resize()
    If m_SplitterOrientation = socOrientEastWest Then
        If m_SplitterPosition + CTL_STP > ScaleHeight Then
            lblResize.Top = ScaleHeight - SPLIT_MINMAX
            Call AdjustPanels
            Exit Sub
        End If
    Else
        If m_SplitterPosition + CTL_STP > ScaleWidth Then
            lblResize.Left = ScaleWidth - SPLIT_MINMAX
            Call AdjustPanels
            Exit Sub
        End If
    End If

    Call OrientSplitter
    Call AdjustPanels
    Call lblResize_Paint
End Sub

'// Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_SplitterOrientation = .ReadProperty("SplitterOrientation", def_SplitterOrientation)
        m_SplitterPosition = .ReadProperty("SplitterPosition", def_SplitterPosition)
        
        UserControl.BorderStyle = .ReadProperty("ContainerBorderStyle", UserControl.BorderStyle)
        UserControl.Enabled = .ReadProperty("Enabled", def_Enabled)
    End With
    
    If m_SplitterOrientation = socOrientEastWest Then
        m_SplitterMin = SPLIT_MINMAX
    Else
        m_SplitterMin = SPLIT_MINMAX
    End If
    
    Set lblResize.MouseIcon = imgCur(m_SplitterOrientation)
    Call UserControl_Resize
    
    '// It is black to allow developers to place controls easily :)
    If Ambient.UserMode Then lblResize.BackColor = vbButtonFace
    
    m_Client1Width = PropBag.ReadProperty("Client1Width", def_Client1Width)
    m_Client1Height = PropBag.ReadProperty("Client1Height", def_Client1Height)
    m_Client1Left = PropBag.ReadProperty("Client1Left", def_Client1Left)
    m_Client1Top = PropBag.ReadProperty("Client1Top", def_Client1Top)
    m_Client2Width = PropBag.ReadProperty("Client2Width", def_Client2Width)
    m_Client2Height = PropBag.ReadProperty("Client2Height", def_Client2Height)
    m_Client2Left = PropBag.ReadProperty("Client2Left", def_Client2Left)
    m_Client2Top = PropBag.ReadProperty("Client2Top", def_Client2Top)
    m_SplitterMin = PropBag.ReadProperty("SplitterMin", SPLIT_MINMAX)
    m_PanelEdgeStyle = PropBag.ReadProperty("PanelEdgeStyle", def_PanelEdgeStyle)
    m_HasCaptions = PropBag.ReadProperty("HasCaptions", def_HasCaptions)
    m_Captions(CONTAINER_1) = PropBag.ReadProperty("Caption1", def_Caption1)
    m_Captions(CONTAINER_2) = PropBag.ReadProperty("Caption2", def_Caption2)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
    If m_HasCaptions And (m_SplitterOrientation = socOrientEastWest) Then
        m_CaptionOffset = CAPTION_HT
        pCaption(CONTAINER_1).Visible = True
        pCaption(CONTAINER_2).Visible = True
    End If
End Sub

'// Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("SplitterOrientation", m_SplitterOrientation, def_SplitterOrientation)
        Call .WriteProperty("SplitterPosition", m_SplitterPosition, def_SplitterPosition)
        Call .WriteProperty("ContainerBorderStyle", UserControl.BorderStyle, def_ContainerBorderStyle)
        Call .WriteProperty("Enabled", UserControl.Enabled, def_Enabled)
        Call .WriteProperty("Client1Width", m_Client1Width, def_Client1Width)
        Call .WriteProperty("Client1Height", m_Client1Height, def_Client1Height)
        Call .WriteProperty("Client1Left", m_Client1Left, def_Client1Left)
        Call .WriteProperty("Client1Top", m_Client1Top, def_Client1Top)
        Call .WriteProperty("Client2Width", m_Client2Width, def_Client2Width)
        Call .WriteProperty("Client2Height", m_Client2Height, def_Client2Height)
        Call .WriteProperty("Client2Left", m_Client2Left, def_Client2Left)
        Call .WriteProperty("Client2Top", m_Client2Top, def_Client2Top)
        Call .WriteProperty("SplitterMin", m_SplitterMin, SPLIT_MINMAX)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("PanelEdgeStyle", m_PanelEdgeStyle, def_PanelEdgeStyle)
        Call .WriteProperty("HasCaptions", m_HasCaptions, def_HasCaptions)
        Call .WriteProperty("Caption1", m_Captions(CONTAINER_1), def_Caption1)
        Call .WriteProperty("Caption2", m_Captions(CONTAINER_2), def_Caption2)
    End With
End Sub

