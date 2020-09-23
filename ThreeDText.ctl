VERSION 5.00
Begin VB.UserControl ThreeDText 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF24FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4530
   MaskColor       =   &H00FF24FF&
   ScaleHeight     =   3360
   ScaleWidth      =   4530
   ToolboxBitmap   =   "ThreeDText.ctx":0000
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   375
      TabIndex        =   0
      Top             =   2910
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "ThreeDText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************************
'**                            Transparent 3D Text
'**                               Version 1.0.0
'**                               By Ken Foster
'**                                August 2005
'**                     Freeware--- no copyrights claimed
'*******************************************************************

Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, _
     ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, _
     ByVal lpDrawTextParams As Any) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
     ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
     
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Enum eDirection
   TopLeft = 1
   BottomRight = 2
   TopRight = 3
   BottomLeft = 4
End Enum

Const m_def_Caption = "3D Text"
Const m_def_ColorS = vbRed
Const m_def_ColorE = vbBlack
Const m_def_Direction = 1
Const m_def_Xoffset = 1
Const m_def_Yoffset = 1
Const m_def_Depth = 1


Private m_CaptionRect As RECT
Private m_Flag As Long

Private m_Caption As String
Private m_ColorS As OLE_COLOR
Private m_ColorE As OLE_COLOR
Private m_Direction As Integer
Private m_Xoffset As Integer
Private m_Yoffset As Integer
Private m_Depth As Integer

Dim R As Integer
Dim G As Integer
Dim B As Integer
Dim SRed As Integer
Dim SGreen As Integer
Dim SBlue As Integer
Dim ERed As Integer
Dim EGreen As Integer
Dim EBlue As Integer
Dim ct As Integer
Event Click()

Private Sub UserControl_Initialize()
      m_Depth = m_def_Depth
      m_Xoffset = m_def_Xoffset
      m_Yoffset = m_def_Yoffset
      m_ColorS = m_def_ColorS
      m_ColorE = m_def_ColorE
      m_Direction = m_def_Direction
End Sub

Private Sub UserControl_InitProperties()
     Caption = Extender.Name                                 'assigns Caption name of usercontrol
     UserControl.FontSize = 30                               'font size as a start
     UserControl.FontBold = True                             'make it bold so it is easier to read caption
     ColorS = vbRed
     ColorE = vbBlack
     Direction = 1
     Depth = 1
     ReSize
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Enter text to describe action."
     Caption = m_Caption
End Property

Public Property Let Caption(NewCaption As String)
     m_Caption = NewCaption
     PropertyChanged "Caption"
     Doit Direction
End Property

Public Property Get Depth() As Integer
   Depth = m_Depth
End Property

Public Property Let Depth(NewDepth As Integer)
   If NewDepth <= 0 Then Exit Property                       'no negative numbers
   m_Depth = NewDepth
   PropertyChanged "Depth"
   Doit Direction
End Property

Public Property Get Direction() As eDirection
Attribute Direction.VB_Description = "Skew of  3D text"
     Direction = m_Direction
End Property

Public Property Let Direction(ByVal NewDirection As eDirection)
     m_Direction = NewDirection
     PropertyChanged "Direction"
     Doit Direction
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Selects font to display text"
     Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal NewFont As Font)
     Set UserControl.Font = NewFont
     PropertyChanged "Font"
     ReSize
     Doit Direction
End Property

Public Property Get ColorS() As OLE_COLOR
Attribute ColorS.VB_Description = "Start color of gradient"
     ColorS = m_ColorS
End Property

Public Property Let ColorS(ByVal NewColorS As OLE_COLOR)
     m_ColorS = NewColorS
     PropertyChanged "ColorS"
     Doit Direction
End Property

Public Property Get ColorE() As OLE_COLOR
Attribute ColorE.VB_Description = "End color of gradient"
     ColorE = m_ColorE
End Property

Public Property Let ColorE(ByVal NewColorE As OLE_COLOR)
     m_ColorE = NewColorE
     PropertyChanged "ColorE"
   Doit Direction
End Property

Public Property Get Xoffset() As Integer
   Xoffset = m_Xoffset
End Property

Public Property Let Xoffset(NewXoffset As Integer)
   If NewXoffset <= 0 Then Exit Property                     'no negative numbers
   m_Xoffset = NewXoffset
   PropertyChanged "Xoffset"
   Doit Direction
End Property

Public Property Get Yoffset() As Integer
   Yoffset = m_Yoffset
End Property

Public Property Let Yoffset(NewYoffset As Integer)
   If Yoffset <= 0 Then Exit Property                        'no negative numbers
   m_Yoffset = NewYoffset
   PropertyChanged "Yoffset"
   Doit Direction
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     With PropBag
          Caption = .ReadProperty("Caption", Extender.Name)
          ColorS = .ReadProperty("ColorS", m_def_ColorS)
          ColorE = .ReadProperty("ColorE", m_def_ColorE)
          Direction = .ReadProperty("Direction", m_def_Direction)
          Xoffset = PropBag.ReadProperty("Xoffset", m_def_Xoffset)
          Yoffset = PropBag.ReadProperty("Yoffset", m_def_Yoffset)
          Depth = PropBag.ReadProperty("Depth", m_def_Depth)
          Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
     End With
     Doit Direction
End Sub

Private Sub UserControl_Resize()
   ReSize
   Doit Direction
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     With PropBag
          Call .WriteProperty("Caption", m_Caption, Extender.Name)
          Call .WriteProperty("ColorS", m_ColorS, m_def_ColorS)
          Call .WriteProperty("ColorE", m_ColorE, m_def_ColorE)
          Call .WriteProperty("Direction", m_Direction, m_def_Direction)
          Call .WriteProperty("Xoffset", m_Xoffset, m_def_Xoffset)
          Call .WriteProperty("Yoffset", m_Yoffset, m_def_Yoffset)
          Call .WriteProperty("Depth", m_Depth, m_def_Depth)
          Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
     End With
End Sub

Private Sub Doit(Index As Integer)
   Dim i As Integer
   Dim X As Single
   Dim Y As Single
   Dim RChange As Integer
   Dim GChange As Integer
   Dim BChange As Integer
   
   On Error Resume Next
   UserControl.Cls
   GetColor ColorS, 0, 0, 0, ColorE, 0, 0, 0
   For i = 0 To 254 Step Depth
   
      Select Case Index
        Case 1                                               'TopLeft
            X = X - Xoffset
            Y = Y - Yoffset
            UserControl.CurrentX = (300 * Xoffset) + X       'determines where text will print (Right/Left)
            UserControl.CurrentY = Y + (300 * Yoffset)       'also used to determine where text will print (Up/Down)
         Case 2                                              'BottomRight
            X = X + Xoffset
            Y = Y + Yoffset
            UserControl.CurrentX = (50 * Xoffset) + X        'determines where text will print (Right/Left)
            UserControl.CurrentY = Y + (50 * Yoffset)        'also used to determine where text will print (Up/Down)
         Case 3                                              'TopRight
            X = X + Xoffset
            Y = Y - Yoffset
            UserControl.CurrentX = (50 * Xoffset) + X        'determines where text will print (Right/Left)
            UserControl.CurrentY = Y + (300 * Yoffset)       'also used to determine where text will print (Up/Down)
         Case 4                                              'BottomLeft
            X = X - Xoffset
            Y = Y + Yoffset
            UserControl.CurrentX = (300 * Xoffset) + X       'determines where text will print (Right/Left)
            UserControl.CurrentY = Y + (20 * Yoffset)        'also used to determine where text will print (Up/Down)
      End Select
      
      RChange = RChange + (ERed - SRed) / 255                'start of gradient colors
      GChange = GChange + (EGreen - SGreen) / 255
      BChange = BChange + (EBlue - SBlue) / 255
      R = SRed + RChange
      G = SGreen + GChange
      B = SBlue + BChange
      UserControl.ForeColor = RGB(R, G, B)                   'set text color
      
      If i >= 240 Then UserControl.ForeColor = m_ColorE      'adds a shadow effect with minor adjustment so all match TopLeft button
      If i >= 250 Then UserControl.ForeColor = m_ColorS      'highlights start text
   
      UserControl.Print m_Caption
      Next
      UserControl.MaskPicture = UserControl.Image
      ReSize
   End Sub

Private Sub GetColor(ByVal LngCol As Long, R1 As Integer, G1 As Integer, B1 As Integer, LngCol1 As Long, R2 As Integer, G2 As Integer, B2 As Integer)
   R1 = LngCol Mod 256
   G1 = (LngCol And vbGreen) / 256
   B1 = (LngCol And vbBlue) / 65536
   
   R2 = LngCol1 Mod 256
   G2 = (LngCol1 And vbGreen) / 256
   B2 = (LngCol1 And vbBlue) / 65536
   
   SRed = R2
   SGreen = G2
   SBlue = B2
   ERed = R1
   EGreen = G1
   EBlue = B1
End Sub

Private Sub ReSize()
     With Label1
        .Caption = m_Caption
        .FontSize = UserControl.FontSize
        .Font = UserControl.Font
      End With
      
      UserControl.Width = Label1.Width + (450 * Xoffset)
      UserControl.Height = Label1.Height + (300 * Yoffset)
End Sub
