VERSION 5.00
Begin VB.UserControl TrillianFrame 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2250
   ControlContainer=   -1  'True
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   150
   ToolboxBitmap   =   "TrillianFrame.ctx":0000
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      ScaleHeight     =   1
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   96
      TabIndex        =   1
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label lblmain 
      BackColor       =   &H00E7A27B&
      BackStyle       =   0  'Transparent
      Caption         =   "Trillian Frame"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   1335
   End
   Begin VB.Shape Shpmain 
      BackColor       =   &H00E7A27B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E7A27B&
      FillColor       =   &H00E7A27B&
      Height          =   225
      Index           =   0
      Left            =   15
      Top             =   15
      Width           =   2220
   End
   Begin VB.Shape Shpmain 
      BorderColor     =   &H00E7A27B&
      BorderWidth     =   4
      Height          =   3255
      Index           =   1
      Left            =   45
      Top             =   285
      Width           =   2175
   End
   Begin VB.Shape Shpmain 
      BackColor       =   &H00E7A27B&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E7A27B&
      FillColor       =   &H00E7A27B&
      Height          =   15
      Index           =   2
      Left            =   15
      Top             =   240
      Width           =   2220
   End
End
Attribute VB_Name = "TrillianFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal color As Long) As Byte
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Const dark_blue As Long = 14053681
Const light_blu As Long = 15180411
Const dark_grey As Long = 14074813
Private Border As Long, Bind_Font As Boolean, Bind_Border As Boolean
Private Function minWidth() As Long
    minWidth = (Border + Shpmain(1).BorderWidth) * 2 + 5
End Function
Private Function minHeight() As Long
    minHeight = minWidth + picmain.Height + Shpmain(0).Height
End Function
Private Function SysToLNG(ByVal lColor As Long) As Long
'Special thanks to redbird77 for this code and realizing what the bug was
SysToLNG = lColor ' If hi-bit if hi-byte is set, then it is a system color.
If (lColor And &H80000000) Then SysToLNG = GetSysColor(lColor And &HFFFFFF)
End Function
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", UserControl.Name)
    BorderColor = PropBag.ReadProperty("BorderColor", light_blu)
    ForeColor = PropBag.ReadProperty("ForeColor", light_blu)
    BackColor = PropBag.ReadProperty("BackColor", vbWhite)
    FontColor = PropBag.ReadProperty("FontColor", vbWhite)
    BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    GradientHeight = PropBag.ReadProperty("GradientHeight", 1)
    GradientWidth = PropBag.ReadProperty("GradientWidth", 96)
    FontSize = PropBag.ReadProperty("FontSize", 8)
    FrameSize = PropBag.ReadProperty("FrameSize", 4)
    BindFontColor = PropBag.ReadProperty("BindFontColor", False)
    BindBorderColor = PropBag.ReadProperty("BindBorderColor", False)
End Sub

Public Property Let FrameSize(temp As Long)
    Shpmain(1).BorderWidth = temp + temp Mod 2
    UserControl_Resize
End Property
Public Property Get FrameSize() As Long
    FrameSize = Shpmain(1).BorderWidth
End Property

Public Property Let BindFontColor(temp As Boolean)
    Bind_Font = temp
    If temp Then FontColor = BackColor
End Property
Public Property Get BindFontColor() As Boolean
    BindFontColor = Bind_Font
End Property
Public Property Let BindBorderColor(temp As Boolean)
    Bind_Border = temp
    If temp Then BorderColor = ForeColor
End Property
Public Property Get BindBorderColor() As Boolean
    BindBorderColor = Bind_Border
End Property

Public Property Let FontSize(temp As Long)
    lblmain.FontSize = temp
    lblmain.Height = temp + 5
    Shpmain(0).Height = temp + 7
    UserControl_Resize
End Property
Public Property Get FontSize() As Long
    FontSize = lblmain.FontSize
End Property

Public Property Let BorderWidth(temp As Long)
    If temp < 0 Then temp = 0
    Border = temp
    UserControl_Resize
End Property
Public Property Get BorderWidth() As Long
    BorderWidth = Border
End Property

Public Property Let GradientHeight(temp As Long)
    If temp < 1 Then temp = 1
    picmain.Height = temp
    Shpmain(2).Height = temp
    UserControl_Resize
    RefreshGradient
End Property
Public Property Get GradientHeight() As Long
    GradientHeight = picmain.Height
End Property
Public Property Let GradientWidth(temp As Long)
    If temp < 1 Then temp = 1
    If temp > Shpmain(0).Width Then temp = Shpmain(0).Width
    picmain.Width = temp
    UserControl_Resize
    RefreshGradient
End Property
Public Property Get GradientWidth() As Long
    GradientWidth = picmain.Width
End Property

Public Property Let BorderColor(temp As OLE_COLOR)
    Shpmain(0).BorderColor = SysToLNG(temp)
End Property
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = Shpmain(0).BorderColor
End Property

Public Property Let ForeColor(ByVal temp As OLE_COLOR)
    temp = SysToLNG(temp)
    If Bind_Font Then BorderColor = temp
    Shpmain(0).BackColor = temp
    Shpmain(1).BorderColor = temp
    Shpmain(2).BackColor = temp
    Shpmain(2).BorderColor = temp
    Shpmain(2).FillColor = temp
    RefreshGradient
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Shpmain(0).BackColor
End Property
Public Property Get FontColor() As OLE_COLOR
    FontColor = lblmain.ForeColor
End Property
Public Property Let FontColor(temp As OLE_COLOR)
    lblmain.ForeColor = temp
End Property
Public Property Let BackColor(ByVal temp As OLE_COLOR)
    temp = SysToLNG(temp)
    If Bind_Font Then FontColor = temp
    Shpmain(1).BackColor = temp
    UserControl.BackColor = temp
    lblmain.ForeColor = temp
    RefreshGradient
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = Shpmain(1).BackColor
End Property
Public Sub RefreshGradient()
        Dim temp As Long, temp2 As Long, temp3 As Long
        picmain.Cls
        For temp = 0 To picmain.Width - 1
            temp2 = AlphaBlend(ForeColor, BackColor, (temp + 1) / picmain.Width)
            For temp3 = 0 To picmain.Height - 1
                SetPixelV picmain.hDC, temp, temp3, temp2
            Next
        Next
End Sub
Private Sub UserControl_Resize()
    Dim temp As Long, wid As Long
    wid = UserControl.Width / 15
    If wid < minWidth Then
        UserControl.Width = minWidth * 15
        wid = minWidth
    End If
    If UserControl.Height < minHeight * 15 Then UserControl.Height = minHeight * 15
    
    Shpmain(0).Move Border, Border, wid - (Border * 2)
    Shpmain(2).Move Border, Border + Shpmain(0).Height, Shpmain(0).Width
    picmain.Move Border, Shpmain(0).Height + Border
    lblmain.Move Border + 2, Border + 1, Shpmain(0).Width - 2
    temp = Shpmain(1).BorderWidth / 2
    Shpmain(1).Move Border + temp, picmain.Top + temp + picmain.Height, Shpmain(0).Width - Shpmain(1).BorderWidth + 1
    Shpmain(1).Height = (UserControl.Height / 15) - temp - Shpmain(1).Top - (Border - 1)
End Sub

Public Property Let Caption(text As String)
    lblmain.Caption = text
End Property
Public Property Get Caption() As String
    Caption = lblmain.Caption
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", lblmain.Caption, UserControl.Name
    PropBag.WriteProperty "BorderColor", Shpmain(0).BorderColor, light_blu
    PropBag.WriteProperty "ForeColor", Shpmain(0).BorderColor, light_blu
    PropBag.WriteProperty "BackColor", Shpmain(1).BackColor, vbWhite
    PropBag.WriteProperty "FontColor", lblmain.BackColor, vbWhite
    PropBag.WriteProperty "BorderWidth", Border, 1
    PropBag.WriteProperty "GradientHeight", picmain.Height, 1
    PropBag.WriteProperty "GradientWidth", picmain.Width, 96
    PropBag.WriteProperty "FontSize", lblmain.FontSize, 8
    PropBag.WriteProperty "FrameSize", Shpmain(1).BorderWidth, 4
    PropBag.WriteProperty "BindFontColor", Bind_Font, False
    PropBag.WriteProperty "BindBorderColor", Bind_Border, False
End Sub

Private Function Red(color As Long)
    Red = color Mod 256
End Function

Private Function Green(color As Long)
    Green = ((color And &HFF00) / 256) Mod 256
End Function

Private Function Blue(color As Long)
    Blue = (color And &HFF0000) / 65536
End Function

Private Function AlphaBlend(colorA As Long, colorB As Long, Alpha As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = blend(Red(colorA), Red(colorB), Alpha)
    g = blend(Green(colorA), Green(colorB), Alpha)
    b = blend(Blue(colorA), Blue(colorB), Alpha)
    AlphaBlend = RGB(r, g, b)
End Function

Private Function blend(colorA As Long, colorB As Long, Alpha As Double) As Long
    blend = Abs((colorA - colorB) * Alpha + colorB) Mod 256
End Function

