Attribute VB_Name = "mdlGUI"
Public Const COLOR_3DDKSHADOW = 21
Public Const COLOR_3DFACE = 15
Public Const COLOR_3DHIGHLIGHT = 20
Public Const COLOR_3DHILIGHT = 20
Public Const COLOR_3DLIGHT = 22
Public Const COLOR_3DSHADOW = 16
Public Const COLOR_ACTIVEBORDER = 10
Public Const COLOR_ACTIVECAPTION = 2
Public Const COLOR_ADD = 712
Public Const COLOR_ADJ_MAX = 100
Public Const COLOR_ADJ_MIN = -100
Public Const COLOR_APPWORKSPACE = 12
Public Const COLOR_BACKGROUND = 1
Public Const COLOR_BLUE = 708
Public Const COLOR_BLUEACCEL = 728
Public Const COLOR_BOX1 = 720
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNHIGHLIGHT = 20
Public Const COLOR_BTNHILIGHT = 20
Public Const COLOR_BTNSHADOW = 16
Public Const COLOR_BTNTEXT = 18
Public Const COLOR_CAPTIONTEXT = 9
Public Const COLOR_CURRENT = 709
Public Const COLOR_CUSTOM1 = 721
Public Const COLOR_DESKTOP = 1
Public Const COLOR_ELEMENT = 716
Public Const COLOR_GRADIENTACTIVECAPTION = 27
Public Const COLOR_GRADIENTINACTIVECAPTION = 28
Public Const COLOR_GRAYTEXT = 17
Public Const COLOR_GREEN = 707
Public Const COLOR_GREENACCEL = 727
Public Const COLOR_HIGHLIGHT = 13
Public Const COLOR_HIGHLIGHTTEXT = 14
Public Const COLOR_HOTLIGHT = 26
Public Const COLOR_HUE = 703
Public Const COLOR_HUEACCEL = 723
Public Const COLOR_HUESCROLL = 700
Public Const COLOR_INACTIVEBORDER = 11
Public Const COLOR_INACTIVECAPTION = 3
Public Const COLOR_INACTIVECAPTIONTEXT = 19
Public Const COLOR_INFOBK = 24
Public Const COLOR_INFOTEXT = 23
Public Const COLOR_LUM = 705
Public Const COLOR_LUMACCEL = 725
Public Const COLOR_LUMSCROLL = 702
Public Const COLOR_MATCH_VERSION = &H200
Public Const COLOR_MENU = 4
Public Const COLOR_MENUTEXT = 7
Public Const COLOR_MIX = 719
Public Const COLOR_NO_TRANSPARENT = &HFFFFFFFF
Public Const COLOR_PALETTE = 718
Public Const COLOR_RAINBOW = 710
Public Const COLOR_RED = 706
Public Const COLOR_REDACCEL = 726
Public Const COLOR_SAMPLES = 717
Public Const COLOR_SAT = 704
Public Const COLOR_SATACCEL = 724
Public Const COLOR_SATSCROLL = 701
Public Const COLOR_SAVE = 711
Public Const COLOR_SCHEMES = 715
Public Const COLOR_SCROLLBAR = 0
Public Const COLOR_SOLID = 713
Public Const COLOR_SOLID_LEFT = 730
Public Const COLOR_SOLID_RIGHT = 731
Public Const COLOR_TUNE = 714
Public Const COLOR_WINDOW = 5
Public Const COLOR_WINDOWFRAME = 6
Public Const COLOR_WINDOWTEXT = 8

Public Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCPAINT = &HEE0086
Public Const SRCINVERT = &H660046

Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function LoByte Lib "TLBINF32" Alias "lobyte" (ByVal Word As Integer) As Byte
Private Declare Function HiByte Lib "TLBINF32" Alias "hibyte" (ByVal Word As Integer) As Byte
Private Declare Function loword Lib "TLBINF32" (ByVal DWord As Long) As Integer
Private Declare Function hiword Lib "TLBINF32" (ByVal DWord As Long) As Integer
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long


Sub GetWindowRGB(ByRef R&, ByRef G&, ByRef B&, ColInd&)
Dim Color&: Color = GetSysColor(ColInd&)
 R = LoByte(loword(Color))
  G = HiByte(loword(Color))
   B = LoByte(hiword(Color))
End Sub

Public Sub VGrad(CallingWnd As Object, ByVal R&, ByVal G&, ByVal B&, Optional b2w As Boolean = True, Optional DoEv As Boolean = True, Optional UseButtonFace As Boolean = False, Optional DefColInd& = COLOR_BTNFACE)
Dim i&, j&, ColFlg As Integer: ColFlg = 200
 If UseButtonFace = True Then GetWindowRGB R, G, B, DefColInd&
  For j = 0 To CallingWnd.ScaleWidth
   If DoEv = True Then DoEvents
    For i = 0 To CallingWnd.ScaleHeight
     SetPixel CallingWnd.hDC, j, i, GetCColor(i, CallingWnd.ScaleHeight, R, G, B, b2w)
    Next i
  Next j
End Sub

Public Sub HGrad(CallingWnd As Object, ByVal R&, ByVal G&, ByVal B&, Optional b2w As Boolean = True, Optional fade As Boolean = False, Optional DoEv As Boolean = True, Optional UseButtonFace As Boolean = False, Optional DefColInd& = COLOR_BTNFACE, Optional RefreshA As Boolean = False)
On Error Resume Next
Dim i&, j&, ColFlg As Integer: ColFlg = 200
 If UseButtonFace = True Then GetWindowRGB R, G, B, DefColInd&
  For j = 0 To CallingWnd.ScaleHeight
   If RefreshA = True Then object.Refresh
    If DoEv = True Then DoEvents
     If fade = True Then FadePixels j, CallingWnd.ScaleHeight, R, G, B
      For i = 0 To CallingWnd.ScaleWidth
       SetPixel CallingWnd.hDC, i, j, GetCColor(i, CallingWnd.ScaleWidth, R, G, B, b2w)
      Next i
  Next j
End Sub

Function FadePixels(Pos As Long, PosMax As Long, ByRef R&, ByRef G&, ByRef B&)
Dim tr&, tg&, tb&
  tr& = (Pos / PosMax) * (255 - R)
   tg& = (Pos / PosMax) * (255 - G)
    tb& = (Pos / PosMax) * (255 - B)
     R = R + tr
      G = G + tg
       B = B + tb
End Function

Function GetCColor(Pos As Long, PosMax As Long, StartingR As Long, StartingG As Long, StartingB As Long, Optional b2w As Boolean = True) As Long
Dim R&, G&, B&
 If b2w = True Then
  R& = (Pos / PosMax) * (255 - StartingR)
   G& = (Pos / PosMax) * (255 - StartingG)
    B& = (Pos / PosMax) * (255 - StartingB)
     GetCColor = RGB(StartingR + R, StartingG + G, StartingB + B)
 Else
  R& = (Pos / PosMax) * (255 - StartingR)
   G& = (Pos / PosMax) * (255 - StartingG)
    B& = (Pos / PosMax) * (255 - StartingB)
     GetCColor = RGB(255 - R, 255 - G, 255 - B)
 End If
End Function
