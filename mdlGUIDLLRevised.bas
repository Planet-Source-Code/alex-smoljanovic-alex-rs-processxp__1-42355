Attribute VB_Name = "mdlGUI"
Option Explicit
'***********************************************************************
'This application and its components were explicitly developed for
'PSC(Planet Source Code) Users as Open Source Projects.
'This code and the code of its components are property of their author.
'
'If you compile this application you may not redistribute it.
'However, you may use any of this code in you're own application(s).
'
'Alex Smoljanovic, Salex Software (c) 2001-2003
'salex_software@shaw.ca
'***********************************************************************


Private Declare Function VerticalGrad Lib "ProcXPGUI.dll" (ByVal hdc As Long, ByVal Height As Long, ByVal Width As Long, ByVal SysColInd As Long, ByVal b2w As Long) As Integer
Private Declare Function HorizontalGrad Lib "ProcXPGUI.dll" (ByVal hdc As Long, ByVal Height As Long, ByVal Width As Long, ByVal SysColInd As Long, ByVal b2w As Long, ByVal fade As Long) As Integer
Private Declare Function HRadialGrad Lib "ProcXPGUI.dll" (ByVal hdc As Long, ByVal Height As Long, ByVal Width As Long, ByVal SysColInd As Long) As Integer

Global GUIMissing As Boolean
'This flag evaluates to true when the failure to load the ProcXPGui dll occurs

Public Enum enFade
 NOFADE = 0
 EQUALITYFADE = 1
 UNEQUALITYFADE = 2
End Enum
'See ProcXPGUI dll project source files for information on how these flags are used

Public Enum enSysCol
 COLOR_3DDKSHADOW = 21
 COLOR_3DFACE = 15
 COLOR_3DHIGHLIGHT = 20
 COLOR_3DHILIGHT = 20
 COLOR_3DLIGHT = 22
 COLOR_3DSHADOW = 16
 COLOR_ACTIVEBORDER = 10
 COLOR_ACTIVECAPTION = 2
 COLOR_ADD = 712
 COLOR_ADJ_MAX = 100
 COLOR_ADJ_MIN = -100
 COLOR_APPWORKSPACE = 12
 COLOR_BACKGROUND = 1
 COLOR_BLUE = 708
 COLOR_BLUEACCEL = 728
 COLOR_BOX1 = 720
 COLOR_BTNFACE = 15
 COLOR_BTNHIGHLIGHT = 20
 COLOR_BTNHILIGHT = 20
 COLOR_BTNSHADOW = 16
 COLOR_BTNTEXT = 18
 COLOR_CAPTIONTEXT = 9
 COLOR_CURRENT = 709
 COLOR_CUSTOM1 = 721
 COLOR_DESKTOP = 1
 COLOR_ELEMENT = 716
 COLOR_GRADIENTACTIVECAPTION = 27
 COLOR_GRADIENTINACTIVECAPTION = 28
 COLOR_GRAYTEXT = 17
 COLOR_GREEN = 707
 COLOR_GREENACCEL = 727
 COLOR_HIGHLIGHT = 13
 COLOR_HIGHLIGHTTEXT = 14
 COLOR_HOTLIGHT = 26
 COLOR_HUE = 703
 COLOR_HUEACCEL = 723
 COLOR_HUESCROLL = 700
 COLOR_INACTIVEBORDER = 11
 COLOR_INACTIVECAPTION = 3
 COLOR_INACTIVECAPTIONTEXT = 19
 COLOR_INFOBK = 24
 COLOR_INFOTEXT = 23
 COLOR_LUM = 705
 COLOR_LUMACCEL = 725
 COLOR_LUMSCROLL = 702
 COLOR_MATCH_VERSION = &H200
 COLOR_MENU = 4
 COLOR_MENUTEXT = 7
 COLOR_MIX = 719
 COLOR_NO_TRANSPARENT = &HFFFFFFFF
 COLOR_PALETTE = 718
 COLOR_RAINBOW = 710
 COLOR_RED = 706
 COLOR_REDACCEL = 726
 COLOR_SAMPLES = 717
 COLOR_SAT = 704
 COLOR_SATACCEL = 724
 COLOR_SATSCROLL = 701
 COLOR_SAVE = 711
 COLOR_SCHEMES = 715
 COLOR_SCROLLBAR = 0
 COLOR_SOLID = 713
 COLOR_SOLID_LEFT = 730
 COLOR_SOLID_RIGHT = 731
 COLOR_TUNE = 714
 COLOR_WINDOW = 5
 COLOR_WINDOWFRAME = 6
 COLOR_WINDOWTEXT = 8
End Enum

Public Enum enB2W
 Black2White = 0
 white2black = 1
End Enum

Public Declare Function StretchBlt Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
'StretchBlt function stretches one DeviceContext to the specified dimensions and copies it to the destination Device Context
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H440328
Public Const SRCPAINT = &HEE0086
Public Const SRCINVERT = &H660046

Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
'TransparentBlt copies the source device context, and copies it to the destination device context omitting all pixels of the color value specified in the crTransparent argument
Declare Function SetPixel Lib "GDI32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
'Set pixel sets the color of a specific pixel in the specified device context handle

Public Function VGrad(hdc&, Height&, Width&, SysCol As enSysCol, b2w As enB2W)
On Error GoTo errh 'If an error occurs jump to label errh
 If GUIMissing = False Then VerticalGrad hdc, Height, Width, SysCol, b2w
 'If guimissing evaluates to false(so far not missing) then call VerticalGrad to draw a vertical gradient starting or ending with the specific system defined color
 'See ProcXPGUI Dll c++ project source files for more info...
  Exit Function
errh:
 If Err.Number = 53 Then
 'DLL wasn't found
  MsgBox "A required Dynamic-Link-Library file (""ProcXPGUI.dll"") is missing." & vbCrLf & "Please re-install ProcessXP." & vbCrLf & vbCrLf & "ProcessXP will be unable to properly draw the window and component's GUI.", vbCritical, "Error - ProcXPGUI.dll"
   GUIMissing = True
 End If
End Function

Public Function HGrad(hdc&, Height&, Width&, SysCol As enSysCol, b2w As enB2W, fade As enFade)
On Error GoTo errh
 If GUIMissing = False Then HorizontalGrad hdc, Height, Width, SysCol, b2w, fade
 'If guimissing evaluates to false(so far not missing) then call VerticalGrad to draw a vertical gradient starting or ending with the specific system defined color
 'See ProcXPGUI Dll c++ project source files for more info...
  Exit Function
errh:
 If Err.Number = 53 Then
 'DLL wasn't found
  MsgBox "A required Dynamic-Link-Library file (""ProcXPGUI.dll"") is missing." & vbCrLf & "Please re-install ProcessXP." & vbCrLf & vbCrLf & "ProcessXP will be unable to properly draw the window and component's GUI.", vbCritical, "Error - ProcXPGUI.dll"
   GUIMissing = True
 End If
End Function

Public Function HRGrad(hdc&, Height&, Width&, SysCol As enSysCol)
On Error GoTo errh
 If GUIMissing = False Then HRadialGrad hdc, Height, Width, SysCol
 'If guimissing evaluates to false(so far not missing) then call VerticalGrad to draw a vertical gradient starting or ending with the specific system defined color
 'See ProcXPGUI Dll c++ project source files for more info...
  Exit Function
errh:
 If Err.Number = 53 Then
 'DLL wasn't found
  MsgBox "A required Dynamic-Link-Library file (""ProcXPGUI.dll"") is missing." & vbCrLf & "Please re-install ProcessXP." & vbCrLf & vbCrLf & "ProcessXP will be unable to properly draw the window and component's GUI.", vbCritical, "Error - ProcXPGUI.dll"
   GUIMissing = True
 End If
End Function
