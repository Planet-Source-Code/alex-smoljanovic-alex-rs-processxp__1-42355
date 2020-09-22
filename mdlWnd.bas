Attribute VB_Name = "mdlWnd"
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

Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function SetFocusA Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
  Public Const SHOW_OPENWINDOW = 1
  Public Const SHOW_ICONWINDOW = 2
  Public Const SHOW_FULLSCREEN = 3
   Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function SetWindowPos Lib "user32" _
(ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_ALPHA = &H2&

Private Const ES_NUMBER = &H2000&

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_CLOSE = &H10

Global TargetPID&, numChild&, numWind&, tmpParent&

Public Function EnumWindowsProc(ByVal hWnd&, ByVal lParam&) As Boolean
'Enumerate Windows Call back function
Dim tmpProcID& 'dimensionalize tmpProcID as long type
 GetWindowThreadProcessId hWnd, tmpProcID
 'retrieve the specified window's process id
  If tmpProcID = TargetPID Then
  'if the current window's process id evaluates to the target process id then the window belongs to the target process...
    numWind& = numWind& + 1: tmpParent& = hWnd
    'increment numWind variable, initialize tmpParent
     numChild& = 0: EnumChildWindows hWnd, AddressOf CntChildProc, ByVal 0&
     'refresh numChild count, call EnumChildWindows to enumerate this windows child windows for the purpose of determining the number of child windows
      frmSearchThread.RetEnum hWnd, False
      'See frmSearchThread's RetEnum function for more info...
       EnumChildWindows hWnd, AddressOf EnumChildProc, ByVal 0&
       'enumerate the child windows of this window, see EnumChildProc for more info...
  End If
   EnumWindowsProc = True 'Return true so the window enumeration continues...
End Function

Public Function TermEnumWindows(ByVal hWnd&, ByVal TermTargetPID&) As Boolean
'This function is called specifically to terminate an application,
'this will enumerate through every window, if the window belongs to the target process
'to terminate the windows message WM_CLOSE will be sent to the window to close it
'the function that call EnumWindows with this call back function specified will
'then wait a maximum of three second for the process to terminate itself
'if the process hasn't terminated it's self with in that time the user is asked
'if they wish to force the termination of the non responsive process...
Dim tmpProcID& 'dimensionalize tmpProcID as long data type
 GetWindowThreadProcessId hWnd, tmpProcID
 'retrieve the process id of the current window
  If tmpProcID = TermTargetPID Then
  'if the current windows process id evaluates to the target process id then..
   PostMessage hWnd, WM_CLOSE, ByVal 0&, ByVal 0&
   'post the windows message WM_CLOSE, most windows will terminate once
   'receiving this message with the exception of some CONSOLE Windows and other windows that don't validate this message
  End If
   TermEnumWindows = True
   'return true to continue the enumeration
   'this will continue to enumerate all top-level windows, eventually requesting that
   'every window of the specified process be closed
End Function

Public Function CntChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim tmpProcID& 'dimensionalize tmpProcID as long data type
 GetWindowThreadProcessId hWnd, tmpProcID
 'retrive the process id of the current window
  If tmpProcID = TargetPID Then
  'if tmpProcID evaluates to TargetPID then..
   numChild& = numChild& + 1
   'increment numChild
  End If
   CntChildProc = 1
   'return 1(true) to continue enumeration...
End Function

Public Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
Dim tmpProcID& 'dimensionalize tmpProcID as long data type
 GetWindowThreadProcessId hWnd, tmpProcID
 'retrive the process id of the current window
  If tmpProcID = TargetPID Then
  'if tmpProcID evaluates to TargetPID then..
   frmSearchThread.RetEnum hWnd, True
   'see frmSearchThread's function RetEnum for more info...
  End If
   EnumChildProc = 1
   'return 1(true) to continue enumeration
End Function

Public Sub Add_ES_Number(hWnd&)
'this function adds the ES_NUMBER (EditStyle_Number) style constant to the specified edit window(text box)
Dim oStyle&, nStyle& 'dimensionalize oStyle(old) as long type, nStyle(new) as long type
 If IsWindow(hWnd) = 0 Then Exit Sub
 'if the specified window handle is not a window then exit this sub routine
  oStyle = GetWindowLong(hWnd, GWL_STYLE)
  'return the current window style of the specified window
   If oStyle = -1 Or oStyle = 0 Then Exit Sub
   'GetWindowLong function failed to return the window style, exit sub routine
    nStyle = oStyle Or ES_NUMBER
    'add Edit Style ES_NUMBER to the oStyle variable(current style of window)
     SetWindowLong hWnd, GWL_STYLE, nStyle
     'set the specified windows new window style
     'this edit style allows only numerical data to be entered into the control...
End Sub

Public Sub TransPrep(hWnd&)
On Error Resume Next
'on the event of an error resume execution on the next line of this procedure
Dim oStyle&, nStyle&, tStyle&
'dimensionalize oStyle as long type, nStyle as long type, tStyle as long type
  If KeepTrans = True Then
  'if globale variable KeepTrans evaluates to true then...
   oStyle& = GetWindowLong(hWnd&, GWL_EXSTYLE)
   'retrive the current window style of the specified window
    tStyle = oStyle And (Not WS_EX_LAYERED)
    'initialize tStyle(temporary) with the current window style with the window style WS_EX_LAYERED removed if it even existed
     If tStyle = oStyle Then
     'if tStyle evaluates to oStyle as it will when when the extended window style WS_EX_LAYERED is not yet specified in the extended window style of the window
     '* WS_EX_LAYERED has not yet been added to this window...
      SetWindowLong hWnd&, GWL_EXSTYLE, oStyle& Or WS_EX_LAYERED
      'add the extended window style constant WS_EX_LAYERED to the window
     End If
      SetLayeredWindowAttributes hWnd&, 0, WinTrans, LWA_ALPHA
      'set the transparency of the window to the global WinTrans variable value(byte; 0 to 255;255 = 100% visible)
  Else
  'keepwintrans evaluates to false
   oStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
   'return the current extended window style of the specified window
    nStyle = oStyle And (Not WS_EX_LAYERED)
    'remove the extended window style WS_EX_LAYERED from nStyle(the windows current extended window style)
     SetWindowLong hWnd&, GWL_EXSTYLE, nStyle
     'set the windows new extended window style
  End If
End Sub


Public Sub UpdateWinPos(hWnd&)
Dim WinRCT As RECT: GetWindowRect hWnd, WinRCT
'dimensionalize WinRCT as Rect type structure, initialize it with the window rectangle of the specified window
 If KeepWinTop = True Then
 'if global variable KeepWinTop evaluates to true than...
  SetWindowPos hWnd&, HWND_TOPMOST, WinRCT.Left, WinRCT.Top, WinRCT.Right, WinRCT.Bottom, flags
  'set the windows Z axis to top most(above all NOTTOPMOST windows)
 Else
  SetWindowPos hWnd&, HWND_NOTOPMOST, WinRCT.Left, WinRCT.Top, WinRCT.Right, WinRCT.Bottom, flags
  'set the windows Z axis to not top most
 End If
End Sub
