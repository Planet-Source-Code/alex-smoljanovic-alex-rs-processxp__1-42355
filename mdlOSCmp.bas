Attribute VB_Name = "mdlOSCmp"
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


Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetVersionEx& Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO)
 Private Const VER_PLATFORM_WIN32_NT = 2
 Private Const VER_PLATFORM_WIN32_WINDOWS = 1
 Private Const VER_PLATFORM_WIN32s = 0
 Const PLANES = 14
 Const BITSPIXEL = 12

Public Enum enOSSpec
 Win95 = 0
 Win98 = 1
 Win98SE = 2
 WinME = 3
 WinNT4 = 4
 Win32s = 7
End Enum

Public Enum enWinVer
 Win2K = 2
 fullCompatibleOS = 1
 inCompatibleOS = 0
End Enum

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type ConvertedOSInfo
    OperatingSystem As enWinVer
    OSSpecs As enOSSpec
    OSSpecsEx As String
    OSBuild As Long
    DispDevDescription As String
    DispDevBits As Long
End Type

Global HostOS As ConvertedOSInfo

Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Function GetDeviceColors(oDesc$) As Long
GetDeviceColors = GetDeviceCaps(GetDC(GetDesktopWindow), BITSPIXEL)  ' Returns the device capability specified by nIndex(BITPIXEL) 'Bits Per Pixel'
'returns the device context's color bit-depth
 oDesc$ = CStr(GetDeviceColors) & "-bit (" & FormatNumber(CSng(2 ^ (GetDeviceCaps(GetDC(GetDesktopWindow), PLANES) * GetDeviceCaps(GetDC(GetDesktopWindow), BITSPIXEL))), 0, , , vbTrue) & " colors)"
 'return the string "#-bit (# colors)"
End Function

Public Function RetOSInf() As Boolean
On Error GoTo errh
'on the event of an erro jump to label errh
Dim osvi As OSVERSIONINFO: osvi.dwOSVersionInfoSize = 148 'initialize variable
 If GetVersionEx(osvi) <> 0 Then
 'if the function returned succesfully then...
   Select Case osvi.dwPlatformId
    Case VER_PLATFORM_WIN32s: 'if osvi(operating system version information)'s dwPlatformID evaluates to the value of the VER_PLATFORM_WIN32s constant then...
     HostOS.OperatingSystem = inCompatibleOS 'update flag, see mdlInit for more info...
      HostOS.OSSpecs = Win32s 'update flag, see frmIOS for more info...
    Case VER_PLATFORM_WIN32_WINDOWS:
     If osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 0 Then
     'if the OS(operating system)'s MajorVersion evaluates to 4,
     'and the OS's MinorVersion evaluates to 0 then OS is Win95...
      HostOS.OperatingSystem = inCompatibleOS 'See mdlInit for more info..
      'The OperatingSystem member will evaluate to Win2k or FullCompatibleOS
      'If the operating system is Win2k then we will draw graphics differently
      ', change background colors of certain object and so on
      'This member will evaluate to FullCompatibleOS if the OS is WindowsXP or greater
      'We will keep track of this flag through out many of the graphically related procedures...
       HostOS.OSSpecs = Win95 'update operating system specs flag
       'This flag is only used when loading frmIOS, this form is only loaded
       'when the OperatingSystem flag evaluates to inCompatibleOS
        If LCase(osvi.szCSDVersion) = "c" Or LCase(osvi.szCSDVersion) = "b" Then HostOS.OSSpecsEx = " OSR2"
        'if szCSDVersion evaluates to "c" or "b" then the OS is Windows 95 OSR2
     ElseIf osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 10 Then
     'Windows 98
      HostOS.OperatingSystem = inCompatibleOS '..
       HostOS.OSSpecs = Win98 '..
        If LCase(osvi.szCSDVersion) = "a" Then HostOS.OSSpecs = Win98SE
        'if szCSDVersion evaluates to "a" then the OS is Window 98 Second Edition
     ElseIf osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 90 Then
     'Windows Mellenium edition
      HostOS.OperatingSystem = inCompatibleOS '..
       HostOS.OSSpecs = WinME '..
     End If
    Case VER_PLATFORM_WIN32_NT:
     If osvi.dwMajorVersion = 4 Then HostOS.OperatingSystem = inCompatibleOS: HostOS.OSSpecs = WinNT4
     'Windows NT 4(<)
      If osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 0 Then HostOS.OperatingSystem = Win2K
      'Windows 2000, OS is compatible
       If osvi.dwMajorVersion = 5 And osvi.dwMinorVersion >= 1 Or osvi.dwMajorVersion > 5 Then HostOS.OperatingSystem = fullCompatibleOS
       'Windows XP(>), OS is compatible
   End Select
 End If
  HostOS.OSBuild = osvi.dwBuildNumber 'build number of the Operating System version information
   RetOSInf = True 'no errors have occured, return true
    HostOS.DispDevBits = GetDeviceColors(HostOS.DispDevDescription)
    'Return the display adapters Bits Per Pixel, copy the display settings description to the DispDevDescription type member
   Exit Function 'discontinue execution of this procedure
errh: 'label errh
 RetOSInf = False 'an error occured, return false
  HostOS.DispDevBits = GetDeviceColors(HostOS.DispDevDescription) 'make sure the display adapters capabilites are still returned despite the error
  'Return the display adapters Bits Per Pixel, copy the display settings description to the DispDevDescription type member
End Function
