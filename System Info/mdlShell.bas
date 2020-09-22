Attribute VB_Name = "mdlShell"
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


Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Public Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SHACF_FILESYSTEM = &H1
Public Declare Sub SHAutoComplete Lib "shlwapi.dll" (ByVal hwndEdit As Long, ByVal dwFlags As Long)
 
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

'ShellExec Errors <= 32 ret
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const ERROR_BAD_FORMAT = 11&
Public Const SE_ERR_ACCESSDENIED = 5
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_DLLNOTFOUND = 32
Public Const SE_ERR_FNF = 2
Public Const SE_ERR_NOASSOC = 31
Public Const SE_ERR_OOM = 8
Public Const SE_ERR_PNF = 3
Public Const SE_ERR_SHARE = 26

Public Function FormatSEError(FncRet&, oBuffer$) As Boolean
 If FncRet > 32 Then FormatSEError = False: Exit Function
  Select Case FncRet
   Case 0:
    oBuffer = "The operating system is out of memory or resources."
   Case ERROR_FILE_NOT_FOUND:
    oBuffer = "The specified file was not found."
   Case ERROR_PATH_NOT_FOUND:
    oBuffer = "The specified path was not found."
   Case ERROR_BAD_FORMAT:
    oBuffer = "The .EXE file is not a valid Microsoft Win32® PE Header File, or an error has occured in the executable image."
   Case SE_ERR_ACCESSDENIED:
    oBuffer = "The operating system denied access to the specified file."
   Case SE_ERR_ASSOCINCOMPLETE:
    oBuffer = "The file name association is incomplete or invalid."
   Case SE_ERR_DDEBUSY:
    oBuffer = "The Dynamic Data Exchange (DDE) transaction could not be completed because other DDE transactions were being processed."
   Case SE_ERR_DDEFAIL:
    oBuffer = "The DDE transaction failed."
   Case SE_ERR_DDETIMEOUT:
    oBuffer = "The DDE transaction could not be completed because the request timed out."
   Case SE_ERR_DLLNOTFOUND:
    oBuffer = "The specified dynamic-link library (DLL) was not found."
   Case SE_ERR_FNF:
    oBuffer = "The specified file was not found."
   Case SE_ERR_NOASSOC:
    oBuffer = "There is no application associated with the given file name extension."
   Case SE_ERR_OOM:
    oBuffer = "There was not enough memory to complete the operation."
   Case SE_ERR_PNF:
    oBuffer = "The specified path was not found."
   Case SE_ERR_SHARE:
    oBuffer = "A sharing violation occurred."
  End Select
   FormatSEError = True
End Function

Public Function FormatWEError(FncRet&, oBuffer$) As Boolean
 If FncRet > 31 Then FormatWEError = False: Exit Function
  Select Case FncRet
   Case 0:
    oBuffer = "The operating system is out of memory or resources."
   Case ERROR_FILE_NOT_FOUND:
    oBuffer = "The specified file was not found."
   Case ERROR_PATH_NOT_FOUND:
    oBuffer = "The specified path was not found."
   Case ERROR_BAD_FORMAT:
    oBuffer = "The .EXE file is not a valid Microsoft Win32® PE Header File, or an error has occured in the executable image."
   Case Else:
    oBuffer = "An error was generated, but the exact cause of the error is un-known."
  End Select
   FormatWEError = True
End Function

