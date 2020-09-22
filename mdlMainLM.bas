Attribute VB_Name = "mdlMain"
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


Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = (STATUS_WAIT_0 + 0)
'Private Const WAIT_TIMEOUT = 258&
'Private Const WAIT_FAILED = &HFFFFFFFF

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
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
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SHACF_FILESYSTEM = &H1
Public Const SHACF_AUTOAPPEND_FORCE_OFF = &H80000000
Public Const SHACF_AUTOAPPEND_FORCE_ON = &H40000000
Public Const SHACF_AUTOSUGGEST_FORCE_OFF = &H20000000
Public Const SHACF_AUTOSUGGEST_FORCE_ON = &H10000000
Public Const SHACF_FILESYS_ONLY = &H10
Public Const SHACF_USETAB = &H8

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



Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
 Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
 Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
 Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
 Private Const FORMAT_MESSAGE_FROM_STRING = &H400
 Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
 Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
 Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF


Public Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal Process As Long, ByRef ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Thread32First Lib "kernel32" (ByVal hSnapshot As Long, ByRef lpte As THREADENTRY32) As Long
Private Declare Function Thread32Next Lib "kernel32" (ByVal hSnapshot As Long, ByRef lpte As THREADENTRY32) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long

Public Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer

End Type

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const PrSYNCHRONIZE = &H100000


Private Const MAX_PATH As Integer = 260

'GFN
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH '260
End Type

Private Type MODULEENTRY32
  dwSize As Long
  th32ModuleID As Long
  th32ProcessID As Long
  GlblcntUsage As Long
  ProccntUsage As Long
  modBaseAddr As Long
  modBaseSize As Long
  hModule As Long
  szModule As String * 256
  szExePath As String * 260
End Type

Private Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
End Type


Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT = &H80000000
Private Const THREAD_ALL_ACCESS = &H1F03FF

Private Const THREAD_PRIORITY_IDLE = -15
Private Const THREAD_PRIORITY_LOWEST = -2
Private Const THREAD_PRIORITY_BELOW_NORMAL = -1
Private Const THREAD_PRIORITY_NORMAL = 0
Private Const THREAD_PRIORITY_ABOVE_NORMAL = 1
Private Const THREAD_PRIORITY_HIGHEST = 2
Private Const THREAD_PRIORITY_TIME_CRITICAL = 15


Public Type fProcMod
 ProcessEXE As String
 ProcessPath As String
 eProcessID As Long
 eParentID As Long
 isMT As Boolean
 numThreads As Long
  Modules() As String
  ModulesPath() As String
  ModulesCnt As Long
   ThreadID() As String
   ThreadOwnerPID() As String
   ThreadBasePri() As String
   ThreadDeltaPri() As String
   ThreadCnt As Long
End Type


Public Enum InfoFrame
 ProcessInfo = 0
 VersionInfo = 1
End Enum

Public Enum TermPriv
 OnlyAdmin = 0
 AllUsers = 1
End Enum

Public Enum TermReturn
 Terminated = 0
 Failed = 1
 Cancelled = 2
End Enum

'The global function declarations of this module consist of mainly process related functions

Global TerminationPriv As TermPriv
Global AllCanEnumWin As Boolean
Global AllCanManipWin As Boolean
'Security flags...

Global CurrentInfo As InfoFrame, MemUpdateOn As Boolean
Global Starting As Boolean
'Global flags

Global fProc() As fProcMod 'This array's elements will store the information of the enumerated processes
Global ProcCnt&, ModCnt&
'ProcCnt will provide an additional method of determining how many elements exist in the fProc array besided the UBound function which will determine the largest element index

Global ShowProcIcon As Boolean, ShowModIcon As Boolean
Global ProcRelation As Boolean, ModRelation As Boolean
'List populational flags, see frmMains RefreshLibrary to see how they are used...
Global KeepWinTop As Boolean, WinTrans&, KeepTrans As Boolean
'General window flags... Used by TransPrep and UpdateWinPos functions...

Global SICommandLine As Boolean, sndClass As clsMain, playSounds As Boolean, sndSupported As Boolean, playedTerm As Boolean
'dimensionalize sndClass as the Sound Servers main interactive class
'playSounds, sndSupported and playedTerm flags are used to determine how and if a sound should be played on certian events and while certain functions are called

Public Sub CountProcesses()
'This procedure is used only to return the number of processes, this information will be useful while redimensionalizing are main process array(fProc)
'If the ammount of processes are unknown, then our array fProc must be over dimensionalized this will waste memory
On Error GoTo errh 'On the event of an error jump to the label errh
Dim hSnapshot&, uProcess As PROCESSENTRY32, rProc&, prCnt&
'Dimensionalize hSnapShot as long data type(Variable names who's prefixes are 'h' is usually used as handled(long))
'uProcess as PROCESSENTRY32 type, rProc(Variable names who's prefixes are 'r' is usually used as a fnct return value)
 hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
 'initialize hSnapShot with the handle snapshot handle returned by CreateToolHelp32SnapShot
  uProcess.dwSize = Len(uProcess) 'Initialize uProcess who's size will evaluate to the size of the structure
   rProc = Process32First(hSnapshot, uProcess) 'Select the first process in the snapshot handle
   'initialze rProc with the first(in the snapshot) process's handle
    Do While rProc
    'Do While loop; Loops while rProc evaluates to a non zero value
     DoEvents 'Yield execution
      prCnt& = prCnt& + 1 'increment prCnt by one
       rProc = Process32Next(hSnapshot, uProcess)
       'return the Next process's handle in the snapshot to rProc
    Loop 'check loop conditions, loop again if loop expression evaluates to true
     ProcCnt = prCnt& - 1 'set ProcCnt to the value of prCnt minus 1
      CloseHandle hSnapshot 'Closes the handle
       Exit Sub 'discontinue execution of this procedure
errh: 'label errh
 If Err.Number = 10 Then MsgBox "The enumeration of processes has temporarily ceased." & vbCrLf & vbCrLf & "This may have been caused if a process is currently initializing." & vbCrLf & "If that is the case; The enumeration of processes or the population of the process list will resume on that processes completion of initialization.", vbExclamation, "ProcessXP - Idle": Exit Sub
  MsgBox "An un-expected error has occured, while this error was not fatal it is suggested that you restart ProcessXP." & vbCrLf & vbCrLf & "If this problem still occurs please report this error to Salex Software by using Bug Reporting tool.", vbCritical, "ProcessXP - Error"
End Sub

Public Function enumProcesses()
On Error GoTo errh 'on the event of an error jump to label errh
Dim CurrentInd&, hSnapshot&, uProcess As PROCESSENTRY32, tmpMod&, rProc&
'Dimensionalize CurrentInd as long type, hSnapShot as long type,
'uProcess as PROCESSENTRY32 type struct., tmpMode as long type, rProc as long type
CountProcesses 'Call CountProcesses function to determine the number of processes that will be enumerated
 hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, 0&)
 'return hSnapshot to the SnapShot handle returned by the CreateToolHel32Snapshot functions...
  uProcess.dwSize = Len(uProcess) 'initialize structure...
   rProc = Process32First(hSnapshot, uProcess) 'Return the first process handle in the snapshot
    Do While rProc 'Loop while rProc(Process's Handle Return) is non zero
    DoEvents 'Yield execution to other procedures
     ReDim Preserve fProc(0 To CurrentInd&) As fProcMod
     'redimensionalize array fProc with the Preserve keyword(this preserves the existing elements in the arry, otherwise they are deleted) to
      fProc(CurrentInd&).ProcessEXE = Left$(uProcess.szExeFile, IIf(InStr(1, uProcess.szExeFile, Chr$(0)) > 0, InStr(1, uProcess.szExeFile, Chr$(0)) - 1, 0))
      'Parse the process's module name
       fProc(CurrentInd&).eParentID = uProcess.th32ParentProcessID
       'set the element of index CurrentInd's eParentID member in the fProc array to the current process's parent process id
        fProc(CurrentInd&).eProcessID = uProcess.th32ProcessID
        'set the element of index CurrentInd's eProcessID member in the fProc array to the current process's process id(PID)
         If Starting = True Then frmSpash.UpdateProgress CurrentInd&, ProcCnt&, "Enumerating Processes", fProc(CurrentInd&).ProcessEXE: DoEvents
         'If starting flag evaluates to true as it will when this procedure
         'is being called while frmSpash has not yet finished loading all
         'the forms(more specifically frmMain) then update the progress display
          If Starting = False Then frmMain.lblProcListTmp.Caption = "Enumerating Processes (" & CStr(CLng((CurrentInd& / ProcCnt) * 100)) & "% Complete)"
           If uProcess.cntThreads > 0 Then fProc(CurrentInd&).isMT = True: fProc(CurrentInd&).numThreads = uProcess.cntThreads
           'If this process's number of threads is greater than zero than this process is multithreaded, set the current elements isMT member to true otherwise it will remain false, set its Thread count member to the number of threads which belong to this process
            enumModules uProcess.th32ProcessID, fProc(CurrentInd&).Modules, fProc(CurrentInd&).ModulesPath, fProc(CurrentInd&).ModulesCnt
            'enumerate this processes modules; see enumModules for more info...
             fProc(CurrentInd).ProcessPath = GetProcModPath(uProcess.th32ProcessID, 0, fProc(CurrentInd&).ProcessEXE)
             'see GetProcModPath function for more info...
              rProc = Process32Next(hSnapshot, uProcess)
              'Select the next process in the process snapshot
               CurrentInd& = CurrentInd& + 1
               'increment CurrentInd by one
    Loop 'Check loop conditions, if loop condition expression evaluates to true then loop...
     ProcCnt = CurrentInd& - 1 'update the Processs count flag
      CloseHandle hSnapshot 'close the snapshot's handle
       Exit Function 'discontinue execution of this procedure
errh:
 If Err.Number = 10 Then MsgBox "The enumeration of processes has temporarily ceased." & vbCrLf & vbCrLf & "This may have been caused if a process is currently initializing." & vbCrLf & "If that is the case; The enumeration of processes or the population of the process list will resume on that processes completion of initialization.", vbExclamation, "ProcessXP - Idle": Exit Function
 'Since during this procedures enumeration loop DoEvents function is called, it is possible for the user to attempt to refresh the list again, several error would occur since the arrays are temporarily locked, and can't be redimensionalize
  MsgBox "An un-expected error has occured, while this error was not fatal it is suggested that you restart ProcessXP." & vbCrLf & vbCrLf & "If this problem still occurs please report this error to Salex Software by using Bug Reporting tool.", vbCritical, "ProcessXP - Error"
End Function


Private Function enumModules(ByVal ProcID&, ByRef ModArray() As String, ByRef ModulesPth() As String, ByRef fModCnt As Long)
'This function will enumerate the modules of the process specified by its PID(Process ID)
'the byref(By Reference(reference to actual variable) rather than By Value(only the variables value is passed)) array argument are a reference to the array passed by the paramater
If ProcID = 0 Then fModCnt = 0: Exit Function
'If the Process ID evaluates to zero then update the module count variable and exit this procedure
 Dim mProcess As MODULEENTRY32, rMod&, tmpCnt&, hSnapshot&
 'dimensionalize mProcess as MODULEENTRY32 type struct., rMode as long type, tmpCnt as long type, hSnapShot as long type
   hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, ProcID&)
   'return handle to a module snapshot(module consits of only module entries of the specified process id)
    mProcess.dwSize = Len(mProcess) 'initialize structure...
     rMod = Module32First(hSnapshot, mProcess)
     'select the first module in the module snapshot
      Do While rMod 'loop while rMode evaluate's to a non-zero value
      DoEvents 'yield execution of this procedure to other procedures
       tmpCnt& = tmpCnt& + 1: ModCnt = ModCnt + 1
       'increment counting flags..
        If tmpCnt& = 1 Then GoTo fFrst
        'If the module is the first module in the snapshot(tmpCnt = 0) then
        'then the current module is the actual processes module handle, we won't return it the module array
         fModCnt& = fModCnt& + 1 'increment fModCnt...
          ReDim Preserve ModArray(0 To fModCnt&)
          'add a new element to the array preserving the existing elements
           ReDim Preserve ModulesPth(0 To fModCnt&)
           'add a new element to the array preserving the existing elements
            ModArray(fModCnt - 1) = Left(mProcess.szModule, InStr(mProcess.szModule, Chr(0)) - 1)
            'return the modules EXE name to the specified array element(no path)
              ModulesPth(fModCnt - 1) = GetProcModPath(ProcID, mProcess.hModule, ModArray(fModCnt - 1))
              'return the modules file path to the specified array element
              'see GetProcModPath function for more info...
fFrst: ' fFrst label; jumped to when the currently select module of the module snapshot
                     'is the actual processes module handle
                rMod = Module32Next(hSnapshot, mProcess)
                'return the handle to the next module in the module snapshot
      Loop 'check loop conditions, if loop condition expression evaluates to false then loop...
       ModCnt = ModCnt& - 1: fModCnt = fModCnt - 1
       'decrement counting flags as the first added element is counted as 1 yet it is of index o in the array
        CloseHandle hSnapshot 'close the handle to snapshot
End Function


Public Function GethModule(ProcID&) As Long
'Function returns the Processes Module handle, this function is used to determine a processes path
 Dim mProcess As MODULEENTRY32, rMod&, tmpCnt&, hSnapshot&
 'dimensionalize mProcess as MODULEENTRY32 type stuct., rMod as long type, tmpCnt as long type, hSnapShot as long type
   hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, ProcID&)
   'return a handle to the new module snapshot(consists of only module handles of the specified process)
    mProcess.dwSize = Len(mProcess) 'initialize structure...(function which uses this structure determines the amount of memory to allocate)
     rMod = Module32First(hSnapshot, mProcess)
     'return the first module's handle in the module snapshot
      GethModule = mProcess.hModule 'return the processes module handle
       CloseHandle hSnapshot 'close the snapshots handle(clean up)
End Function

'************************************
'This function is used while debugging(Returns the error message of system errors)
Private Function FormatSystemError() As String
 FormatSystemError = Space$(200)
  FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, 0&, FormatSystemError, 200, ByVal 0&
   FormatSystemError = RTrim$(FormatSystemError)
End Function
'************************************

Public Function TerminateProc(ProcID&, ProcPath$, Description$) As TermReturn
'Under Win32, the operating system promises to clean up
'resources owned by a process when it shuts down. This does not,
'however, mean that the process itself has had the opportunity to do any
'final flushes of information to disk, any final communication over a remote
'connection, nor does it mean that the process' DLL's will have the opportunity
'to execute their PROCESS_DETACH code. This is why it is generally preferable
'to avoid terminating an application under Windows 95 and Windows NT.

'Visit either of the following URL's for more info
'http://msdn.microsoft.com/msdnmag/issues/02/06/debug/default.aspx for even more information....
'http://support.microsoft.com/default.aspx?scid=KB;EN-US;q178893&ID=KB;EN-US;q178893

'This procedure will enumerate every top-level window thread which belongs to the process specified by the ProcID argument
'If the application has not terminated within three seconds, confirmation is requested from the user to force the process to terminate...

Dim lngSize&, lngHwndProcess&, lngReturn&, cProcPath$, strModuleName$, hProcess&, strProcessname$
'dimensionalize lngSize as long type, lngHwndProcess as long type, lngReturn as long type, cProcPath as string type, strModuleName as string type, hProcess as long type
Dim ModAr() As String, ModArP() As String, ModCnt&, firstMod&, retWait&
'dimensionalize ModAr as a one dimensional array of strings, ModArP as a one dimensional array of strings
'ModCnt as long type, firstMod as long type
strModuleName = Space(MAX_PATH) 'allocate memory for this variable(MAX_PATH constant specifies the maximimum length of characters a path can consist of...)
 lngSize = 500 'initialize variable
  lngHwndProcess = OpenProcess(PrSYNCHRONIZE Or PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcID&)
  'Since we will be using WaitForSingleObject we must specify the PrSYNCHRONIZE constant in the access rights argument...
  'initialize lngHwndProcess with the Process's handle returned by OpenProcess function...
   EnumWindows AddressOf TermEnumWindows, ByVal ProcID
   'Enumerate top-level windows, the enumeration call back function will be function TermEnumWindows(see mdlWnd for more info)
   'If the window(Window Handle) sent to this procedure(TermEnumWindows) belongs to the specified process, window message WM_CLOSE will be sent to it
    If WaitForSingleObject(ByVal lngHwndProcess, 3000) <> WAIT_OBJECT_0 Then GoTo DoKillProc
    'The WaitForSingleObject function returns when one of the following occurs:
    'The specified object is in the signaled state.
    'The time-out interval elapses.
    'This functions return will evaluate to WAIT_OBJECT_0 if the process terminated it's self
    'Otherwise jump to label DoKillProc as the process is not responding to the window messages...
     TerminateProc = Terminated
     'Return Terminate as the process was terminated
      CloseHandle lngHwndProcess 'close the process's handle
       Exit Function 'discontinue execution of this procedure
DoKillProc: 'label DoKillProc, jumped to when the application is not responding
 If sndSupported = True And playedTerm = False Then sndClass.doTerm: playedTerm = True
 'If sound is supported(which is determined by frmSpash by testing the existance of the sndServer dll) and playterm equals fals then play the termination explanation sound and set playterm to true so that this sound isn't played more than once
  If MsgBox("The application is not responding." & vbCrLf & vbCrLf & "Forcing a process to terminate can leave the system unstable, this will not notify any attached DLLs that the process is terminating. This method should be used as a last resort." & vbCrLf & vbCrLf & "Force the termination of """ & Description & """?", vbQuestion + vbYesNo, "Terminate Process") = vbYes Then
  'Request user confirmation to terminate the processes...
   lngReturn = GetModuleFileNameExA(lngHwndProcess&, 0, strModuleName$, lngSize&)
   'copies the path of the process to pre-initialized strModuleName buffer, and returns the length of characters copied
    strProcessname = Left(strModuleName, lngReturn) 'return the number of characters in the strProcessName variable specified by lngReturn
     cProcPath$ = UCase$(Trim$(strProcessname))
     'return the uppercase value of the variable strProcessName whos trailing and leading spaces are removed
     'this will ensure than a textual comparison is performed rather than a binary comparison
      If ProcPath = cProcPath$ Then
      'if ProcPath evaluates to cProcPath(ensure the right process is about to be terminated) then...
       hProcess = OpenProcess(1&, -1&, ProcID&)
       'open the process specified with termination rights and return its handle
        If TerminateProcess(hProcess&, 0&) = 0 Then
         TerminateProc = Failed
         'if the function call TerminateProcess returned 0, then the process was not terminated
        Else
         TerminateProc = Terminated
         'the process was terminated
        End If
         CloseHandle hProcess 'close the handle to the opened process
      Else
       TerminateProc = Failed
       'The process specified wasn't found...
      End If
   Else
    TerminateProc = Cancelled
    'User decided not to force the termination of the process...
   End If
    CloseHandle lngHwndProcess 'close the handle...
End Function

Private Function GetProcModPath(fProcModID&, fModuleID&, Optional EXEName$) As String
On Error Resume Next 'on the event of an error resume execution on the next line
Dim strModuleName$, lngSize&, lngHwndProcess&, lngReturn&, strProcessname$
'dimensionalize strModuleName as string type, lngSize as long type, lngHwndProcess as long type, lngReturn as long type
strModuleName = Space(MAX_PATH) 'allocate memory(MAX_PATH constant specifies the maximum length of characters a path can consist of..)
 lngSize = 500 'initialize variable
  lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, fProcModID&)
  'open the specified process's module, and return its handle
   lngReturn = GetModuleFileNameExA(lngHwndProcess, fModuleID&, strModuleName, lngSize)
   'copy the modules file name to the buffer strModuleName, and return the length of characters
    strProcessname = Left(strModuleName, lngReturn) 'again, this variable is used to verify the module specified belongs to the same process who's path is requested
     GetProcModPath = UCase$(Trim$(strProcessname))
     'return the uppercase and trimed(leading and trailing spaces are removed) value of the path
      If InStr(1, GetProcModPath, "\SYSTEMROOT\", 1) > 0 Then
      'if the substring "\SYSTEMROOT\" exists with in the path buffer then...
       GetProcModPath = Replace(GetProcModPath, "\SYSTEMROOT\System32", GetSystemRoot, , , 1)
       'replace "\SYSTEMROOT\" with the actual system root path
       'see GetSystemRoot function...
       'If you are using Visual Basic 5, then the replace function is not available to you
       'add the following to this module
       
       'Place this enumeration in the general declarations section of this module
'        Public Enum enCompareMethod
'         TextComparison = 1 'Ignores case
'         BinaryComparison = 0 'uses binary value of character for comparison
'        End Enum
'
'      Place this function below or above function GetProcModPath
'        Public Function Replace(ByVal sExp$, sFind$, sReplaceAs$, Optional start& = 1, Optional count& = -1, Optional CompareMethod As enCompareMethod) As String
'        Dim sPos&
'         Do While InStr(start, sExp, sFind, CompareMethod) > 0
'          sPos = InStr(start, sExp, sFind, CompareMethod)
'           Mid$(sExp, sPos, sPos + Len(sFind)) = sReplaceAs
'         Loop
'          Replace = sExp
'        End Function
'
'       or:
'        GetProcModPath =  GetSystemRoot & mid$(GetProcModPath, instr(1,GetProcModPath, "\SYSTEMROOT\System32",1) + len("\SYSTEMROOT\System3"))

      End If
       If GetProcModPath <> "" And InStr(1, GetProcModPath, ":\", 1) <> 2 Then
       'determines if the drive part of this path is the fist character in the string
        GetProcModPath = Mid$(GetProcModPath, InStr(1, GetProcModPath, ":\", 1) - 1)
        'removes characters before the drive part of this string
       End If
        If GetProcModPath = "" And EXEName <> "" Then
        'If no other path information can be retrieved and the EXEName doesn't evaluate to nothing("") then..
         If Dir(GetSystemRoot & "\" & EXEName$) <> "" Then
         'test if this file exists in the system root
         'I can't gaurantee the accuracy of this, as this file could reside in another special directory
         'just because a file with this same name exists in the system directory doesn't mean that the process's file path is neccesarily in the system directory(possible file name coincidence)
          GetProcModPath = GetSystemRoot & "\" & EXEName$
          'Return the system directory with the file name appended to it
         End If
        End If
         CloseHandle lngHwndProcess 'close the opened process's handle
End Function

Public Function GetShortPath(strFileName As String) As String
'This function returns the Short File Path of a path(DOS Path)
Dim lngRes&, strPath$: strPath = String$(MAX_PATH, 0)
'dimensionalize lngRes as long type, strPath as string type
'allocate memory to strPath(MAX_PATH specifies the maximum length of a path)
 lngRes = GetShortPathName(strFileName, strPath, MAX_PATH)
 'copy the short path name of the specified file to variable strPath, and return the length of the path
  GetShortPath = Left$(strPath, lngRes) 'return only the length of the path returned by the call to the GetShortPathName function
End Function

Public Function GetProcTime(ProcID&) As String
'Retrieves and formats the time a process started processing...
Dim hProcess&, fTime As FILETIME, fNull As FILETIME, sTime As SYSTEMTIME
'dimensionalize hProcess as long type, fTime as FILETIME type struct., fNull as FILETIME struct, sTime as SYSTEMTIME type struct.
 hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcID&)
 'Open process and return its handle to variable hProcess...
  GetProcessTimes hProcess, fTime, fNull, fNull, fNull
  'Call GetProcessTimes to retrive the starting time of the specified process
  'function uses a reference to fTime to copy memory
   FileTimeToLocalFileTime fTime, fTime
   'convert FileTime format to LocalFileTime format
    FileTimeToSystemTime fTime, sTime 'copies our fTime structure to sTime variable
     GetProcTime = Format$(CStr(sTime.wHour), "00") & ":" & Format$(CStr(sTime.wMinute), "00") & ":" & Format$(CStr(sTime.wSecond), "00") & " " & Format$(CStr(sTime.wDay), "00") & "/" & Format$(CStr(sTime.wMonth), "00") & "/" & Format$(CStr(sTime.wYear), "00")
     'Format the string... (0#:##:## ##/##/##)
      CloseHandle hProcess 'close the open process's handle
End Function

Public Function GetMemory(ProcID&) As String
'dimensionalize byteSize as doube data type, tmpBuffer as string data type, hProcess as long datat type, ProcMem as PROCESS_MEMORY_COUNTERS
Dim byteSize As Double, tmpBuffer$, hProcess&, ProcMem As PROCESS_MEMORY_COUNTERS: ProcMem.cb = LenB(ProcMem)
'initialize ProcMem with the size in bytes in which is required for storing its self
 hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcID&)
 'open the specified process and returns its handle
  If hProcess <= 0 Then GetMemory = "Can't Determine": Exit Function
  'If the process wasn't open(hProcess evaluates to less than or equal to 0), then exit this procedure
   GetProcessMemoryInfo hProcess, ProcMem, ProcMem.cb
   'call GetProcessMemoryInfo to copy the memory information about the specific process
   'to a reference to our ProcMem variable
    byteSize = ProcMem.WorkingSetSize / 1024: tmpBuffer = "KB's"
    'initialize byteSize with the value of ProcMem's WorkingSetSize memory divided by 1024(1024 bytes per kilobyte)
     If byteSize / 1024 >= 1 Then byteSize = byteSize / 1024: tmpBuffer = "MB's"
     'if byteSize divided by 1024 evaluates to or greater than 1 then divide it by 1024(1024 KiloBytes's per megabyte)
      If byteSize / 1024 >= 1 Then byteSize = byteSize / 1024: tmpBuffer = "GB's"
      'if byteSize divided by 1024 evaluates to or greater than 1 then divide it by 1024(1024 megabyte's per gigbyte)
       GetMemory = FormatNumber$(byteSize, 0, , , vbTrue) & " " & tmpBuffer
       'Return the formatted number(#,###) with the measurement unit appended
        CloseHandle hProcess 'close the handle
End Function

Public Function GetMemoryLong(ProcID&) As Long
'Function is used to retrieve the memory utilization of a specific process as long value rather than a formatted string
'function is only used by the search(frmSearch) algorithm when the memory utilization flag is not "Any Amount"

'for reference on this function see GetMemory function...
Dim byteSize As Double, hProcess&, ProcMem As PROCESS_MEMORY_COUNTERS: ProcMem.cb = LenB(ProcMem)
 hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcID&)
  If hProcess <= 0 Then GetMemoryLong = -1: Exit Function
   GetProcessMemoryInfo hProcess, ProcMem, ProcMem.cb
    GetMemoryLong = (ProcMem.WorkingSetSize / 1024) / 1024
     CloseHandle hProcess
End Function

Public Sub GetModuleInformation(ByVal ProcID&, ByVal ModName$, ByRef GlbUsage&)
'This function is used to retrieve the amount of times a module is globally loaded
If ProcID = 0 Then GlbUsage& = 0: Exit Sub
'If the process id isn't specified then exit the sub routine
 Dim mProcess As MODULEENTRY32, rMod&, tmpBuffer$, hSnapshot&
 'dimensionalize mProcess as MODULEENTRY32 type struct., rMod as long type, tmpBuffer as string type, hSnapShot as long data type
   hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, ProcID&)
   'initialize hSnapShot with the handle to a module snapshot returned by CreateToolHelp32Snapshot function
    mProcess.dwSize = Len(mProcess) 'initialize structure, retrieve the length of this variable for sufficient memory allocation
     rMod = Module32First(hSnapshot, mProcess)
     'return the handle to the first module in the snapshot
      Do While rMod
      'Do While loop; loops while rMod's value is non zero
      DoEvents 'Yield execution to other procedures
       tmpBuffer$ = Left(mProcess.szModule, InStr(mProcess.szModule, Chr(0)) - 1)
       'initialize tmpBuffer with the return of the left function which will return all the characters before the terminating character, minus the first terminating character
        If Trim$(tmpBuffer) = Trim$(ModName) Then
        'If tmpBuffer evaluates to ModName then... (The module who's information is requested is the current module handle in the snapshot)
         GlbUsage& = mProcess.GlblcntUsage
         'Initialize GlbUsage variable with the modules global count usage
           CloseHandle hSnapshot 'close the snapshot's handle
            Exit Sub 'exit sub routine
        End If
         rMod = Module32Next(hSnapshot, mProcess)
         'return the handle to the next module in the snapshot
      Loop 'check conditions, if loops condition express evaluates to true then continue the loop
       CloseHandle hSnapshot 'close the snapshot's handle
End Sub

Public Function FormatSEError(FncRet&, oBuffer$) As Boolean
'This function determines if the return value of the ShellExecute function is an error value
 If FncRet > 32 Then FormatSEError = False: Exit Function
 'If argument FncRet(Functions Return value) is greater than 32 no error has occured, exit function
  Select Case FncRet
   Case 0: 'If FncRet evaluates to 0 then...
    oBuffer = "The operating system is out of memory or resources."
    'initialize oBuffer(outBuffer) with the appropriate error message
   Case ERROR_FILE_NOT_FOUND: 'FncRet evaluates to the constant ERROR_FILE_NOT_FOUND's value then...
    oBuffer = "The specified file was not found." '...
   Case ERROR_PATH_NOT_FOUND: '...
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
   'Return true since the return value specifies an error
End Function

Public Function FormatWEError(FncRet&, oBuffer$) As Boolean
'This function determines if the return value of a WinExec function call is an error
 If FncRet > 31 Then FormatWEError = False: Exit Function
 'No error has occured, exit this procedure
  Select Case FncRet
   Case 0: 'if FncRet evaluates to 0 then...
    oBuffer = "The operating system is out of memory or resources."
    'copy the appropriate error message to the argument oBuffer$(variable reference;ByRef keyword isn't included in the argument statement, but all arguments are passed as ByRef args. by default(by variable reference))...
   Case ERROR_FILE_NOT_FOUND: '...
    oBuffer = "The specified file was not found."
   Case ERROR_PATH_NOT_FOUND:
    oBuffer = "The specified path was not found."
   Case ERROR_BAD_FORMAT:
    oBuffer = "The .EXE file is not a valid Microsoft Win32® PE Header File, or an error has occured in the executable image."
   Case Else:
    oBuffer = "An error was generated, but the exact cause of the error is un-known."
  End Select
   FormatWEError = True
   'return true since the return value is an error
   'see frmShell for an example use of this function and function FormatSEError...
End Function

Public Function GetSystemRoot() As String
'Returns the system's System directory(special folder)
Dim strBuffer$, slen&: strBuffer = String(MAX_PATH, 0)
'dimensionalize strBuffer as string data type, slen as long data type
'initialize strBuffer with the return of the string function which will return a string of the specified length(MAX_PATH constant specifies the maximum amount of characters that a path can consist of)
 slen = GetSystemDirectory(strBuffer, MAX_PATH)
 'copy the system directory path to strBuffer, function returns the strings length
  GetSystemRoot = Left$(strBuffer, slen)
  'return the system directory( Left function returns the specified length of a string, in this case the length of the directory)
End Function

Public Function glbHex(Data As Variant) As String
On Error GoTo errh 'on the event of an error jump to label errh
If VarType(Data) = vbString Then Data = Val(Data)
'if the variable type of variable data evaluates to String type, then set variable Data's value to the number value of the variables value
 glbHex = Hex$(CLng(Data))
 'return the hexidecimal string of the number
  Do Until Len(glbHex) = 8
  'do until loop; loops until the length of glbHex evaluates to eight
   glbHex = "0" & glbHex
   'add a zero
  Loop 'next iteration if loops condition is true
   Exit Function 'exit this procedure
errh: 'label errh
 glbHex = "00000000"
 'return 0, as an error occured
End Function

