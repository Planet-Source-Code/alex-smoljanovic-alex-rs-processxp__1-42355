Attribute VB_Name = "mdlInit"
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

Public Declare Function IsDebuggerPresent Lib "kernel32.dll" () As Long
'IsDebuggerPresent function determines whether the calling process is running under the context of a debugger...

Public Sub Main()
On Error Resume Next 'On the event of an error resume execution of this procedure on the next line
If IsDebuggerPresent Then End 'If this process is being debugged, terminate...
 Dim rICc& 'Dimensionalize rICc as long data type
 rICc = InitCommonControls 'Initialize Common Controls
 'See the comments in the frmSpash object for explanations about InitCommonControls
  If RetOSInf() = False Then 'See function RetOSInf(Return Operating Systen Information)
   If MsgBox("ProcessXP must determine what Operating System it is running on." & vbCrLf & vbCrLf & "While it attempted to retreive the systems operating system information an error occured." & vbCrLf & "It might be unsafe to run ProcessXP, it is compatible with Windows 2000, and Windows XP." & vbCrLf & vbCrLf & "Do you wish to run ProcessXP anyway?", vbQuestion + vbYesNo, "Initialization Error") = vbNo Then End
   'An error occured while trying to retreive OS info...
  End If
   If HostOS.DispDevBits < 16 Then
   'if the display adapter color setting is less than 16-bit(65,536 colors) then...
    If MsgBox("Please change you're Display Color Quality setting to atleast support 16-bit (65,536) colors." & vbCrLf & "You're display device is currently only supporting " & HostOS.DispDevDescription & "." & vbCrLf & vbCrLf & "Do you want to change you're Display's Color Quality setting?", vbQuestion + vbYesNo, "ProcessXP Requires 16-bit Colors") = vbYes Then
     ShowDispSet
     'Show Display Settings Dialog...
    End If
     MsgBox "Is it required that you restart ProcessXP after you have changed you're display settings.", vbInformation + vbOKOnly, "ProcessXP"
      End 'Terminate
   End If
    If HostOS.OperatingSystem = inCompatibleOS Then
    'If OperatingSystem member of HostOS evaluates to inCompatibleOS then ...
     Err.Clear
      Load frmIOS 'Load Incompatible Operating System dialog
       If Err.Number = 339 Then MsgBox "Incompatible Operating System." & vbCrLf & "Additionally, an error occured while initializing a component." & vbCrLf & vbCrLf & Err.Description, vbCritical, "ProcessXP": End
       'Additionally an error occured while loaded the dialog...
        Exit Sub
    End If
     Err.Clear
      Load frmSpash 'Load the splash dialog...
       If Err.Number = 339 Then MsgBox "A component failed to initialize, or wasn't found." & vbCrLf & "Please re-install ProcessXP.", vbExclamation, "ProcessXP Component Error": Set sndClass = Nothing: End
       'A component failed to initialize, this application will terminate
End Sub


Private Sub ShowDispSet()
On Error GoTo errh 'on the event of an error jump to label errh
Dim iShell As Shell, iFolder As Folder, iFIV As FolderItemVerbs, i&, j&
'dimensionalize iShell as Shell type structure, iFolder as Folder type structure
'iFIV as FolderItemVerbs type structure, i and j as long data type
 Set iShell = New Shell 'Initialize iShell with a new instance of the Shell class
  Set iFolder = iShell.NameSpace(ssfCONTROLS)
  'initialize iFolder with the Folder type return of iShell's NameSpace method
   For i = 0 To iFolder.Items.Count - 1
   'For Next loop; i starts at 0, loops until i is equal to the number of folder items minus one, incrementing i by one each iteration
    If InStr(1, LCase(iFolder.Items.Item(i).Name), "display", 1) > 0 Then
    'determine the position of a substring within another string
     Set iFIV = iFolder.Items.Item(i).Verbs
     'initialize iFIV with the return of the specified folder item's verbs
      For j = 0 To iFIV.Count - 1
      'loop through each of the items verbs
       If InStr(1, iFIV.Item(j).Name, "open", 1) > 0 Then
       'if the substrin "open" exists in the verb's name then execute it
        iFIV.Item(j).DoIt
         Exit For
       End If
      Next j 'increment j, next iteration
       Exit For
    End If
   Next i 'increment i, next iteration
    Exit Sub
errh:
 MsgBox "An error has occured while attempting to show Display Settings." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Display Setting Error"
 'Display error message
End Sub
