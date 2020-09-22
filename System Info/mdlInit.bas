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
Public Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public Sub Main()
On Error Resume Next
Dim rICc&
 rICc = InitCommonControls
  If RetOSInf() = False Then
   If MsgBox("System Info must determine what Operating System it is running on." & vbCrLf & vbCrLf & "While it attempted to retreive the systems operating system information an error occured." & vbCrLf & "It might be unsafe to run System Info, it is compatible with Windows 2000, and Windows XP." & vbCrLf & vbCrLf & "Do you wish to run System Info anyway?", vbQuestion + vbYesNo, "Initialization Error") = vbNo Then End
  End If
   If HostOS.DispDevBits < 16 Then
    If MsgBox("Please change you're Display Color Quality setting to atleast support 16-bit (65,536) colors." & vbCrLf & "You're display device is currently only supporting " & HostOS.DispDevDescription & "." & vbCrLf & vbCrLf & "Do you want to change you're Display's Color Quality setting?", vbQuestion + vbYesNo, "System Info Requires 16-bit Colors") = vbYes Then
     ShowDispSet
    End If
     MsgBox "Is it required that you restart System Info after you have changed you're display settings.", vbInformation + vbOKOnly, "System Info"
      End
   End If
    If HostOS.OperatingSystem = inCompatibleOS Then
     Err.Clear
      Load frmIOS
       If Err.Number = 339 Then MsgBox "Incompatible Operating System." & vbCrLf & "Additionally, an error occured while initializing a component." & vbCrLf & vbCrLf & Err.Description, vbCritical, "System Info": End
        Exit Sub
    End If
     Err.Clear
      Load frmSpash
       If Err.Number = 339 Then MsgBox "A component failed to initialize, or wasn't found." & vbCrLf & "Please re-install System Info, or use the ProcessXP Installation 'Repair' feature.", vbExclamation, "System Info Component Error": End
End Sub


Private Sub ShowDispSet()
On Error GoTo errh
Dim iShell As Shell, iFolder As Folder, iFIV As FolderItemVerbs, i&, j&
 Set iShell = New Shell
  Set iFolder = iShell.NameSpace(ssfCONTROLS)
   For i = 0 To iFolder.Items.Count - 1
    If InStr(1, LCase(iFolder.Items.Item(i).Name), "display", 1) > 0 Then
     Set iFIV = iFolder.Items.Item(i).Verbs
      For j = 0 To iFIV.Count - 1
       If InStr(1, iFIV.Item(j).Name, "open", 1) > 0 Then
        iFIV.Item(j).DoIt
         Exit For
       End If
      Next j
       Exit For
    End If
   Next i
    Exit Sub
errh:
 MsgBox "An error has occured while attempting to show Display Settings." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Display Setting Error"
End Sub
