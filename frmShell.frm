VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmShell 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Task"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   Icon            =   "frmShell.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   117
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   279
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbShellExec 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      ItemData        =   "frmShell.frx":08CA
      Left            =   1260
      List            =   "frmShell.frx":08D4
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1140
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Create Process"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2940
      TabIndex        =   7
      Top             =   1140
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   3060
      Picture         =   "frmShell.frx":08E4
      ScaleHeight     =   60
      ScaleWidth      =   1125
      TabIndex        =   6
      Top             =   1020
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   720
      Width           =   435
   End
   Begin VB.TextBox txtPEFile 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   660
      TabIndex        =   4
      Top             =   720
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3540
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   180
      Picture         =   "frmShell.frx":0CB8
      ScaleHeight     =   45
      ScaleWidth      =   3780
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   60
      Picture         =   "frmShell.frx":15D8
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Explain"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   3660
      MouseIcon       =   "frmShell.frx":221A
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblSE 
      BackStyle       =   0  'Transparent
      Caption         =   "Send command:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PE File:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This will create a new process specified by the file you choose."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   180
      Width           =   3435
   End
End
Attribute VB_Name = "frmShell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Private Enum ShellType
 eShellExec = 0
  eWinExec = 1
End Enum

Dim ShellAction As ShellType
'General Declarations...

Private Sub cmdRun_Click()
Dim WExecPath$, SExErr$, WExErr$, sOp$: WExecPath$ = txtPEFile.Text: sOp$ = LCase$(cmbShellExec.List(cmbShellExec.ListIndex))
'Dimensionalize and initialize variables....
 If ShellAction = eWinExec Then
 'This variable determines which method should be used to perform the desired action
 'If the user has specified an executable image the variable ShellAction will evaluate to eWinExec(prefix e specifies Enumeration Member, this helps preventing ambigous variable/object names)
 'If the user has specified a file which depends on an executable image, this variable will evaluate to eShellExec. The shell exec function will execute the program associated with the file specified to 'Shell' with the Shell formatted command line specifying the file to open
   If Left$(WExecPath$, 2) <> "\""" Then WExecPath$ = "\""" & WExecPath$
   'If the first to characters of the WinExec File path don't evaluate to '\"' then...
    If InStr(1, WExecPath, ".exe ", 1) <> 0 Then
    'If ".exe " exists with in the WinExec file path then...
     WExecPath = Left$(WExecPath, InStr(1, WExecPath, ".exe ", 1) + 4) & "\""" & Mid$(WExecPath, InStr(1, WExecPath, ".exe ", 1) + 6)
     '*Read the commented statements below for an explanation...
    Else
     If LCase$(Right$(WExecPath, 4)) = ".exe" Then
     'If the path doesn't include a command line append the following...
      WExecPath = WExecPath & "\"""
     Else
      MsgBox "Invalid Path Specified." & vbCrLf & vbCrLf & "To specify a valid path conform to the following guidlines:" & vbCrLf & "DirectoryPath/File.exe CommandLine...", vbExclamation, "Invalid Path"
     End If
     'The Commandline Paramater of the WinExec error will be formatted as follows;
     'Pointer to a null-terminated character string that contains the command line (file name plus optional parameters) for the application to be executed. If the name of the executable file in the lpCmdLine parameter does not contain a directory path, the system searches for the executable file in this sequence:
     'The directory from which the application loaded.
     'The current directory.
     'The Windows system directory. The GetSystemDirectory function retrieves the path of this directory.
     'The Windows directory. The GetWindowsDirectory function retrieves the path of this directory.
     'The directories listed in the PATH environment variable.
     'Use environ function to retrieve the environmental variables...
     
     'The reason the commandline paramater is formated with the backslash characters is to prevent a security risk
     'The executable name is treated as the first white space-delimited string in command line argument. If the executable or path name has a space in it, there is a risk that a different executable could be run because of the way the function parses spaces. "WinExec("C:\Program Files\MyApp", ...)" is dangerous because the function will attempt to run "Program.exe", if it exists, instead of "MyApp.exe".
     
     'If a malicious user were to create an application called "Program.exe" on a system, any program that incorrectly calls WinExec using the Program Files directory will run this application instead of the intended application.
     'To avoid this problem, use CreateProcess rather than WinExec. However, if you must use WinExec for legacy reasons, make sure the application name is enclosed in quotation marks as shown in the example below.
     'WinExec("\"C:\Program Files\MyApp.exe\" -L -S", ...)

     
    End If
     If FormatWEError(WinExec(txtPEFile.Text, SW_SHOWNORMAL), WExErr$) = True Then
     'Function FormatWEError(Format WinExec Error) validates the return of the WinExec function, see FormatWEError(mdlMain) function for more info...
      MsgBox WExErr$, vbExclamation, "Error"
      'Display formatted error message...
     Else
      If MsgBox("The action was successful, would you like to refresh the process list?", vbQuestion + vbYesNo, "Refresh Process List") = vbYes Then
      'No error occured...
       enumProcesses 'Enumerate processes; see function enumProcess for more info...
        frmMain.RefreshLibrary 'Re-populate the process list, see frmMain -> RefreshLibrary function for more info...
      End If
     End If
 Else
   If FormatSEError(ShellExecute(Me.hwnd, sOp$, txtPEFile, ByVal 0&, ByVal 0&, SW_SHOWNORMAL), SExErr$) = True Then
   'ShellExec method will be used, call FormatSEError to validate the return of ShellExecute function to determine if an error occured, if an error did occur then create an error message
   'See FormatSEError(mdlMain) function for more info...
    MsgBox SExErr$, vbExclamation, "Error"
    'Display formatted error message
   Else
    If MsgBox("The action was successful, would you like to refresh the process list?", vbQuestion + vbYesNo, "Refresh Process List") = vbYes Then
     enumProcesses 'Enumerate processes; see function enumProcess for more info...
      frmMain.RefreshLibrary 'Re-populate the process list, see frmMain -> RefreshLibrary function for more info...
    End If
   End If
 End If
End Sub

Private Sub Command1_Click()
On Error GoTo errh
'On the event of an error jump to label errh
 CD.Filter = "PE File (*.exe)|*.exe|Associated Program|*.*"
 'Set the filter flags of the common dialog
 'Format: Description1|Extension1|Description2|Extension2;Extension3;Extension4|...
  CD.ShowOpen 'Show the dialog
   If CD.FilterIndex = 1 Then
   'If the filter's selection index when the file was selected evaluates to 1 then ...
     txtPEFile.Text = GetShortPath(CD.FileName)
     'See GetShortPath function for more info...(Returns DOS-Short File Names)
      cmbShellExec.Visible = False: lblSE.Visible = False
      'Update window...
        ShellAction = eWinExec
        'Update ShellAction flag, see sub routine cmdRun_Click for more info on this flag...
   Else
    txtPEFile.Text = CD.FileName
    'Update text box's text property
     cmbShellExec.Visible = True: lblSE.Visible = True
     'Update window
       ShellAction = eShellExec
       'Update ShellAction flag, see sub routine cmdRun_Click for more info on this flag...
   End If
  Exit Sub
errh:
 If Err.Number = 32755 Then Exit Sub
 'If the error 32755 occured then the user cancelled the dialog, don't update variables or flags. Exit this sub routine
  MsgBox Err.Description, vbCritical, "Error [" & Err.Number & "]"
  'Inform user of the un-expected error
End Sub

Private Sub Form_Activate()
 TransPrep Me.hwnd 'See function TransPrep for more info...
End Sub

Private Sub Form_Load()
 HGrad Me.hdc, Me.ScaleHeight, Me.ScaleWidth, COLOR_BTNFACE, white2black, NOFADE
 'See HGrad function for more info...(mdlGUI, this function formats the arguments to call the actual procedure in the ProcXP GUI DLL)
  cmbShellExec.ListIndex = 0
  'Select the first item in the combo box
   SHAutoComplete txtPEFile.hwnd, SHACF_FILESYSTEM Or SHACF_USETAB Or SHACF_AUTOAPPEND_FORCE_ON
   'SHAutoComplete performs actions when text is typed into the text box..
   'The SHACF_FILESYSTEM flag specifies the information to query from whilst the user is typing
   'Other flags include:
    'SCOPE Const SHACF_AUTOAPPEND_FORCE_OFF = &H80000000
    'SCOPE Const SHACF_AUTOAPPEND_FORCE_ON = &H40000000
    'SCOPE Const SHACF_AUTOSUGGEST_FORCE_OFF = &H20000000
    'SCOPE Const SHACF_AUTOSUGGEST_FORCE_ON = &H10000000
    'SCOPE Const SHACF_DEFAULT = &H0
    'SCOPE Const SHACF_FILESYS_ONLY = &H10
    'SCOPE Const SHACF_FILESYSTEM = &H1
    'SCOPE Const SHACF_URLALL = (SHACF_URLHISTORY Or SHACF_URLMRU)
    'SCOPE Const SHACF_URLHISTORY = &H2
    'SCOPE Const SHACF_URLMRU = &H4
    'SCOPE Const SHACF_USETAB = &H8
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Me.Visible = True Then Cancel = 1: Me.Hide
 'If this form is visible, it is not to be unloaded, hide it, this will bring back focus to the window which showed this form as a modal dialog
End Sub

Private Sub Label3_Click()
 If sndSupported = True Then sndClass.doShell
 'See sound server project files for more info...
  frmExplain.DoDlg "shell", Me
  'Show the Explain dialog, so this object DoDlg method for more info...
End Sub

Private Sub txtPEFile_Change()
'Determine if the file currently specified is a file which is to be executed using the WinExec or ShellExecute method...
 If InStr(1, txtPEFile.Text, ".exe ", 1) = 0 Then
 'Determine the position of a substring within a string
  If LCase$(Right$(txtPEFile.Text, 4)) = ".exe" Then
   ShellAction = eWinExec
   'Update flag
    cmbShellExec.Visible = False: lblSE.Visible = False
    'Update window
  Else
   ShellAction = eShellExec
   'Update shell flag
    cmbShellExec.Visible = True: lblSE.Visible = True
    'Update window
  End If
 Else
  ShellAction = eWinExec 'Update shell flag
   cmbShellExec.Visible = False: lblSE.Visible = False 'Update window
 End If
End Sub
