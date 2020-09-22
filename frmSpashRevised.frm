VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSpash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1860
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView ListView1 
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   2220
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   217
      TabIndex        =   2
      Top             =   780
      Width           =   3255
      Begin VB.Label lblCurProc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Processing ""VB6.exe"""
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
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label lblProgress 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading ProcessXP Registry Keys.."
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
         TabIndex        =   4
         Top             =   60
         Width           =   2385
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salex Software © 2001 - 2003"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   600
         TabIndex        =   3
         Top             =   900
         Width           =   2595
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   219
      TabIndex        =   1
      Top             =   585
      Width           =   3285
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      Picture         =   "frmSpashRevised.frx":0000
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   220
      TabIndex        =   0
      Top             =   0
      Width           =   3300
   End
End
Attribute VB_Name = "frmSpash"
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


Private Sub Form_Initialize()
  Dim rICc&
   rICc = InitCommonControls
   'The only purpose of this function call is to request that windows check for this application manifest file to determine which version of Microsoft's Common Controls to use
   'If the version determined by the manifest(XML) file is the version which utilizes the UxTheme library than the controls will be drawn according to the current Theme Data
   'NOTE: To maintain the same functionality while developing you're WindowsXP theme compatible application in c++ refer to MSDN's online XP Theme documentation or the documents in the Software Developement Kit which discuss the UxTheme system
   'I have personally dealt with many frustrating problems which arrise from this,
   'basically I have determined that its a good idea to call this function during the intialization of the first loaded form
   'and even if none of the common controls are being initialized on the first form loaded, create a common control on that form. If you will not be using it then set its visibility property to false...
   'Also, do not try to call any functions which show modal dialogs before the first common control has been initialized.
   'And make sure if the manifest file is included with the distributed version of you're application file that you call InitCommonControls or you're application will fail to start when executed on WindowsXP
   'The window control's visual face-lift is well worth it ;)
End Sub

Private Sub Form_Load()
On Error Resume Next
'On the event of an error resume execution on the next line
 Starting = True 'This flag is used by other procedures of other forms, for example when the process list in frmMain is being populated, it evaluates the value of this flag, if its true it will call the sub UpdateProgress of this object to visually inform the user the progress of the list population
  VGrad picFrame.hdc, picFrame.ScaleHeight, picFrame.ScaleWidth, COLOR_BTNFACE, white2black
   HGrad Picture2.hdc, Picture2.ScaleHeight, Picture2.ScaleWidth, COLOR_ACTIVECAPTION, Black2White, UNEQUALITYFADE: Me.Show
   'See this functions in mdlGUI for more information...
   If GetSetting("ProcessXP", "Startup", "CI", "0") = "0" Then CurrentInfo = ProcessInfo Else CurrentInfo = VersionInfo
    If LCase$(GetSetting("ProcessXP", "Startup", "MU", "1")) = "false" Then MemUpdateOn = False Else MemUpdateOn = True
     ProcRelation = Val(GetSetting("ProcessXP", "Options", "ProcRelation", "0"))
      ModRelation = Val(GetSetting("ProcessXP", "Options", "ModRelation", "1"))
       ShowProcIcon = Val(GetSetting("ProcessXP", "Options", "ProcIcon", "1"))
        ShowModIcon = Val(GetSetting("ProcessXP", "Options", "ModIcon", "1"))
         KeepWinTop = Val(GetSetting("ProcessXP", "Options", "WinOnTop", "1"))
          WinTrans = Val(GetSetting("ProcessXP", "Options", "TransVal", "230"))
           KeepTrans = Val(GetSetting("ProcessXP", "Options", "KeepTrans", "1"))
           'Retrieve information stored in registry keys, they keys are unique to each user who runs this application since these keys are stored with in HKEY_CURRENT_USER
            TransPrep Me.hwnd 'See TransPrep function..
             UpdateWinPos Me.hwnd 'See UpdateWinPos function
              Err.Clear 'Clear the current error info if any...
               playSounds = CBool(GetSetting("ProcessXP", "Options", "EnableSoundServer", "1"))
               'Retrieve registry key value...
                If playSounds = True Then Set sndClass = New clsMain Else sndSupported = True
                'If playSounds evaluates to true then create an instance of the sound servers main class
                 sndSupported = playSounds 'sndSupported flag is evaluated before the attempt of playing sounds..
                  If Err.Number = 429 Then sndSupported = False: MsgBox "A required Dynamic-Link-Library file (""sndServer.dll"") is missing, please re-install ProcessXP." & vbCrLf & vbCrLf & "ProcessXP will be unable to fully utilize it's ability to play sounds.", vbExclamation, "Error - sndServer.dll"
                  'The component wasn't found, set the sndSupported flag to false so that the methods of the sndServer are called which would raise additional errors
                   Err.Clear 'Clear error info if any...
        If Trim(GetString(HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "TP")) = "" Then
         TerminationPriv = OnlyAdmin
        Else
         TerminationPriv = Val(Trim(GetString(HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "TP")))
        End If
         If Trim(GetString(HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "ACEW")) = "" Then
          AllCanEnumWin = False
         Else
          AllCanEnumWin = Val(Trim(GetString(HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "ACEW")))
         End If
          If Trim(GetString(HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "ACMW")) = "" Then
           AllCanManipWin = False
          Else
           AllCanManipWin = Val(Trim(GetString(HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "ACMW")))
          End If
          'Retrieve Security related registry keys which are non-user dependant(restricted users are effected by these flags)
           lblProgress.Caption = "Drawing Components GUI..." '...
            Load frmSearchThread: Load frmExplain
            If Err.Number = 339 Then GoTo DepNF
            'Error 339 indicates a component wasn't initialized(most likely improperly installed[wrong version])
            'if the error occured then jump to label DepNF
             Load frmPointWindow
             If Err.Number = 339 Then GoTo DepNF
              Load frmShell
              If Err.Number = 339 Then GoTo DepNF
               Load frmSearch
               If Err.Number = 339 Then GoTo DepNF
                Load frmTmrInt
                If Err.Number = 339 Then GoTo DepNF
                 Load frmAbout
                 If Err.Number = 339 Then GoTo DepNF
                  Load frmOptions
                  If Err.Number = 339 Then GoTo DepNF
                   Load frmSec
                   If Err.Number = 339 Then GoTo DepNF
                    lblProgress.Caption = "Loading Main Window..."
                     Load frmMain
                     If Err.Number = 339 Then GoTo DepNF
                      Starting = False 'Application is no longer loading...
                       Unload Me 'Unload this object...
                        Exit Sub
                        'Discontinue the execution of this procedure...

DepNF: 'DepNF label
 Dim form As form
  For Each form In Forms
  'For Each loop; enumerates through all form elements in the Forms collection
  'The Forms collection consists of only loaded form objects...
   If LCase(form.Name) <> "frmspash" Then
   'If the lower case value of the current form's name is unequal to this forms name then...
    Unload form 'Unload the form
   End If
  Next form 'select the next element in the forms collection
   MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found"
    End 'delete all memory and loaded objects
End Sub

Public Sub UpdateProgress(Optional cValue& = -1, Optional cMax& = -1, Optional NewTitle$ = "", Optional NewProc$ = "")
'This procedure is called by other form procedures when the Starting flag evaluates to true
On Error Resume Next
'On the event of an error resume execution on the next line
Dim tPercent& 'dimensionalize tPercent as long data type
 If NewProc$ <> "" Then lblCurProc.Caption = "Processing """ & NewProc & """...": lblCurProc.Visible = True
 'update label's caption property
  If NewProc = "" Then lblCurProc.Visible = False 'update label's caption property
   If cValue& <> -1 And cMax& <> -1 Then tPercent = (cValue / cMax) * 100
    If NewTitle$ <> "" And cValue& <> -1 And cMax& <> -1 Then
    'Calculate the percent complete of the procedure if requested(if the optional arguments aren't empty ofcourse)
     lblProgress.Caption = NewTitle$ & "... " & CStr(tPercent) & "%"
     'update label's caption property
    Else
     If NewTitle$ <> "" Then lblProgress.Caption = NewTitle$ & "..."
     'update label's caption property
    End If
End Sub

