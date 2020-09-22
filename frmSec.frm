VERSION 5.00
Begin VB.Form frmSec 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Security"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
   Icon            =   "frmSec.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Window Thread Allowances"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   4
      Top             =   1140
      Width           =   4275
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   60
         ScaleHeight     =   615
         ScaleWidth      =   4095
         TabIndex        =   5
         Top             =   240
         Width           =   4095
         Begin VB.CheckBox chkAllManipWindow 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Users can manipulate Window Threads"
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
            Left            =   60
            TabIndex        =   7
            Top             =   240
            Width           =   3975
         End
         Begin VB.CheckBox chkAllEnumWindow 
            BackColor       =   &H00FFFFFF&
            Caption         =   "All Users can enumerate a Processes Window Threads"
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
            Left            =   60
            TabIndex        =   6
            Top             =   0
            Width           =   3975
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Process Allowances"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   4275
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   60
         ScaleHeight     =   615
         ScaleWidth      =   4035
         TabIndex        =   1
         Top             =   240
         Width           =   4035
         Begin VB.OptionButton optAllTerminate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Allow all users to Terminate Processes"
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
            Left            =   0
            TabIndex        =   3
            Top             =   300
            Width           =   3975
         End
         Begin VB.OptionButton optAdminTerminate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Allow only Administrators to Terminate Processes"
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
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Value           =   -1  'True
            Width           =   3975
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   296
      TabIndex        =   8
      Top             =   0
      Width           =   4440
   End
   Begin VB.PictureBox picBot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1500
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   9
      Top             =   1980
      Width           =   3075
      Begin VB.Label Label1 
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
         Left            =   2400
         MouseIcon       =   "frmSec.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSec"
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


Private Sub chkAllEnumWindow_Click()
'Update global security flags...
 If chkAllEnumWindow.Value = 1 Then AllCanEnumWin = True Else AllCanEnumWin = False
  If AllCanEnumWin = False Then chkAllManipWindow.Enabled = False Else chkAllManipWindow.Enabled = True
End Sub

Private Sub chkAllManipWindow_Click()
'Update global security flags...
 If chkAllManipWindow.Value = 1 Then AllCanManipWin = True Else AllCanManipWin = False
End Sub

Private Sub Form_Activate()
 TransPrep Me.hwnd 'See TransPrep function for more info...
End Sub

Private Sub Form_Load()
'Set the initial states of the check boxes and radio buttons based upon the values of the global security flags
 If TerminationPriv = OnlyAdmin Then
  optAdminTerminate.Value = True: optAllTerminate.Value = False
 ElseIf TerminationPriv = AllUsers Then
  optAllTerminate.Value = True: optAdminTerminate.Value = False
 End If
  If AllCanEnumWin = True Then chkAllEnumWindow.Value = 1 Else chkAllEnumWindow.Value = 0
   If AllCanManipWin = True Then chkAllManipWindow.Value = 1 Else chkAllManipWindow.Value = 0
     HGrad Picture3.hdc, Picture3.ScaleHeight, Picture3.ScaleWidth, COLOR_ACTIVECAPTION, Black2White, UNEQUALITYFADE
      HRGrad picBot.hdc, picBot.ScaleHeight, picBot.ScaleWidth, COLOR_ACTIVECAPTION
      'See HGrad and HRGrad(mdlGUI) 'intermediate' functions for more info...
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Me.Visible = True Then Cancel = 1: Me.Hide: Exit Sub
 'If this window is visible it is not to be unloaded, hide it. it will be unloaded when frmMain is unloaded
  If TerminationPriv = AllUsers Then SaveString HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "TP", "1" Else SaveString HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "TP", "0"
   If AllCanEnumWin = True Then SaveString HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "ACEW", "1" Else SaveString HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "ACEW", "0"
    If AllCanManipWin = True Then SaveString HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "ACMW", "1" Else SaveString HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcXP\Sec", "ACMW", "0"
    'Save the security flags in the registry.
    'Since the registry key stored and retrieved by SaveSetting and GetSetting are User Dependant(HKEY_CURRENT_USER), we will save these settings in the registry to keys which can be retrieved by all users
End Sub


Private Sub Label1_Click()
 If sndSupported = True Then sndClass.doSec
 'For more information on the sound server, see the sndServer Active-X DLL Project which should have been included in the compressed ZIP file from which this project was extracted
  frmExplain.DoDlg "sec", Me
  'Show the Explain dialog, see this objects DoDlg method for more info...
End Sub

Private Sub optAdminTerminate_Click()
 If optAdminTerminate.Value = True Then TerminationPriv = OnlyAdmin
 'Update global variable
End Sub

Private Sub optAllTerminate_Click()
 If optAllTerminate.Value = True Then TerminationPriv = AllUsers
 'Update global variable
End Sub
