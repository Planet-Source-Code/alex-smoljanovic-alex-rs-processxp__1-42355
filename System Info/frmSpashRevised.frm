VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSpash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1875
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3255
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ListView ListView1 
      Height          =   435
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   767
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   495
      Left            =   60
      TabIndex        =   6
      Top             =   1980
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
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

Private Sub Form_Load()
On Error Resume Next
Dim Form As Form
If HostOS.OperatingSystem = Win2K Then StretchBlt picLogo.hdc, 0, 0, picLogo.ScaleWidth, picLogo.ScaleHeight, picLogo.hdc, 0, 0, picLogo.ScaleWidth, picLogo.ScaleHeight, SRCCOPY
  VGrad picFrame.hdc, picFrame.ScaleHeight, picFrame.ScaleWidth, COLOR_BTNFACE, White2Black
   HGrad Picture2.hdc, Picture2.ScaleHeight, Picture2.ScaleWidth, COLOR_ACTIVECAPTION, Black2White, UNEQUALITYFADE: Me.Show
    lblProgress.Caption = "Drawing Components GUI..."
     Load frmAbout
      lblProgress.Caption = "Loading Main Window..."
       Load frmSysInfo
        frmSysInfo.Show
         If Err.Number = 339 Then GoTo DepNF
          Exit Sub
DepNF:
  For Each Form In Forms
   If LCase(Form.Name) <> "frmspash" Then
    Unload Form
   End If
  Next Form
   MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found"
    End
End Sub


