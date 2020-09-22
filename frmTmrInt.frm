VERSION 5.00
Begin VB.Form frmTmrInt 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gloabal Memory Monitor"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   71
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   267
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   2340
      TabIndex        =   6
      Top             =   720
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
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
      Left            =   3180
      TabIndex        =   5
      Top             =   720
      Width           =   795
   End
   Begin VB.TextBox txtInt 
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
      Left            =   2100
      TabIndex        =   3
      Top             =   420
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "milliseconds"
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
      Left            =   3060
      TabIndex        =   4
      Top             =   480
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update Memory Status every"
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
      TabIndex        =   2
      Top             =   480
      Width           =   1980
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000 milliseconds is equal to 1 second."
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
      TabIndex        =   1
      Top             =   180
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note that the interval is in Milliseconds"
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
      TabIndex        =   0
      Top             =   60
      Width           =   2760
   End
End
Attribute VB_Name = "frmTmrInt"
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


Dim uInterval As Long

Private Sub cmdCancel_Click()
 uInterval = frmMain.tmrMem.Interval
  txtInt.Text = CStr(uInterval)
   Me.Hide
End Sub

Private Sub cmdSave_Click()
 frmMain.tmrMem.Interval = uInterval
  Me.Hide
End Sub

Private Sub Form_Activate()
 TransPrep Me.hwnd
End Sub

Private Sub Form_Load()
 uInterval = CLng(GetSetting("ProcessXP", "Settings", "TMI", "1500"))
  txtInt.Text = CStr(uInterval)
   VGrad Me.hdc, Me.ScaleHeight, Me.ScaleWidth, COLOR_BTNFACE, white2black
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Me.Visible = True Then Cancel = 1: Me.Hide: frmMain.tmrMem.Interval = uInterval: Exit Sub
  SaveSetting "ProcessXP", "Settings", "TMI", CStr(uInterval)
End Sub

Private Sub txtInt_Change()
 txtInt.Text = Val(txtInt.Text)
  uInterval = CLng(txtInt.Text)
End Sub
