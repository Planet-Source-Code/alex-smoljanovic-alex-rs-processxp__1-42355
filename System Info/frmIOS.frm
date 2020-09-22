VERSION 5.00
Begin VB.Form frmIOS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Incompatible Operating System"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5160
   Icon            =   "frmIOS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   1020
      Width           =   915
   End
   Begin VB.Label lblOS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You're Operating System is incompatible"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   60
      TabIndex        =   4
      Top             =   960
      Width           =   3075
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 2003 .NET Family Server"
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
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   2355
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows XP (NT 5.1)"
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
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 2000 (NT 5)"
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
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Information is compatible only on:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3180
   End
End
Attribute VB_Name = "frmIOS"
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

Private Sub Command1_Click()
 End
End Sub

Private Sub Form_Load()
 Unload frmSpash
  Me.Show
    If HostOS.OSSpecs = Win32s Then
     lblOS.Caption = "Windows 3.x " & "build " & HostOS.OSBuild
    ElseIf HostOS.OSSpecs = Win95 Then
     lblOS.Caption = "Windows 95" & HostOS.OSSpecsEx & " build " & HostOS.OSBuild
    ElseIf HostOS.OSSpecs = Win98 Then
     lblOS.Caption = "Windows 98 " & "build " & HostOS.OSBuild
    ElseIf HostOS.OSSpecs = Win98SE Then
     lblOS.Caption = "Windows 98 SE " & "build " & HostOS.OSBuild
    ElseIf HostOS.OSSpecs = WinME Then
     lblOS.Caption = "Windows ME " & "build " & HostOS.OSBuild
    ElseIf HostOS.OSSpecs = WinNT4 Then
     lblOS.Caption = "Windows NT 4.x " & "build " & HostOS.OSBuild
    End If
End Sub


