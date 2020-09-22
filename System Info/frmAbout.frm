VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBillBoard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   2940
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   268
      TabIndex        =   0
      Top             =   0
      Width           =   4050
      Begin VB.Label lblCompiled 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000005&
         Height          =   150
         Left            =   60
         TabIndex        =   13
         Top             =   0
         Width           =   3870
      End
   End
   Begin VB.PictureBox picBot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   470
      TabIndex        =   10
      Top             =   3600
      Width           =   7050
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":2E898
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   435
         Left            =   60
         TabIndex        =   11
         Top             =   120
         Width           =   6915
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   249
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   1
      Top             =   0
      Width           =   3075
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   60
         Picture         =   "frmAbout.frx":2E9A8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":2F272
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   60
         TabIndex        =   12
         Top             =   1020
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System Information"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   120
         Width           =   1950
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.02.13"
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
         Left            =   2340
         TabIndex        =   8
         Top             =   0
         Width           =   540
      End
      Begin VB.Label lblName 
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
         Left            =   60
         TabIndex        =   7
         Top             =   3300
         Width           =   2850
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registered to:"
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
         TabIndex        =   6
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label lblOrg 
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
         Left            =   60
         TabIndex        =   5
         Top             =   3480
         Width           =   2460
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salex Software© 2003"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F27E53&
         Height          =   150
         Left            =   1350
         MouseIcon       =   "frmAbout.frx":2F37C
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   2460
         Width           =   1545
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DigiScene Studios© 2003"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F27E53&
         Height          =   150
         Left            =   1080
         MouseIcon       =   "frmAbout.frx":2F686
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   2640
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmAbout"
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
Dim rUser$, rCompany$
 VGrad picLeft.hdc, picLeft.ScaleHeight, picLeft.ScaleWidth, COLOR_BTNFACE, White2Black
  VGrad picBot.hdc, picBot.ScaleHeight, picBot.ScaleWidth, COLOR_BTNFACE, Black2White
   lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
    If Mid$(App.Path, Len(App.Path) - 1) = "\" Then
     lblCompiled.Caption = "Compiled on " & FileDateTime(App.Path & App.EXEName)
    Else
     lblCompiled.Caption = "Compiled on " & FileDateTime(App.Path & "\" & App.EXEName & ".exe")
    End If
     rUser = GetString(HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcessXP\RegUser", "UserName")
      rCompany = GetString(HKEY_LOCAL_MACHINE, "Software\Salex Software\ProcessXP\RegUser", "Company")
       If rUser <> "" Then lblName.Caption = rUser Else lblName.Caption = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
        If rCompany <> "" Then lblOrg.Caption = rCompany Else lblOrg.Caption = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")
         If lblName.Caption = "" Then lblName.Caption = "Can't Retreive Name": lblName.ForeColor = vbRed
          If lblOrg.Caption = "" Then lblOrg.Caption = "Can't Retreive Company": lblOrg.ForeColor = vbRed
End Sub

Private Sub Label1_Click()
 Me.Hide
End Sub

Private Sub Label3_Click()
 Me.Hide
End Sub

Private Sub Label5_Click()
 Me.Hide
End Sub

Private Sub Label6_Click()
 ShellExecute Me.hwnd, "open", "mailto:salex_software@shaw.ca?subject=ProcessXP", vbNullString, vbNullString, vbNormal
End Sub

Private Sub Label7_Click()
 Me.Hide
End Sub

Private Sub Label9_Click()
 ShellExecute Me.hwnd, "open", "mailto:salex_software@shaw.ca?subject=ProcessXP", vbNullString, vbNullString, vbNormal
End Sub

Private Sub lblCompiled_Click()
 Me.Hide
End Sub

Private Sub lblName_Click()
 Me.Hide
End Sub

Private Sub lblOrg_Click()
 Me.Hide
End Sub

Private Sub lblVersion_Click()
 Me.Hide
End Sub

Private Sub picBillBoard_Click()
 Me.Hide
End Sub

Private Sub picBot_Click()
 Me.Hide
End Sub

Private Sub picLeft_Click()
 Me.Hide
End Sub

Private Sub Picture2_Click()
 Me.Hide
End Sub
