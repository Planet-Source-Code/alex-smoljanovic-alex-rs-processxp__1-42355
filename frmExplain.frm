VERSION 5.00
Begin VB.Form frmExplain 
   BackColor       =   &H80000018&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Explain"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4455
   ForeColor       =   &H80000017&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picBot 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   1
      Top             =   1995
      Width           =   4455
   End
   Begin VB.TextBox txtExp 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000017&
      Height          =   1995
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmExplain"
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


Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
'general declarations

Public Sub DoDlg(wExplain$, hParent As form)
 UpdateCaption wExplain$, hParent 'See UpdateCaption
  SetParent Me.hwnd, GetDesktopWindow
  'Set the specified windows parent, similar to an objects container property
   Me.Move 0, 0
   'Move this window to the coordinates specified in the left, and top arguments
    SetParent Me.hwnd, hParent.hwnd
    'Set this windows new parent(calling window)
     ShowWindow Me.hwnd, 1
     'Ensure that this window is now visible
      SetFocusA Me.hwnd
      'Set focus to this window
End Sub

Private Sub UpdateCaption(wExplain$, hParent As form)
Dim pcRct As RECT: GetClientRect hParent.hwnd, pcRct
'Get the client area of this windows parent window
Me.Width = (pcRct.Right - pcRct.Left) * Screen.TwipsPerPixelX: Me.Height = (pcRct.Bottom - pcRct.Top) * Screen.TwipsPerPixelY
'Set this windows width and height properties to match the parents client width and height
 Select Case LCase(wExplain)
  Case "gpbw":
  'get process by window
   txtExp.Text = "Get Process by Window allows you to retreive the Process ID by moving you're cursor over a window or child window. This is a useful feature as it saves you time from manually searching the window threads of each processing module to find the module which owns a specific window, or to determine the Window class of a specific window or child window."
   'Update the explanation
  Case "sec":
  'security
   txtExp.Text = "The security features of ProcessXP ensure that restricted users of this system don't terminate processes that need to run for system security reasons or otherwise. This feature also ensures that restricted users do not manipulate a window object of a thread either to throw an exception in the parent module or to bypass features of the module that they would otherwise not be able to."
   'Update the explanation
  Case "shell":
  'New task
   txtExp.Text = "This feature allows you either to execute a valid PE file, or to execute a files associated program to view its information. The WinExec, and ShellExec application program interface functions are used in-order to perform these actions."
   'Update the explanation
 End Select
End Sub

Private Sub Form_Resize()
Dim tcRct As RECT: GetClientRect Me.hwnd, tcRct
'Get client area of this window
 txtExp.Width = (tcRct.Right - tcRct.Left) * Screen.TwipsPerPixelX
  txtExp.Height = ((tcRct.Bottom - tcRct.Top) * Screen.TwipsPerPixelY) - picBot.Height
  'Set txtExp's width/height to this windows client area width/height
   If GUIMissing = False Then HGrad picBot.hdc, picBot.ScaleHeight * 2, picBot.ScaleWidth, COLOR_INFOBK, Black2White, UNEQUALITYFADE Else: picBot.Visible = False: txtExp.Height = txtExp.Height + picBot.Height
   'If the GUI dll is supported, then draw a horizontal gradient who's COLOR's intesity will fade
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Me.Visible = True Then ShowWindow Me.hwnd, 0: SetParent Me.hwnd, GetDesktopWindow
 'If this window is visible then cancel the unloading of this form,
 'set this form's visibility property to false(hidden), set this windows parent to its initial parent to prepare for the next time this explanation window is shown
End Sub


