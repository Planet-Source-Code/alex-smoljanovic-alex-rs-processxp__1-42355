VERSION 5.00
Begin VB.Form frmPointWindow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Get Process By Window"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   Icon            =   "frmPointWindow.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmPointWindow.frx":000C
   ScaleHeight     =   2700
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   3675
      TabIndex        =   12
      Top             =   180
      Width           =   3675
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPointWindow.frx":08D6
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         TabIndex        =   13
         Top             =   0
         Width           =   3615
      End
   End
   Begin VB.PictureBox Picture2 
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
      TabIndex        =   11
      Top             =   0
      Width           =   4455
   End
   Begin VB.CheckBox chkHideWindow 
      BackColor       =   &H80000005&
      Caption         =   "Hide Main Window while selecting"
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
      Left            =   720
      TabIndex        =   10
      Top             =   780
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Select"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   660
      Picture         =   "frmPointWindow.frx":0960
      ScaleHeight     =   45
      ScaleWidth      =   3780
      TabIndex        =   3
      Top             =   1140
      Width           =   3780
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Caption         =   "Window/Process Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   4335
      Begin VB.TextBox txtProcID 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1140
         Locked          =   -1  'True
         MouseIcon       =   "frmPointWindow.frx":1280
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox txtWindowClass 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtWindowText 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   300
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process ID:"
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
         TabIndex        =   8
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Window Class:"
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
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Window Text:"
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
         Top             =   300
         Width           =   945
      End
   End
   Begin VB.PictureBox picTarget 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   500
      Left            =   60
      Picture         =   "frmPointWindow.frx":158A
      ScaleHeight     =   495
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
   Begin VB.PictureBox picBot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   14
      Top             =   2520
      Width           =   3075
      Begin VB.Label Label5 
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
         Left            =   2460
         MouseIcon       =   "frmPointWindow.frx":1E54
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   0
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmPointWindow"
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


Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Dim RethWnd As Boolean, HideWindow As Boolean
Dim CurPos As POINTAPI, wPID&, WndRect As RECT, OldWndRect As RECT, OldwHwnd&, wHwnd&, wndText$, wndClass$, rWC&, i&
'General Declerations

Private Sub chkHideWindow_Click()
 HideWindow = chkHideWindow.Value
 'Update global variable
End Sub

Private Sub Command1_Click()
If GetParentProcess(txtProcID.Text) = True Then
 'If the parent process was found and selected then unload this window
 Unload Me
Else
 If MsgBox("Couldn't find the process." & vbCrLf & vbCrLf & "Would you like to refresh the process list?", vbQuestion + vbYesNo, "Refresh Process List") = vbYes Then
  enumProcesses 'Enumerate Processes, see enumProcess function for more info...
   frmMain.RefreshLibrary 'Call RefreshLibrary, see RefreshLibrary for more information...
    If GetParentProcess(txtProcID.Text) = True Then
    'If parent process was found and selected then unload this window
     Unload Me
    Else
     MsgBox "Still can't find the process.", vbExclamation, "Can't find process"
    End If
 End If
End If
End Sub


Private Sub Form_Activate()
 TransPrep Me.hwnd 'See function TransPrep for more info...
End Sub

Private Sub Form_Load()
 HGrad Picture2.hdc, Picture2.ScaleHeight, Picture2.ScaleWidth, COLOR_ACTIVECAPTION, Black2White, UNEQUALITYFADE
  HRGrad picBot.hdc, picBot.ScaleHeight, picBot.ScaleWidth, COLOR_ACTIVECAPTION
  'Call ProcXP GUI DLL functions Horizontal gradient with UNEQUALITYFADE flag passed in the last argument of the first function call
  'See functions HRGrad(mdlGUI), and HGgrad(mdlGUI), and ProcXP GUI DLL C++ Project for more info...
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If RethWnd = False Then Exit Sub 'If RethWnd(Return Window Handle flag) evaluates to false(not selecting window) then exit sub routine
 wndText = String(250, 0): wndClass = String(250, 0) 'Inititate variables
 GetCursorPos CurPos 'call function GetCurosPos with a pointer to CurPos variable
  wHwnd& = WindowFromPoint(CurPos.X, CurPos.Y) 'Function returns the topmost window at these coordinates
  If X > 0 And X <= Me.Width And Y > 0 And Y <= Me.Height Or wHwnd = frmMain.hwnd Or wHwnd = Me.hwnd Then Exit Sub
  'Determines if the window being selected is this window, or our Main window, if so then exit this sub routine
   If wHwnd = 0 Or OldwHwnd& = wHwnd Then Exit Sub 'If the window handle evaluates to the last selected window handle then exit sub routine
    If OldwHwnd& <> 0 Then InvertRect GetDC(OldwHwnd&), ConvertRect(OldWndRect)
    'If old window handle(previously selected window) is not zero then call InvertRect to invert the pixels of the specified window handles device context handle to restore the hDc
    'See ConvertRect(Which converts the rect as screen coordinates to just the windows client area dimensions) for more information
     GetWindowRect wHwnd&, WndRect 'Return the Window Rect of the window specified by its window handle
      InvertRect GetDC(wHwnd&), ConvertRect(WndRect) 'Invert the newly selected windows rectangle coordinates
       GetWindowText wHwnd, wndText, GetWindowTextLength(wHwnd) + 1 'initialize wndText with the specified window's text
        wndText = Left$(wndText, GetWindowTextLength(wHwnd) + 1) 'remove the chr(0) characters from the buffer
         rWC& = GetClassName(wHwnd, wndClass, 250) 'Retrieve window class of the specified window
          wndClass = Left$(wndClass, rWC) 'remove the chr(0) characters from the buffer
           txtWindowText.Text = wndText '..
            txtWindowClass.Text = wndClass '..
             GetWindowThreadProcessId wHwnd, wPID& ' Retreive the window threads(specified by its handle) process ID
              txtProcID.Text = CStr(wPID) '..
               wndText = "": wndClass = "" '..
                OldwHwnd& = wHwnd& 'Update OldwHwnd with the newly selected window
                 GetWindowRect OldwHwnd, OldWndRect
End Sub

Private Function ConvertRect(oRect As RECT) As RECT
'This functions converts the rectangle from screen coordinate to just the rectangle dimensions
 ConvertRect.Right = oRect.Right - oRect.Left 'Retrieve the width of the of the rectangle as it's left property will be set to 0
  ConvertRect.Bottom = oRect.Bottom - oRect.Top
   ConvertRect.Left = 0 'Set the left prop. to zero
    ConvertRect.Top = 0
    'Example results of this function
    'Step 1
    '-------------|
    ' ____        |
    ' |__|        |  <- Screen
    ' Rectangle   |
    '             |
    '_____________|
    
    'Step 2
    '-------------|
    '|   |        |
    '|___|        |  <- Screen
    ' Rectangle   |
    '             |
    '_____________|
    
    'Step 3
    '-------------|
    '|__|         |
    'Rectangle    |  <- Screen
    '             |
    '             |
    '_____________|
    
End Function

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'Restore this windows initial mouse pointer
  Me.MousePointer = 0
   RethWnd = False 'Set RethWnd flag to false(not returning hWnd)
    ReleaseCapture 'Release mouse capture
     If HideWindow = True Then ShowWindow frmMain.hwnd, 1 'If main window was hidden, then show it
      If OldwHwnd <> 0 Then InvertRect GetDC(OldwHwnd&), ConvertRect(OldWndRect)
      'Invert the last selected window, this should revert the initial invert of its hDc(Device Context Handle)
       OldwHwnd = 0 'Set variable OldwHwnd to zero so when the next time the user selects a window the last selected window won't be inverted again
        picTarget.Visible = True 'Set picturebox picTargets visiblity property to true
End Sub

Private Sub Form_Unload(Cancel As Integer)
 ShowWindow frmMain.hwnd, 1 'Ensure the main window is visible
  If Me.Visible = True Then Cancel = 1: Me.Hide
  'Cancel the unload if this window is visible, call Hide method to hide this window and active its parent
End Sub

Private Sub Label5_Click()
 If sndSupported = True Then sndClass.doPBW
 'If sndSupported flag evaluates to true, call sndClass's doPBW(Process By Window) method to play the explanation sound
  frmExplain.DoDlg "gpbw", Me
  'Call frmExplain's DoDlg function to setup up the information, see this procedure for more information
End Sub

Private Sub picTarget_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.MousePointer = 99 'Set this windows mousepointer property to custom, custom specifies to use the form's MouseIcon property
   SetCapture Me.hwnd 'Set capture to this form, despite cursor position, all mouse events(messages) will be sent to this window
    RethWnd = True 'Set RethWnd flag to true (Return Window Handle)
     If HideWindow = True Then ShowWindow frmMain.hwnd, 0
     'If HideWindow evaluates to true(set by the Hide Main Window check box) then hide the main window, don't use the form's Hide method as an error will occur since that window(frmMain) is displaying a modal dialog
      'RethWnd = True
       picTarget.Visible = False 'set picTarget's visibility property to false
End Sub

Private Sub txtProcID_Click()
If GetParentProcess(txtProcID.Text) = True Then
'if the parent process was found and selected unload this window
 Unload Me
Else
 If MsgBox("Couldn't find the process." & vbCrLf & vbCrLf & "Would you like to refresh the process list?", vbQuestion + vbYesNo, "Refresh Process List") = vbYes Then
  enumProcesses 'Enumerate Processes, see enumProcess for more info..
   frmMain.RefreshLibrary 'Refresh library(repopulate treeview control), see RefreshLibrary function for more info...
    If GetParentProcess(txtProcID.Text) = True Then
     Unload Me
     'If a node representing the specified process was found and selected than unload this form
    Else
     MsgBox "Still can't find the process.", vbExclamation, "Can't find process"
    End If
 End If
End If
End Sub

Private Function GetParentProcess(ByVal ProcID$) As Boolean
On Error Resume Next
'On the event of an error resume execution of this procedure on the next line
Dim i&, tmpPID$, itmX As Node: ProcID$ = Trim$(ProcID$)
'Dimensionalize variable i as long, tmpPID as string data type, itmX as node structure
Dim tmpBuf$: tmpBuf = ProcID$ 'dimensionalize tmpBuf as string type, initialize it
 For i = 2 To frmMain.tvList.Nodes.Count
 'for next loop; i starts at 2, loops until i equals to the amount of nodes in the treeview control incrementing i by one each iteration
 DoEvents 'yield execution
  tmpPID$ = Mid$(frmMain.tvList.Nodes(i).Key, InStr(1, frmMain.tvList.Nodes(i).Key, "PID ", 1) + 4, InStr(1, frmMain.tvList.Nodes(i).Key, "|/", 1) - 2)
   tmpPID$ = Left$(tmpPID$, InStr(1, tmpPID$, "|", 1) - 1)
   'Parse the PID(ProcessID) part of the key
    If Trim(tmpPID$) = Trim(ProcID$) Then
    'If the parsed PID equals to the ProcessID specified in the ProcID argument then...
     frmMain.tvList.Nodes(i).Selected = True 'select the node
      Set itmX = frmMain.tvList.Nodes(i) 'initialize itmX with the return of the Node in the collection nodes specified by index i
       Call frmMain.tvList_NodeClick(itmX) 'See this sub routine for more info...
        GetParentProcess = True 'return true
         Exit Function 'discontinue execution since the parent node was found
    End If
Conti: 'Conti label
 Next i 'next Node index(increment i);loop
End Function
