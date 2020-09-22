VERSION 5.00
Begin VB.Form frmSpash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   2
      Top             =   660
      Width           =   2715
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
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lblProgress 
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
         Top             =   120
         Width           =   2565
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
         Left            =   60
         TabIndex        =   3
         Top             =   600
         Width           =   2595
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   -1140
      Picture         =   "frmSpash.frx":0000
      ScaleHeight     =   45
      ScaleWidth      =   3780
      TabIndex        =   1
      Top             =   600
      Width           =   3780
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   0
      Picture         =   "frmSpash.frx":0920
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   0
      Width           =   2700
   End
End
Attribute VB_Name = "frmSpash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
 If RetOSInf() = False Then
  If MsgBox("ProcessXP must determine what Operating System it is running on." & vbCrLf & vbCrLf & "While it attempted to retreive the systems operating system information an error occured." & vbCrLf & "It might be unsafe to run ProcessXP allthough it is compatible with Win95, Win98, WinME, Win2K, WinXP." & vbCrLf & vbCrLf & "Do you wish to run ProcessXP anyway?", vbQuestion + vbYesNo, "Initiation Error") = vbNo Then End
 End If
  If HostOS.DispDevColors < 65536 Then
   If MsgBox("Please change you're Display Color Quality setting to atleast support 16-bit (65,536) colors." & vbCrLf & "You're display device is currently only supporting " & CStr(FormatNumber(HostOS.DispDevColors, 0, , , vbTrue)) & " colors." & vbCrLf & vbclrf & "Do you want to change you're Display's Color Quality setting?", vbQuestion + vbYesNo, "ProcessXP Requires 16-bit Colors") = vbYes Then
    ShowDispSet
   End If
    MsgBox "Is it required that you restart ProcessXP after you have changed you're display settings.", vbInformation + vbOKOnly, "ProcessXP"
     End
  End If
   If HostOS.OperatingSystem = inCompatibleOS Then
    Load frmIOS
     Exit Sub
   End If
    If HostOS.OperatingSystem = Win2K Then StretchBlt picLogo.hDC, 0, 0, picLogo.ScaleWidth, picLogo.ScaleHeight, picLogo.hDC, 0, 0, picLogo.ScaleWidth, picLogo.ScaleHeight, SRCCOPY
 Starting = True
  VGrad picFrame, 239, 235, 222, False, , True: Me.Show
   If GetSetting("ProcessXP", "Startup", "CI", "0") = "0" Then CurrentInfo = ProcessInfo Else CurrentInfo = VersionInfo
    If LCase$(GetSetting("ProcessXP", "Startup", "MU", "1")) = "false" Then MemUpdateOn = False Else MemUpdateOn = True
     ProcRelation = Val(GetSetting("ProcessXP", "Options", "ProcRelation", "0"))
      ModRelation = Val(GetSetting("ProcessXP", "Options", "ModRelation", "1"))
       ShowProcIcon = Val(GetSetting("ProcessXP", "Options", "ProcIcon", "1"))
        ShowModIcon = Val(GetSetting("ProcessXP", "Options", "ModIcon", "1"))
         KeepWinTop = Val(GetSetting("ProcessXP", "Options", "WinOnTop", "1"))
          WinTrans = Val(GetSetting("ProcessXP", "Options", "TransVal", "230"))
           KeepTrans = Val(GetSetting("ProcessXP", "Options", "KeepTrans", "1"))
            TransPrep Me.hWnd
             UpdateWinPos Me.hWnd
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
           lblProgress.Caption = "Drawing Components GUI..."
            Load frmSearchThread
            If Err.Number = 339 Then GoTo DepNF
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
                      Me.Hide
                       Starting = False
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

Private Sub Form_Initialize()
Dim rICc&
 rICc = InitCommonControls
End Sub

Public Sub UpdateProgress(Optional cValue& = -1, Optional cMax& = -1, Optional NewTitle$ = "", Optional NewProc$ = "")
On Error Resume Next
Dim tPercent&
 If NewProc$ <> "" Then lblCurProc.Caption = "Processing """ & NewProc & """...": lblCurProc.Visible = True
  If NewProc = "" Then lblCurProc.Visible = False
   If cValue& <> -1 And cMax& <> -1 Then tPercent = (cValue / cMax) * 100
    If NewTitle$ <> "" And cValue& <> -1 And cMax& <> -1 Then
     lblProgress.Caption = NewTitle$ & "... " & CStr(tPercent) & "%"
    Else
     If NewTitle$ <> "" Then lblProgress.Caption = NewTitle$ & "..."
    End If
End Sub

