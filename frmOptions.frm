VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5850
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraGO 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   180
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CheckBox chkSoundServer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enable Sound Server"
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
         Left            =   240
         TabIndex        =   21
         Top             =   1380
         Width           =   4275
      End
      Begin VB.TextBox txtTrans 
         Appearance      =   0  'Flat
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
         Left            =   3360
         TabIndex        =   19
         Text            =   "100"
         Top             =   780
         Width           =   435
      End
      Begin VB.CheckBox chkKeepTrans 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Keep ProcessXP windows transparent at"
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
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   3135
      End
      Begin VB.CheckBox chkWinTop 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Keep ProcessXP windows at the Top of other windows"
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
         Left            =   240
         TabIndex        =   16
         Top             =   540
         Width           =   4635
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   540
         Picture         =   "frmOptions.frx":000C
         ScaleHeight     =   30
         ScaleWidth      =   3780
         TabIndex        =   15
         Top             =   1200
         Width           =   3780
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   540
         Picture         =   "frmOptions.frx":092C
         ScaleHeight     =   30
         ScaleWidth      =   3780
         TabIndex        =   14
         Top             =   2160
         Width           =   3780
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":124C
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   420
         TabIndex        =   22
         Top             =   1620
         Width           =   4875
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% transparency"
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
         Left            =   3840
         TabIndex        =   20
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Window Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   180
         TabIndex        =   17
         Top             =   240
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   2940
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   4680
      TabIndex        =   10
      Top             =   2940
      Width           =   1035
   End
   Begin VB.Frame fraLO 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   5535
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   540
         Picture         =   "frmOptions.frx":12EF
         ScaleHeight     =   30
         ScaleWidth      =   3780
         TabIndex        =   9
         Top             =   2160
         Width           =   3780
      End
      Begin VB.CheckBox chkModIcon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display Module's File Icon"
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
         Left            =   240
         TabIndex        =   7
         Top             =   1860
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkProcIcon 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display Process's File Icon"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   540
         Picture         =   "frmOptions.frx":1C0F
         ScaleHeight     =   30
         ScaleWidth      =   3780
         TabIndex        =   5
         Top             =   1140
         Width           =   3780
      End
      Begin VB.CheckBox chkModRelation 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display Process's Modules according to parent/child relationship"
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.CheckBox chkProcRelation 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display Processes according to parent/child relationship"
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
         Left            =   240
         TabIndex        =   2
         Top             =   540
         Value           =   1  'Checked
         Width           =   4275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Uses more memory, list population takes more time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   1680
         TabIndex        =   12
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Graphical Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List Hierarchy Options"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1560
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5953
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "List Options"
            Key             =   "listoptions"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "General"
            Key             =   "gen"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
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


Dim tmpProcR As Boolean, tmpModR As Boolean, ProcI As Boolean, ModI As Boolean
Dim tmpWinTop As Boolean, tmpWinTrans&, tmpKeepTrans As Boolean, tmpSndServer As Boolean
'Temporary values who's values will be set during the DoDlg function, and will restore the global variables if the user cancels

'general declarations

Public Function DoDlg()
'This functions main purpose is to copy all the global variables which it edits to temporary variable so that if the user cancels the initial values of the variables are restored
 tmpProcR = ProcRelation
  tmpModR = ModRelation
   ProcI = ShowProcIcon
    ModI = ShowModIcon
     tmpWinTop = KeepWinTop
      tmpWinTrans = WinTrans
       tmpKeepTrans = KeepTrans
       'Copy initial values
        txtTrans.Text = FormatNumber((WinTrans / 255) * 100, 0)
        'Variable WinTrans specifies a byte value, this byte value is then used to set the transparency of the ProcessXP windows
        'FormatNumber will return the percent, so that instead of the user choosing a value from 1 to 255, they choose a transparency value from 1 to 100
         tmpSndServer = playSounds
         'Copy initial value...
End Function

Private Sub chkKeepTrans_Click()
 KeepTrans = chkKeepTrans.Value
 'Update the global variable, a checkbox control's Value is a boolean value(0 or 1)
  If chkKeepTrans.Value = 0 Then txtTrans.Enabled = False Else txtTrans.Enabled = True
End Sub

Private Sub chkModIcon_Click()
 ShowModIcon = chkModIcon.Value
 'Update global boolean variable
End Sub

Private Sub chkModRelation_Click()
 ModRelation = chkModRelation.Value
 'Update global boolean variable
End Sub

Private Sub chkProcIcon_Click()
 ShowProcIcon = chkProcIcon.Value
 'Update global boolean variable
End Sub

Private Sub chkProcRelation_Click()
 ProcRelation = chkProcRelation.Value
 'Update global boolean variable
End Sub

Private Sub chkWinOnTop_Click()
 KeepWinTop = chkWinTop.Value
 'Update global boolean variable
End Sub

Private Sub chkSoundServer_Click()
 playSounds = chkSoundServer.Value
 'Update global boolean variable
End Sub

Private Sub chkWinTop_Click()
 KeepWinTop = chkWinTop.Value
 'Update global boolean variable
End Sub

Private Sub cmdCancel_Click()
 ProcRelation = tmpProcR
  ModRelation = tmpModR
   ShowProcIcon = ProcI
    ShowModIcon = ModI
     KeepWinTop = tmpWinTop
      WinTrans = tmpWinTrans
       KeepTrans = tmpKeepTrans
        chkProcRelation.Value = cBoolInt(ProcRelation) 'See cBoolInt function for more info
         chkModRelation.Value = cBoolInt(ModRelation) 'See cBoolInt function for more info
          chkProcIcon.Value = cBoolInt(ShowProcIcon) 'See cBoolInt function for more info
           chkModIcon.Value = cBoolInt(ShowModIcon) 'See cBoolInt function for more info
            chkWinTop.Value = cBoolInt(KeepWinTop) 'See cBoolInt function for more info
             chkKeepTrans.Value = cBoolInt(KeepTrans) 'See cBoolInt function for more info
              txtTrans.Text = FormatNumber((WinTrans / 255) * 100, 0)
               playSounds = tmpSndServer
               'Restore the global variables since the user cancelled
                Me.Hide
                'Call Hide method to hide this form and activate its parent
End Sub


Private Sub cmdSave_Click()
On Error Resume Next
Dim form As form
'on the event of an error resume execution on the next line of this procedure
 Me.Hide 'Hide this form
  For Each form In Forms 'Enumerate through each form in collection Forms(The collection Forms consist of only Loaded Forms)
   TransPrep form.hwnd 'See TransPrep function for more info
  Next form 'Select Next form in Forms collection
   If Val(txtTrans.Text) < 10 Then txtTrans.Text = 10
   'If user has specified a transparency percent of less than 10, set the percent to 10 as the user will no long be able to see an ProcessXP windows
    If Val(txtTrans.Text) > 100 Then txtTrans.Text = 100 'If greater than 100, set to 100 since the byte value can be no larger than 255 which is what a value of greater than 100 will evaluate to
     If playSounds = True Then
     'If global variable playSounds evaluates to true then...
      Err.Clear 'Clear current error if any
       If sndClass Is Nothing Then Set sndClass = New clsMain
       'If sndClass has not yet been initialized, initialize it
        If Err.Number = 429 Then chkSoundServer.Enabled = False: sndSupported = False Else sndSupported = True: chkSoundServer.Enabled = True
        'If err 429 occurs(component can't be found/initialized), disable the check box which gives the user the ability to load the snd server
     Else
     'If global variable playSounds evaluates to false then...
      If Not (sndClass Is Nothing) Then Set sndClass = Nothing
      'If sndClass does not evaluate to nothing then terminate it...
       sndSupported = False
       'set sndSupported flag to false to prevent procedures which attempt to call sndserver methods to play sounds from calling the methods...
       'This flag is also set if the sndServer dll fails to initialize...
     End If
End Sub

Private Sub Form_Activate()
 TransPrep Me.hwnd 'See TransPrep function for more info...
End Sub

Private Sub Form_Load()
If Starting = True Then frmSpash.UpdateProgress , , "Loading Option Settings"
'If starting evaluates to true which it will when this form is being loaded by frmSpash then call its UpdateProgress method to update the loading progress
 chkProcRelation.Value = cBoolInt(ProcRelation) 'see cBoolInt function for more info..
  chkModRelation.Value = cBoolInt(ModRelation) 'see cBoolInt function for more info..
   chkProcIcon.Value = cBoolInt(ShowProcIcon) 'see cBoolInt function for more info..
    chkModIcon.Value = cBoolInt(ShowModIcon) 'see cBoolInt function for more info..
     chkWinTop.Value = cBoolInt(KeepWinTop) 'see cBoolInt function for more info..
      chkKeepTrans.Value = cBoolInt(KeepTrans) 'see cBoolInt function for more info..
       chkSoundServer.Value = cBoolInt(playSounds) 'see cBoolInt function for more info..
        Add_ES_Number txtTrans.hwnd 'This function adds a window style to this windows existing style that allows only number to be type while this window has focused
        'See Add_ES_Number sub routine for more information
         If HostOS.OperatingSystem = Win2K Then
         'If the Host operating system evaluates to Windows 2000 then...
          fraLO.BackColor = Me.BackColor: fraGO.BackColor = Me.BackColor
           chkProcRelation.BackColor = Me.BackColor
            chkModRelation.BackColor = Me.BackColor
             chkProcIcon.BackColor = Me.BackColor
              chkModIcon.BackColor = Me.BackColor
               chkKeepTrans.BackColor = Me.BackColor
                chkWinTop.BackColor = Me.BackColor
                'Change objects background value from white to system color Button Face(or Window Background)
                 Picture1.Visible = False: Picture3.Visible = False
                  Picture2.Visible = False: Picture4.Visible = False
                  'Hide the horizontal gradient pictures
                   chkSoundServer.BackColor = Me.BackColor '...
        End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Me.Visible = True Then Cancel = 1: Me.Hide: Exit Sub
 'If this window is visible then cancel the unload, and call its Hide method to hide this window and activate its parent
 'This window will only be unloaded when the main window is unloaded
  SaveSetting "ProcessXP", "Options", "ProcRelation", CStr(chkProcRelation.Value)
   SaveSetting "ProcessXP", "Options", "ModRelation", CStr(chkModRelation.Value)
    SaveSetting "ProcessXP", "Options", "ProcIcon", CStr(chkProcIcon.Value)
     SaveSetting "ProcessXP", "Options", "ModIcon", CStr(chkModIcon.Value)
      SaveSetting "ProcessXP", "Options", "KeepTrans", CStr(chkKeepTrans.Value)
       SaveSetting "ProcessXP", "Options", "TransVal", CStr(WinTrans)
        SaveSetting "ProcessXP", "Options", "WinOnTop", CStr(chkWinTop.Value)
         SaveSetting "ProcessXP", "Options", "EnableSoundServer", CStr(chkSoundServer.Value)
         'Save options in the registry
End Sub

Private Function cBoolInt(bValue As Boolean) As Long
'This function converts Boolean values to Long values
 If bValue = True Then cBoolInt = 1 Else cBoolInt = 0
End Function

Private Sub TabStrip1_Click()
 Select Case TabStrip1.SelectedItem.Key
 'Select case statement, see MSDN Help System, or MSDN online @ http://msdn.microsoft.com
  Case "listoptions": 'If the selected expression(TabStrip1.SelectedItem.Key) evaluates to "listoptions"...
   fraLO.Visible = True
    fraGO.Visible = False
  Case "gen":
   fraGO.Visible = True
    fraLO.Visible = False
 End Select
  'Update frames visiblility according to the tab selected
End Sub

Private Sub txtTrans_Change()
On Error GoTo errh
'On the event of an error jump to label errh
Dim TransVal%, tmpNum&: txtTrans.Text = CStr(Val(txtTrans.Text)): tmpNum& = Val(txtTrans.Text)
'Dimensionalize TransVal as integer type, tmpNum as long type
 If tmpNum& < 10 Then tmpNum& = 10
 'if tmpNum evaluates to less than 10 then set tmpNum to 10 as this is the percent of window transparency, if its to low the user will no long be able to see the windows unless they manually change the byte value of the registry key in which this value is stored...
  If tmpNum& > 100 Then tmpNum& = 100 'If greater than 100, set it to 100
   TransVal% = (tmpNum& / 100) * 255 'Determine the percent of 255
    WinTrans = TransVal% 'Update global variable
     Exit Sub 'discontinue execution of this procedure
errh: 'label errh
 WinTrans = 255
 'An unexpected error has occured, set the global variable to 255(100% visible)
End Sub
