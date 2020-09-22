VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSearchThread 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Threads - [Program.exe]"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   Icon            =   "frmSearchThread.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picExtended 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   5955
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   5950
      Begin VB.CommandButton Command1 
         Caption         =   "&Hide"
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
         Left            =   5100
         TabIndex        =   46
         Top             =   1860
         Width           =   795
      End
      Begin VB.ListBox lstWindowStyles 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   60
         TabIndex        =   31
         Top             =   180
         Width           =   2775
      End
      Begin VB.ListBox lstExtWindowStyles 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   60
         TabIndex        =   30
         Top             =   1260
         Width           =   2775
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Window Procedure:"
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
         Left            =   2940
         TabIndex        =   45
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label lblWinProc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         Left            =   4440
         TabIndex        =   44
         Top             =   180
         Width           =   450
      End
      Begin VB.Label lblHwnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         Left            =   4440
         TabIndex        =   43
         Top             =   360
         Width           =   450
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Window Handle:"
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
         Left            =   2940
         TabIndex        =   42
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label lblhInst 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         Left            =   4440
         TabIndex        =   41
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instance Handle:"
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
         Left            =   2940
         TabIndex        =   40
         Top             =   540
         Width           =   1170
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Data:"
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
         Left            =   2940
         TabIndex        =   39
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblUserData 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         Left            =   4440
         TabIndex        =   38
         Top             =   720
         Width           =   450
      End
      Begin VB.Label lblWinRct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         Left            =   4440
         TabIndex        =   37
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Window Rectangle:"
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
         Left            =   2940
         TabIndex        =   36
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Rectangle:"
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
         Left            =   2940
         TabIndex        =   35
         Top             =   1440
         Width           =   1170
      End
      Begin VB.Label lblClientRect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label4"
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
         Left            =   4440
         TabIndex        =   34
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Window Styles: []"
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
         TabIndex        =   33
         Top             =   0
         Width           =   1920
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extended Window Styles: []"
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
         TabIndex        =   32
         Top             =   1080
         Width           =   1920
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Window Properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   60
      TabIndex        =   2
      Top             =   2280
      Width           =   5835
      Begin VB.PictureBox picUpdate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2820
         MouseIcon       =   "frmSearchThread.frx":000C
         MousePointer    =   99  'Custom
         Picture         =   "frmSearchThread.frx":0316
         ScaleHeight     =   270
         ScaleWidth      =   240
         TabIndex        =   25
         Top             =   1260
         Width           =   240
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
         Height          =   225
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1500
         Width           =   1755
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
         Height          =   225
         Left            =   1140
         TabIndex        =   20
         Top             =   1320
         Width           =   1635
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2010
         Left            =   3120
         Picture         =   "frmSearchThread.frx":06BA
         ScaleHeight     =   134
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   2
         TabIndex        =   18
         Top             =   0
         Width           =   30
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   -300
         Picture         =   "frmSearchThread.frx":0B2C
         ScaleHeight     =   45
         ScaleWidth      =   3780
         TabIndex        =   13
         Top             =   1200
         Width           =   3780
      End
      Begin VB.PictureBox picRefresh 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   5340
         MouseIcon       =   "frmSearchThread.frx":144C
         MousePointer    =   99  'Custom
         Picture         =   "frmSearchThread.frx":1756
         ScaleHeight     =   540
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   120
         Width           =   480
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   -480
         Picture         =   "frmSearchThread.frx":251A
         ScaleHeight     =   2
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   216
         TabIndex        =   10
         Top             =   720
         Width           =   3240
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Extended Window Info..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         MouseIcon       =   "frmSearchThread.frx":2E3A
         MousePointer    =   99  'Custom
         TabIndex        =   28
         Top             =   420
         Width           =   2055
      End
      Begin VB.Label lblDestroyWindow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Close Window (WM_CLOSE)"
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
         Left            =   3240
         MouseIcon       =   "frmSearchThread.frx":3144
         MousePointer    =   99  'Custom
         TabIndex        =   26
         Top             =   1500
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Window Thread ID:"
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
         TabIndex        =   24
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblThreadID 
         Alignment       =   2  'Center
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
         Height          =   195
         Left            =   1680
         TabIndex        =   23
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label9 
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
         Height          =   195
         Left            =   60
         TabIndex        =   21
         Top             =   1500
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         Height          =   195
         Left            =   60
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblWndState 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4320
         MouseIcon       =   "frmSearchThread.frx":344E
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   1260
         Width           =   1155
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Window State:"
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
         Left            =   3240
         TabIndex        =   16
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label lblVisible 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         MouseIcon       =   "frmSearchThread.frx":3758
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Visible:"
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
         Left            =   3240
         TabIndex        =   14
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblEnabled 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3900
         MouseIcon       =   "frmSearchThread.frx":3A62
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Enabled:"
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
         Left            =   3240
         TabIndex        =   11
         Top             =   900
         Width           =   675
      End
      Begin VB.Label lblParent 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1500
         MouseIcon       =   "frmSearchThread.frx":3D6C
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent Window:"
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
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblNumChildren 
         Alignment       =   2  'Center
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
         Height          =   195
         Left            =   1500
         TabIndex        =   7
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Children:"
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
         TabIndex        =   6
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label lblThreadOwner 
         Alignment       =   2  'Center
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
         Height          =   195
         Left            =   1680
         TabIndex        =   5
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Window Thread Owner:"
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
         TabIndex        =   4
         Top             =   300
         Width           =   1755
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   2160
      Picture         =   "frmSearchThread.frx":4076
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   252
      TabIndex        =   1
      Top             =   2220
      Width           =   3780
   End
   Begin ComctlLib.ListView lstWindows 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      SmallIcons      =   "imgIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Window Class"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Window Text"
         Object.Width           =   3074
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Process"
         Object.Width           =   1535
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Thread ID"
         Object.Width           =   882
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   27
      Top             =   3975
      Width           =   5955
   End
   Begin ComctlLib.ImageList imgIcons 
      Left            =   60
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "frmSearchThread"
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

Private arStyles(0 To 81) As tStyle
Private arExStyles(0 To 24) As tStyle

Private Type tStyle
 Desc As String
  Value As Long
End Type

Private Const GWL_HINSTANCE = (-6)
Private Const GWL_HWNDPARENT = (-8)
Private Const GWL_ID = (-12)
Private Const GWL_USERDATA = (-21)
Private Const GWL_WNDPROC = (-4)

Dim ProcID&, ProcName$, CurHwnd&, initWindText$
'general declarations

Public Function DoDlg(iProcID&, iProcName$)
'This function is called prior to it being shown as a modal dialog
Me.Caption = "Search Threads - [" & iProcName & "]"
'User Friendly feature, prevents user confusion
 ProcID = iProcID: ProcName$ = iProcName$
  TargetPID = ProcID&
   'Initialize variables
   EnumWindows AddressOf EnumWindowsProc, ByVal 0&
   'This API function will enumerate all windows, it will then call the window procedure specified in the lpEnumFunc(Long Pointer to a window procedure) argument
   'See function EnumWindowsProc(mdlWin) for a more detailed explanation
End Function

Public Sub RetEnum(hWnd&, Childh As Boolean)
'This function is called by EnumWindowsProc function which is the callback procedure for the EnumWindows API function. This function is called when the window handle returned by the enumeration function's ProcessID is the target ProcessID
'See EnumWindowsProc(mdlWin) function for more info...
Dim ThreadID&, ProcessID& 'Dimensionalize variable ThreadID as long data type(Thread Handle)
ThreadID = GetWindowThreadProcessId(hWnd&, ProcessID)
'Initialize ThreadID with the Thread ID returned by function GetWindowThreadProcessId
'This function returns the Thread ID of the specified window(Handle) additionally it sets the variable's address passed in the lpdwProcessID(Long Pointer to DWORD variable) argument to the process which owns the thread
'For more information on threads, in the MSDN help system or at MSDN online
Dim wndText$, itmX As ListItem, wndClass$, rCN&: wndText = Space$(GetWindowTextLength(hWnd) + 1): wndClass = String(255, 0)
'dimensionalize wndText as string type, itmX as ListItem type, wndClass as string type, rCN as long type(Used as the return when calling GetClassName, this evaluates to the length of the string associated with the class of a window)
'Initialize wndText, the function GetWindowText uses a pointer to a variables memory, the String(Length, Character) function will allocate the memory needed by the function to return the full window text of the specified window
 GetWindowText hWnd, wndText, GetWindowTextLength(hWnd) + 1
 'Call GetWindowText function to retrieve the Window text of a window object,the last argument specifies the length of the buffer to return
  rCN = GetClassName(hWnd, wndClass, 255)
  'Get the class name of the specified window, rCN is set to the buffer length returned by this function
   wndClass = Left$(wndClass, rCN)
   'The length of the variable wndClass is 255 character, return only the length of the window class name
    If Childh = True Then
    'The window being returned is a child window, the reason why we are keeping track of this is so when this item being added is clicked we are able to determine its a child, because of this we can add functionality to allow users to jump to the parent window of this child window when they click the ParentID label
     Set itmX = lstWindows.ListItems.Add(, CStr(hWnd) & "|CHLD*|" & CStr(tmpParent&), wndClass)
    Else
    'The window being returned is a parent window
     Set itmX = lstWindows.ListItems.Add(, CStr(hWnd) & "|PRNT*|" & CStr(numChild), wndClass)
    End If
      itmX.SubItems(1) = wndText
       itmX.SubItems(2) = ProcName
        itmX.SubItems(3) = glbHex(ThreadID)
        'Update the rest of the Listview controls columns...
         If lstWindows.ListItems.Count = 1 Then lstWindows_ItemClick itmX: itmX.Selected = True: itmX.EnsureVisible
         'If the list contains atleast one item, then select it to update the caption properties of the text boxes and labels which provide information to the user about the windows enumerated
End Sub

Private Sub Command1_Click()
 picExtended.Visible = False
 'set the object picExtended's visible property to false
End Sub

Private Sub Form_Activate()
 TransPrep Me.hWnd 'See TransPrep function for more info...
End Sub

Private Sub Form_Load()
 UpdateWinPos Me.hWnd 'See UpdateWinPos function for more info...
  HRGrad Picture2.hdc, Picture2.ScaleHeight, Picture2.ScaleWidth, COLOR_ACTIVECAPTION
  'This function is an intermediate function to the actual external function of the ProcessXP GUI dll
  'See HRGrad function(mdlGUI) for more info...
'This array's elements will store the values of general window style constants
'see procedure getWindowStyle for more info on how this arrays elements are used
arStyles(0).Desc = "WS_ACTIVECAPTION"
arStyles(0).Value = &H1
 arStyles(1).Desc = "WS_BORDER"
 arStyles(1).Value = &H800000
  arStyles(2).Desc = "WS_CAPTION"
  arStyles(2).Value = &HC00000
   arStyles(3).Desc = "WS_CHILD"
   arStyles(3).Value = &H40000000
    arStyles(4).Desc = "WS_CHILDWINDOW"
    arStyles(4).Value = (&H40000000)
     arStyles(5).Desc = "WS_CLIPCHILDREN "
     arStyles(5).Value = &H2000000
      arStyles(6).Desc = "WS_CLIPSIBLINGS"
      arStyles(6).Value = &H4000000
       arStyles(7).Desc = "WS_DISABLED "
       arStyles(7).Value = &H8000000
        arStyles(8).Desc = "WS_DLGFRAME"
        arStyles(8).Value = &H400000
         arStyles(9).Desc = "WS_GROUP"
         arStyles(9).Value = &H20000
          arStyles(10).Desc = "WS_GT"
          arStyles(10).Value = &H20000 Or &H10000
           arStyles(11).Desc = "WS_HSCROLL "
           arStyles(11).Value = &H100000
            arStyles(12).Desc = "WS_ICONIC"
            arStyles(12).Value = &H1000000
             arStyles(13).Desc = "WS_MAXIMIZE"
             arStyles(13).Value = &H1000000
              arStyles(14).Desc = "WS_MAXIMIZEBOX"
              arStyles(14).Value = &H10000
               arStyles(15).Desc = "WS_MINIMIZE"
               arStyles(15).Value = &H20000000
                arStyles(16).Desc = "WS_MINIMIZEBOX"
                arStyles(16).Value = &H20000
                 arStyles(17).Desc = "WS_OVERLAPPED"
                 arStyles(17).Value = &H0&
                  arStyles(18).Desc = "WS_OVERLAPPEDWINDOW"
                  arStyles(18).Value = (&H0& Or &HC00000 Or &H80000 Or &H40000 Or &H20000 Or &H10000)
                   arStyles(19).Desc = "WS_POPUP"
                   arStyles(19).Value = &H80000000
                    arStyles(20).Desc = "WS_POPUPWINDOW"
                    arStyles(20).Value = (&H80000000 Or &H800000 Or &H80000)
                     arStyles(21).Desc = "WS_SIZEBOX"
                     arStyles(21).Value = &H40000
                      arStyles(22).Desc = "WS_SYSMENU"
                      arStyles(22).Value = &H80000
                       arStyles(23).Desc = "WS_TABSTOP"
                       arStyles(23).Value = &H10000
arStyles(24).Desc = "WS_THICKFRAME"
arStyles(24).Value = &H40000
 arStyles(25).Desc = "WS_TILED"
 arStyles(25).Value = &H0&
  arStyles(26).Desc = "WS_TILEDWINDOW"
  arStyles(26).Value = (&H0& Or &HC00000 Or &H80000 Or &H40000 Or &H20000 Or &H10000)
   arStyles(27).Desc = "WS_VISIBLE"
   arStyles(27).Value = &H10000000
    arStyles(28).Desc = "WS_VSCROLL"
    arStyles(28).Value = &H200000
     arStyles(29).Desc = "ES_AUTOHSCROLL"
     arStyles(29).Value = &H80&
      arStyles(30).Desc = "ES_AUTOVSCROLL"
      arStyles(30).Value = &H40&
       arStyles(31).Desc = "ES_CENTER"
       arStyles(31).Value = &H1&
        arStyles(32).Desc = "ES_CONTINUOUS"
        arStyles(32).Value = &H80000000
         arStyles(33).Desc = "ES_DISABLENOSCROLL"
         arStyles(33).Value = &H2000
          arStyles(34).Desc = "ES_DISPLAY_REQUIRED"
          arStyles(34).Value = &H2
           arStyles(35).Desc = "ES_EX_NOCALLOLEINIT"
           arStyles(35).Value = &H1000000
            arStyles(36).Desc = "ES_LEFT"
            arStyles(36).Value = &H0&
             arStyles(37).Desc = "ES_LOWERCASE"
             arStyles(37).Value = &H10&
              arStyles(38).Desc = "ES_MULTILINE"
              arStyles(38).Value = &H4&
               arStyles(39).Desc = "ES_NOHIDESEL"
               arStyles(39).Value = &H100&
                arStyles(40).Desc = "ES_NOIME"
                arStyles(40).Value = &H80000
                 arStyles(41).Desc = "ES_NOOLEDRAGDROP"
                 arStyles(41).Value = &H8
                  arStyles(42).Desc = "ES_NUMBER"
                  arStyles(42).Value = &H2000&
                   arStyles(43).Desc = "ES_OEMCONVERT"
                   arStyles(43).Value = &H400&
                    arStyles(44).Desc = "ES_PASSWORD"
                    arStyles(44).Value = &H20&
                     arStyles(45).Desc = "ES_READONLY"
                     arStyles(45).Value = &H800&
                      arStyles(46).Desc = "ES_RIGHT"
                      arStyles(46).Value = &H2&
                       arStyles(47).Desc = "ES_SAVESEL"
                       arStyles(47).Value = &H8000
                        arStyles(48).Desc = "ES_SELECTIONBAR"
                        arStyles(48).Value = &H1000000
                         arStyles(49).Desc = "ES_SELFIME"
                         arStyles(49).Value = &H40000
                          arStyles(50).Desc = "ES_SUNKEN"
                          arStyles(50).Value = &H4000
                           arStyles(51).Desc = "ES_SYSTEM_REQUIRED"
                           arStyles(51).Value = (&H1)
                            arStyles(52).Desc = "ES_UPPERCASE"
                            arStyles(52).Value = &H8&
arStyles(53).Desc = "ES_USER_PRESENT"
arStyles(53).Value = (&H4)
 arStyles(54).Desc = "ES_VERTICAL"
 arStyles(54).Value = &H400000
  arStyles(55).Desc = "ES_WANTRETURN"
  arStyles(55).Value = &H1000&
   arStyles(56).Desc = "ESB_DISABLE_BOTH"
   arStyles(56).Value = &H3
    arStyles(57).Desc = "ESB_DISABLE_DOWN"
    arStyles(57).Value = &H2
     arStyles(58).Desc = "ESB_DISABLE_LEFT"
     arStyles(58).Value = &H1
      arStyles(59).Desc = "ESB_DISABLE_LTUP"
      arStyles(59).Value = &H1
       arStyles(60).Desc = "ESB_DISABLE_RIGHT"
       arStyles(60).Value = &H2
        arStyles(61).Desc = "ESB_DISABLE_RTDN"
        arStyles(61).Value = &H2
         arStyles(62).Desc = "ESB_DISABLE_UP"
         arStyles(62).Value = &H1
          arStyles(63).Desc = "ESB_ENABLE_BOTH"
          arStyles(63).Value = &H0
           arStyles(64).Desc = "TPM_BOTTOMALIGN"
           arStyles(64).Value = &H20&
            arStyles(65).Desc = "TPM_CENTERALIGN"
            arStyles(65).Value = &H4&
             arStyles(66).Desc = "TPM_HORIZONTAL"
             arStyles(66).Value = &H0&
              arStyles(67).Desc = "TPM_HORNEGANIMATION"
              arStyles(67).Value = &H800&
               arStyles(68).Desc = "TPM_HORPOSANIMATION"
               arStyles(68).Value = &H400&
                arStyles(69).Desc = "TPM_LEFTALIGN"
                arStyles(69).Value = &H0&
                 arStyles(70).Desc = "TPM_LEFTBUTTON"
                 arStyles(70).Value = &H0&
                  arStyles(71).Desc = "TPM_NOANIMATION"
                  arStyles(71).Value = &H4000&
                   arStyles(72).Desc = "TPM_NONOTIFY"
                   arStyles(72).Value = &H80&
                    arStyles(73).Desc = "TPM_RECURSE"
                    arStyles(73).Value = &H1&
                     arStyles(74).Desc = "TPM_RETURNCMD"
                     arStyles(74).Value = &H100&
                      arStyles(75).Desc = "TPM_RIGHTALIGN"
                      arStyles(75).Value = &H8&
                       arStyles(76).Desc = "TPM_RIGHTBUTTON"
                       arStyles(76).Value = &H2&
                        arStyles(77).Desc = "TPM_TOPALIGN"
                        arStyles(77).Value = &H0&
                         arStyles(78).Desc = "TPM_VCENTERALIGN"
                         arStyles(78).Value = &H10&
                          arStyles(79).Desc = "TPM_VERNEGANIMATION"
                          arStyles(79).Value = &H2000&
                           arStyles(80).Desc = "TPM_VERPOSANIMATION"
                           arStyles(80).Value = &H1000&
                            arStyles(81).Desc = "TPM_VERTICAL"
                            arStyles(81).Value = &H40&
arExStyles(0).Desc = "WS_EX_ACCEPTFILES"
arExStyles(0).Value = &H10&
 arExStyles(1).Desc = "WS_EX_APPWINDOW"
 arExStyles(1).Value = &H40000
  arExStyles(2).Desc = "WS_EX_CLIENTEDGE"
  arExStyles(2).Value = &H200&
   arExStyles(3).Desc = "WS_EX_CONTEXTHELP"
   arExStyles(3).Value = &H400&
    arExStyles(4).Desc = "WS_EX_CONTROLPARENT"
    arExStyles(4).Value = &H10000
     arExStyles(5).Desc = "WS_EX_DLGMODALFRAME"
     arExStyles(5).Value = &H1&
      arExStyles(6).Desc = "WS_EX_LAYERED"
      arExStyles(6).Value = &H80000
       arExStyles(7).Desc = "WS_EX_LAYOUTRTL"
       arExStyles(7).Value = &H400000
        arExStyles(8).Desc = "WS_EX_LEFT"
        arExStyles(8).Value = &H0&
         arExStyles(9).Desc = "WS_EX_LEFTSCROLLBAR"
         arExStyles(9).Value = &H4000&
          arExStyles(10).Desc = "WS_EX_LTRREADING"
          arExStyles(10).Value = &H0&
           arExStyles(11).Desc = "WS_EX_MDICHILD"
           arExStyles(11).Value = &H40&
            arExStyles(12).Desc = "WS_EX_NOACTIVATE"
            arExStyles(12).Value = &H8000000
             arExStyles(13).Desc = "WS_EX_NOINHERITLAYOUT"
             arExStyles(13).Value = &H100000
              arExStyles(14).Desc = "WS_EX_NOPARENTNOTIFY"
              arExStyles(14).Value = &H4&
                arExStyles(15).Desc = "WS_EX_OVERLAPPEDWINDOW"
                arExStyles(15).Value = (&H100& Or &H200&)
                 arExStyles(16).Desc = "WS_EX_PALETTEWINDOW"
                 arExStyles(16).Value = (&H100& Or &H80& Or &H8&)
                  arExStyles(17).Desc = "WS_EX_RIGHT"
                  arExStyles(17).Value = &H1000&
                   arExStyles(18).Desc = "WS_EX_RIGHTSCROLLBAR"
                   arExStyles(18).Value = &H0&
                    arExStyles(19).Desc = "WS_EX_RTLREADING"
                    arExStyles(19).Value = &H2000&
                     arExStyles(20).Desc = "WS_EX_STATICEDGE"
                     arExStyles(20).Value = &H20000
                      arExStyles(21).Desc = "WS_EX_TOOLWINDOW"
                      arExStyles(21).Value = &H80&
                       arExStyles(22).Desc = "WS_EX_TOPMOST"
                       arExStyles(22).Value = &H8&
                        arExStyles(23).Desc = "WS_EX_TRANSPARENT"
                        arExStyles(23).Value = &H20&
                         arExStyles(24).Desc = "WS_EX_WINDOWEDGE"
                         arExStyles(24).Value = &H100&
End Sub


Public Sub RetStyles(hWnd&)
Dim WindowRCT As RECT 'dimensionalize WindowRCT as RECT type structure
Dim sWinStyle&, eWinStyle&, winCap$ 'dimensionalize sWinStyle as long type, eWinStyle as long type, winCap as string type
 sWinStyle = GetWindowLong(hWnd, GWL_STYLE)
 'initialize sWinStyle variable with the return of the handle to the windows style information
  eWinStyle& = GetWindowLong(hWnd, GWL_EXSTYLE)
  'initialize eWinStyle& variable with the return of the handle to the windows extended style information
   Label12 = "Standard Window Styles: [" & glbHex(sWinStyle) & "]"
   'update label caption to contain the hexidecimal string of the window styles handle
    Label10 = "Extended Window Styles: [" & glbHex(eWinStyle) & "]"
    'update label caption to contain the hexidecimal string of the window's extended styles handle
      getWindowStyle sWinStyle
      'see getWindowStyle for more info...
       getExWindowStyle eWinStyle 'see getExWindowStyle for more info...
        lblWinProc = glbHex(GetWindowLong(hWnd, GWL_WNDPROC))
        'see glbHex function in mdlMain(This procedure prefixes the hexidecimal string of a number with zero's until the strings length is eight)
        'update labels caption with the return handle to the window's Window Procedure(Window Procedures are the main procedure which receives Window Messages)
         lblHwnd = glbHex(hWnd) 'update labels caption with the hexidecimal string of the handle to the window
          lblhInst = glbHex(GetWindowLong(hWnd, GWL_HINSTANCE))
          'update labels caption with the hexidecimal string of handle to the applications instance
           lblUserData = glbHex(GetWindowLong(hWnd, GWL_USERDATA))
           'update labels caption property to the hexidecimal string of the handle to the windows user data information
            GetWindowRect hWnd, WindowRCT
            'initialize WindowRCT with the rectangular dimensions of the specified window
             lblWinRct = (WindowRCT.Right - WindowRCT.Left) & " x " & (WindowRCT.Bottom - WindowRCT.Top) & " pixels"
             'update labels caption...
              GetClientRect hWnd, WindowRCT
              'retrieve the windows client area dimensions
               lblClientRect = (WindowRCT.Right - WindowRCT.Left) & " x " & (WindowRCT.Bottom - WindowRCT.Top) & " pixels"
               'update labels caption property...
End Sub

Sub getExWindowStyle(WinStyle&)
Dim i% 'dimensionalize i as integer
lstExtWindowStyles.Clear 'remove all current elemens from the listitem control
 For i = 0 To UBound(arExStyles)
 'for next loop; i starts at 0, loops until i evaluates to the largest element index of the array
 DoEvents 'yeild events
  If (WinStyle And arExStyles(i).Value) = arExStyles(i).Value Then
  'determine if the Windows Extended Style specifies the Extended Window Style specified by the element index of i in the arExStyles array
   lstExtWindowStyles.AddItem arExStyles(i).Desc
   'add the elements description value to the list as the style was found in the window style
  End If
 Next i
End Sub

Sub getWindowStyle(WinStyle&)
Dim i% 'dimensionalize i as integer
lstWindowStyles.Clear 'remove all current elemens from the listitem control
 For i = 0 To UBound(arStyles)
 'for next loop; i starts at 0, loops until i evaluates to the largest element index of the array
 DoEvents 'yeild events
  If (WinStyle And arStyles(i).Value) = arStyles(i).Value Then
  'determine if the Windows Style specifies the Window Style represented by the element index of i in the arStyles array
   lstWindowStyles.AddItem arStyles(i).Desc
   'add the elements description value to the list as the style was found in the window style
  End If
 Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Me.Visible = True Then Cancel = 1: lstWindows.ListItems.Clear: picExtended.Visible = False: Me.Hide
 'If this window is visible then cancel and hide the window
 'This window will remain loaded for speed, it will only be unloaded when the object frmMain is unloaded
End Sub

Private Sub Label7_Click()
 picExtended.Visible = True
 'update the object picExtended's visibility property
End Sub

Private Sub lblDestroyWindow_Click()
If AllCanManipWin = False And IsNTAdmin(0&, 0&) = 0 Then MsgBox "You must be an Administrator to manipulate Window Threads.", vbInformation, "Security": Exit Sub
'If All Can Manipulate Window Threads flag is false(Only Admins) and the current user isn't an administrator than inform the user they have no privellage to do so...
 If MsgBox("Are you sure you want to close this window?", vbQuestion + vbYesNo, "Destroy Window") = vbYes Then
 'Confirm the users desire to send the Window Message WM_CLOSE, same window message sent when the close menu item in a windows system menu is clicked
  SendMessage CurHwnd&, WM_CLOSE, 0&, 0&
  'Call SendMessage function to send the WindowMessage constant to the specified window
  'For more information of registering custom window messages, and the way in which such messages are communicated see the MSDN Help System, or MSDN online at http://msdn.microsoft.com
 End If
End Sub

Private Sub lblEnabled_Click()
If AllCanManipWin = False And IsNTAdmin(0&, 0&) = 0 Then MsgBox "You must be an Administrator to manipulate Window Threads.", vbInformation, "Security": Exit Sub
'Again, determine if this user has been given the privellage to perform the action which he or she is attempting to perform, these flags are set by an administrator using the security dialog
 If LCase(lblEnabled.Caption) = "false" Then
  EnableWindow CurHwnd&, 1
  'The window is currently disabled, enable it
   lblEnabled.Caption = "True"
   'Update enabled label caption
 Else
  EnableWindow CurHwnd&, 0
  'The window is currently enabled, disable it
   lblEnabled.Caption = "False"
   'Update enabled label caption
 End If
End Sub

Private Sub lblParent_Click()
Dim tmpBuffer$, i& 'Dimensionalize tmpBuffer as string type, i as long type
 If lblParent.Caption = "None" Then Exit Sub
 'If the window has no parent the discontinue this sub routine
  For i = 1 To lstWindows.ListItems.Count
  'For Next loop; i starts at 1, loops until i equals to the amount of items in the list view control incrementing i by one each iteration...
   If InStr(1, lstWindows.ListItems(i).Key, "|PRNT*|", 1) >= 1 Then
   'if the string "|PRNT*|" exists with with in the listitem's key specified by index i then...
    tmpBuffer = Mid$(lstWindows.ListItems(i).Key, 1, InStr(1, lstWindows.ListItems(i).Key, "|", 1) - 1)
    'initialize tmpBuffer with the substring returned by Mid function, this will return every character starting at position one to the position of the first "|" character
     If glbHex(tmpBuffer) = lblParent.Caption Then
     'If the the current item(specified by its index:i) in the list represents the parent window of the selected window representing listitem then...
      lstWindows.ListItems(i).Selected = True
      'Select the item as it represents the parent window of the selected window(represented by a listitem in the ListView control)
       lstWindows.ListItems(i).EnsureVisible
       'Ensure the specified list item is visible in the control
        lstWindows_ItemClick lstWindows.ListItems(i)
        'See lstWindows_ItemClick sub routine for more info..
     End If
   End If
  Next i
End Sub

Private Sub lblVisible_Click()
If AllCanManipWin = False And IsNTAdmin(0&, 0&) = 0 Then MsgBox "You must be an Administrator to manipulate Window Threads.", vbInformation, "Security": Exit Sub
'Again, determine if this user has been given the privellage to perform the action which he or she is attempting to perform, these flags are set by an administrator using the security dialog
 If LCase$(lblVisible.Caption) = "false" Then
  ShowWindow CurHwnd&, 1
  'The window is hidden, show it (1 = Normal)
   lblVisible.Caption = "True"
   'Update label
 Else
  ShowWindow CurHwnd&, 0
  'The window is visible, hide it(0 = Hide)
   lblVisible.Caption = "False"
   'Update label
 End If
 'Use SW constants in the nCmdShow argument:
    'SCOPE Const SW_FORCEMINIMIZE = 11
    'SCOPE Const SW_HIDE = 0
    'SCOPE Const SW_INVALIDATE = &H2
    'SCOPE Const SW_MAX = 10
    'SCOPE Const SW_MAXIMIZE = 3
    'SCOPE Const SW_MINIMIZE = 6
    'SCOPE Const SW_NORMAL = 1
    'SCOPE Const SW_OTHERUNZOOM = 4
    'SCOPE Const SW_OTHERZOOM = 2
    'SCOPE Const SW_PARENTCLOSING = 1
    'SCOPE Const SW_PARENTOPENING = 3
    'SCOPE Const SW_RESTORE = 9
    'SCOPE Const SW_SCROLLCHILDREN = &H1
    'SCOPE Const SW_SHOW = 5
    'SCOPE Const SW_SHOWDEFAULT = 10
    'SCOPE Const SW_SHOWMAXIMIZED = 3
    'SCOPE Const SW_SHOWMINIMIZED = 2
    'SCOPE Const SW_SHOWMINNOACTIVE = 7
    'SCOPE Const SW_SHOWNA = 8
    'SCOPE Const SW_SHOWNOACTIVATE = 4
    'SCOPE Const SW_SHOWNORMAL = 1
    'SCOPE Const SW_SMOOTHSCROLL = &H10
End Sub

Private Sub lblWndState_Click()
If AllCanManipWin = False And IsNTAdmin(0&, 0&) = 0 Then MsgBox "You must be an Administrator to manipulate Window Threads.", vbInformation, "Security": Exit Sub
'Again, determine if this user has been given the privellage to perform the action which he or she is attempting to perform, these flags are set by an administrator using the security dialog
 If LCase$(lblWndState.Caption) = "minimized" Then
  ShowWindow CurHwnd, SHOW_OPENWINDOW
  'Window is minimized(ICONIC), show it as normal
   lblWndState.Caption = "Normal"
 ElseIf LCase$(lblWndState.Caption) = "normal" Then
  ShowWindow CurHwnd, SHOW_FULLSCREEN
  'Window is normal(Restored), show it as maximized
   lblWndState.Caption = "Maximized"
 ElseIf LCase$(lblWndState.Caption) = "maximized" Then
  ShowWindow CurHwnd, SHOW_ICONWINDOW
  'Window is maximized(full size), show it as Minimized(Iconic)
   lblWndState.Caption = "Minimized"
 End If
  'Use SW constants in the nCmdShow argument:
    'SCOPE Const SW_FORCEMINIMIZE = 11
    'SCOPE Const SW_HIDE = 0
    'SCOPE Const SW_INVALIDATE = &H2
    'SCOPE Const SW_MAX = 10
    'SCOPE Const SW_MAXIMIZE = 3
    'SCOPE Const SW_MINIMIZE = 6
    'SCOPE Const SW_NORMAL = 1
    'SCOPE Const SW_OTHERUNZOOM = 4
    'SCOPE Const SW_OTHERZOOM = 2
    'SCOPE Const SW_PARENTCLOSING = 1
    'SCOPE Const SW_PARENTOPENING = 3
    'SCOPE Const SW_RESTORE = 9
    'SCOPE Const SW_SCROLLCHILDREN = &H1
    'SCOPE Const SW_SHOW = 5
    'SCOPE Const SW_SHOWDEFAULT = 10
    'SCOPE Const SW_SHOWMAXIMIZED = 3
    'SCOPE Const SW_SHOWMINIMIZED = 2
    'SCOPE Const SW_SHOWMINNOACTIVE = 7
    'SCOPE Const SW_SHOWNA = 8
    'SCOPE Const SW_SHOWNOACTIVATE = 4
    'SCOPE Const SW_SHOWNORMAL = 1
    'SCOPE Const SW_SMOOTHSCROLL = &H10
End Sub

Private Sub lstWindows_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim tmpKey$, tmpBuffer$, ThreadID&, ProcessID&: tmpKey = Item.Key
Dim iHwnd&, hParent&, NumOfChildred$
'Dimensionalize variables...
 If InStr(1, tmpKey$, "CHLD*", 1) > 1 Then
 'Determine the position of a substring with in a string...
   iHwnd = Left$(tmpKey$, InStr(1, tmpKey$, "|", 1) - 1): CurHwnd& = CLng(iHwnd)
   'Parse the window handle information from the key of the item clicked, initialize CurHwnd with the long value of the string iHwnd
    If IsWindow(iHwnd) = 0 Then
    'If the handle specified is a window handle, and not a handle to another device or not a handle at all then...
     MsgBox "The windows information can't be retreived as it appears not to be a window.", vbInformation, "Not a Window Handle"
      lstWindows.ListItems.Remove Item.Index
      'This item doesn't represent a valid window object, remove it from the list...
       Exit Sub 'Discontinue execution of this sub routine...
    End If
    Frame1.Caption = "Window [" & glbHex(iHwnd) & "]'s Properties"
    hParent = CLng(Mid$(tmpKey$, InStr(1, tmpKey$, "CHLD*|", 1) + 6))
    'Initialize hParent(Parent Window's Handle) with the Parent hWnd parsed from the items key
     ThreadID = GetWindowThreadProcessId(iHwnd&, ProcessID)
     'Initialize ThreadID with the return of the GetWindowThreadProcessID function which will return the thread id of the specified window object, additionally it returns the ProcessID to which the window belongs...
      lblThreadOwner = glbHex(CStr(ProcessID))
      'Update label's caption property...
       lblThreadID = glbHex(CStr(ThreadID))
       'Update label's caption property...
        lblNumChildren.Caption = 0
        'Update label's caption property...
         lblParent.Caption = glbHex(CStr(hParent))
         'Update label's caption property...
          txtWindowText.Text = Item.SubItems(1): initWindText$ = Item.SubItems(1)
          'Update text boxes text property
           txtWindowClass.Text = Item.Text
           'Update text boxes text property
            If IsWindowEnabled(iHwnd) = 0 Then
             lblEnabled.Caption = "False"
             'Update label's caption property...
            Else
             lblEnabled.Caption = "True"
             'Update label's caption property...
            End If
              If IsWindowVisible(iHwnd) = 0 Then
               lblVisible.Caption = "False"
               'Update label's caption property...
              Else
               lblVisible.Caption = "True"
               'Update label's caption property...
              End If
                If IsIconic(iHwnd) = 0 And IsZoomed(iHwnd) = 0 Then
                'NOTE: Zoomed is equivelant to Maximized...
                 lblWndState.Caption = "Normal"
                 'Update label's caption property...
                ElseIf IsIconic(iHwnd) = 1 Then
                 lblWndState.Caption = "Minimized"
                 'Update label's caption property...
                ElseIf IsZoomed(iHwnd) = 1 Then
                 lblWndState.Caption = "Maximized"
                 'Update label's caption property...
                End If
 
 ElseIf InStr(1, tmpKey$, "PRNT*", 1) Then
 'Determine the position of a substring with in a string
  iHwnd = Left$(tmpKey$, InStr(1, tmpKey$, "|", 1) - 1): CurHwnd& = CLng(iHwnd)
  'Parse the window handle information from the items key property
   If IsWindow(iHwnd) = 0 Then
   'If the handle doesn't belong to a valid window object...
    MsgBox "The windows information can't be retreived as it appears not to be a window.", vbInformation, "Not a Window Handle"
     lstWindows.ListItems.Remove Item.Index
     'Remove the item as it doesn't represent a window object
      Exit Sub 'exit this sub routine
   End If
    hParent = 0 '...
    Frame1.Caption = "Window [" & glbHex(iHwnd) & "]'s Properties"
     ThreadID = GetWindowThreadProcessId(iHwnd&, ProcessID)
     'Initialize ThreadID with the return of the GetWindowThreadProcessID function which will return the thread id of the specified window object, additionally it returns the ProcessID to which the window belongs...
      lblThreadOwner = glbHex(CStr(ProcessID))
      'Update labels caption prop.
       lblThreadID = glbHex(CStr(ThreadID))
       'Update labels caption prop.
        lblNumChildren.Caption = CLng(Mid$(tmpKey$, InStr(1, tmpKey$, "PRNT*|", 1) + 6))
        'Update labels caption prop.
         lblParent.Caption = "None" '...
          txtWindowText.Text = Item.SubItems(1): initWindText$ = Item.SubItems(1)
          '...
           txtWindowClass.Text = Item.Text '...
            If IsWindowEnabled(iHwnd) = 0 Then
            'If the window specified by its handle isn't enabled then...
             lblEnabled.Caption = "False"
             'Update labels caption prop.
            Else
             lblEnabled.Caption = "True"
             'Update labels caption prop.
            End If
              If IsWindowVisible(iHwnd) = 0 Then
              'If the window specified by its handle isn't visible then...
               lblVisible.Caption = "False"
               'Update labels caption prop.
              Else
               lblVisible.Caption = "True"
               'Update labels caption prop.
              End If
 
                If IsIconic(iHwnd) = 0 And IsZoomed(iHwnd) = 0 Then
                'If the window specified by its handle is neither iconic(minimized) or Zoomed(Maximized) then...
                 lblWndState.Caption = "Normal"
                 'Update labels caption prop.
                ElseIf IsIconic(iHwnd) = 1 Then
                 lblWndState.Caption = "Minimized"
                 'Update labels caption prop.
                ElseIf IsZoomed(iHwnd) = 1 Then
                 lblWndState.Caption = "Maximized"
                 'Update labels caption prop.
                End If
 End If
  RetStyles CurHwnd
End Sub

Private Sub picRefresh_Click()
 lstWindows.ListItems.Clear 'Remove all items from the list
  EnumWindows AddressOf EnumWindowsProc, ByVal 0&
   'This API function will enumerate all windows, it will then call the window procedure specified in the lpEnumFunc(Long Pointer to a window procedure) argument
   'See function EnumWindowsProc(mdlWin) for a more detailed explanation
End Sub

Private Sub picUpdate_Click()
 Call txtWindowText_KeyPress(13)
 'See txtWindowText_KeyPress sub routine for more info...
End Sub

Private Sub txtWindowText_KeyPress(KeyAscii As Integer)
On Error Resume Next
'on the event of an error resume execution of this procedure on the next line
 If KeyAscii = 13 Then
 'KeyAscii argument is the ascii value of the key being pressed, 13 evaluates to carriege return
  If AllCanManipWin = False And IsNTAdmin(0&, 0&) = 0 Then MsgBox "You must be an Administrator to manipulate Window Threads.", vbInformation, "Security": Exit Sub
  'Again, determine if this user has been given the privellage to perform the action which he or she is attempting to perform, these flags are set by an administrator using the security dialog
  If MsgBox("Do you want to change this windows text to" & vbCrLf & """" & txtWindowText.Text & """?", vbQuestion + vbYesNo, "Change Window Text") = vbYes Then
   SetWindowText CurHwnd&, txtWindowText
   'Set the window's new window text specified by the textbox txtWindowText's text property
  Else
   txtWindowText.Text = initWindText$
   'User decided not to update the window objects window text, restore the initial window text to the text box's text property
  End If
 End If
End Sub
