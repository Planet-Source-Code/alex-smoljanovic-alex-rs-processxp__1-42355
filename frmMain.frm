VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProcessXP"
   ClientHeight    =   4740
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8940
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox tmpProcIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2220
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   111
      Top             =   2220
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox tmpModIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1920
      Picture         =   "frmMain.frx":0E54
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   110
      Top             =   2220
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox tmpProcessIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1620
      Picture         =   "frmMain.frx":13DE
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   109
      Top             =   2220
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picProcList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   900
      ScaleHeight     =   3735
      ScaleWidth      =   4875
      TabIndex        =   106
      Top             =   0
      Visible         =   0   'False
      Width           =   4875
      Begin VB.Label lblProcListTmp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Processing List"
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
         TabIndex        =   107
         Top             =   1680
         Width           =   4755
      End
   End
   Begin VB.PictureBox picMemBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2280
      Picture         =   "frmMain.frx":1968
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   180
      TabIndex        =   74
      Top             =   2700
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Timer tmrMem 
      Interval        =   1500
      Left            =   2580
      Top             =   3120
   End
   Begin VB.PictureBox picOff 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   600
      Left            =   1740
      Picture         =   "frmMain.frx":3518
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   70
      Top             =   3120
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picOn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   600
      Left            =   1140
      Picture         =   "frmMain.frx":448C
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   69
      Top             =   3120
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox picOnOff 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   1080
      MouseIcon       =   "frmMain.frx":5400
      MousePointer    =   99  'Custom
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   68
      ToolTipText     =   "Memory Stats"
      Top             =   3840
      Width           =   540
   End
   Begin VB.PictureBox picPage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3540
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   67
      Top             =   4140
      Width           =   2235
      Begin VB.Label lblPage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   0
         TabIndex        =   73
         Top             =   10
         Width           =   2175
      End
   End
   Begin VB.PictureBox picVirtual 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3540
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   65
      Top             =   3960
      Width           =   2235
      Begin VB.Label lblVirtual 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   0
         TabIndex        =   72
         Top             =   10
         Width           =   2175
      End
   End
   Begin VB.PictureBox picPhysical 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3540
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   149
      TabIndex        =   63
      Top             =   3780
      Width           =   2235
      Begin VB.Label lblPhysical 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   0
         TabIndex        =   71
         Top             =   10
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   2280
      Picture         =   "frmMain.frx":570A
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   61
      Top             =   3720
      Width           =   2250
   End
   Begin VB.PictureBox tmpLargeIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   3465
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   50
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin ComctlLib.TreeView tvList 
      Height          =   3735
      Left            =   900
      TabIndex        =   0
      Top             =   0
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   6588
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   26
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgTVIcons"
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
   End
   Begin VB.PictureBox picIconTmp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3720
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picR 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4425
      Left            =   5775
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   211
      TabIndex        =   3
      Top             =   0
      Width           =   3165
      Begin VB.PictureBox tmpHelp 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   1380
         MouseIcon       =   "frmMain.frx":5E5E
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":6168
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   42
         Top             =   3735
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox tmpWeb 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   2520
         MouseIcon       =   "frmMain.frx":6F2C
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":7236
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   36
         TabIndex        =   20
         Top             =   3735
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox tmpSec 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   180
         MouseIcon       =   "frmMain.frx":81AA
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":84B4
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   17
         Top             =   3735
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Process/Module Information"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   60
         TabIndex        =   4
         Top             =   0
         Width           =   3075
         Begin VB.PictureBox picVersion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3435
            Left            =   60
            ScaleHeight     =   3435
            ScaleWidth      =   2955
            TabIndex        =   94
            Top             =   240
            Visible         =   0   'False
            Width           =   2955
            Begin VB.ComboBox cmbProperty 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   270
               ItemData        =   "frmMain.frx":9278
               Left            =   60
               List            =   "frmMain.frx":92A0
               Style           =   2  'Dropdown List
               TabIndex        =   103
               Top             =   1300
               Width           =   2835
            End
            Begin VB.TextBox txtInfo 
               BackColor       =   &H8000000F&
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   60
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   102
               Top             =   1560
               Width           =   2835
            End
            Begin VB.PictureBox Picture15 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   30
               Left            =   300
               Picture         =   "frmMain.frx":935A
               ScaleHeight     =   30
               ScaleWidth      =   2250
               TabIndex        =   98
               Top             =   3180
               Width           =   2250
            End
            Begin VB.PictureBox Picture14 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   30
               Left            =   360
               Picture         =   "frmMain.frx":9AAE
               ScaleHeight     =   2
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   150
               TabIndex        =   97
               Top             =   660
               Width           =   2250
            End
            Begin VB.TextBox txtVFilePath 
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
               Left            =   660
               Locked          =   -1  'True
               TabIndex        =   96
               ToolTipText     =   "Short File Path (DOS Path)"
               Top             =   240
               Width           =   2235
            End
            Begin VB.PictureBox picVFileIcon 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   95
               Top             =   60
               Width           =   540
            End
            Begin VB.Label lblOpenProp 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Open Advanced Properties Dialog"
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
               Left            =   60
               MouseIcon       =   "frmMain.frx":A202
               MousePointer    =   99  'Custom
               TabIndex        =   108
               Top             =   2820
               Width           =   2325
            End
            Begin VB.Label lblProductVersion 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "3.0.0.1"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Left            =   1260
               TabIndex        =   105
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label lblProductName 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ProcessXP"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   135
               Left            =   1260
               TabIndex        =   104
               Top             =   780
               Width           =   1695
            End
            Begin VB.Label lblProcessInformation 
               BackStyle       =   0  'Transparent
               Caption         =   "<< View Process Information"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00F27E53&
               Height          =   255
               Left            =   60
               MouseIcon       =   "frmMain.frx":A50C
               MousePointer    =   99  'Custom
               TabIndex        =   101
               ToolTipText     =   "View this files embedded version information"
               Top             =   3240
               Width           =   2775
            End
            Begin VB.Label Label26 
               BackStyle       =   0  'Transparent
               Caption         =   "Product Version:"
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
               TabIndex        =   100
               Top             =   960
               Width           =   1695
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Product Name:"
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
               TabIndex        =   99
               Top             =   780
               Width           =   1020
            End
         End
         Begin VB.PictureBox picMod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3435
            Left            =   60
            ScaleHeight     =   3435
            ScaleWidth      =   2955
            TabIndex        =   80
            Top             =   240
            Visible         =   0   'False
            Width           =   2955
            Begin VB.TextBox txtMdlDesc 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   60
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   93
               Top             =   2160
               Width           =   2895
            End
            Begin VB.PictureBox picModIcon 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   85
               Top             =   60
               Width           =   540
            End
            Begin VB.TextBox txtMdlLPath 
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
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   84
               Text            =   "C:\asdas"
               ToolTipText     =   "Complete File Path"
               Top             =   1020
               Width           =   2835
            End
            Begin VB.TextBox txtMdlSPath 
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
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   83
               ToolTipText     =   "Short File Path (DOS Path)"
               Top             =   720
               Width           =   2835
            End
            Begin VB.PictureBox Picture13 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   30
               Left            =   360
               Picture         =   "frmMain.frx":A816
               ScaleHeight     =   2
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   150
               TabIndex        =   82
               Top             =   1320
               Width           =   2250
            End
            Begin VB.PictureBox Picture12 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   30
               Left            =   300
               Picture         =   "frmMain.frx":AF6A
               ScaleHeight     =   30
               ScaleWidth      =   2250
               TabIndex        =   81
               Top             =   3180
               Width           =   2250
            End
            Begin VB.Label lblGloballyUsed 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "PROCESSID"
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
               Left            =   1200
               TabIndex        =   92
               ToolTipText     =   "Unique Identifier of this Process"
               Top             =   1680
               Width           =   1695
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Globally Used:"
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
               TabIndex        =   91
               Top             =   1680
               Width           =   990
            End
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Module's Description:"
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
               TabIndex        =   90
               Top             =   1980
               Width           =   1695
            End
            Begin VB.Label lblModName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "WinNTLoging.exe"
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
               Left            =   720
               TabIndex        =   89
               Top             =   240
               Width           =   2220
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Parents Process:"
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
               TabIndex        =   88
               Top             =   1500
               Width           =   1155
            End
            Begin VB.Label lblmdlProcID 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "PROCESSID"
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
               Left            =   1200
               MouseIcon       =   "frmMain.frx":B6BE
               MousePointer    =   99  'Custom
               TabIndex        =   87
               ToolTipText     =   "Unique Identifier of this Process"
               Top             =   1500
               Width           =   1695
            End
            Begin VB.Label lblMViewVersionHeader 
               BackStyle       =   0  'Transparent
               Caption         =   "View Files Version Header >>"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00F27E53&
               Height          =   255
               Left            =   60
               MouseIcon       =   "frmMain.frx":B9C8
               MousePointer    =   99  'Custom
               TabIndex        =   86
               ToolTipText     =   "View this files embedded version information"
               Top             =   3240
               Width           =   2775
            End
         End
         Begin VB.PictureBox picProc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3435
            Left            =   60
            ScaleHeight     =   3435
            ScaleWidth      =   2955
            TabIndex        =   26
            Top             =   240
            Visible         =   0   'False
            Width           =   2955
            Begin VB.PictureBox picProcMemRef 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   2700
               MouseIcon       =   "frmMain.frx":BCD2
               MousePointer    =   99  'Custom
               Picture         =   "frmMain.frx":BFDC
               ScaleHeight     =   270
               ScaleWidth      =   240
               TabIndex        =   79
               Top             =   2880
               Width           =   240
            End
            Begin VB.PictureBox Picture8 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   30
               Left            =   300
               Picture         =   "frmMain.frx":C380
               ScaleHeight     =   30
               ScaleWidth      =   2250
               TabIndex        =   59
               Top             =   3180
               Width           =   2250
            End
            Begin VB.PictureBox Picture7 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   30
               Left            =   360
               Picture         =   "frmMain.frx":CAD4
               ScaleHeight     =   2
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   150
               TabIndex        =   58
               Top             =   1320
               Width           =   2250
            End
            Begin VB.TextBox lblParentID 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   840
               Locked          =   -1  'True
               MouseIcon       =   "frmMain.frx":D228
               MousePointer    =   99  'Custom
               TabIndex        =   51
               Text            =   "PROCESSID"
               ToolTipText     =   "Unique Identifier of this Process's Parent Process"
               Top             =   1740
               Width           =   2115
            End
            Begin VB.TextBox txtFullPath 
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
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   49
               ToolTipText     =   "Short File Path (DOS Path)"
               Top             =   720
               Width           =   2835
            End
            Begin VB.TextBox txtPathName 
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
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   29
               Text            =   "C:\asdas"
               ToolTipText     =   "Complete File Path"
               Top             =   1020
               Width           =   2835
            End
            Begin VB.PictureBox picProcIcon 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   540
               Left            =   60
               ScaleHeight     =   36
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   36
               TabIndex        =   28
               Top             =   60
               Width           =   540
            End
            Begin VB.Label lblProcMem 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               Left            =   1140
               TabIndex        =   78
               ToolTipText     =   "Memory allocated for this process"
               Top             =   2940
               Width           =   1575
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Process Memory:"
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
               TabIndex        =   77
               Top             =   2940
               Width           =   1170
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Time Started:"
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
               TabIndex        =   76
               Top             =   2700
               Width           =   915
            End
            Begin VB.Label lblProcTime 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               Left            =   960
               TabIndex        =   75
               ToolTipText     =   "The time this process started execution"
               Top             =   2700
               Width           =   1935
            End
            Begin VB.Label numModules 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               Left            =   960
               TabIndex        =   41
               ToolTipText     =   "This process's number of Modules (Referenced DLL's, OCX's, ect.)"
               Top             =   2460
               Width           =   1935
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Modules:"
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
               TabIndex        =   40
               Top             =   2460
               Width           =   630
            End
            Begin VB.Label lblVersionInfo 
               BackStyle       =   0  'Transparent
               Caption         =   "View Files Version Header >>"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00F27E53&
               Height          =   255
               Left            =   60
               MouseIcon       =   "frmMain.frx":D532
               MousePointer    =   99  'Custom
               TabIndex        =   39
               ToolTipText     =   "View this files embedded version information"
               Top             =   3240
               Width           =   2775
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Parents ID:"
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
               TabIndex        =   38
               Top             =   1740
               Width           =   780
            End
            Begin VB.Label lblNumThreads 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "0"
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
               Left            =   960
               TabIndex        =   37
               ToolTipText     =   "This process's number of threads"
               Top             =   2220
               Width           =   1935
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Threads:"
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
               TabIndex        =   36
               Top             =   2220
               Width           =   585
            End
            Begin VB.Label lblBoolMT 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "TRUE/FALSE"
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
               Left            =   960
               TabIndex        =   35
               ToolTipText     =   "Does this process have multiple threads?"
               Top             =   1980
               Width           =   1935
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Multi-threaded:"
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
               TabIndex        =   34
               Top             =   1980
               Width           =   1080
            End
            Begin VB.Label lblProcID 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "PROCESSID"
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
               Left            =   900
               TabIndex        =   33
               ToolTipText     =   "Unique Identifier of this Process"
               Top             =   1500
               Width           =   1995
            End
            Begin VB.Label Label11 
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
               TabIndex        =   32
               Top             =   1500
               Width           =   795
            End
            Begin VB.Label lblProcName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "WinNTLoging.exe"
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
               Left            =   720
               TabIndex        =   27
               Top             =   240
               Width           =   2220
            End
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Please select a process or module"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   120
            TabIndex        =   31
            Top             =   1620
            Width           =   2865
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No information"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   1245
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
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
         Left            =   1500
         TabIndex        =   44
         Top             =   4260
         Width           =   315
      End
      Begin VB.Label btnHelp 
         BackStyle       =   0  'Transparent
         Height          =   675
         Left            =   1320
         MouseIcon       =   "frmMain.frx":D83C
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label btnWeb 
         BackStyle       =   0  'Transparent
         Height          =   675
         Left            =   2460
         MouseIcon       =   "frmMain.frx":DB46
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   3720
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web"
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
         Left            =   2640
         TabIndex        =   22
         Top             =   4260
         Width           =   300
      End
      Begin VB.Label btnSec 
         BackStyle       =   0  'Transparent
         Height          =   675
         Left            =   120
         MouseIcon       =   "frmMain.frx":DE50
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Security"
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
         Left            =   180
         TabIndex        =   18
         Top             =   4260
         Width           =   570
      End
   End
   Begin VB.PictureBox picL 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   0
      ScaleHeight     =   295
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   2
      Top             =   0
      Width           =   915
      Begin VB.PictureBox tmpPicProg 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   60
         Picture         =   "frmMain.frx":E15A
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   53
         TabIndex        =   60
         Top             =   3600
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3825
         Left            =   870
         Picture         =   "frmMain.frx":E9BE
         ScaleHeight     =   3825
         ScaleWidth      =   30
         TabIndex        =   57
         Top             =   300
         Width           =   30
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   -90
         Picture         =   "frmMain.frx":F1FA
         ScaleHeight     =   30
         ScaleWidth      =   1125
         TabIndex        =   56
         Top             =   3240
         Width           =   1125
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   -90
         Picture         =   "frmMain.frx":F5CE
         ScaleHeight     =   30
         ScaleWidth      =   1125
         TabIndex        =   55
         Top             =   2400
         Width           =   1125
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   -90
         Picture         =   "frmMain.frx":F9A2
         ScaleHeight     =   30
         ScaleWidth      =   1125
         TabIndex        =   54
         Top             =   1680
         Width           =   1125
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   -90
         Picture         =   "frmMain.frx":FD76
         ScaleHeight     =   30
         ScaleWidth      =   1125
         TabIndex        =   53
         Top             =   840
         Width           =   1125
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   -90
         Picture         =   "frmMain.frx":1014A
         ScaleHeight     =   30
         ScaleWidth      =   1125
         TabIndex        =   52
         Top             =   4140
         Width           =   1125
      End
      Begin VB.PictureBox picBtnRefresh 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   180
         MouseIcon       =   "frmMain.frx":1051E
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":10828
         ScaleHeight     =   540
         ScaleWidth      =   480
         TabIndex        =   45
         ToolTipText     =   "Refresh Process List"
         Top             =   3360
         Width           =   480
      End
      Begin VB.PictureBox picPB 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   53
         TabIndex        =   14
         Top             =   4200
         Width           =   795
         Begin VB.Label pbText 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   15
            TabIndex        =   15
            Top             =   0
            Width           =   735
         End
         Begin VB.Shape pbValue 
            BorderColor     =   &H80000012&
            DrawMode        =   6  'Mask Pen Not
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   225
            Left            =   -15
            Top             =   0
            Visible         =   0   'False
            Width           =   15
         End
      End
      Begin VB.PictureBox btnPPrint 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   180
         MouseIcon       =   "frmMain.frx":115EC
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":118F6
         ScaleHeight     =   540
         ScaleWidth      =   480
         TabIndex        =   12
         ToolTipText     =   "Print Process List"
         Top             =   2460
         Width           =   480
      End
      Begin VB.PictureBox btnPTerm 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   180
         MouseIcon       =   "frmMain.frx":126BA
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":129C4
         ScaleHeight     =   540
         ScaleWidth      =   480
         TabIndex        =   10
         ToolTipText     =   "Terminate Selected Process"
         Top             =   1680
         Width           =   480
      End
      Begin VB.PictureBox btnPOptions 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   180
         MouseIcon       =   "frmMain.frx":13788
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":13A92
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   8
         ToolTipText     =   "View Options"
         Top             =   900
         Width           =   540
      End
      Begin VB.PictureBox tmpSearch 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   180
         MouseIcon       =   "frmMain.frx":14A06
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":14D10
         ScaleHeight     =   36
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   6
         ToolTipText     =   "Search for a Process or Module"
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label btnRefresh 
         BackStyle       =   0  'Transparent
         Height          =   675
         Left            =   120
         MouseIcon       =   "frmMain.frx":15AD4
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   3360
         Width           =   675
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Refresh"
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
         Left            =   180
         TabIndex        =   47
         Top             =   3900
         Width           =   510
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   120
         MouseIcon       =   "frmMain.frx":15DDE
         MousePointer    =   99  'Custom
         TabIndex        =   46
         Top             =   3420
         Width           =   555
      End
      Begin VB.Label btnPrint 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   180
         MouseIcon       =   "frmMain.frx":160E8
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   2460
         Width           =   555
      End
      Begin VB.Label btnTerm 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   120
         MouseIcon       =   "frmMain.frx":163F2
         MousePointer    =   99  'Custom
         TabIndex        =   24
         Top             =   1740
         Width           =   675
      End
      Begin VB.Label btnOptions 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   120
         MouseIcon       =   "frmMain.frx":166FC
         MousePointer    =   99  'Custom
         TabIndex        =   23
         Top             =   900
         Width           =   615
      End
      Begin VB.Label btnSearch 
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   180
         MouseIcon       =   "frmMain.frx":16A06
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   60
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Print"
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
         Left            =   240
         TabIndex        =   13
         Top             =   3000
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terminate"
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
         TabIndex        =   11
         Top             =   2220
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
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
         TabIndex        =   9
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
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
         Top             =   600
         Width           =   555
      End
   End
   Begin ComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   4425
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   14473
            Text            =   "Populating List..."
            TextSave        =   "Populating List..."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1217
            MinWidth        =   176
            TextSave        =   "7:59 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Available Page File"
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
      Left            =   1680
      TabIndex        =   66
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Available Virtual Memory"
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
      Left            =   1680
      TabIndex        =   64
      Top             =   4020
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Available Physical Memory"
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
      Left            =   1620
      TabIndex        =   62
      Top             =   3840
      Width           =   1875
   End
   Begin ComctlLib.ImageList imgButtons 
      Left            =   4680
      Top             =   3255
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":16D10
            Key             =   "search"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":17962
            Key             =   "flags"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":185B4
            Key             =   "terminate"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":19206
            Key             =   "print"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":19E58
            Key             =   "help"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1AAAA
            Key             =   "security"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1B6FC
            Key             =   "web"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgTVIcons 
      Left            =   4080
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1C34E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1C668
            Key             =   "MODULE"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imgToolbar 
      Left            =   1140
      Top             =   3060
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewTaskSh 
         Caption         =   "&New Task..."
      End
      Begin VB.Menu l8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "S&ystem"
      Begin VB.Menu mnuSysInfo 
         Caption         =   "&System Info"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuGetProcBySelWin 
         Caption         =   "&Get Process by Window"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchList 
         Caption         =   "Search &List for Process/Module"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu l11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContact 
         Caption         =   "&Contact Us"
      End
      Begin VB.Menu l10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents..."
      End
   End
   Begin VB.Menu mnuListCM 
      Caption         =   "ListCM"
      Visible         =   0   'False
      Begin VB.Menu mnuAdvProc 
         Caption         =   "Advanced Process Information"
      End
      Begin VB.Menu l14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewModule 
         Caption         =   "View Module's Information"
      End
      Begin VB.Menu mnuViewProcess 
         Caption         =   "View Process's Information"
      End
      Begin VB.Menu mnuViewFileVersion 
         Caption         =   "View File's Version Information"
      End
      Begin VB.Menu l5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewWindowThreads 
         Caption         =   "View Window Threads"
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTermProc 
         Caption         =   "&Terminate Process"
      End
      Begin VB.Menu mnuTermPar 
         Caption         =   "&Terminate Parent Process"
      End
   End
   Begin VB.Menu mnuMemCM 
      Caption         =   "MEMCM"
      Visible         =   0   'False
      Begin VB.Menu mnuUpdateInterval 
         Caption         =   "&Change Update Interval"
      End
   End
End
Attribute VB_Name = "frmMain"
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



Private Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long

Private Type KeyInfo
 IsProcess As Boolean
 FilePath As String
 ProcID As String
 ParentID As String
 isMultiThreaded As Boolean
 numThreads As Long
 numModules As Long
End Type
'This structure/type is used by the ParseKey function, see ParseKey function for more info...

Dim CFVersion As VERHEADER 'See VERHEADER structure for more info...
Dim PhyP&, VirP&, PagP&, CurProcID&, CurSelPath$
'Dimensionalize PhyP, VirP, PagP, and CurProcID as long data type, CurSelPath as string data type
'PhyP, VirP, and PagP are used for storing previous Global Physical, Virtual, and PageFile Memory information.
'See the tmrMem_Timer() sub routine for more detailed information...

'general declarations...

Private Sub btnPOptions_Click()
On Error GoTo errh
'On the event of an error the execution point will continue at the errh label
 UpdateWinPos frmOptions.hWnd 'UpdateWindow's Z-axis position based upon global settings, see this function for more info
  frmOptions.DoDlg 'Call this objects DoDlg method, this procedure prepares this objects variables, see this procedure for more info...
   frmOptions.Show vbModal, Me 'Show this form as a modal dialog
    UpdateWinPos Me.hWnd 'Update this windows Z-axis, see this procedure for more info...
     Exit Sub 'Exit this sub routine as the only time the execution point should be at the next line is if an error occured
errh:
 If Err.Number = 339 Then MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'If error 339 occured then a control on this form wasn't initialized, allthough this error shouldn't occur since the controls on this form are initialized when they are loaded by the frmSpash form
  MsgBox "An error has been raised." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Un expected Error"
  'An un-expected error has occured, inform the user...
End Sub

Private Sub btnPTerm_Click()
 Call btnTerm_Click
 'See btnTerm_Click sub routine for more information...
End Sub

Private Sub btnRefresh_Click()
 enumProcesses 'Enumerate Process, see this procedure in mdlMain module for more information...
  RefreshLibrary 'Refresh the TreeView control, and format the nodes, see this procedure for more info...
End Sub

Private Sub btnSearch_Click()
On Error GoTo errh
 UpdateWinPos frmSearch.hWnd 'See this procedure for more info...  frmSearch.Show vbModal, Me 'Show frmSearch as a modal dialog
  frmSearch.Show vbModal, Me 'show form as a modal dialog
   UpdateWinPos Me.hWnd 'See this procedure for more info...
    Exit Sub 'Exit sub routine so execution isn't continued at label errh
errh:
 If Err.Number = 339 Then MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'Control was initialized, this error will not occur during this sub routine, this is just 'fool proof' error handling
  MsgBox "An error has been raised." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Un expected Error"
  'Inform user of the error which occured, if an error occured it would be caused by the operating system
End Sub

Private Sub btnSec_Click()
On Error GoTo errh
'If an error is raised, continue execution at label errh
If IsNTAdmin(0&, 0&) = 0 Then MsgBox "You must be the administrator to use this feature.", vbExclamation, "Administrator Privellages Needed":    Exit Sub
'If the user is not an administrator, then exit this sub
 UpdateWinPos frmSec.hWnd 'See this procedure for more info...
  frmSec.Show vbModal, Me 'Show this form as a modal dialog
   UpdateWinPos Me.hWnd 'See this procedure for more info...
    Exit Sub 'Exit this sub routine as we don't need to continue execution of this sub
errh: 'errh label
 If Err.Number = 339 Then MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'A control caused an error, this error will not occur, frmSpash would have handled this error
  MsgBox "An error has been raised." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Un expected Error"
  'Fool proof error handling, this error would have been caused by the operating system, inform user
End Sub

Private Sub btnTerm_Click()
On Error Resume Next
'On the event of an error continue execution of this procedure at the next line
If TerminationPriv = OnlyAdmin And IsNTAdmin(0&, 0&) = 0 Then MsgBox "You must be an Administrator to terminate a process.", vbInformation, "Security": Exit Sub
'if TerminationPriv equals to OnlyAdmin (security options are set by user using frmSec dialog) and the current user isn't an administrator exit the sub routine
If tvList.SelectedItem Is Nothing Then Exit Sub 'If this controls selecteditem (property returns Node structure) is nothing(not initialized) then exit sub
Dim NodeKey As KeyInfo, SelNode As Node, TerminateReturn As TermReturn: Set SelNode = tvList.SelectedItem
'TerminateReturn returns either Terminated, Failed or cancelled, see TerminateProc(mdlMain) for more info...
'Dimensionalize NodeKey as KeyInfo structure, SelNode as Node structure, initialize SelNode by setting it to the selectedItem Node
 ParseKey NodeKey, SelNode.Key 'This function parses the Nodes Key property and sets NodeKey's properties. See this function for more info...
  If NodeKey.IsProcess = False Then
  'If the selected node isn't a Process then... (This is determined by the key of the specified Node, see the RefreshLibrary sub routine for more info...)
    If MsgBox("You can't terminate a process's module." & vbCrLf & vbCrLf & "Do you wish to terminate " & tvList.SelectedItem.Parent.Text & " process?", vbQuestion + vbYesNo, "Terminate Parent Process") = vbNo Then Exit Sub
     TerminateReturn = TerminateProc(CLng(NodeKey.ProcID), NodeKey.FilePath, tvList.SelectedItem.Parent.Text)
     If TerminateReturn = Failed Then
     'See TerminateProc for more info...
      MsgBox "The process couldn't be terminated.", vbInformation, "Termination Incomplete"
     ElseIf TerminateReturn = Terminated Then
      tvList.Nodes.Remove (tvList.SelectedItem.Parent.Key)
      'Process was terminated, remove it from the Tree View control
     End If
  Else
    If MsgBox("Terminate """ & SelNode.Text & """?" & vbCrLf & vbCrLf & "The application will be terminated cleanly.", vbQuestion + vbYesNo, "Terminate Process") = vbYes Then
     TerminateReturn = TerminateProc(CLng(NodeKey.ProcID), NodeKey.FilePath, tvList.SelectedItem.Text)
     If TerminateReturn = Failed Then
     'See TerminateProc for more info...
       MsgBox "The process couldn't be terminated.", vbInformation, "Termination Incomplete"
     ElseIf TerminateReturn = Terminated Then
       tvList.Nodes.Remove (tvList.SelectedItem.Key)
       'Process was terminated, remove it from the Tree View control
     End If
   End If
  End If
End Sub

Private Sub cmbProperty_Change()
 Call cmbProperty_Click 'See this sub routine for more info...
End Sub

Private Sub cmbProperty_Click()
On Error Resume Next
'On the event of an error, continue execution at the next line of this procedure
 Select Case cmbProperty.ListIndex
 'Select Case statement: See MSDN Help System or MSDN Online for more information...
 'Select Case Expression (Expression to be evaluated)
 ' Case VALUE1(If expression evaluated above is equal to this value): Expression
 ' Case Else: Expression
 'End Select
  Case 0:
  'The ListIndex property returns the selected item's index, in this case index 0 evaluates to "Company Name" in the Combo Box control
   txtInfo.Text = CFVersion.CompanyName
   'CFVersion is a general declaration of this form (privately scoped global)
   'The properties of this variable are set when an item of the TreeView control is clicked, see tvList_ItemClick(..) sub routine for more info...
  Case 1:
   txtInfo.Text = CFVersion.FileDescription
  Case 2:
   txtInfo.Text = CFVersion.FileVersion
  Case 3:
   txtInfo.Text = CFVersion.InternalName
  Case 4:
   txtInfo.Text = CFVersion.LegalCopyright
  Case 5:
   txtInfo.Text = CFVersion.OrigionalFileName
  Case 6:
   txtInfo.Text = CFVersion.ProductName
  Case 7:
   txtInfo.Text = CFVersion.ProductVersion
  Case 8:
   txtInfo.Text = CFVersion.Comments
  Case 9:
   txtInfo.Text = CFVersion.LegalTradeMarks
  Case 10:
   txtInfo.Text = CFVersion.PrivateBuild
  Case 11:
   txtInfo.Text = CFVersion.SpecialBuild
 End Select
End Sub

Private Sub Form_Load()
On Error Resume Next
'If an error occurs, continue execution of this procedure on the next line
frmSpash.UpdateProgress , , "Drawing Main Window GUI"
'see object frmSpash's method UpdateProgress for more information...
If HostOS.OperatingSystem = fullCompatibleOS Then HGrad picL.hdc, picL.ScaleHeight / 2, picL.ScaleWidth / 2, COLOR_BTNFACE, Black2White, UNEQUALITYFADE: picL.Refresh: picL.ScaleMode = 1
'If the host operating system is fully compatible (WinXP) then draw a horizonatal gradient
'If the operating system is not fully compatible(Win2K), we will not draw a gradient because we won't be using transparentBlt to draw some of our button icons
 If HostOS.OperatingSystem = fullCompatibleOS Then HRGrad picR.hdc, picR.ScaleHeight, picR.ScaleWidth, COLOR_BTNFACE: picR.Refresh: picR.ScaleMode = 1
  frmSpash.UpdateProgress , , "Drawing Bitmap Graphics"
  'See frmSpash - > UpdateProgress for more info...
   If HostOS.OperatingSystem = Win2K Then
   'If the host operating system is Windows 2000, then we will be using the actual picture boxes as the buttons as opposed to using the hDc property of the picture boxes to transparentBlt them to the form
    tmpSearch.Visible = True
     tmpSec.Visible = True
      tmpWeb.Visible = True
       tmpHelp.Visible = True
   Else
   'Since the operating system is WinXP then we will be using TransparentBlt to draw the icons overlapping the radial gradient in the top left, and the horizontal gradiens in the lower right
   'The last argument of this function is the Mask Color, this function omits pixels of this specific color
   TransparentBlt picL.hdc, 12, 7, 32, 36, tmpSearch.hdc, 0, 0, 32, 36, vbWhite
    TransparentBlt picR.hdc, 12, 248, 32, 36, tmpSec.hdc, 0, 0, 32, 36, vbWhite
     TransparentBlt picR.hdc, 168, 248, 36, 36, tmpWeb.hdc, 0, 0, 36, 36, vbWhite
      TransparentBlt picR.hdc, 92, 248, 32, 36, tmpHelp.hdc, 0, 0, 32, 36, vbWhite
    End If
     If MemUpdateOn = True Then picOnOff.Picture = picOn.Picture Else picOnOff.Picture = picOff.Picture
     'Evaluate MemUpdateOn(boolean Update Memory Status global) and use the according image...
      tmrMem.Enabled = MemUpdateOn
      'Set the timers enabled property to the MemUpdateOn variable's boolean value
       enumProcesses 'Enumerate processes, see this procedure for more info...
        RefreshLibrary 'Populate the Tree View control, see this procedure for more info...
         cmbProperty.ListIndex = 0 'Select the first item in the combo box
          tmrMem.Interval = CLng(GetSetting("ProcessXP", "Settings", "TMI", "1500"))
          'set Time tmrMem's interval to the registry key "TMI", if the key doesn't exits use the value 1500
           Call tmrMem_Timer 'Call the sub routine tmrMem_Time despite its enabled property so that the memory progress bars are set atleast once
            TransPrep Me.hWnd 'Prepare this window for transparency, see this procedure for more information
             Me.Show 'Show this window
              UpdateWinPos Me.hWnd  'Update the Window position, see this procedure for more info...
End Sub

Public Function RefreshLibrary()
On Error Resume Next
'On the even of an error continue execution at the next line in this procedure
picBtnRefresh.Enabled = False: btnRefresh.Enabled = False: picProcList.Visible = True: tvList.Visible = False
'Disable the refresh button, as the list is being populated
  sbMain.Panels(1).Text = "Populating List..."
  'update the first panel specified by its index(1) with the current status
   Dim itmX As Node, itmMod As Node, i&, j%: tvList.Nodes.Clear: tvList.Enabled = False
   'Dimensionalize itmX as Node structure, itmMod as Node Structure, i as long data type. Clear the tree view control(removes all nodes), disable the tree view control even though it won't be visible to the user
    imgTVIcons.ListImages.Clear
    'Remove all items in the ImageList imgTVIcons...
     imgTVIcons.ListImages.Add , , tmpProcessIcon.Image
     'Add the Process Icon used as the Root Node's image to the Image List
      imgTVIcons.ListImages.Add , "MODULE", tmpModIcon.Image
      'Add the Module icon used as the Module Node(Parent to other modules of any process) to the ImageList
       imgTVIcons.ListImages.Add , , tmpProcIcon.Image
       'Add the Generic Process Icon used as the Process Icon if the user has decided to not display the Processes Icon as this saves time
        tvList.Nodes.Add , , "ROOT", ProcCnt & " Processes [" & ModCnt & " Modules]", 1
        'Add the Root Node to which all other nodes will be a child to
         For i = 0 To ProcCnt
         'For next loop, there will be as many iterations as there are Processes(specified by the global variable ProcCnt)
         DoEvents: PbHandle ((i / ProcCnt) * 100)
         'DoEvents function yeilds execution to other procedures processing asynchronously
         'call PHandle function which will determine the percent complete and update the population progress bar(Bottom Left)
          If ShowProcIcon = True Then GetIcon fProc(i).ProcessPath, picIconTmp
          'If ShowProcIcon evaluates to true then call GetIcon function to return the icon handle(index:0) of the specified file, and draw it to a temporary picture box(To make this more efficient, instead of drawing the Icon to a picture box you could store the hDc(Handle to a Device Context) in memory; ex: Dim tmpHdc&
           If ShowProcIcon = True Then imgTVIcons.ListImages.Add , fProc(i).ProcessEXE, picIconTmp.Image
           'If ShowProcIcon evaluates to true then add the temporaty picture boxes image property to the Image List with the Key of the Process Name(EXE Name), this is done so that the Image uses a unique identifier, but to save memory will only be used on for each node who's process name is identical
            If ShowProcIcon = True Then
             picIconTmp.Cls 'Clear this temporary picture which is used to store the drawn icon of the file
              Set itmX = tvList.Nodes.Add("ROOT", tvwChild, fProc(i).ProcessPath & "|P*|PID" & Str$(fProc(i).eProcessID) & "|/ParID" & Str$(fProc(i).eParentID) & "|/MT" & Str$(fProc(i).isMT) & "|/nTh" & fProc(i).numThreads & "|/" & Str$(fProc(i).ModulesCnt), fProc(i).ProcessEXE, fProc(i).ProcessEXE)
              'initialize variable itmX(Node Structure) with the return of the objects Add method wich returns the Node added
              'The key of this node stores information such as ModulePath,PID(ProcessID), boolean MultiThreaded, number of threads and so on...
              'When a procedure need to retrieve this information is calls the ParseKey function which parses this information
            Else
             Set itmX = tvList.Nodes.Add("ROOT", tvwChild, fProc(i).ProcessPath & "|P*|PID" & Str$(fProc(i).eProcessID) & "|/ParID" & Str$(fProc(i).eParentID) & "|/MT" & Str$(fProc(i).isMT) & "|/nTh" & fProc(i).numThreads & "|/" & Str$(fProc(i).ModulesCnt), fProc(i).ProcessEXE, 3)
              'initialize variable itmX(Node Structure) with the return of the objects Add method wich returns the Node added
              'The key of this node stores information such as ModulePath,PID(ProcessID), boolean MultiThreaded, number of threads and so on...
              'When a procedure need to retrieve this information is calls the ParseKey function which parses this information
            End If
             If Starting = True Then frmSpash.UpdateProgress i, ProcCnt, "Populating Process List": DoEvents
             'If starting evaluates to true as it will when this form is being loaded by frmSpash update its progress by calling its UpdateProgress sub routine
              If Starting = False Then lblProcListTmp.Caption = "Populating Process List (" & CStr(CLng((i / ProcCnt) * 100)) & "% Complete)"
              'If starting evaluates to true as it will when this form is being loaded by frmSpash update its progress by calling its UpdateProgress sub routine
               If fProc(i).ModulesCnt <= 0 Then GoTo SkipM
               'If this processes modules count equals to 0 then jump to label SkipM (Skip Modules), otherwise continue to populate the modules of this process
                If ModRelation = True Then Set itmMod = tvList.Nodes.Add(itmX, tvwChild, "MODULEPARENT" & CStr(i), "Modules", "MODULE")
                'Add the Modules Parent node to the treeview control, this is the Parent node to which the actuals modules will be child to in relation
                 For j = 0 To fProc(i).ModulesCnt
                 'For next loop; there will be as many iterations as there are modules
                 DoEvents 'Yield execution to other procedures processing asynchronously
                  If ShowModIcon = True Then
                  'if ShowModIcon evaluates to true then the user desires to get and display the icons for the enumerated modules, otherwise the generic module icon is used as the node image.
                   GetIcon fProc(i).ModulesPath(j), picIconTmp
                   'call GetIcon function to return the icon handle(index:0) of the specified file, and draw it to a temporary picture box
                    imgTVIcons.ListImages.Add , fProc(i).Modules(j), picIconTmp.Image
                    'Add the image to the ImageList
                     picIconTmp.Cls 'clear the temporary picture box
                      If ModRelation = True Then
                      'If ModRelation evaluates to true then the modules will be add to the treeview control based upon their parent child relationship(the nodes representing modules will be added to the Parent MODULE node of the current Process Node
                       tvList.Nodes.Add itmMod, tvwChild, fProc(i).ModulesPath(j) & "|PID" & Trim$(Str$(fProc(i).eProcessID)), fProc(i).Modules(j), fProc(i).Modules(j)
                      Else
                       tvList.Nodes.Add "ROOT", tvwChild, fProc(i).ModulesPath(j) & "|PID" & Trim$(Str$(fProc(i).eProcessID)), fProc(i).Modules(j), fProc(i).Modules(j)
                       'Add the node to the Root node(no parent child relationship)
                      End If
                  Else
                   'User wishes not to use the icons of the actual module file
                   If ModRelation = True Then
                    'If ModRelation evaluates to true then the modules will be add to the treeview control based upon their parent child relationship(the nodes representing modules will be added to the Parent MODULE node of the current Process Node
                    tvList.Nodes.Add itmMod, tvwChild, fProc(i).ModulesPath(j) & "|PID" & Trim$(Str$(fProc(i).eProcessID)), fProc(i).Modules(j), "MODULE"
                   Else
                    tvList.Nodes.Add "ROOT", tvwChild, fProc(i).ModulesPath(j) & "|PID" & Trim$(Str$(fProc(i).eProcessID)), fProc(i).Modules(j), "MODULE"
                    'Add the node to the Root node(no parent child relationship)
                   End If
                  End If
                 Next j
SkipM: 'Label: SkipM, jumped to when a process has no enumerated modules
        Next i 'Next process
         sbMain.Panels(1).Text = "Ready"
         'Update the statusbars first panel(index 1) to the status of the list population which is now complete
          tvList.Nodes.Item("ROOT").Expanded = True
          'Expand the ROOT node
           tvList.Enabled = True
           'Enable the TreeView control window
            picBtnRefresh.Enabled = True: btnRefresh.Enabled = True: picProcList.Visible = False: tvList.Visible = True
            'Enable the Refresh Button, set the TreeView controls visiblility property to true(visible)
             SortProcesses 'Call SortProcesses; See this procedure for more info...
             'SortProcesses function manipulates the Nodes which represent processes parent child relationship properties
End Function

Private Function SortProcesses()
If ProcRelation = True Then
'if ProcRelation evaluates to true...
Dim tPkey As KeyInfo, tMKey As KeyInfo, cNode As Node, npNode As Node, CurInd&, i%
'Dimensionalize tPkey as KeyInfo structure, tMKey as KeyInfo structure, cNode as Node structure, npNode as Node structure, and CurInd as long data type
 For i = tvList.Nodes.Count To 1 Step -1
 'For next loop; there will be as many iterations as there are nodes
 'In this For, Next loop the step keyword is used, the variable i will be decremented by one each iteration
   If tvList.Nodes.Item(i).Key = "ROOT" Or tvList.Nodes.Item(i).Key = "MODULEPARENT" Then GoTo skipaction
   'If the current node is not the Root node, or a MODULEPARENT node (which is the parent node to all modules if the ModRelation ship flag was true during the last list population then jump to the label skipaction
    If InStr(1, tvList.Nodes.Item(i).Key, "|P*|PID", 1) > 0 Then
    'Instr function determines the position of one string with in another string
    'If the current node is a Node representing a process then(if P* exists in the Key)...
     ParseKey tPkey, tvList.Nodes.Item(i).Key
     'Parse the nodes key, see this procedure for more info...
      If GetParentNode(tPkey.ParentID, npNode, tPkey.ProcID) = False Then GoTo skipaction
      'GetParentNode functions determines if the specified node is a child node to another Node representing a process, if it isn't then jump to label SkipAction
       Set cNode = tvList.Nodes(i)
       'Initialize the variable Node cNode with the current node of index i
        MoveNode cNode, npNode, False
        'Call MoveNode function to Move the node to the specified node
        'This function is recursive as it must also enumerate through the specified nodes children who's amount isn't constant
        'See MoveNode function for more info...
    End If
skipaction: 'Label skipaction
 Next i 'Next node index
End If
End Function

Function MoveNode(mNode As Node, pNode As Node, Optional enumSib As Boolean = True)
On Error Resume Next
'On the event of an error, resume execution of this procedure on the next line
Dim NodeKey$, i&, NewNode As Node, tmpNode As Node, tmpNewNode As Node
'Dimensionalize NodeKey as string data type, NewNode as Node Struct., tmpNode as Node struct., tmpNewNode as Node stuct.
 mNode.Key = "*" & mNode.Key
 'Since the key of a node must be unique, we will add an asterik to the node as we will be copying this node
  tvList.Nodes.Add pNode, tvwChild, Mid$(mNode.Key, 2), mNode.Text, mNode.Image
  'Copy the Node to the new parent node
   If mNode.Children > 0 Then
   'If the current node has children nodes(modules, other processes)...
    Set tmpNode = mNode.Child 'Initialize tmpNode with the return of the mNode.Child property who's return is Node structure...
     For i = 1 To mNode.Children
     'For, Next loop; starting at 1, iterations will discontinue when i evaluates to the number of children of the specified node
      tmpNode.Key = "*" & tmpNode.Key
      'Again, we will use an asterek and append the old key value so that this node's key property will be unique
       tvList.Nodes.Add Mid$(mNode.Key, 2), tvwChild, Mid$(tmpNode.Key, 2), tmpNode.Text, tmpNode.Image
       'Add the new node with the unique key to the new parent(the copied process node)
        If tmpNode.Children > 0 Then MoveNode tmpNode.Child, tvList.Nodes(Mid$(tmpNode.Key, 2))
        'if the child node which was just copied and moved has children nodes...
         If i <= mNode.Children Then Set tmpNode = tmpNode.Next
         'if i evaluates to less than or equal to the amount of Children nodes then set tmpNode to the return of Next function which will return the Next sibling node
     Next i 'Next Node index
   End If
    If enumSib = True Then
    'if enumSib (Boolean Enumerate Siblings) evaluates to true then...
     If mNode.Parent.Children > 0 Then
     'If mNode(Module Node)'s parent ammount of children is greater than zero...
      Set tmpNode = mNode.Parent.Child: Set tmpNode = tmpNode.Next
      'set tmpNode the next Processes child which would be a previously moved process
       For i = 1 To mNode.Parent.Children - 1
       'For next loop, there will be as many iterations as there are (modules parent) children nodes minus one
        tmpNode.Key = "*" & tmpNode.Key
        'Again use an asterik to ensure this key will be unique as the original node will remain in the list as the new node will be moved to a new parent...
         tvList.Nodes.Add pNode.Key, tvwChild, Mid$(tmpNode.Key, 2), tmpNode.Text, tmpNode.Image
         'Move the newly copied node...
          If tmpNode.Children > 0 Then MoveNode tmpNode.Child, tvList.Nodes(Mid$(tmpNode.Key, 2))
          'if tmpNode's ammount of children is greater than zero...
           If i <= mNode.Parent.Children Then Set tmpNode = tmpNode.Next
           'if current index is less than or equal to the mNode(module node)'s parent number of children nodes then set tmpNode to the return of the next sibling node
       Next i 'Next Node index
      End If
    End If
     tvList.Nodes.Remove mNode.Key
     'Remove the orignal node, including its children nodes...
End Function

Private Function GetParentNode(ByVal ParProcID$, ByRef opNode As Node, ByVal ProcID$) As Boolean
On Error Resume Next
'On the event of an error resume execution on the next line in this procedure
Dim i&, tmpPID$, itmX As Node: ParProcID$ = Trim$(ParProcID$)
'Dimensionalize i as long data type, tmpPID as string data type, itmX as Node Structure
'Call function Trim($ as string) to remove leading and trailing space characters
Dim tmpBuf$: tmpBuf = ParProcID$
'dimensionalize tmpBuf as string data type, initialize this variable with the value of ParProcID
 For i = 1 To tvList.Nodes.Count
 'For Next loop, there will be as many iterations as there are nodes in the Tree View control
  If InStr(1, tvList.Nodes(i).Key, "P*", 1) = 0 Or Left$(tvList.Nodes(i).Key, 1) = "1" Then GoTo Conti
  'If the "P*" exists in the current node's key prop. or the first character is 1 then...
   tmpPID$ = Mid$(tvList.Nodes(i).Key, InStr(1, tvList.Nodes(i).Key, "PID ", 1) + 4, InStr(1, tvList.Nodes(i).Key, "|/", 1) - 2)
    tmpPID$ = Left$(tmpPID, InStr(1, tmpPID, "|", 1) - 1)
    'Parse the string after 'PID ' and before '|/' which should return the embedded(in key) process id(PID)
     If Trim(tmpPID$) = Trim(ParProcID$) And Trim(tmpPID$) <> Trim(ProcID$) Then
     'If the string comparison of ProcessID(specified by the ParProcID argument) is equal to the current nodes process id then...
      Set opNode = tvList.Nodes(i)
      'Initialize opNode with the return of the Nodes specified in the collection with the index of i
      'This is a ByRef(by reference) argument and so the procedure which is calling this function assumes the argument opNode as the Node return(this is the parent node of the specified ProcessID)
       GetParentNode = True
       'Since the Parent node was found return true
        Exit Function 'Exit this function
     End If
Conti: 'Label Conti
 Next i 'Next Node Index
  GetParentNode = False
  'If the parent node was found this function was terminated, so if the execution point is at this line then we can assume that the parent node was node found
  'Return false
End Function

Private Sub Form_Resize()
On Error Resume Next
'On the event of an error resume at the next line in this procedure
 tvList.Width = (Me.Width - 4155)
 'Use this sub routine if you wish to resize the frmMain object while edititing its controls
 'This procedure is called when the form is loaded
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim Form As Form
'On the event of an error resume at the next line in this procedure
Cancel = 0 'Do not cancel
 SaveSetting "ProcessXP", "Startup", "CI", CStr(CurrentInfo)
  SaveSetting "ProcessXP", "Startup", "MU", CStr(MemUpdateOn)
  'Save settings in the registry
    For Each Form In Forms
    'for next loop; enumerate through each loaded form(forms collection consists of loaded forms only)
     If Form.Name <> "frmMain" Then Unload Form
     'If the current form is not this form, then unload it
    Next Form 'select next form in collection
     ReDim fProc(0 To 1)
     'Preserve keyword is omitted(as is the data type), this is only significant while debugging
      If Not (sndClass Is Nothing) Then Set sndClass = Nothing
      'Terminate the sndClass(Sound Server), this will purge the temporary wav files from the temporary folder
       End
       'Unload all objects, and free memory
End Sub

Private Sub Label3_Click()
 Call btnTerm_Click 'See btnTerm_Click sub routine for more information...
End Sub

Private Sub lblmdlProcID_Click()
'User wishes to jump to the process who owns the currently selected module
Dim tmpBuffer$ 'Dimensionalize tmpBuffer as string data type
 If lblmdlProcID.Caption <> "" Then
 'if the caption of this label(who value should be the ProcessId of this modules owning process)
  tmpBuffer = Mid$(lblmdlProcID.Caption, 1, InStr(1, lblmdlProcID.Caption, " ", 1) - 1)
  'The caption of this label control is "ProcName - PID#", parse the PID# part of this label's caption property
   GetParentProcess tmpBuffer, True
   'Call GetParentProcess with the optional GotoParentProcess argument set to true so this function selects the specified process if found rather than just returning its parsed key information
 End If
End Sub

Private Sub lblMViewVersionHeader_Click()
Dim itmX As Node: Set itmX = tvList.SelectedItem
'Dimensionalize itmX as Node Structure, initialize itmX with the return of the tvList object's SelectedItem function which returns the selected node
 CurrentInfo = VersionInfo
 'Set currentinfo which uses the InfoFrame enumeration, this is a flag used to determine which frame to display(Process information, or Module information)
   Call tvList_NodeClick(itmX)
   'Call tvList_NodeClick() with the NodeItem argument selected
   'See tvList_NodeClick sub routine for more information...
End Sub

Private Sub lblOpenProp_Click()
On Error Resume Next
'On the event of an error resume execution on the next line of this procedure
Dim SEI As SHELLEXECUTEINFO, R&
'Dimensionalize SEI as SHELLEXECUTEINFO structure, R as long data type
 With SEI
  .cbSize = Len(SEI) 'Initialize this variables size with the length of the variable SEI
   .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
   'set the fMask constant flags, see MSDN Help System or MSDN on-line(http://msdn.microsoft.com) for more information
    .hWnd = Me.hWnd
    'Set this variables .hwnd property to the Window Handle of this window, the dialog created by this function's parent window will be the window with this handle
    'Window Handle properties are unique, every window including child windows and windows with no visibility have this property which can be retreived by several application program interface functions including but not limited to FindWindow, FindWindowEx, GetParent, ...
     .lpVerb = "properties"
     'The verb(open, print, explore, edit, properties) specified will be the context under which a document is executed by Shell
      .lpFile = CurSelPath$
      'Specify the File(Full Path) on which the action will be executed
       .lpParameters = vbNullChar
       'Extended paramaters property flags: See Microsoft Developers Network for more infromation on this and other Shell methods
        .lpDirectory = vbNullChar
        'The directory in which to start(start in)
         .nShow = 0 'See MSDN help system or MSDN online
          .hInstApp = 0 'See MSDN help system or MSDN online
           .lpIDList = 0 'See MSDN help system or MSDN online
 End With
  Err.Clear 'Clear the current error information if any...
   If ShellExecuteEx(SEI) = 0 Then MsgBox "The following error has occured.", vbInformation, "Can't Show Dialog"
    If Err Then MsgBox "There has been an error." & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Error"
    'if there was an error than it was unexpected, inform the user...
     Err.Clear 'Clear the error if any...
End Sub

Private Sub lblPage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button <> 2 Then Exit Sub 'If the button is not equal to two, then exit this sub routine
  PopupMenu mnuMemCM
  'Display a context menu which will allow users to change the interval of the memory status update timer
End Sub

Private Sub lblParentID_Click()
On Error Resume Next
'On the event of an error resume execution on the next line of this procedure...
Dim tPID$: tPID = Left$(lblParentID.Text, InStr(1, lblParentID.Text, " ", 1) - 1)
'Dimensionalize tPID as string data type
'Initialize tPID with the return of the function Left which will return all the character in the specified string before the first space character
 If tPID <> "" Then GetParentProcess tPID$, True
 'if tPID's value does not equal to "", then attempt to select the node representing this processes parent process
End Sub

Private Sub lblParentID_GotFocus()
'Since this is actualy a Textbox despite its 'lbl' prefix we don't want a Caret to be displayed or for there to be any sub string selected
 HideCaret lblParentID.hWnd 'Hides the caret in the specified window(specified by its window handle)
  lblParentID.SelLength = 0 'Ensure that no sub string is selected in the textbox control
End Sub

Private Sub lblPhysical_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button <> 2 Then Exit Sub 'If the button is not equal to two, then exit this sub routine
  PopupMenu mnuMemCM
  'Display a context menu which will allow users to change the interval of the memory status update timer
End Sub

Private Sub lblProcessInformation_Click()
Dim itmX As Node: Set itmX = tvList.SelectedItem
'dimensionalize itmX as Node structure, intialize variable itmX with the return of the selecteditem property(returns the selected node)
 CurrentInfo = ProcessInfo: Call tvList_NodeClick(itmX)
 'Update the CurrentInfo variable to ProcessInfo which is a flag used to determine the information to be displayed when a node is clicked
 'Call tvList_NodeClick(itmX), see this procedure for more information...
End Sub

Private Sub lblVersionInfo_Click()
'Dimensionalize itmX as Node structure, initialize itmX with the node returned by the selecteditem property
Dim itmX As Node: Set itmX = tvList.SelectedItem
 CurrentInfo = VersionInfo 'Update CurrentInfo flag
   Call tvList_NodeClick(itmX) 'See tvList_NodeClick() procedure for more info...
End Sub

Private Sub lblVirtual_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button <> 2 Then Exit Sub 'If the button is not equal to two, then exit this sub routine
  PopupMenu mnuMemCM
  'Display a context menu which will allow users to change the interval of the memory status update timer
End Sub

Private Sub mnuAbout_Click()
On Error GoTo errh
'On the even of an error jump to the label errh
 UpdateWinPos frmAbout.hWnd 'See UpdateWinPos function for more info...
  frmAbout.Show vbModal, Me 'Show frmAbout as a modal dialog(remains active until closed or hidden)
   UpdateWinPos Me.hWnd 'See UpdateWinPos function for more info...
    Exit Sub 'Exit this sub routine as an has not occured
errh: 'label errh
 If Err.Number = 339 Then MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'error 339 signifies a control initiation error, this error will not occur since this error would have been handled when this form is loaded by frmSpash
  MsgBox "An error has been raised." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Un expected Error"
  'Unexpected error, inform user
End Sub

Private Sub mnuAdvProc_Click()
Dim selKey As KeyInfo: If InStr(1, tvList.SelectedItem.Key, "P*", 1) > 0 Then ParseKey selKey, tvList.SelectedItem.Key Else Exit Sub
'dimensionalize selKey as KeyInfo structure, if the selected node is a node representing a process then parse the node's key value, if its not a node representing a process then exit this sub routine
UpdateWinPos frmAdvProc.hWnd 'See UpdateWinPos function for more info...
 frmAdvProc.DoDlg CLng(selKey.ProcID), Trim(GetParentProcess(CStr(selKey.ParentID), False))
 'Call DoDlg to determine the Process who's advanced information the user wishes to retrieve
  frmAdvProc.Show vbModal, Me 'show this form as a modal dialog
   UpdateWinPos Me.hWnd 'See updatewinpos for more information...
End Sub

Private Sub mnuExit_Click()
 Unload Me 'Unload this form
End Sub

Private Sub mnuGetProcBySelWin_Click()
On Error GoTo errh
'on the event of an error jump to the label errh
 UpdateWinPos frmPointWindow.hWnd 'See UpdateWinPos for more info..
  frmPointWindow.Show vbModal, Me 'Show this form as a modal dialog
   UpdateWinPos Me.hWnd 'See UpdateWinPos for more info...
    Exit Sub 'Exit sub routine as further execution is un-needed as an error hasn't occured
errh:
 If Err.Number = 339 Then MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'This error should not occur, it would have been handled by the object who loaded it(frmSpash)
 'Fool Proof error handling...
  MsgBox "An error has been raised." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Un expected Error"
  'Unexpected error, most likely an error raised by the operating system, inform user
End Sub

Private Sub mnuNewTaskSh_Click()
On Error GoTo errh
'on the event of an error jump to the label errh
 UpdateWinPos frmShell.hWnd 'See UpdateWinPos for more info..
  frmShell.Show vbModal, Me 'Show this form as a modal dialog
   UpdateWinPos Me.hWnd 'See UpdateWinPos for more info..
    Exit Sub
errh:
 If Err.Number = 339 Then MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'This error should not occur, it would have been handled by the object who loaded it(frmSpash)
 'Fool Proof error handling...
  MsgBox "An error has been raised." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Un expected Error"
  'Unexpected error, most likely an error raised by the operating system, inform user
End Sub

Private Sub mnuSearchList_Click()
On Error GoTo errh
'on the event of an error jump to the label errh
 UpdateWinPos frmSearch.hWnd 'See UpdateWinPos for more info...
  frmSearch.Show vbModal, Me 'Show this form as a modal dialog
   UpdateWinPos Me.hWnd 'See UpdateWinPos for more info...
    Exit Sub
errh: 'label errh
 If Err.Number = 339 Then MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'This error should not occur, it would have been handled by the object who loaded it(frmSpash)
 'Fool Proof error handling...
  MsgBox "An error has been raised." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Un expected Error"
  'Unexpected error, most likely an error raised by the operating system, inform user
End Sub

Public Sub mnuSysInfo_Click()
On Error GoTo errh
'on the event of an error jump to the label errh
 UpdateWinPos frmSysInf.hWnd 'Updates Window Position and automatically loads frmSysInf
  frmSysInf.Show vbModal, Me 'Show this form as a modal dialog
   UpdateWinPos Me.hWnd 'See UpdateWinPos for more info...
    Exit Sub 'Exit sub to discontinue execution since no error has occured
errh: 'label errh
 If Err.Number = 339 Then MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'if this error has occured, than a component of this window can't be initialized
 'During the loading of this form unlike several of the other forms its even more important to error handle as this form does not remain loaded and so this might be the first time this form was loaded
  MsgBox "An error has been raised." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Un expected Error"
  'Unexpected error, most likely an error raised by the operating system, inform user
End Sub

Private Sub mnuTermPar_Click()
 Call btnTerm_Click 'see btnTerm_Click sub routine
End Sub

Private Sub mnuTermProc_Click()
 Call btnTerm_Click 'see btnTerm_Click sub routine
End Sub

Private Sub mnuUpdateInterval_Click()
On Error GoTo errh
'on the event of an error jump to the label errh
 UpdateWinPos frmTmrInt.hWnd 'See UpdateWinPos for more info...
  frmTmrInt.Show vbModal, Me 'Show this form as a modal dialog
   UpdateWinPos Me.hWnd 'See UpdateWinPos for more info...
    Exit Sub 'Exit sub to discontinue execution since no error has occured
errh: 'label errh
 If Err.Number = 339 Then MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'This error should not occur, it would have been handled by the object who loaded it(frmSpash)
 'Fool Proof error handling...
  MsgBox "An error has been raised." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Un expected Error"
  'Unexpected error, most likely an error raised by the operating system, inform user
End Sub

Private Sub mnuViewFileVersion_Click()
Dim itmX As Node: Set itmX = tvList.SelectedItem
'Dimensionalize itmX as Node struct., initilize itmX with the return of the tree view's selecteditme property
CurrentInfo = VersionInfo 'Update CurrentInfo flag, used to determine the information to be displayed when an item has been clicked
 Call tvList_NodeClick(itmX) 'See tvList_nodeClick() for more info...
End Sub

Private Sub mnuViewModule_Click()
Dim itmX As Node: Set itmX = tvList.SelectedItem
'Dimensionalize itmX as Node struct., initilize itmX with the return of the tree view's selecteditme property
CurrentInfo = ProcessInfo 'Update CurrentInfo flag, used to determine the information to be displayed when an item has been clicked
 Call tvList_NodeClick(itmX) 'See tvList_nodeClick() for more info...
End Sub

Private Sub mnuViewProcess_Click()
Dim itmX As Node: Set itmX = tvList.SelectedItem
'Dimensionalize itmX as Node struct., initilize itmX with the return of the tree view's selecteditme property
CurrentInfo = ProcessInfo 'Update CurrentInfo flag, used to determine the information to be displayed when an item has been clicked
 Call tvList_NodeClick(itmX) 'See tvList_nodeClick() for more info...
End Sub

Private Sub mnuViewWindowThreads_Click()
On Error Resume Next
'On the event of an error resume execution on the next line of this procedure
If AllCanEnumWin = False And IsNTAdmin(0&, 0&) = 0 Then MsgBox "You must be an Administrator to enumerate a Processes Window Threads.", vbInformation, "Security": Exit Sub
'Evaluate global security flags to determine if this user is allowed to enumerate a processes window threads
Dim NodeInfo As KeyInfo
'Dimensionalize NodeIndo as KeyInfo struct.
 ParseKey NodeInfo, tvList.SelectedItem.Key
 'Pase the specified node's key property
  Err.Clear 'clear the current error information if any
   frmSearchThread.DoDlg CLng(NodeInfo.ProcID), tvList.SelectedItem.Text
   'Call DoDlg sub routine to prepare the dialog before the user sees it, see this objects DoDlg method for more information...
    If Err.Number = 339 Then GoTo errh
    'If an error has occured and its number is 339 then jump to label errh (see the code following this label for more info..)
     UpdateWinPos frmSearchThread.hWnd 'See this function for more information...
      frmSearchThread.Show vbModal, Me 'Show this form as a modal dialog
       UpdateWinPos Me.hWnd 'See this function for more information...
        Exit Sub 'Discontinue execution of this procedure to avoid displaying an error message when no error has occured
errh: 'label: errh
 MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'Inform user of the error and how they might solve it
End Sub

Private Sub picBtnRefresh_Click()
 Call btnRefresh_Click 'see btnRefresh_Click sub routine for more info...
End Sub

Private Sub picOnOff_Click()
MemUpdateOn = Not (MemUpdateOn)
'Update Globl variable MemUpdateOn
 If MemUpdateOn = True Then
  picOnOff.Picture = picOn.Picture
  'Update picture property, on picture appears sunken, off picture appears raised
 Else
  picOnOff.Picture = picOff.Picture
  'Update picture property, on picture appears sunken, off picture appears raised
 End If
  tmrMem.Enabled = MemUpdateOn
  'Update tmrMem's enabled property.
  'If MemUpdateOn evaluates to true, the the timer will be enabled, its interval is specified by the global Memory Status Update Inteval
End Sub

Private Function ConvertBytes(SizeByte&, Max&) As String
'This function is used to return the appropriate measurment of Free Global Memory and the percent utilized
On Error GoTo errh
'On the event of an error jump to the label errh
Dim tmpBuffer&, strBuffer$
'dimensionalize tmpBuffer as string data type, strBuffer as string data type
 tmpBuffer = SizeByte / 1024: strBuffer$ = "KB's"
 'Initialize tmpBuffer with the value returned by the mathematical division of argument by 1024(1,024 bytes per KB)
  If tmpBuffer / 1024 > 1 Then tmpBuffer = tmpBuffer / 1024: strBuffer$ = "MB's"
  'if tmpBuffer devided by 1024 evaluates to greater than one, then return in MB's
   If tmpBuffer / 1024 > 1 Then tmpBuffer = tmpBuffer / 1024: strBuffer$ = "GB's"
   'if tmpBuffer devided by 1024 evaluates to greater than one, then return in GB's
    strBuffer = Str$(tmpBuffer) & " " & strBuffer
    'set strBuffer to the Division return, and append the measurement unit
     tmpBuffer = (SizeByte / Max) * 100
     'calculate the percent: (Val / Number) * 100
      ConvertBytes = Str$(tmpBuffer) & "% (" & Trim$(strBuffer) & ")"
      'return the formatted data
       Exit Function 'Discontinue execution of this function
errh:
 ConvertBytes = CStr(SizeByte&) & " bytes"
 'An error occured while calculating, return the inital value
End Function

Private Sub picOnOff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button <> 2 Then Exit Sub 'If button doesn't equal to two then exit this sub routine
  PopupMenu mnuMemCM
  'Display context menu to allow user to change the Global Memory Status update interval
End Sub

Private Sub picProcMemRef_Click()
 'Refresh the memory utilization of the selected process
 lblProcMem.Caption = GetMemory(CurProcID) 'See GetMemory function in mdlMain for more info...
End Sub

Private Sub tmpSearch_Click()
On Error GoTo errh
'On the event of an error jump to label errh
 UpdateWinPos frmSearch.hWnd 'See UpdateWinPos function for more info...
  frmSearch.Show vbModal, Me 'Show this form as a modal dialog
   UpdateWinPos Me.hWnd 'See UpdateWinPos function for more info...
    Exit Sub 'Discontinue execution of this procedure as an error has not occured
errh: 'label errh
 If Err.Number = 339 Then MsgBox Err.Description & vbCrLf & vbCrLf & "A required component couldn't be initiated or wasn't found." & vbCrLf & "Please re-install ProcessXP to solve this problem.", vbCritical, "Component Not Found": Exit Sub
 'This error should not have occured since frmSpash would have handled it and performed actions accordingly
  MsgBox "An error has been raised." & vbCrLf & vbCrLf & """" & Err.Description & """", vbCritical, "Un expected Error"
  'Unexpected error, inform user...
End Sub

Private Sub tmpSec_Click()
 Call btnSec_Click 'See btnSec_Click Sub routine for more info...
End Sub

Private Sub tmrMem_Timer()
On Error GoTo errh
'On the event of an error jump to label errh
Dim MemStat As MEMORYSTATUS, tPercent As Integer, i&
'dimensionalize MemStat as MEMORYSTATUS structure, tPercent as integer(Note: To return the Decimal use a double dimensionalized variable, use Fix() to remove the decimal), i as long data type
 MemStat.dwLength = Len(MemStat) 'Initiate MemStat, set its Length property to the length of its structure
  GlobalMemoryStatus MemStat 'Return the Global Memory Status information, this function uses a pointer to the variable MemStat
    With MemStat
     lblPhysical.Caption = ConvertBytes(.dwAvailPhys, .dwTotalPhys)
     'Update lblPhysical's caption property to the return of the ConvertBytes function
     'See ConvertBytes function for more info..
       Dim Phytmp&: Phytmp = .dwTotalPhys / .dwAvailPhys
       'Dimensionalize Phytmp as long data type, intitialize it with the return of Total Physical Memory devided by the Available Physical Memory
       'Note: operator / performs a devision of two numbers, operator \ returns only the remainder of the devision
        tPercent = picPhysical.ScaleWidth / Phytmp
        'Initialize tPercent with the return of picPhysical's scalewidth devided by the percent of physcial memory utilization
         If tPercent > PhyP Then
         'PhyP's value equals the percent evaluated by the last time tmrMem_Timer was called
         'If tPercent is greater than PhyP which means our memory availability has increased
          For i = PhyP To tPercent
          'For next loop; i will be incremented by one each iteration
           DoEvents 'yeild execution
            picPhysical.Cls 'Clear the picture box
             StretchBlt picPhysical.hdc, 0, 0, i, picPhysical.ScaleHeight, picMemBar.hdc, 0, 0, picMemBar.ScaleWidth, picMemBar.ScaleHeight, SRCCOPY
             'Use stretchblt to stretch the source destination by the specified dimensions, and draw to the Destination hDc(Device Context Handle)
           Next i 'increment i by 1
          Else
           'Our physical memory availability has decreased
           For i = PhyP To tPercent Step -1
           'For next loop; i will be decremented by one each iteration
           DoEvents 'yeild execution
            picPhysical.Cls 'Clear the picture box
             StretchBlt picPhysical.hdc, 0, 0, i, picPhysical.ScaleHeight, picMemBar.hdc, 0, 0, picMemBar.ScaleWidth, picMemBar.ScaleHeight, SRCCOPY
             'Use stretchblt to stretch the source destination by the specified dimensions, and draw to the Destination hDc(Device Context Handle)
           Next i 'decrement i by 1
          End If
           PhyP = tPercent 'Store the current Physical Memory Availablility to perform the animation when this sub is called again

            Dim Virtmp&: Virtmp = .dwTotalVirtual / .dwAvailVirtual
             lblVirtual.Caption = ConvertBytes(.dwAvailVirtual, .dwTotalVirtual)
              picVirtual.Cls
               tPercent = picVirtual.ScaleWidth / Virtmp
                If tPercent > VirP Then
                 For i = VirP To tPercent
                 DoEvents
                  picVirtual.Cls
                   StretchBlt picVirtual.hdc, 0, 0, i, picVirtual.ScaleHeight, picMemBar.hdc, 0, 0, picMemBar.ScaleWidth, picMemBar.ScaleHeight, SRCCOPY
                 Next i
                Else
                 For i = VirP To tPercent Step -1
                 DoEvents
                  picVirtual.Cls
                   StretchBlt picVirtual.hdc, 0, 0, i, picVirtual.ScaleHeight, picMemBar.hdc, 0, 0, picMemBar.ScaleWidth, picMemBar.ScaleHeight, SRCCOPY
                 Next i
                End If
                 VirP = tPercent
                    
                    Dim Pagtmp&: Pagtmp = .dwTotalPageFile / .dwAvailPageFile
                    lblPage.Caption = ConvertBytes(.dwAvailPageFile, .dwTotalPageFile)
                     picPage.Cls
                      tPercent = picPage.ScaleWidth / Pagtmp
                        If tPercent > PagP Then
                         For i = PagP To tPercent
                         DoEvents
                          picPage.Cls
                           StretchBlt picPage.hdc, 0, 0, i, picPage.ScaleHeight, picMemBar.hdc, 0, 0, picMemBar.ScaleWidth, picMemBar.ScaleHeight, SRCCOPY
                         Next i
                        Else
                         For i = PagP To tPercent Step -1
                         DoEvents
                          picPage.Cls
                           StretchBlt picPage.hdc, 0, 0, i, picPage.ScaleHeight, picMemBar.hdc, 0, 0, picMemBar.ScaleWidth, picMemBar.ScaleHeight, SRCCOPY
                         Next i
                        End If
                         PagP = tPercent
    End With
     Exit Sub 'discontinue execution of this procedure as no error has occured
errh: 'label errh
 tmrMem.Enabled = False 'An error occured, disable the timer so this sub isn't call again
  picOnOff.Picture = picOff.Picture 'Update picture to the raised button image
   MsgBox "There has been an error while retreiving the systems global memory status." & vbCrLf & vbCrLf & "The memory update interval has been disabled.", vbInformation, "ProcessXP"
   'Inform user...
End Sub

Private Sub tvList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'On the event of an error resume execution on the next line of this procedure
Dim itmX As Node: Set itmX = tvList.HitTest(X, Y): Err.Clear
'dimensionalize itmX as node structure, initialize itmX with the return of the object tvList's HitTest which returns the node at the specified coordinates(x,y)
 itmX.Selected = True 'Select the node
  If Err.Number = 91 Then GoTo SkipSSel
  'If error 91 occured then itmX is nothing, jump to label SkipSSel
   Call tvList_NodeClick(itmX) 'See tvList_NodeClick sub for more info...
SkipSSel: 'Label SkipSSel
Dim NodeInfo As KeyInfo
'Dimensionalize NodeInfo as KeyInfo structure
 If Button = 2 Then
 'If button evaluates to two...
  If tvList.Nodes.Count <= 0 Then Exit Sub
  'if the treeview control contains zero nodes then exit this sub
    ParseKey NodeInfo, tvList.SelectedItem.Key
    'Parse the selected node's key property
     If NodeInfo.IsProcess = True Then
     'If the selected node is a node representing a process then...
      mnuViewProcess.Visible = True
       mnuViewModule.Visible = False
        mnuTermPar.Visible = False
         mnuTermProc.Visible = True
         'Only menus which perform actions on processes will be visible
          If CurrentInfo = ProcessInfo Then mnuViewProcess.Enabled = False: mnuViewFileVersion.Enabled = True
           If CurrentInfo = VersionInfo Then mnuViewProcess.Enabled = True: mnuViewFileVersion.Enabled = False
           'Determine which Frame(ProcessInfo, or ModuleInfo) is currently displayed based upon the value of global variable CurrentInfo
            mnuViewWindowThreads.Enabled = True
             mnuAdvProc.Enabled = True
             'Enable the menu items which are disabled when the currently selected item is a module
              PopupMenu mnuListCM, , , , mnuViewProcess
              'Display context menu, with the Default flag set to mnuViewProcess. This menu item text will appear bold
     Else
      mnuViewProcess.Visible = False
       mnuViewModule.Visible = True
        mnuTermPar.Visible = True
         mnuTermProc.Visible = False
          mnuViewWindowThreads.Enabled = False
          'Only menus which perform actions on modules will be visible or enabled
          If CurrentInfo = ProcessInfo Then mnuViewModule.Enabled = False: mnuViewFileVersion.Enabled = True
           If CurrentInfo = VersionInfo Then mnuViewModule.Enabled = True: mnuViewFileVersion.Enabled = False
           'Determine which Frame(ProcessInfo, or ModuleInfo) is currently displayed based upon the value of global variable CurrentInfo
            mnuAdvProc.Enabled = False
            'Disble the menu items which are enabled when the currently selected item is a process
             PopupMenu mnuListCM, , , , mnuViewModule
             'Display context menu, with the Default flag set to mnuViewModule. This menu item text will appear bold
    End If
 End If
End Sub

Private Sub tvList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This procedure determines which node is under the cursor, and sets the treview window's tooltiptext property accordingly
On Error Resume Next
'On the event of an error resume execution on the next line of this procedure
Dim itmH As Node 'dimensionalize itmH as node structure
 Set itmH = tvList.HitTest(X, Y)
 'initialize itmH with the return of the HitTest function which returns the Node at the specified coordinates
  If itmH <> Empty Then
  'If itmH is not empty(not initialized) then...
   If itmH.Key <> "ROOT" And InStr(1, itmH.Key, "MODULEPARENT", 1) = 0 Then
   'If the node is not the ROOT or a ModuleParent then...
    If InStr(1, itmH.Key, "|", 1) > 0 Then
    'In the Node key the path should be before the first "|" character
     tvList.ToolTipText = Left$(itmH.Key, InStr(1, itmH.Key, "|", 1) - 1)
     'Set tvList's ToolTipText property to the return of the Left function wich will return the specified ammount of characters from left to right
    Else
     tvList.ToolTipText = itmH.Key
    End If
   End If
  End If
End Sub

Public Sub tvList_NodeClick(ByVal Node As ComctlLib.Node)
On Error Resume Next
'On the event of an error resume execution on the next line of this procedure
If LCase$(Node.Key) = "root" Or InStr(1, Node.Key, "MODULEPARENT", 1) = 1 Then Exit Sub
'If the node specified in the Node paramater is the Root node or a module parent then discontinue execution of this sub routine
If CurrentInfo = ProcessInfo Then
'If Global CurrentInfo evaluates to ProcessInfo then show process info frame...
Dim NodeKey As KeyInfo, GlobalMDLUsage&, tmpBuffer$, mdlPath$, mdlFileInfo As VERHEADER
'dimensionalize NodeKey as KeyInfo structure, GlobalMDLUsage as long data type, tmpBuffer as string data type, mdlPath as string data type, mdlFileInfo as VERHEADER(Version Header) structure
 ParseKey NodeKey, Node.Key
 'Parse the Node's Key property. See ParseKey for more info...
   If NodeKey.IsProcess = True Then
   'If the Node is representing a process then...
    picProc.Visible = True: picMod.Visible = False: picVersion.Visible = False
    'Update the visibility of the appropriate frame(Process Info, Module Info, or Version info)
     picProcIcon.Cls: GetLargeIcon NodeKey.FilePath, picProcIcon
     'Clear picProcIcon picture box, retreive the handle of the files large icon, and draw it
      txtPathName.Text = NodeKey.FilePath: CurSelPath$ = NodeKey.FilePath
      'Update the textbox txtPathName text's property to the file path of the file
       lblProcID.Caption = glbHex(NodeKey.ProcID)
       'Update the label's caption property with the ProcessID
        lblParentID.Text = GetParentProcess(NodeKey.ParentID)
        'Update the label's caption property with the Parents Process ID
         If NodeKey.isMultiThreaded = True Then lblBoolMT.Caption = "True" Else lblBoolMT.Caption = "False"
          lblNumThreads.Caption = NodeKey.numThreads '...
           numModules.Caption = NodeKey.numModules '...
            lblProcName.Caption = Node.Text '...
             txtFullPath.Text = GetShortPath(NodeKey.FilePath)
             'See GetShorPath function for more info...
              lblProcTime.Caption = GetProcTime(CLng(NodeKey.ProcID))
              'See GetProcTime function for more info...
               lblProcMem.Caption = GetMemory(CLng(NodeKey.ProcID))
               'See GetMemory function for more info...
                CurProcID = NodeKey.ProcID '...
   Else
   'if the node represents a module then...
   picProc.Visible = False: picMod.Visible = True: picVersion.Visible = False '...
    tmpBuffer = NodeKey.ProcID '...
    mdlPath$ = NodeKey.FilePath '...
     GetModuleInformation CLng(tmpBuffer), Node.Text, GlobalMDLUsage&
     'see GetModuleInformation for more information...
      lblmdlProcID.Caption = GetParentProcess(tmpBuffer)
      'See GetParentProcess function for more information...
       If GlobalMDLUsage >= 2 Then
        lblGloballyUsed.Caption = FormatNumber(GlobalMDLUsage, 0, , , vbTrue) & " times"
        'FormatNumber function will group the digits(#,###)
       ElseIf GlobalMDLUsage = 1 Then
        lblGloballyUsed.Caption = GlobalMDLUsage & " time"
       ElseIf GlobalMDLUsage = 0 Then
        lblGloballyUsed.Caption = GlobalMDLUsage & " times"
       End If
       'update lblGloballyUsed caption with a grammatically correct statement
        picModIcon.Cls: GetLargeIcon mdlPath$, picModIcon 'see GetLargeIcon for more info...
        'Clear picModIcon picture property, draw the files large icon to it
         lblModName.Caption = Node.Text '...
          txtMdlSPath.Text = GetShortPath(mdlPath$) 'See GetShorPath for more info...
           txtMdlLPath.Text = mdlPath$: CurSelPath$ = mdlPath$ '...
            GetVerHeader mdlPath$, mdlFileInfo 'see GetVerHeader for more info
             txtMdlDesc.Text = mdlFileInfo.FileDescription '...
   End If
End If
If CurrentInfo = VersionInfo Then
'Display File's Version information...
picVersion.Visible = True: picProc.Visible = False: picMod.Visible = False
 Dim vNodeKey As KeyInfo, vmdlPath$
 'Dimensionalize vNodeKey as KeyInfo structure, vmdlPath as string data type
  ParseKey vNodeKey, Node.Key 'Parse the Node's Key property
    If NodeKey.IsProcess = True Then
    'If the Node represents a process then...
     GetVerHeader NodeKey.FilePath, CFVersion 'See GetVerHeader for more information
      picVFileIcon.Cls 'Clear the picture property of this picture box
       GetLargeIcon vNodeKey.FilePath, picVFileIcon: CurSelPath$ = vNodeKey.FilePath
       'Draw the Large Icon of the specified file to picVFileIcon, ...
        lblProductName.Caption = CFVersion.ProductName
        'See GetVerHeader for more info...
         lblProductVersion.Caption = CFVersion.ProductVersion '...
          txtVFilePath.Text = GetShortPath(vNodeKey.FilePath) 'See GetShortPath for more info...
           Call cmbProperty_Click 'See cmbProperty_Click sub routine for more info...
    Else
    'The node must represent a module...
    vmdlPath$ = vNodeKey.FilePath 'Initialize vmdlPath with the return of vNodeKey's File Path value...
     GetVerHeader vmdlPath$, CFVersion 'See GetVerHeader for more info...
      picVFileIcon.Cls 'Clear the picture box
       GetLargeIcon vmdlPath$, picVFileIcon 'Draw the files(specified in the icPath argument) to the picturebox picVFileIcon
        lblProductName.Caption = CFVersion.ProductName '...
         lblProductVersion.Caption = CFVersion.ProductVersion '...
          txtVFilePath.Text = GetShortPath(vmdlPath$): CurSelPath$ = vmdlPath$
          'See GetShortPath function for more info...
          'Update the publically scoped delcaration CurSelPath to the current file path
           Call cmbProperty_Click 'see cmbProperty_Click sub routine for more information
    End If
End If
End Sub

Private Function GetParentProcess(ByVal ProcID$, Optional GoToParent As Boolean = False) As String
On Error Resume Next
'On the event of an error resume execution on the next line of this procedure
Dim i&, ParsedNode As KeyInfo, tmpPID$, itmX As Node: ProcID$ = Trim$(ProcID$)
'Dimensionalize i as long data type, ParsedNode as KeyInfo structure, tmpPID as string data type, itmX as node structure
'Remove ProcID's leading and trailing spaces
Dim tmpBuf$: tmpBuf = ProcID$ 'Initialize tmpBuf to ProcID's value
 For i = 2 To tvList.Nodes.Count
 'For next loop; i starts at 2, loops until i equals to the amount of Nodes in the TreeVeiw control tvList incrementing i by one each iteration
 DoEvents 'Yeild execution to asynchronously processing procedures
  ParseKey ParsedNode, tvList.Nodes(i).Key
  'Parse the Node's Key property
   If Trim(ParsedNode.ProcID) = Trim(ProcID$) Then
   'If the the parent node was found...
    If GoToParent = True Then tvList.Nodes(i).Selected = True: Set itmX = tvList.Nodes(i): Call tvList_NodeClick(itmX): Exit Function
    'If GoToParent argument evaluates to true, then select the node
     tmpBuf = tmpBuf & " (" & tvList.Nodes(i).Text & ")"
     'Append the Process's EXE Name to variable tmpBuf
      GetParentProcess = tmpBuf
      'Return tmpBuf
        Exit For 'The node was found, exit the for next loop
    End If
Conti: 'Conti label
 Next i 'next Node index(increment i)
End Function

Public Function PbHandle(uPercent As Long)
On Error Resume Next
'On the event of an error resume execution on the next line of this procedure
Dim tPercent As Double 'Dimensionalize tPercent as Double data type
 If uPercent > 100 Then uPercent = 100 'Ensure the uPercent argument is less than or equal to 100
  tPercent = CDbl(uPercent) / 100 'initialize tPercent to the evaluation of the division of the double value of uPercent by one-hundred
   pbText.Caption = Str$(uPercent) & "%" '..
    picPB.Cls 'Clear picture box
     StretchBlt picPB.hdc, 0, 0, (tPercent * picPB.ScaleWidth), picPB.ScaleHeight, tmpPicProg.hdc, 0, 0, tmpPicProg.ScaleWidth, tmpPicProg.ScaleHeight, SRCCOPY
     'Stretch the original progress bar's picture to the progress bar picture box
     'The width argument will be a percent of the destination picturebox's width equal to the percent done
     'To increase the quality of this type of graphical progress bar follow these guidlines:
     'For a progress bar with rounded corners;
     'The image would consist of three seperate images(Front, Middle, End). Do not stretch an image horizontally if it contains dominant vertical lines, don't stretch an image vertically if it contains dominant horizontal lines as the resoulution is decreased.
     'Stretch only the middle part of the picture, as it's res. will not decrease since it only contains lines of the angle the image is being stretched(horizontally)
     'After you have stretched the middle(horizontal) part, add the front and end using bitblt...
End Function

Private Function ParseKey(ByRef rKI As KeyInfo, ByVal iKey As String) As Boolean
On Error Resume Next
'Option Compare Text
'On the even of an error continue execution on the next line of this procedure
Dim tmpCpy$, tmpNumT$: tmpCpy = iKey
'Dimensionalize tmpCpy as string data type, tmpNumT as string data type, initialize tmpCpy
 If InStr(1, iKey, "|", 1) <= 0 Then Exit Function
 'If the key does not contain the character "|" which means its the Root or a ModuleParent node, then exit this function
 If InStr(InStr(1, iKey, "|", 1), iKey, "P*", 1) >= 1 Then rKI.IsProcess = True
 'If the string "P*" exists with in the key then the Node type is a node which represents a Process
 If InStr(InStr(1, iKey, "|", 1), iKey, "M*", 1) >= 1 Then rKI.IsProcess = False
 'If the string "M*" exists with in the key then the Node type is a node which represents a Module
  If rKI.IsProcess = True Then
  'If IsProcess(Is a Process) evaluates to true then...
   With rKI 'With Block
    .FilePath = Left$(iKey, InStr(1, iKey, "|", 1) - 1)
    'Parse the Files Path, the string before the first "|" char
     iKey = Mid(iKey, InStr(1, iKey, "|", 1))
     'Remove the FilePath and the FilePath terminating character from the temporary key
      iKey = Mid(iKey, InStr(1, iKey, "|PID ", 1) + 4)
      'Parse the PID(Process ID) information
       rKI.ProcID = Left$(iKey, InStr(1, iKey, "|/", 1) - 1)
       'Update the ProcID value with the previously parsed PID value..
        iKey = Mid(iKey, InStr(1, iKey, "|/ParID ", 1) + 7)
        'Remove the ProcessID and the ProcessID terminating character(s) from the temporary key
         rKI.ParentID = Left$(iKey, InStr(1, iKey, "|/MT", 1) - 1)
         'Update the Parent Process ID value
          iKey = Mid(iKey, InStr(1, iKey, "|/MT ", 1) + 4)
          'Remove the Parent Process ID and the Parent Process ID terminating character from the temporary key
           rKI.isMultiThreaded = Trim(LCase$(iKey)) Like "true|/nth*"
           'set isMultiThreaded to the return of the text comparison Key(which will either be "true|/nth..." or "false|/nth..." ) Like "true|/nth*"
           'In this comparison; if true is found in this pattern the result is true other its false
           'Pattern example:
           '"a1B2!Blah", one of several patterns which would evaluate this string and return true would be:
           '[a-z]#[A-B]#?B*h
           '  |  |  |  |||||-h
           ' a-z |  |  ||||-- Multi Character Wildcard(*)
           '    num |  |||- B
           '       A-B ||-Single Character wildcard(?)
           '          num
           'See MSDN Help System or MSDN Online for more information on Binary and Textual data comparisons using Like and other Text comparing operators
           iKey = Mid(iKey, InStr(1, iKey, "|/nTh", 1) + 5)
           'Remove MultiThreaded part and the MultiThreaded part terminating character from the temporary key
             rKI.numThreads = CLng(Left$(iKey, InStr(1, iKey, "|/", 1) - 1))
             'Parse the number of threads part from the key
             iKey = Mid(iKey, InStr(1, iKey, "|/ ", 1) + 3)
             'Parse the Number of Modules part of the key
               rKI.numModules = CLng(iKey) '...
   End With 'End with block
  Else
   With rKI
   'Node represents a Module
    .ProcID = Mid$(iKey, InStr(1, iKey, "|PID", 1) + 4)
    'Parse its owning process's processid from the key
     .FilePath = Left$(iKey, InStr(1, iKey, "|", 1) - 1)
     'Parse the file path part of this key
   End With
  End If
   ParseKey = True
   'Return true
    Exit Function 'discontinue execution of this function since no error has occured
errh: 'label errh
 ParseKey = False
 'return false as an error has occured
End Function
