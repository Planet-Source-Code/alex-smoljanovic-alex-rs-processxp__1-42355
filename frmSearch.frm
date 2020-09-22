VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   264
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   566
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   4200
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   28
      Top             =   3540
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh Process's"
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
      Left            =   60
      TabIndex        =   27
      Top             =   3540
      Width           =   1515
   End
   Begin VB.PictureBox picMTemp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4080
      Picture         =   "frmSearch.frx":08CA
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   25
      Top             =   3180
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox tmpIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3660
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   24
      Top             =   1620
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CheckBox chkModule 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search &Modules"
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
      Left            =   2220
      TabIndex        =   21
      Top             =   720
      Value           =   1  'Checked
      Width           =   1395
   End
   Begin VB.CommandButton cmdSelectIn 
      Caption         =   "&Select Item in Process List"
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
      Left            =   6300
      TabIndex        =   17
      Top             =   2760
      Width           =   1995
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Matching Items"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   4200
      TabIndex        =   16
      Top             =   180
      Width           =   4155
      Begin ComctlLib.TreeView tvFound 
         Height          =   2235
         Left            =   60
         TabIndex        =   18
         Top             =   240
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   3942
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imglstTVIcons"
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
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   240
         Picture         =   "frmSearch.frx":0E54
         ScaleHeight     =   105
         ScaleWidth      =   3840
         TabIndex        =   19
         Top             =   2460
         Width           =   3840
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   3735
      TabIndex        =   13
      Top             =   2820
      Width           =   3735
      Begin VB.OptionButton optRetSelect 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select found items in the Main Process list"
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
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton optRetDisplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display Results in a new list to the Right"
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
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3195
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Find Next"
      Enabled         =   0   'False
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
      Left            =   1800
      TabIndex        =   12
      Top             =   3540
      Width           =   1095
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
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
      Left            =   2940
      TabIndex        =   11
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   3975
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   180
         Picture         =   "frmSearch.frx":1774
         ScaleHeight     =   240
         ScaleWidth      =   3720
         TabIndex        =   29
         Top             =   240
         Width           =   3720
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select a criteria on which to base you're query"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   30
            Top             =   30
            Width           =   3240
         End
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1680
         Picture         =   "frmSearch.frx":1CFE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   23
         Top             =   540
         Width           =   240
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3600
         Picture         =   "frmSearch.frx":2288
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   22
         Top             =   540
         Width           =   240
      End
      Begin VB.CheckBox chkProcess 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search &Processes"
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
         Left            =   180
         TabIndex        =   20
         Top             =   540
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.PictureBox picSearch 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3420
         Picture         =   "frmSearch.frx":2812
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   120
         Picture         =   "frmSearch.frx":30DC
         ScaleHeight     =   30
         ScaleWidth      =   3780
         TabIndex        =   8
         Top             =   2340
         Width           =   3780
      End
      Begin VB.ComboBox cmbMemUsage 
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
         ItemData        =   "frmSearch.frx":39FC
         Left            =   1140
         List            =   "frmSearch.frx":3A39
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   2715
      End
      Begin VB.TextBox txtFilePath 
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
         TabIndex        =   5
         Text            =   "?:\program files\*\  ;  ?:\winnt\*"
         Top             =   1440
         Width           =   2955
      End
      Begin VB.TextBox txtFileTitle 
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
         TabIndex        =   3
         Text            =   "gdi##.???"
         Top             =   960
         Width           =   2955
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Left            =   60
         Picture         =   "frmSearch.frx":3B41
         ScaleHeight     =   30
         ScaleWidth      =   3780
         TabIndex        =   1
         Top             =   840
         Width           =   3780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "With results:"
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
         TabIndex        =   9
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Memory Use:"
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
         TabIndex        =   6
         Top             =   1980
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Path:"
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
         TabIndex        =   4
         Top             =   1500
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Title:"
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
         TabIndex        =   2
         Top             =   1020
         Width           =   630
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   26
      Top             =   0
      Width           =   4095
   End
   Begin ComctlLib.ImageList imglstTVIcons 
      Left            =   4560
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSearch.frx":4461
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSearch"
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


Option Compare Text

Private Type SearchCriteria
 rProcesses As Boolean
 rModules As Boolean
 FileTitle As String
 FilePath As String
 MemUsageLower As Long
 MemUsageUpper As Long
End Type

Private Type SearchRet
 PID As Long
 EXEName As String
 Index As Long
End Type

Private Type wpSearchRet
 PID As Long
 EXEName As String
 Index As Long
 PEPath As String
End Type

Private Enum plNodeType
 Exclusion = 0
 Process = 1
 module = 2
End Enum

Dim OldSearch As SearchCriteria, RefreshingList As Boolean
Dim oSearchRet() As SearchRet, tmpSearchRet() As wpSearchRet, CurInd&, CntSearchRet&
'dimensionalize OldSearch as enumeration SearchCriteria, RefreshingLins as boolean
'OsearchRet Array as enumeration SearchRet, tmpSearchRet array as enumeration wpSearchRet, CurInd as long type, cntSearchRet as long type
'general declerations

Private Sub cmdFind_Click()
On Error Resume Next
If RefreshingList = True Then MsgBox "Please wait until the list is refreshed.", vbExclamation, "Refreshing Process List": Exit Sub
'Can't search through the Nodes if is list is being populated...
Dim TitleStrings(0 To 25) As String, TitleCnt&, i&, tmpBuffer$, tQuery$, PathStrings(0 To 25) As String, PathCnt&, pQuery$
'Dimensionalize TitleStrings as a one dimensional array as string data types, TitleCnt as long type, tmpBuffer as string type, PathStrings as one dimensional array as string type, PathCnt as long type and pQuery as string data type
'User can seperate multiple names, and file paths so ";" is used as a seperater, The spaces before and after this character are purged from the string
tQuery = txtFileTitle.Text: pQuery = txtFilePath.Text
'Initialize tQuery(Title Query), and pQuery(Path Query)
 If Mid$(tQuery, Len(tQuery) - 1) = ";" Then tQuery = Mid$(tQuery, Len(tQuery) - 1)
 'If the title query string is terminated with the ';' character, then remove it...
  If Mid$(pQuery, Len(pQuery) - 1) = ";" Then pQuery = Mid$(pQuery, Len(pQuery) - 1)
  'If the path query string is terminated with the ';' character, then remove it...

If InStr(1, tQuery, ";", 1) = 0 Then GoTo DoPaths
'If the seperator character doesn't exist withing the title query then jump to DoPaths label
 For i = 0 To 25
 'For next loop; i starts at 0, loop continues until i is equal to 25 incrementing i by one each iteration
  If InStr(1, tQuery, ";", 1) = 0 And TitleCnt& = 0 Then GoTo DoPaths
  'If ";" doesn't exist with in the Title query, and TitleCnt evaluates to zero then jump to DoPaths
  If InStr(1, tQuery, ";", 1) = 0 And TitleCnt& > 0 Then TitleStrings(i) = tQuery: GoTo DoPaths
  'If ";" doesn't exist with in the Title query and TitleCnt evals to zero, then
   tmpBuffer = Mid$(tQuery, 1, InStr(1, tQuery, ";", 1) - 1)
   'intitialize tmpBuffer with the substring before the first ";"
    If InStr(1, tmpBuffer, ";", 1) > 0 Then
    'if the seperator character ";" exists in the tmpBuffer variable then...
     tmpBuffer = Left$(tmpBuffer, InStr(1, tmpBuffer, ";", 1) - 1)
     'remove the seperator character...
    End If
     TitleStrings(i) = tmpBuffer
     'Set the element of index i of the array TitleStrings to the tmpBuffer variables value
      tQuery = Mid$(tQuery, InStr(1, tQuery, ";", 1) + 1)
      'Remove the first title in the string(the string before ";", and the string terminating character ";")
       TitleCnt& = TitleCnt& + 1
       'Increment TitleCnt variable by one
 Next i 'Increment i by one, next iteration
  TitleCnt& = TitleCnt& - 1
  'Since TitleCnt is incremented after each Title string is determined, TitleCnt will be 1 greater than they are actual Title Strings
  'Decrement TitleCnt
  'This part of the procedure parsed all of the Title Strings into an arry
  'ex: "Shell32.dll;Explore;nt*.dll
  'TitleStrings arrary would now consist of three elements, 'Shell32.dll', 'Explore', and 'nt*.dll'
DoPaths: 'DoPaths label
   If TitleCnt = 0 And tQuery <> "" Then
   'If no Title Strings were parsed but the title query doesn't equal "" then...
    TitleStrings(TitleCnt) = Trim(txtFileTitle.Text)
    'Set the first element in the TitleStrings array to the Title Qeuery
   End If
    If InStr(1, pQuery, ";", 1) = 0 Then GoTo DoSkip
    'If ";" does not exist in the Path query than jump to label DoSkip...
     For i = 0 To 25
     'For next loop;i starts at 0, loops until i is equal to 25
      If InStr(1, pQuery, ";", 1) = 0 And PathCnt& = 0 Then GoTo DoSkip
      'If ";" doesn't exist in pQuery and no paths have been parsed then jump to label DoSkip
       If InStr(1, pQuery, ";", 1) = 0 And PathCnt& > 0 Then PathStrings(i) = pQuery: GoTo DoSkip
       'If ";" doesnt exist in pQery and PathCnt is greater than 0 as it will when the last Path from the query is being parsed then jump to label DoSkip
        tmpBuffer = Mid$(pQuery, 1, InStr(1, pQuery, ";", 1) - 1)
        'set tmpBuffer to the first path specified before the first ";" character in the string
         If InStr(1, tmpBuffer, ";", 1) > 0 Then
         'if ";" exists in tmpBuffer then ...
          tmpBuffer = Left$(tmpBuffer, InStr(1, tmpBuffer, ";", 1) - 1)
          'set tmpBuffer to the string before the first ";" character in tmpBuffer...
         End If
          PathStrings(i) = tmpBuffer 'Set element i of the PathStrings array to tmpBuffer...
           pQuery = Mid$(pQuery, InStr(1, pQuery, ";", 1) + 1)
           'remove everything before the first ";" character including the first ";" character from pQuery
            PathCnt& = PathCnt& + 1
            'Increment PathCnt by one
     Next i 'increment i by one, next iteration
      PathCnt& = PathCnt& - 1 'Decrement PathCnt by one so that PathCnt evaluates to the number of elements(Path Queries) set in the array
DoSkip: 'label DoSkip
       If PathCnt = 0 And pQuery <> "" Then
       'if no paths have been parsed and path query doesn't equal to "" then..
        PathStrings(PathCnt) = txtFilePath.Text
        'Set the first element in the PathStrings array to the Path Qeuery
       End If
        dSearch TitleStrings(), TitleCnt, PathStrings(), PathCnt
        'See dSearch function for more info...
End Sub

Private Function MatchString(strAr() As String, strMatch$, strCnt&, Optional exStrMatch As String) As Boolean
Dim i& 'Dimensionalize i as long data type
 For i = 0 To strCnt
 'For next loop; i starts at 0, loop continues until i evaluates to argument strCnt incrementing i by one each iteration
  If LCase(Trim(strMatch$)) Like LCase(Trim(IIf(Right(Trim(strAr(i)), 1) = "\" Or Right(Trim(strAr(i)), 1) = "?" Or Right(Trim(strAr(i)), 1) = "*" Or Trim(exStrMatch) = "", Trim(strAr(i)), Trim(strAr(i)) & "\") & Trim(exStrMatch))) = True Then
  'This function matches the titles of two files, if they match this function is called again to match the paths.For example:
  'Title Match: "gdi32.dll" with "gdi##.???", since they match...
  'Path Match: "C:\winnt\system32\gdi32.dll" with "?:\winnt\*\gdi##.???"(exStrMatch appended to the strAr() element)
  'In this case the pattern matching it checking if gdi##.??? is in a next level sub-directory of winnt
  
  'Match patterns, if evaluates to true then return true
   MatchString = True:  Exit For
  End If
 Next i 'increment i, next iteration
End Function

Public Function dSearch(TitleQuery() As String, cntTQ&, PathQuery() As String, cntPQ&)
On Error Resume Next
Dim i&, cNodeType As plNodeType, cMem&
Dim tmpMemU$, tmpMemL$, tmpMemBuffer$, UseMUpper As Boolean, UseMLower As Boolean
Dim FileTitle$, FilePath$, tPID$
ReDim tmpSearchRet(0 To frmMain.tvList.Nodes.Count) As wpSearchRet
CntSearchRet& = 0: CurInd& = 0
 With frmMain.tvList
  For i = 1 To .Nodes.Count
  DoEvents
   GetNodeType .Nodes(i).Key, cNodeType
    If cNodeType = Exclusion Then GoTo NoAdd
     If cNodeType = Process Then
      If chkProcess.Value = 0 Then GoTo NoAdd
       tmpMemBuffer = cmbMemUsage.List(cmbMemUsage.ListIndex)
        
        If cmbMemUsage.List(cmbMemUsage.ListIndex) = "Any Amount" Then
         UseMUpper = False: UseMLower = False
        Else
         tmpMemL = Left$(tmpMemBuffer, InStr(1, tmpMemBuffer, " ", 1) - 1)
          tmpMemU = Mid$(tmpMemBuffer, InStr(1, tmpMemBuffer, "to ", 1) + 3)
           tmpMemU = Mid$(tmpMemU, 1, InStr(1, tmpMemU, " ", 1) - 1)
            If tmpMemL = "*" Then UseMLower = False Else UseMLower = True
             If tmpMemU = "*" Then UseMUpper = False Else UseMUpper = True
        End If
         FileTitle = .Nodes(i).Text: FilePath = Left$(.Nodes(i).Key, InStr(1, .Nodes(i).Key, "|", 1) - 1)
          If txtFileTitle.Text <> "" Then
           If MatchString(TitleQuery(), FileTitle, cntTQ) = False Then GoTo NoAdd
          End If
           If txtFilePath.Text <> "" Then
            If MatchString(PathQuery(), FilePath, cntPQ, FileTitle) = False Then GoTo NoAdd
             If Trim(FilePath) = "" Then GoTo NoAdd
           End If
            tPID$ = Mid$(.Nodes(i).Key, InStr(1, .Nodes(i).Key, "P*|PID", 1) + 6)
            tPID = Mid$(tPID, 1, InStr(1, tPID, "|/ParID", 1) - 1)
             cMem = GetMemoryLong(CLng(tPID))
              If UseMUpper = True And UseMLower = True Then
               If cMem < CLng(tmpMemL) Or cMem > CLng(tmpMemU) Then GoTo NoAdd
              ElseIf UseMUpper = False And UseMLower = True Then
               If cMem < CLng(tmpMemL) Then GoTo NoAdd
              ElseIf UseMLower = False And UseMUpper = True Then
               If cMem > CLng(tmpMemU) Then GoTo NoAdd
              End If
SkipCMem:
               tmpSearchRet(CntSearchRet).PID = CLng(tPID)
                tmpSearchRet(CntSearchRet).EXEName$ = .Nodes(i).Text
                 tmpSearchRet(CntSearchRet).Index& = i
                  tmpSearchRet(CntSearchRet).PEPath = FilePath
     Else
      If chkModule.Value = 0 Then GoTo NoAdd
       FileTitle = .Nodes(i).Text: FilePath = Left$(.Nodes(i).Key, InStr(1, .Nodes(i).Key, "|", 1) - 1)
        If txtFileTitle.Text <> "" Then
         If MatchString(TitleQuery(), FileTitle, cntTQ) = False Then GoTo NoAdd
        End If
         If txtFilePath.Text <> "" Then
          If MatchString(PathQuery(), FilePath, cntPQ, FileTitle) = False Then GoTo NoAdd
           If Trim(FilePath) = "" Then GoTo NoAdd
         End If
          tPID$ = Mid$(.Nodes(i).Key, InStr(1, .Nodes(i).Key, "|PID", 1) + 4)
           tmpSearchRet(CntSearchRet).PID = CLng(tPID)
            tmpSearchRet(CntSearchRet).EXEName$ = .Nodes(i).Text
             tmpSearchRet(CntSearchRet).Index& = i
              tmpSearchRet(CntSearchRet).PEPath = FilePath
     End If
      CntSearchRet = CntSearchRet + 1
NoAdd:
  Next i
 End With
  CntSearchRet = CntSearchRet - 1
   Me.Caption = "Search - Found " & CStr(CntSearchRet + 1) & " results"
    If CntSearchRet >= 0 Then
     ReDim oSearchRet(0 To CntSearchRet) As SearchRet
      For i = 0 To CntSearchRet '- 1
       oSearchRet(i).EXEName = tmpSearchRet(i).EXEName
        oSearchRet(i).Index = tmpSearchRet(i).Index
         oSearchRet(i).PID = tmpSearchRet(i).PID
          imglstTVIcons.ListImages.Add , "ROOT", picMTemp.Image
           If optRetDisplay.Value = True Then
            tmpIcon.Cls
             GetIcon tmpSearchRet(i).PEPath, tmpIcon
              imglstTVIcons.ListImages.Add , Trim(tmpSearchRet(i).EXEName), tmpIcon.Image
           End If
      Next i
       If optRetSelect.Value = True Then
        Dim itmX As Node
         frmMain.tvList.Nodes(oSearchRet(0).Index).Selected = True: frmMain.tvList.SelectedItem.EnsureVisible
          Set itmX = frmMain.tvList.Nodes.Item(oSearchRet(0).Index)
           frmMain.tvList_NodeClick itmX
            Me.Width = 4230
             CurInd& = CurInd& + 1
       Else
        tvFound.Nodes.Clear
         tvFound.Nodes.Add , , "ROOT", "Found " & CStr(CntSearchRet + 1) & " Items", "ROOT"
          Me.Width = 8580
           For i = 0 To CntSearchRet
            DoEvents
             tvFound.Nodes.Add "ROOT", tvwChild, CStr(oSearchRet(i).Index) & "|" & CStr(oSearchRet(i).PID), oSearchRet(i).EXEName, Trim(oSearchRet(i).EXEName)
           Next i
            tvFound.Nodes("ROOT").Expanded = True
       End If
        If CntSearchRet& >= 0 Then cmdNext.Enabled = True
    Else
     tvFound.Nodes.Clear
      tvFound.Nodes.Add , , "ROOT", "Found No Matches", "ROOT"
    End If
End Function

Private Function GetNodeType(ByVal NKey$, ByRef nType As plNodeType)
Dim tmpBuffer$
 If InStr(1, NKey, "ROOT", 1) >= 1 Or InStr(1, NKey, "MODULEPARENT", 1) >= 1 Then nType = Exclusion: Exit Function
  If InStr(1, NKey, "|P*|PID", 1) > 0 Then nType = Process: Exit Function
   If InStr(1, NKey, "|PID", 1) > 0 Then nType = module: Exit Function
    nType = Exclusion
End Function


Private Sub cmdNext_Click()
On Error GoTo errh
Dim itmX As Node
 frmMain.tvList.Nodes(oSearchRet(CurInd&).Index).Selected = True: frmMain.tvList.SelectedItem.EnsureVisible
  Set itmX = frmMain.tvList.Nodes.Item(oSearchRet(CurInd&).Index)
   frmMain.tvList_NodeClick itmX: tvFound.Nodes.Item(CurInd + 2).Selected = True: tvFound.SelectedItem.EnsureVisible
    If CurInd& >= CntSearchRet& Then CurInd& = 0 Else CurInd& = CurInd& + 1
errh:
End Sub

Private Sub cmdSelectIn_Click()
On Error GoTo errh
Dim ItmInd&, tmpBuffer$, itmX As Node
 tmpBuffer = tvFound.SelectedItem.Key
  If tmpBuffer = "ROOT" Then Exit Sub
   ItmInd = CLng(Mid(tmpBuffer, 1, InStr(1, tmpBuffer, "|", 1) - 1))
    frmMain.tvList.Nodes(ItmInd).Selected = True
     frmMain.tvList.Nodes(ItmInd).EnsureVisible
      Set itmX = frmMain.tvList.Nodes.Item(ItmInd): frmMain.tvList_NodeClick itmX
       If CurInd& >= CntSearchRet& Then CurInd& = 0 Else CurInd& = CurInd& + 1
errh:
End Sub

Private Sub Command1_Click()
 RefreshingList = True
  enumProcesses
   frmMain.RefreshLibrary
    RefreshingList = False
End Sub

Private Sub Form_Activate()
 TransPrep Me.hwnd
End Sub

Private Sub Form_Load()
 Me.Width = 4230
  cmbMemUsage.ListIndex = 0
   SHAutoComplete txtFilePath.hwnd, SHACF_FILESYSTEM
    HGrad Picture7.hdc, Picture7.ScaleHeight, Picture7.ScaleWidth, COLOR_ACTIVECAPTION, Black2White, UNEQUALITYFADE
     HRGrad Picture8.hdc, Picture8.ScaleHeight, Picture8.ScaleWidth, COLOR_ACTIVECAPTION
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Me.Visible = True Then Cancel = 1: Me.Hide
End Sub

