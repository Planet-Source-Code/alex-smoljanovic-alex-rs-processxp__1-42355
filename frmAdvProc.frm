VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAdvProc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Process Information"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   Icon            =   "frmAdvProc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6915
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   45
      Left            =   1320
      Picture         =   "frmAdvProc.frx":08CA
      ScaleHeight     =   45
      ScaleWidth      =   3840
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   3840
   End
   Begin ComctlLib.ListView lstProc 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   7435
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Value"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
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
      Left            =   6000
      TabIndex        =   2
      Top             =   4260
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   467
      TabIndex        =   3
      Top             =   4200
      Width           =   7005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "See System Info > Processes"
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
      Left            =   60
      MouseIcon       =   "frmAdvProc.frx":11EA
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   4500
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To view the advanced process information of every process,"
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
      Top             =   4380
      Width           =   4140
   End
End
Attribute VB_Name = "frmAdvProc"
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


Dim gProcID&, gParProcName$
'Global variables
'Dimensionalize gProcID as long data type, gParProcName as string data type
'These variables are used when the user wishes to refresh the info

Public Function DoDlg(ProcID&, Optional ParProcName$ = "") As Boolean
On Error Resume Next
Dim Item As Object
'On the event of an error, the execution point of this procedure will continue on the next line
gProcID& = ProcID: gParProcName$ = ParProcName
'Initialize global variables
'This function is called before this dialog is show, since this dialog stays loaded, its only unloaded when the application is ended. This increases the memory usage of this program, but when the user wishes to view the advanced process information of a process this dialog will have been allready loaded and so it should be instantly(minus the time needed to retreive the process information) shown
Dim objWMI As Object, objProc As Object, itmX As ListItem
'Dimensionalize objWMI as 'generic' object, objProc as 'generic' object, itmX as ListItem structure
 Set objWMI = GetObject("winmgmts:\\.\root\cimv2") '. = local computer, you can also use the computer name of a computer on a local network
 'Initialize objWMI with the return of the GetObject function, which will return an instance of the WMI service
  Set objProc = objWMI.execquery("Select * from Win32_Process where ProcessID = """ & ProcID & """", , 48)
  'Execute a query, this query is searching through all elements of Win32_Process column where ProcessID is equal to the ProcID specified by the ProcID argument of this function
   lstProc.ListItems.Clear 'Clear the list as it might contain items from a previous search
   For Each Item In objProc
   'Enumerate through each Item in object objProc
   'Note: since the name of some of the elements are non-user friendly this long manual method is used
    If CStr(Item.ProcessID) = "" Or CLng(Item.ProcessID) <> ProcID Then GoTo lblNext
    'Verify the object returned is the data desired
     With lstProc 'Using with block keyword: See With keyword in you're MSDN Help system or MSDN online(http://msdn.microsoft.com) for more info
      .HideColumnHeaders = False
      'Show column headers
       Set itmX = .ListItems.Add(, "Caption", "Caption")
       'Initialize variable itmX with the return of the Add method, this returns a ListItem object
       itmX.SubItems(1) = Item.Caption 'This collection propery(SubItems) validates to the column header specified by index argument
        Set itmX = .ListItems.Add(, , "Command Line")
        itmX.SubItems(1) = CStr(Item.CommandLine)
         Set itmX = .ListItems.Add(, , "Computer")
         itmX.SubItems(1) = Item.CSName
          Set itmX = .ListItems.Add(, , "Description")
          itmX.SubItems(1) = Item.Description
           Set itmX = .ListItems.Add(, , "Executable Path")
           itmX.SubItems(1) = Item.ExecutablePath
            Set itmX = .ListItems.Add(, , "Handle")
            itmX.SubItems(1) = glbHex(Item.Handle)
             Set itmX = .ListItems.Add(, , "Handle Count")
             itmX.SubItems(1) = Item.HandleCount
              Set itmX = .ListItems.Add(, , "Kernel Mode Time")
              itmX.SubItems(1) = CStr(Item.KernelModeTime)
               Set itmX = .ListItems.Add(, , "Max Working Size")
               itmX.SubItems(1) = FormatByteSize(Item.MaximumWorkingSetSize)
               'The FormatByteSize simply determines the Size measurement and size to return
               'Return Table:
               'If size less than 1 KB(1024 bytes) then returns in bytes with " bytes" appended to the return string
               'If size greater than or equal to 1 KB than it returns the size in KB's with " KB's" appended to the return string
               'If size less than 1 MB byte greater than 1 KB then returns in KB's
               'If size greater than or equal to 1 MB then returns in MB's with " MB's" appended to the return string
               '...
                Set itmX = .ListItems.Add(, , "Min Working Size")
                itmX.SubItems(1) = FormatByteSize(Item.MinimumWorkingSetSize)
                 Set itmX = .ListItems.Add(, , "Name")
                 itmX.SubItems(1) = Item.Name
                  Set itmX = .ListItems.Add(, , "Other Operation Count")
                  itmX.SubItems(1) = CStr(Item.OtherOperationCount)
                   Set itmX = .ListItems.Add(, , "Other Transfer Count")
                   itmX.SubItems(1) = CStr(Item.OtherTransferCount)
                    Set itmX = .ListItems.Add(, , "Page Faults")
                    itmX.SubItems(1) = CStr(Item.PageFaults)
                     Set itmX = .ListItems.Add(, , "Page File Usage")
                     itmX.SubItems(1) = FormatByteSize(Item.PagefileUsage)
                      Set itmX = .ListItems.Add(, , "Parent Process")
                      If InStr(1, ParProcName$, "(", 1) > 0 Then
                      'If the Parent Process Name argument of this function is not omitted then append the parent process name
                       itmX.SubItems(1) = glbHex(CStr(Item.ParentProcessId)) & " " & Mid$(ParProcName$, InStr(1, ParProcName$, "(", 1))
                      Else
                       itmX.SubItems(1) = glbHex(CStr(Item.ParentProcessId))
                      End If
                       Set itmX = .ListItems.Add(, , "Peak Page File Usage")
                       itmX.SubItems(1) = CStr(Item.PeakPagefileUsage)
                        Set itmX = .ListItems.Add(, , "Peak Virtual Size")
                        itmX.SubItems(1) = FormatByteSize(Item.PeakVirtualSize)
                         Set itmX = .ListItems.Add(, , "Peak Working Set Size")
                         itmX.SubItems(1) = FormatByteSize(Item.PeakWorkingSetSize)
                          Set itmX = .ListItems.Add(, , "Priority")
                          itmX.SubItems(1) = CStr(Item.Priority)
                           Set itmX = .ListItems.Add(, , "Private Page Count")
                           itmX.SubItems(1) = CStr(Item.PrivatePageCount)
                            Set itmX = .ListItems.Add(, , "Process ID")
                            itmX.SubItems(1) = glbHex(CStr(Item.ProcessID))
                             Set itmX = .ListItems.Add(, , "Quota NonPaged Pool Usage")
                             itmX.SubItems(1) = CStr(Item.QuotaNonPagedPoolUsage)
                              Set itmX = .ListItems.Add(, , "Quota Paged Pool Usage")
                              itmX.SubItems(1) = CStr(Item.QuotaPagedPoolUsage)
                               Set itmX = .ListItems.Add(, , "Quota Peak NonPaged Pool Usage")
                               itmX.SubItems(1) = CStr(Item.QuotaPeakNonPagedPoolUsage)
                                Set itmX = .ListItems.Add(, , "Quota Peak Paged Pool Usage")
                                itmX.SubItems(1) = CStr(Item.QuotaPeakPagedPoolUsage)
                                 Set itmX = .ListItems.Add(, , "Read Operation Count")
                                 itmX.SubItems(1) = CStr(Item.ReadOperationCount)
                                  Set itmX = .ListItems.Add(, , "Read Transfer Count")
                                  itmX.SubItems(1) = CStr(Item.ReadTransferCount)
                                   Set itmX = .ListItems.Add(, , "Session ID")
                                   itmX.SubItems(1) = glbHex(CStr(Item.SessionId))
                                    Set itmX = .ListItems.Add(, , "Status")
                                    itmX.SubItems(1) = Item.SessionId
                                     Set itmX = .ListItems.Add(, , "Thread Count")
                                     itmX.SubItems(1) = CStr(Item.ThreadCount)
                                      Set itmX = .ListItems.Add(, , "User Mode Time")
                                      itmX.SubItems(1) = CStr(Item.UserModeTime)
                                       Set itmX = .ListItems.Add(, , "Virtual Size")
                                       itmX.SubItems(1) = FormatByteSize(Item.VirtualSize)
                                        Set itmX = .ListItems.Add(, , "Windows Version")
                                        itmX.SubItems(1) = Item.WindowsVersion
                                         Set itmX = .ListItems.Add(, , "Working Set Size")
                                         itmX.SubItems(1) = FormatByteSize(Item.WorkingSetSize)
                                          Set itmX = .ListItems.Add(, , "Write Operation Count")
                                          itmX.SubItems(1) = CStr(Item.WriteOperationCount)
                                           Set itmX = .ListItems.Add(, , "Write Transfer Count")
                                           itmX.SubItems(1) = CStr(Item.WriteTransferCount)

     End With
     'end with block
      DoDlg = True
      'return true
       Set objWMI = Nothing
       'Terminate object objWMI
        Set objProc = Nothing
        'Terminate object objProc
         Exit Function
         'Exit the function since the requested process information was found
lblNext:
'Label lblNext
   Next Item
   'select the next item
    'The requested process was not found
    lstProc.HideColumnHeaders = True
    'Hide the controls column headers
     lstProc.ListItems.Add , , "Can't Find Specified Process"
     'add a single item to the control informing user the process wasn't found
      Set objWMI = Nothing
      'Terminate object objWMI
       Set objProc = Nothing
       'Terminte object objProc
End Function

Function FormatByteSize(ByVal bSize, Optional strAppend$ = "") As String
On Error GoTo errh
'On the event of an error, the execution point of this procedure will jump to label: errh
Dim tmpBuffer$ 'dimensionalize tmpBuffer as string data type
If IsNull(bSize) = True Then FormatByteSize = "<Undetermined Size>": Exit Function
'In the database managed by the WMI service, all elements who's value can't be retreived are left un-initialized(null), if the value is null then return the appropriate value ('undetermined')
 If bSize = 0 Then FormatByteSize = "0 Bytes " & strAppend$: Exit Function
 'the size argument is 0, cant perform a division by zero, so return '0 bytes' and exit the function
  If bSize < 1024 Then FormatByteSize = CStr(bSize) & " Bytes " & strAppend$: Exit Function
  'size argument is less than one kilobyte, so return bytes
  If bSize >= 1024 Then bSize = bSize / 1024: tmpBuffer = " KB's " & strAppend$
  'size is greater or equal to 1 KB so return in Kilobytes
   If bSize >= 1024 Then bSize = bSize / 1024: tmpBuffer = " MB's " & strAppend$
   'Size is greater than or equal to 1 Megabyte, return in MB's
    If bSize >= 1024 Then bSize = bSize / 1024: tmpBuffer = " GB's " & strAppend$
    'Size is greater than or equal to 1 gigabyte, return gigabyte's
     If bSize >= 1024 Then bSize = bSize / 1024: tmpBuffer = " TB's " & strAppend$
     'size is greater than or equal to 1 terabyte, return terabyte's
      FormatByteSize = CStr(FormatNumber(bSize, 0, , , vbTrue)) & tmpBuffer: tmpBuffer = ""
      'call formatnumber function to remove the digits after the decimal(eqv fnct: fix()), and group the digits (###,###)
       Exit Function
errh:
 FormatByteSize = "<!ERROR!> " & Err.Description
 'An error has occured with in the scope of this function, return <!ERROR!> with the description of the error appended to the string
End Function

Private Sub Command1_Click()
'Refresh process information
 If DoDlg(gProcID&, gParProcName$) = False Then
  'If function returned false, the process wasn't found
  MsgBox "An error occured while trying to retreive information about Process: " & CStr(gProcID&), vbExclamation, "Error"
 End If
End Sub

Private Sub Form_Activate()
 TransPrep Me.hWnd 'Prepare window for transparency based upon global user transparency settings
End Sub

Private Sub Form_Load()
 If HostOS.OperatingSystem = Win2K Then
  Me.BackColor = &H8000000F
   Picture1.Visible = False
   'For visual/graphic appearance compatibility:
   'set the backcolor property of this form to the default windows background system color
   'set the picture1's visibility property to false as this will look out of place otherwise
 End If
  HGrad Picture2.hdc, Picture2.ScaleHeight, Picture2.ScaleWidth, COLOR_ACTIVECAPTION, Black2White, UNEQUALITYFADE
  'Draw horizontal gradient; see ProcGUIDLL c++ proj for more info...
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Me.Visible = True Then Cancel = 1: Me.Hide
 'If this form is visible, then it shouldn't be unloaded, call Hide method this will reactive the window who showed it as a modal dialog
End Sub

Private Sub Label2_Click()
 Me.Hide 'Hide this form
  Load frmSysInf 'Load system information form, this is one of the few forms of this project who don't stay loaded for memory preservation purposes
   frmSysInf.TabStrip.Tabs.Item("Processes").Selected = True
   'set the specified tab(specified by its key or index)'s selected property to true, this will select the specified tab
    Call frmSysInf.TabStrip_Click
    'call subroutine TabStrip_Click, see this procedure for more info
     Call frmMain.mnuSysInfo_Click
     'call mnuSysInfo_Click subroutine, this subroutine properly shows the system information form, see this sub for more info...
End Sub
