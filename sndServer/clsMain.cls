VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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


Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_MEMORY = &H4
Private Const SND_LOOP = &H8
Private Const SND_NOSTOP = &H10
Private Const SND_FILENAME = &H20000

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260

Private wAddr$, initPath$

Public Function doAbout(Optional LoadFile As Boolean = False, Optional fPurge As Boolean = False) As Boolean
On Error GoTo errh 'on the event of an error jump to label errh
If fPurge = True And Dir(initPath$ & "tmpAbout.wav", vbHidden) <> "" Then SetAttr initPath$ & "tmpAbout.wav", vbNormal: Kill initPath$ & "tmpAbout.wav": Exit Function
'if fPurge evaluates to true and the file exists then purge it from the hard disk
'note: when the file is created it's hidden attribute is set, so to delete it you must first remove using Kill its hidden attribue
 If LoadFile = False Then GoTo pSnd
 'if LoadFile evaluates to false then jump to label pSnd (Play Sound)
  wAddr = StrConv(LoadResData(101, "WAVE"), vbUnicode)
  'initialize wAddr with the Unicode formatted wave file data loaded from resources
   If Dir(initPath$ & "tmpAbout.wav", vbHidden) <> "" Then SetAttr initPath$ & "tmpAbout.wav", vbNormal: Kill initPath$ & "tmpAbout.wav"
   'If the file allready exists, then kill it
    Open initPath$ & "tmpAbout.wav" For Binary Access Write As #1
    'open the file for binary i/o with write access privellages as file number 1 (while manipulating multiple files use Freefile function to return the next available file number)
     Put #1, , wAddr
     'write the wave file data to the opened file
    Close #1 'close the file
     SetAttr initPath$ & "tmpAbout.wav", vbHidden
     'set the files hidden attribute
      doAbout = True 'return true (successful)
       Exit Function 'exit this procedure
pSnd: 'label pSnd
 doAbout = sndPlaySound(initPath$ & "tmpAbout.wav", SND_FILENAME Or SND_ASYNC)
 'play the wav file asynchronously
  Exit Function 'exit this procedure
errh: 'label err
 doAbout = False 'return false
  On Error Resume Next 'ensure the error we will raise won't create an endless loop by jumping back to label errh
   Err.Raise Err.Number, , Err.Description 'raise the same error
End Function

'... see Function doAbout for reference
Public Function doShell(Optional LoadFile As Boolean = False, Optional fPurge As Boolean = False) As Boolean
On Error GoTo errh
If fPurge = True And Dir(initPath$ & "tmpShell.wav", vbHidden) <> "" Then SetAttr initPath$ & "tmpShell.wav", vbNormal: Kill initPath$ & "tmpShell.wav": Exit Function
 If LoadFile = False Then GoTo pSnd
  wAddr = StrConv(LoadResData(103, "WAVE"), vbUnicode)
   If Dir(initPath$ & "tmpShell.wav", vbHidden) <> "" Then SetAttr initPath$ & "tmpShell.wav", vbNormal: Kill initPath$ & "tmpShell.wav"
    Open initPath$ & "tmpShell.wav" For Binary Access Write As #1
     Put #1, , wAddr
    Close #1
     SetAttr initPath$ & "tmpShell.wav", vbHidden
      doShell = True
       Exit Function
pSnd:
 doShell = sndPlaySound(initPath$ & "tmpShell.wav", SND_FILENAME Or SND_ASYNC)
  Exit Function
errh:
End Function

'... see Function doAbout for reference
Public Function doPBW(Optional LoadFile As Boolean = False, Optional fPurge As Boolean = False) As Boolean
On Error GoTo errh
If fPurge = True And Dir(initPath$ & "tmpPBW.wav", vbHidden) <> "" Then SetAttr initPath$ & "tmpPBW.wav", vbNormal: Kill initPath$ & "tmpPBW.wav": Exit Function
 If LoadFile = False Then GoTo pSnd
  wAddr = StrConv(LoadResData(102, "WAVE"), vbUnicode)
   If Dir(initPath$ & "tmpPBW.wav", vbHidden) <> "" Then SetAttr initPath$ & "tmpPBW.wav", vbNormal: Kill initPath$ & "tmpPBW.wav"
    Open initPath$ & "tmpPBW.wav" For Binary Access Write As #1
     Put #1, , wAddr
    Close #1
     SetAttr initPath$ & "tmpPBW.wav", vbHidden
      doPBW = True
       Exit Function
pSnd:
 doPBW = sndPlaySound(initPath$ & "tmpPBW.wav", SND_FILENAME Or SND_ASYNC)
  Exit Function
errh:
End Function

'... see Function doAbout for reference
Public Function doTerm(Optional LoadFile As Boolean = False, Optional fPurge As Boolean = False) As Boolean
On Error GoTo errh
If fPurge = True And Dir(initPath$ & "tmpTerm.wav", vbHidden) <> "" Then SetAttr initPath$ & "tmpTerm.wav", vbNormal: Kill initPath$ & "tmpTerm.wav": Exit Function
 If LoadFile = False Then GoTo pSnd
  wAddr = StrConv(LoadResData(104, "WAVE"), vbUnicode)
   If Dir(initPath$ & "tmpTerm.wav", vbHidden) <> "" Then SetAttr initPath$ & "tmpTerm.wav", vbNormal: Kill initPath$ & "tmpTerm.wav"
    Open initPath$ & "tmpTerm.wav" For Binary Access Write As #1
     Put #1, , wAddr
    Close #1
     SetAttr initPath$ & "tmpTerm.wav", vbHidden
      doTerm = True
       Exit Function
pSnd:
 doTerm = sndPlaySound(initPath$ & "tmpTerm.wav", SND_FILENAME Or SND_ASYNC)
  Exit Function
errh:
End Function

'... see Function doAbout for reference
Public Function doSec(Optional LoadFile As Boolean = False, Optional fPurge As Boolean = False) As Boolean
On Error GoTo errh
If fPurge = True And Dir(initPath$ & "tmpSec.wav", vbHidden) <> "" Then SetAttr initPath$ & "tmpSec.wav", vbNormal: Kill initPath$ & "tmpSec.wav": Exit Function
 If LoadFile = False Then GoTo pSnd
  wAddr = StrConv(LoadResData(105, "WAVE"), vbUnicode)
   If Dir(initPath$ & "tmpSec.wav", vbHidden) <> "" Then SetAttr initPath$ & "tmpSec.wav", vbNormal: Kill initPath$ & "tmpSec.wav"
    Open initPath$ & "tmpSec.wav" For Binary Access Write As #1
     Put #1, , wAddr
    Close #1
     SetAttr initPath$ & "tmpSec.wav", vbHidden
      doSec = True
       Exit Function
pSnd:
 doSec = sndPlaySound(initPath$ & "tmpSec.wav", SND_FILENAME Or SND_ASYNC)
  Exit Function
errh:
End Function

Private Sub Class_Initialize()
'this sub routine is called when an instance of this class is created
Dim gTP& 'dimensionalize gTP as long data type
initPath$ = String(MAX_PATH, 0) 'allocate memory (MAX_PATH defines the length of characters a valid path can consist of)
 gTP& = GetTempPath(MAX_PATH, initPath$) 'retrieve the current users temporary path
 'function returns the path length, copies the path data to initPath
  initPath$ = Left$(initPath$, gTP&) 'return the length of the path only removing the null terminating characters
   doAbout True  'load About wave file
    doShell True 'load Shell wave file
     doPBW True 'load ProcessByWindow wav file
      doTerm True 'load termination wav file
       doSec True 'load security wave file
End Sub

Private Sub Class_Terminate()
'this sub routine is called when an instance of this class is terminated(unloaded)
On Error Resume Next 'on the event of an error resume execution on the next line of this procedure
Call sndPlaySound("", SND_FILENAME Or SND_ASYNC)
'call function sndPlaySound with the file name ommitted forcing the last function call which played a sound to stop
 doAbout , True 'purge About wav file from disk
  doShell , True 'purge Shell wav file from disk
   doPBW , True 'purge ProcessByWindow wav file from disk
    doTerm , True 'purge Termination wav file from disk
     doSec , True 'purge Security wav file from disk
End Sub
