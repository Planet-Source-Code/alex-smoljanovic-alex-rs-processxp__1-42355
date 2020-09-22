Attribute VB_Name = "mdlVersion"
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


Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
    
Public Type VERHEADER
 CompanyName As String
 FileDescription As String
 FileVersion As String
 InternalName As String
 LegalCopyright As String
 OrigionalFileName As String
 ProductName As String
 ProductVersion As String
 Comments As String
 LegalTradeMarks As String
 PrivateBuild As String
 SpecialBuild As String
End Type

'The following are predefined version information Unicode strings:
'Comments, InternalName, ProductName
'CompanyName, LegalCopyright, ProductVersion
'FileDescription, LegalTrademarks, PrivateBuild
'FileVersion, OriginalFilename, SpecialBuild

Public Function GetVerHeader(ByVal fPN$, ByRef oFP As VERHEADER)
'Argument fPN by variable value as string type, oFP (Output) by variable reference as VERHEADER type struct.
Dim lngBufferlen&, lngDummy&, lngRc&, lngVerPointer&, lngHexNumber&, i%
Dim bytBuffer() As Byte, bytBuff(255) As Byte, strBuffer$, strLangCharset$, strVersionInfo(11) As String, strTemp$
'dimensionalize lngBufferlen as long type, lngDummy as long type, lngRc as long type
'lngVerPointer as long type(pointer to a memory reference), lngHexNumber&,i as integer type
'bytBuffer as a one dimensional of bytes, bytBuff as a one dimensional of bytes,
'strBuffer as string type, strLangCharset as string type, one dimensional array strVersion as string type, strTemp as string type
 If Dir(fPN$) = "" Then
 'If the file doesn't exist then...
  oFP.CompanyName = "The file """ & GetShortPath(fPN) & """ doesn't exist."
  'initialize oFP's CompanyName member; see GetShorPath for more info(Returns Old DOS style paths)
   oFP.FileDescription = "The file """ & GetShortPath(fPN) & """ doesn't exist."
   '...
    oFP.FileVersion = "The file """ & GetShortPath(fPN) & """ doesn't exist."
     oFP.InternalName = "The file """ & GetShortPath(fPN) & """ doesn't exist."
      oFP.LegalCopyright = "The file """ & GetShortPath(fPN) & """ doesn't exist."
       oFP.OrigionalFileName = "The file """ & GetShortPath(fPN) & """ doesn't exist."
        oFP.ProductName = "The file """ & GetShortPath(fPN) & """ doesn't exist."
         oFP.ProductVersion = "The file """ & GetShortPath(fPN) & """ doesn't exist."
          oFP.Comments = "The file """ & GetShortPath(fPN) & """ doesn't exist."
           oFP.LegalTradeMarks = "The file """ & GetShortPath(fPN) & """ doesn't exist."
            oFP.PrivateBuild = "The file """ & GetShortPath(fPN) & """ doesn't exist."
             oFP.SpecialBuild = "The file """ & GetShortPath(fPN) & """ doesn't exist."
          Exit Function 'exit this procedure
 End If
   lngBufferlen = GetFileVersionInfoSize(fPN$, lngDummy)
   'GetFileVersionInfoSize function returns the size in bytes of the specified file's version information...
   'initializes lngDummy with a pointer to the return of the file's version resource handle
    If lngBufferlen > 0 Then
    'if the length of the version information is greater than one byte then...
       ReDim bytBuffer(lngBufferlen)
       'redimensionalize array with the appropriate number of elements...
       lngRc = GetFileVersionInfo(fPN$, 0&, lngBufferlen, bytBuffer(0))
       'GetFileVersionInfo function returns the version information of the specified file...
       If lngRc <> 0 Then
       'if the function was successfull then...
        lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
        'retrives the specified version resource block from the variable bytBuffer
        'lpSubBlock argument specifies the version information value to return
        '\ Specifies the root block,
        '\VarFileInfo\Translation specifies the translation array, function retrieves a pointer to an array of language and code page identifiers
        'the identifiers returned will be used to access the language StringTable structure
         If lngRc <> 0 Then
         'if the function returned non zero then it was successful...
          MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
          'move the memory variable to which (Long Pointer to Memory)lngVerPointer is referring to bytBuff array
           lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
            strLangCharset = Hex(lngHexNumber)
             Do While Len(strLangCharset) < 8
              strLangCharset = "0" & strLangCharset
             Loop
             'concenate the language identifier and the codepage
             'for the table of string identifiers see below:
              strVersionInfo(0) = "CompanyName"
               strVersionInfo(1) = "FileDescription"
                strVersionInfo(2) = "FileVersion"
                 strVersionInfo(3) = "InternalName"
                  strVersionInfo(4) = "LegalCopyright"
                   strVersionInfo(5) = "OriginalFileName"
                    strVersionInfo(6) = "ProductName"
                     strVersionInfo(7) = "ProductVersion"
                      strVersionInfo(8) = "Comments"
                       strVersionInfo(9) = "LegalTrademarks"
                        strVersionInfo(10) = "PrivateBuild"
                         strVersionInfo(11) = "SpecialBuild"
                            For i = 0 To 11
                            'For Next loop; loops until i = 11(number of elements in strVersionInfo array;number of version string id's)
                             strBuffer = String$(255, 0) 'allocate memory to buffer
                              strTemp = "\StringFileInfo\" & strLangCharset & "\" & strVersionInfo(i)
                              'build query...
                               lngRc = VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, lngBufferlen)
                               'call version query value to query the specified information from the files resource
                               'lngVerPointer receives the pointer to the value returned
                                If lngRc <> 0 Then
                                'If function was successful then...
                                 lstrcpy strBuffer, lngVerPointer
                                 'copy the memory to which lpString2 argument points to the strBuffer
                                  strBuffer = Mid$(strBuffer, 1, InStr(strBuffer, Chr(0)) - 1)
                                  'Remove the null terminating character
                                   strVersionInfo(i) = strBuffer
                                   'initialize the arrays element with the version information...
                                 Else
                                  strVersionInfo(i) = "This property could not be retreived."
                                  'the function was unsuccessful, set the elements value
                                End If
                            Next i 'increment i
          End If
       End If
    End If
     For i = 0 To 11
     'for next loop; loop until i evaluates to 11 incrementing i by one each iteration
      If Trim(strVersionInfo(i)) = "" Then strVersionInfo(i) = "This property could not be retreived."
      'if the element wasn't initialized set its value to "This property could not be retreived."
     Next i 'increment i by one
      oFP.CompanyName = strVersionInfo(0)
       oFP.FileDescription = strVersionInfo(1)
        oFP.FileVersion = strVersionInfo(2)
         oFP.InternalName = strVersionInfo(3)
          oFP.LegalCopyright = strVersionInfo(4)
           oFP.OrigionalFileName = strVersionInfo(5)
            oFP.ProductName = strVersionInfo(6)
             oFP.ProductVersion = strVersionInfo(7)
              oFP.Comments = strVersionInfo(8)
               oFP.LegalTradeMarks = strVersionInfo(9)
                oFP.PrivateBuild = strVersionInfo(10)
                 oFP.SpecialBuild = strVersionInfo(11)
                 'initialize oFP members...
End Function

'+-------------------------+-------------------------+
'|  SubLanguage ID  |   Primary Language ID   |
'+-------------------------+-------------------------+
'15                    10  9                         0

'0x0000 Language Neutral
'0x007f The language for the invariant locale (LOCALE_INVARIANT). See MAKELCID.
'0x0400 Process or User Default Language
'0x0800 System Default Language
'0x0436 Afrikaans
'0x041c Albanian
'0x0401 Arabic (Saudi Arabia)
'0x0801 Arabic(Iraq)
'0x0c01 Arabic(Egypt)
'0x1001 Arabic(Libya)
'0x1401 Arabic(Algeria)
'0x1801 Arabic(Morocco)
'0x1c01 Arabic(Tunisia)
'0x2001 Arabic(Oman)
'0x2401 Arabic(Yemen)
'0x2801 Arabic(Syria)
'0x2c01 Arabic(Jordan)
'0x3001 Arabic(Lebanon)
'0x3401 Arabic(Kuwait)
'0x3801 Arabic (U.A.E.)
'0x3c01 Arabic(Bahrain)
'0x4001 Arabic(Qatar)
'0x042b Windows 2000/XP: Armenian. This is Unicode only.
'0x042c Azeri(Latin)
'0x082c Azeri(Cyrillic)
'0x042d Basque
'0x0423 Belarusian
'0x0402 Bulgarian
'0x0455 Burmese
'0x0403 Catalan
'0x0404 Chinese(Taiwan)
'0x0804 Chinese(PRC)
'0x0c04 Chinese (Hong Kong SAR, PRC)
'0x1004 Chinese(Singapore)
'0x1404 Windows 98/Me, Windows 2000/XP: Chinese (Macau SAR)
'0x041a Croatian
'0x0405 Czech
'0x0406 Danish
'0x0465 Windows XP: Divehi. This is Unicode only.
'0x0413 Dutch(Netherlands)
'0x0813 Dutch(Belgium)
'0x0409 English (United States)
'0x0809 English (United Kingdom)
'0x0c09 English(Australian)
'0x1009 English(Canadian)
'0x1409 English(New Zealand)
'0x1809 English(Ireland)
'0x1c09 English (South Africa)
'0x2009 English(Jamaica)
'0x2409 English(Caribbean)
'0x2809 English(Belize)
'0x2c09 English(Trinidad)
'0x3009 Windows 98/Me, Windows 2000/XP: English (Zimbabwe)
'0x3409 Windows 98/Me, Windows 2000/XP: English (Philippines)
'0x0425 Estonian
'0x0438 Faeroese
'0x0429 Farsi
'0x040b Finnish
'0x040c French(Standard)
'0x080c French(Belgian)
'0x0c0c French(Canadian)
'0x100c French(Switzerland)
'0x140c French(Luxembourg)
'0x180c Windows 98/Me, Windows 2000/XP: French (Monaco)
'0x0456 Windows XP: Galician
'0x0437 Windows 2000/XP: Georgian. This is Unicode only.
'0x0407 German(Standard)
'0x0807 German(Switzerland)
'0x0c07 German(Austria)
'0x1007 German(Luxembourg)
'0x1407 German(Liechtenstein)
'0x0408 Greek
'0x0447 Windows XP: Gujarati. This is Unicode only.
'0x040d Hebrew
'0x0439 Windows 2000/XP: Hindi. This is Unicode only.
'0x040e Hungarian
'0x040f Icelandic
'0x0421 Indonesian
'0x0410 Italian(Standard)
'0x0810 Italian(Switzerland)
'0x0411 Japanese
'0x044b Windows XP: Kannada. This is Unicode only.
'0x0457 Windows 2000/XP: Konkani. This is Unicode only.
'0x0412 Korean
'0x0812 Windows 95, Windows NT 4.0 only: Korean (Johab)
'0x0440 Windows XP: Kyrgyz.
'0x0426 Latvian
'0x0427 Lithuanian
'0x0827 Windows 98 only: Lithuanian (Classic)
'0x042f FYRO Macedonian
'0x043e Malay(Malaysian)
'0x083e Malay (Brunei Darussalam)
'0x044e Windows 2000/XP: Marathi. This is Unicode only.
'0x0450 Windows XP: Mongolian
'0x0414 Norwegian(Bokmal)
'0x0814 Norwegian(Nynorsk)
'0x0415 Polish
'0x0416 Portuguese(Brazil)
'0x0816 Portuguese(Portugal)
'0x0446 Windows XP: Punjabi. This is Unicode only.
'0x0418 Romanian
'0x0419 Russian
'0x044f Windows 2000/XP: Sanskrit. This is Unicode only.
'0x0c1a Serbian(Cyrillic)
'0x081a Serbian(Latin)
'0x041b Slovak
'0x0424 Slovenian
'0x040a Spanish (Spain, Traditional Sort)
'0x080a Spanish(Mexican)
'0x0c0a Spanish (Spain, Modern Sort)
'0x100a Spanish(Guatemala)
'0x140a Spanish (Costa Rica)
'0x180a Spanish(Panama)
'0x1c0a Spanish (Dominican Republic)
'0x200a Spanish(Venezuela)
'0x240a Spanish(Colombia)
'0x280a Spanish(Peru)
'0x2c0a Spanish(Argentina)
'0x300a Spanish(Ecuador)
'0x340a Spanish(Chile)
'0x380a Spanish(Uruguay)
'0x3c0a Spanish(Paraguay)
'0x400a Spanish(Bolivia)
'0x440a Spanish (El Salvador)
'0x480a Spanish(Honduras)
'0x4c0a Spanish(Nicaragua)
'0x500a Spanish (Puerto Rico)
'0x0430 Sutu
'0x0441 Swahili(Kenya)
'0x041d Swedish
'0x081d Swedish(Finland)
'0x045a Windows XP: Syriac. This is Unicode only.
'0x0449 Windows 2000/XP: Tamil. This is Unicode only.
'0x0444 Tatar(Tatarstan)
'0x044a Windows XP: Telugu. This is Unicode only.
'0x041e Thai
'0x041f Turkish
'0x0422 Ukrainian
'0x0420 Windows 98/Me, Windows 2000/XP: Urdu (Pakistan)
'0x0820 Urdu(India)
'0x0443 Uzbek(Latin)
'0x0843 Uzbek(Cyrillic)
'0x042a Windows 98/Me, Windows NT 4.0 and later: Vietnamese
