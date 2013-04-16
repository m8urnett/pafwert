Attribute VB_Name = "modMain"
#If Pro Then
   Public Const EDITION = "Pro Edition"
#Else
   Public Const EDITION = "Personal Edition"
#End If


Private Const MODULE_SOURCE = "Main"

'CSEH: Pafwert Components
'--------------------------------------------------------------------------------
'    Component:   modMain
'    Filename:    Main.bas
'    Project:     Pafwert
'
'    Description: Common global functions
'
'    Change History:
'
'--------------------------------------------------------------------------------
'    Copyright 2001-2013 Mark Burnett (mb@xato.net)
'
'    Licensed under the Apache License, Version 2.0 (the "License");
'    you may not use this file except in compliance with the License.
'    You may obtain a copy of the License at
'
'    http://www.apache.org/licenses/LICENSE-2.0
'
'    Unless required by applicable law or agreed to in writing, software
'    distributed under the License is distributed on an "AS IS" BASIS,
'    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'    See the License for the specific language governing permissions and
'    limitations under the License.
'
'--------------------------------------------------------------------------------
'
' Future Enhancements
'
'--------------------------------------------------------------------------------
' Options and Compiler Directives
'--------------------------------------------------------------------------------
Option Explicit

'--------------------------------------------------------------------------------
' Declarations
'--------------------------------------------------------------------------------
Public Const ERRSOURCE = ""
Public Const ERRBASE As Long = 8192
Public Const ERR_MISC As Long = vbObjectError + ERRBASE + &H1
Public Const ERR_MISC_DESC As String = "Error parsing pattern"
Public Const ERR_INVALID_PATTERN As Long = vbObjectError + ERRBASE + &H2
Public Const ERR_INVALID_PATTERN_DESC = "Invalid pattern"
Public Const ERR_OPEN_PATTERNS_FILE As Long = vbObjectError + ERRBASE + &H4
Public Const ERR_OPEN_PATTERNS_FILE_DESC As String = "Error opening patterns.cfg"
Public Const ERR_FILE_NOT_FOUND As Long = vbObjectError + ERRBASE + &H8
Public Const ERR_FILE_NOT_FOUND_DESC As String = "Wordlist file not found"
Public Const ERR_INVALID_FILENAME As Long = vbObjectError + ERRBASE + &H10
Public Const ERR_INVALID_FILENAME_DESC As String = "Invalid wordlist filename"
Public Const ERR_WORDLIST_DIR_NOT_FOUND As Long = vbObjectError + ERRBASE + &H12
Public Const ERR_WORDLIST_DIR_NOT_FOUND_DESC As String = "Wordlist directory not found"
Public Const ERR_LOADING_PATTERNS As Long = vbObjectError + ERRBASE + &H14
Public Const ERR_LOADING_PATTERNS_DESC As String = "Error loading patterns"
Public Const ERR_WORDLIST_READ As Long = vbObjectError + ERRBASE + &H16
Public Const ERR_WORDLIST_READ_DESC As String = "Error loading wordlist"
Public Const ERR_PATTERN_MISMATCHED_BRACES As Long = vbObjectError + ERRBASE + &H18
Public Const ERR_PATTERN_MISMATCHED_BRACES_DESC As String = "Mismatched braces in pattern"
Public Const ERR_INVALID_MODIFIER As Long = vbObjectError + ERRBASE + &H20
Public Const ERR_INVALID_MODIFIER_DESC As String = "Invalid modifier"

Private lHCryptprov As Long

Private m_bUseRandAPI As Boolean

'--- Profiling stuff
Public Declare Function timeGetTime _
               Lib "winmm.dll" () As Long

Public Declare Function timeBeginPeriod _
               Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Public Declare Function timeEndPeriod _
               Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Public m_lT      As Long

Public Enum PfCharsets
   pfUcaseLetters = &H1
   pfLowerCase = &H2
   pfNumbers = &H4
   pfSymbols = &H8
   pfSpaces = &H10
End Enum
'--- PassCheck Error and Warning Constants
'--- API Declarations
Declare Function CharLower _
        Lib "user32" _
        Alias "CharLowerA" (ByVal lpsz As String) As String
Declare Function CharUpper _
        Lib "user32" _
        Alias "CharUpperA" (ByVal lpsz As String) As String

Private Declare Function GetLongPathName _
                Lib "kernel32" _
                Alias "GetLongPathNameA" (ByVal lpszShortPath As String, _
                                          ByVal lpszLongPath As String, _
                                          ByVal cchBuffer As Long) As Long
Const MAX_PATH = 255
Declare Function GetFullPathName _
        Lib "kernel32.dll" _
        Alias "GetFullPathNameA" (ByVal lpFileName As String, _
                                  ByVal nBufferLength As Long, _
                                  ByVal lpBuffer As String, _
                                  ByVal lpFilePart As String) As Long

Public Declare Function CryptAcquireContext _
               Lib "advapi32.dll" _
               Alias "CryptAcquireContextA" (phProv As Long, _
                                             pszContainer As String, _
                                             pszProvider As String, _
                                             ByVal dwProvType As Long, _
                                             ByVal dwFlags As Long) As Long

Public Declare Function CryptGenRandom _
               Lib "advapi32.dll" (ByVal hProv As Long, _
                                   ByVal dwLen As Long, _
                                   ByVal pbBuffer As String) As Long
Declare Function GetLastError _
        Lib "kernel32.dll" () As Long

'--- API Constants
Public Const MS_DEF_PROV As String = "Microsoft Base Cryptographic Provider v1.0"
Public Const MS_ENH_PROV As String = "Microsoft Enhanced Cryptographic Provider v1.0"
Public Const MS_STR_PROV As String = "Microsoft Strong Cryptographic Provider"
Public Const PROV_RSA_FULL As Long = 1
Public Const CRYPT_VERIFYCONTEXT As Long = &HF0000000

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       Bracket
' Description:       Brackets a word or phrase between random bracket symbols
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:16:45
'
' Parameters :       Word (String)
'                    sBracketList (String)
'--------------------------------------------------------------------------------
'</CSCM>
'CSEH: ErrRaise
Public Function Bracket(ByVal Word As String, _
                        Optional ByVal sBracketList As String) As String
      '<EhHeader>
      On Error GoTo Bracket_Err
      '</EhHeader>

      Dim sBrackets()      As String
      Dim x                As Long
      If Len(sBracketList) = 0 Then sBracketList = "[ ] < > ( ) ( ) ( ) ( ) ( ) ( ) ( ) ( ) [ ] [ ] | | \ / * * [ ] { } / / \ / / \ \ \ <- -> -> <-"
100   sBrackets() = Split(sBracketList, " ")
102   x = Int(Rand(UBound(sBrackets()) / 2)) * 2
104   Bracket = sBrackets(x) & Word & sBrackets(x + 1)

      '<EhFooter>
      Exit Function

Bracket_Err:
      Err.Raise vbObjectError + 100, "PatternBuilder.modMain.Bracket", "modMain component failure"
      '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       Chance
' Description:       Returns a boolean if the PercentChance randomly occurs
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       PercentChance (Long)
'                    Weight (Long = 1)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Chance(ByVal PercentChance As Long, _
                       Optional ByVal Weight As Long = 1) As Boolean
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".Chance"

         '</EhHeader>

12100    Chance = Rand(100, 1, Weight) <= PercentChance

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       EndTiming
' Description:       Ends the high resolution timer.
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function EndTiming() As Long
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".EndTiming"

         '</EhHeader>

12100    EndTiming = timeGetTime() - m_lT
12110    timeEndPeriod 1

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       Fakeword
' Description:       Returns a fake but pronounceable word
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Word (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Fakeword(ByVal Word As String)
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".Fakeword"

         '</EhHeader>

12100    Dim sPrefix As String
12110    Dim sSuffix As String
12120    Dim sNewWord As String
        
12130    sPrefix = PickOne("a ab abso aca acri admini alpha ambi ana ant ante anti apro aqua archi astro atmo audi auto be bene beta beva bi bio centa chrono circum co co- com con contra counter credo cryo cyber cyclo de deca demo dextro di dia dicto dis double- duo dyna dyno dys e ecto ef endo entre equi euro every ex exo extra fa fan fict fiz flo fore fun gag gamma gap geo gig giga glyco goo gyro he hemi hetero hexa his holo homeo homo hosp hu hydro hyper hypo id identi ig imi in info infra int inter intra intro iso kilo kno la lacto li longi luma ma macro magni mali mega meso meta micro milli mini mis mono multi nano navi neo non non- novi octa octo omni otco over oxy pan para peda penta per peri philo phoni phono physi pico poly post pre pre- pro proto quad re retro sancti semi septo similli steno sub super supra synchro tele tera tetra thermo trans tre tri ultra un una under uni uno vario vita xantho xero")
12140    sSuffix = PickOne("able ad aero alooza any ation be bi bio cate cede ceed cess eting fest fy gram graph iac ible ify ing ism ist ity ize log logue logy maniac ment meter metry ogram ograph oid ology ometer opath opsy osity phile phobe phobia phone super tion tious ty")

12150    If Rand(100, 1) <= 20 Then
12160       If InStr(sPrefix, "-") = 0 Then sPrefix = sPrefix & "-"
12170    End If

         'TODO: If Prexif ends in vowel and word begins with vowell...
12180    Select Case Rand(5, 1)

            Case 1: sNewWord = sPrefix & Word & sSuffix

12190       Case 2: sNewWord = Word & sSuffix

12200       Case Else: sNewWord = sPrefix & Word
12210    End Select

12220    sNewWord = Replace(sNewWord, "aa", "a")
12230    sNewWord = Replace(sNewWord, "ii", "i")
12240    sNewWord = Replace(sNewWord, "hh", "h")
12250    sNewWord = Replace(sNewWord, "jj", "j")
12260    sNewWord = Replace(sNewWord, "kk", "k")
12270    sNewWord = Replace(sNewWord, "qq", "q")
12280    sNewWord = Replace(sNewWord, "uu", "u")
12290    sNewWord = Replace(sNewWord, "ww", "w")
12300    sNewWord = Replace(sNewWord, "xx", "x")
12310    sNewWord = Replace(sNewWord, "yy", "y")
12320    sNewWord = Replace(sNewWord, "zz", "z")
12330    sNewWord = Replace(sNewWord, "eae", "ae")
12340    Fakeword = sNewWord

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function
'
'Sub ListPasswordPolicyInfo(strDomain)
'
'      Dim objComputer
'    Set objComputer = GetObject("WinNT://" & strDomain)
'    WScript.Echo "MinPasswordAge: " & ((objComputer.MinPasswordAge) / 86400)
'    WScript.Echo "MinPasswordLength: " & objComputer.MinPasswordLength
'    WScript.Echo "PasswordHistoryLength: " & objComputer.PasswordHistoryLength
'    WScript.Echo "AutoUnlockInterval: " & objComputer.AutoUnlockInterval
'    WScript.Echo "LockOutObservationInterval: " & objComputer.LockOutObservationInterval
'End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       GetFullPath
' Description:       Returns the full path name
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       FullPath (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetFullPath(ByVal FullPath As String) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".GetFullPath"

         '</EhHeader>

12100    Dim lLen         As Long
12110    Dim sBuffer      As String
        
12120    sBuffer = Space$(MAX_PATH)
12130    lLen = GetFullPathName(FullPath, MAX_PATH, sBuffer, "")

12140    If lLen > 0 And Err.Number = 0 Then
12150       GetFullPath = Left$(sBuffer, lLen)
12160    End If

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       GetLongFileName
' Description:       Returns the long file name of a path.
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       FullPath (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetLongFileName(ByVal FullPath As String) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".GetLongFileName"

         '</EhHeader>

12100    Dim lLen         As Long
12110    Dim sBuffer      As String
12120    sBuffer = String$(MAX_PATH, 0)
12130    lLen = GetLongPathName(FullPath, sBuffer, Len(sBuffer))

12140    If lLen > 0 And Err.Number = 0 Then
12150       GetLongFileName = Left$(sBuffer, lLen)
12160    End If

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       GetNumberPattern
' Description:       Returns a random but patterned series of numbers
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Length (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetNumberPattern(Optional ByVal Length As Long) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".GetNumberPattern"

         '</EhHeader>

12100    Dim i             As Long
12110    Dim iSaved        As Long
12120    Dim iDigit()      As String

12130    If Length = 0 Then Length = 3
12140    ReDim iDigit(Length)
12150    iDigit(1) = Rand(9)

12160    For i = 2 To Length

12170       Select Case Rand(3, 0)

               Case 0:  iDigit(i) = Rand(9)

12180          Case 1:  iDigit(i) = iDigit(Rand(i))

12190          Case 2:  If Val(iDigit(i - 1)) > 1 Then iDigit(i) = Val(iDigit(i - 1)) - 1

12200          Case 3:  If Val(iDigit(i - 1)) < 9 Then iDigit(i) = Val(iDigit(i - 1)) + 1
12210       End Select

12220    Next i

12230    GetNumberPattern = Format(Join(iDigit(), ""), String$(Length, "0"))

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       GetOrdinal
' Description:       Returns the ordnial words for a number
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Num (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetOrdinal(ByVal Num As Long) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".GetOrdinal"

         '</EhHeader>

12100    Dim n      As String
12110    n = CStr(Num)

12120    Select Case Right$(n, 2)

            Case "11", "12", "13"
12130          GetOrdinal = n & "th"

12140       Case Else

12150          Select Case Right$(n, 1)

                  Case "0", "4" To "9"
12160                GetOrdinal = n & "th"

12170             Case "1"
12180                GetOrdinal = n & "st"

12190             Case "2"
12200                GetOrdinal = n & "nd"

12210             Case "3"
12220                GetOrdinal = n & "rd"
12230          End Select
12240    End Select

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       GetPhonetic
' Description:       Returns the phonetic description of a letter
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       sWord (String)
'                    Style (Long = 1)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetPhonetic(sWord As String, _
                             Optional Style As Long = 1)
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".GetPhonetic"

         '</EhHeader>

12100    Dim sWords()        As String
12110    Dim sPhonetic      As String
12120    Dim iLen            As Long
12130    Dim i               As Integer
        
12140    On Error Resume Next

12150    If Style = 1 Or Style = 0 Then
12160       sWords() = Split("Alpha Bravo Charlie Delta Echo Foxtrot Golf Hotel India Juliet Kilo Lima Mike November Oscar Papa Quebec Romeo Sierra Tango Uniform Victor Whiskey X-Ray Yankee Zulu", " ")
12170    Else
12180       sWords() = Split("Adam Baker Charles David Edward Frank George Henry Ida John King Lincoln Mary Nora Ocean Paul Queen Robert Sam Tom Union Victor Wililiam X-Ray Young Zebra")
12190    End If

12200    iLen = Len(sWord)

12210    For i = 1 To iLen
12220       sPhonetic = sPhonetic & sWords(Asc(UCase$(Mid$(sWord, i, 1))) - 65) & " "
12230    Next

12240    GetPhonetic = Trim$(sPhonetic)

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       GetSequence
' Description:       Returns a random sequence of letters or numbers
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Length (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function GetSequence(ByVal Length As Long) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".GetSequence"

         '</EhHeader>

12100    Dim sLetters      As String
12110    Dim sNumbers      As String
12120    Dim sKey1         As String
12130    Dim sKey2         As String
12140    Dim sKey3         As String
12150    Dim sKey4         As String
12160    Dim sKey5         As String
12170    Dim sKey6         As String
12180    Dim sSeq          As String
12190    Dim i             As Long

12200    sLetters = "abcdefghijklmnopqrstuvwxyz"
12210    sNumbers = "1234567890"
12220    sKey1 = "qwertyuiop"
12230    sKey2 = "asdfghjkl"
12240    sKey3 = "zxcvbnm"
12250    sKey4 = "poiuytrewq"
12260    sKey5 = "lkjhgfdsa"
12270    sKey6 = "mnbvcxz"

12280    On Error Resume Next

12290    If Length = 0 Then Length = 3

12300    Select Case Rand(19)

            Case 1: sSeq = Mid$(sLetters, Int(Rnd * (Len(sLetters) - Length) + 1), Length)

12310       Case 2: sSeq = Mid$(sNumbers, Int(Rnd * (Len(sNumbers) - Length) + 1), Length)

12320       Case 3: sSeq = Mid$(sKey1, Int(Rnd * (Len(sKey1) - Length) + 1), Length)

12330       Case 4: sSeq = Mid$(sKey2, Int(Rnd * (Len(sKey2) - Length) + 1), Length)

12340       Case 5: sSeq = Mid$(sKey3, Int(Rnd * (Len(sKey3) - Length) + 1), Length)

12350       Case 6: sSeq = Mid$(sKey4, Int(Rnd * (Len(sKey4) - Length) + 1), Length)

12360       Case 7: sSeq = Mid$(sKey5, Int(Rnd * (Len(sKey5) - Length) + 1), Length)

12370       Case 8: sSeq = Mid$(sKey6, Int(Rnd * (Len(sKey6) - Length) + 1), Length)

12380       Case 9
12390          i = Int(Rnd * 7) + 1
12400          sSeq = Mid$(sKey1, i, Length / 3) & Mid$(sKey2, i, Length / 3) & Mid$(sKey3, i, Length / 3)

12410       Case 10
12420          i = Int(Rnd * 7) + 1
12430          sSeq = Mid$(sKey3, i, Length / 3) & Mid$(sKey2, i, Length / 3) & Mid$(sKey1, i, Length / 3)

12440       Case 11
12450          i = Int(Rnd * 9) + 1

12460          Do
12470             sSeq = sSeq & Mid$(sKey1, i, 1) & Mid$(sKey2, i, 1)
12480          Loop Until Len(sSeq) >= Length

12490       Case 12
12500          i = Int(Rnd * 7) + 1

12510          Do
12520             sSeq = sSeq & Mid$(sKey2, i, 1) & Mid$(sKey3, i, 1)
12530          Loop Until Len(sSeq) >= Length

12540       Case 12
12550          i = Int(Rnd * 7) + 1

12560          Do
12570             sSeq = sSeq & Mid$(sKey3, i, 1) & Mid$(sKey2, i, 1)
12580          Loop Until Len(sSeq) >= Length

12590       Case 13
12600          i = Int(Rnd * 9) + 1

12610          Do
12620             sSeq = sSeq & Mid$(sKey2, i, 1) & Mid$(sKey1, i, 1)
12630          Loop Until Len(sSeq) >= Length

12640       Case 14
12650          i = Int(Rnd * 10) + 1

12660          Do
12670             sSeq = sSeq & Mid$(sKey1, i, 1) & Mid$(sNumbers, i, 1)
12680          Loop Until Len(sSeq) >= Length

12690       Case 15
12700          i = Int(Rnd * 9) + 1

12710          Do
12720             sSeq = sSeq & Mid$(sKey4, i, 1) & Mid$(sKey5, i, 1)
12730          Loop Until Len(sSeq) >= Length

12740       Case 16
12750          i = Int(Rnd * 7) + 1

12760          Do
12770             sSeq = sSeq & Mid$(sKey5, i, 1) & Mid$(sKey6, i, 1)
12780          Loop Until Len(sSeq) >= Length

12790       Case 17
12800          i = Int(Rnd * 10) + 1

12810          Do
12820             sSeq = sSeq & Mid$(sKey1, i, 1) & Mid$(sKey1, 10 - i + 1, 1)
12830          Loop Until Len(sSeq) >= Length

12840       Case 18
12850          i = Int(Rnd * 9) + 1

12860          Do
12870             sSeq = sSeq & Mid$(sKey2, i, 1) & Mid$(sKey2, 9 - i + 1, 1)
12880          Loop Until Len(sSeq) >= Length

12890       Case Else
12900          i = Int(Rnd * 7) + 1

12910          Do
12920             sSeq = sSeq & Mid$(sKey3, i, 1) & Mid$(sKey3, 7 - i + 1, 1)
12930          Loop Until Len(sSeq) >= Length

12940    End Select

12950    GetSequence = Left$(sSeq, Length)

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       IsBounded
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       vntArray (Variant)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function IsBounded(vntArray As Variant) As Boolean
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".IsBounded"

         '</EhHeader>
12100    On Error Resume Next
12110    IsBounded = IsNumeric(UBound(vntArray))
         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       MakeComplex
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Password (String)
'                    MinLen (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function MakeComplex(ByVal Password As String, _
                            ByVal MinLen As Long) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".MakeComplex"

         '</EhHeader>

12100    Do Until Len(Password) >= MinLen

12110       If Len(Password) <= MinLen / 2 Then

               'If its really short
12120          Select Case Rand(11, 1)

                  Case 1: Password = Password & " " & Password

12130             Case 2: Password = Password & " " & StrReverse(Password)

12140             Case 3: Password = Password & " " & PronounceableWord()

12150             Case 4: Password = Password & " " & GetSequence(Rand(MinLen, 1, -1))

12160             Case 5: Password = CStr(GetNumberPattern(CLng(MinLen / 2))) & " " & Password

12170             Case 6: Password = PickOne(LONGMONTHS) & " " & Password

12180             Case 7: Password = PickOne(TWOLETTERWORDS) & " " & Password

12190             Case 8: Password = PickOne(SHORTDAYS) & " " & Password

12200             Case 9: Password = Password & " " & PickOne(THREELETTERWORDS)

12210             Case 10: Password = PronounceableWord() & " " & Password

12220             Case 11: Password = Password & Left$(Password, Rand(5, 1))
12230          End Select

12240       Else

12250          Select Case Rand(9, 1)

                  Case 1: Password = Bracket(Password)

12260             Case 2: Password = Password & NumberCode()

12270             Case 3: Password = Password & " " & PickCharacter(NUMROWFULL)

12280             Case 4: Password = Password & PickOne(DIGRAPHS)

12290             Case 5: Password = Password & PickOne(TWOLETTERWORDS)

12300             Case 6: Password = PickOne(THREELETTERWORDS) & " " & Password

12310             Case 7: Password = Password & PickOne(SMILEYS)

12320             Case 8: Password = PickOne(DIGRAPHS) & " " & Password

12330             Case 9: Password = Left$(Password, 2) & Password
12340          End Select

12350       End If

12360    Loop

12370    If Rand(100) < 10 Then Password = RandomCase(Password)
12380    If Rand(100) < 10 Then Password = Obscure(Password)
12390    MakeComplex = Password

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       NumberCode
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function NumberCode() As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".NumberCode"

         '</EhHeader>
12100    Dim x                 As Long
12110    Dim sCode             As String
12120    Dim sDelim            As String
12130    Dim lRepeatDigit      As Long
12140    lRepeatDigit = Rand(9, 0)
12150    sDelim = PickOne("- - - - - - - - . . . , / \ :")

12160    Do
12170       x = Rand(9, 0)

12180       Do
12190          sCode = sCode & x

12200          If Chance(30) Then
12210             sCode = sCode & lRepeatDigit
12220          ElseIf Chance(40) Then
12230             sCode = sCode & sDelim
12240          End If

12250          If Len(sCode) > 2 Then Exit Do
12260       Loop While Chance(30)

12270       If Len(sCode) > Rand(4, 3) Then Exit Do
12280       If Chance(10) Then sCode = Bracket(sCode)
12290    Loop Until Chance(15) And Len(sCode) > 2

12300    If IsNumeric(Right$(sCode, 1)) = False Then
12310       sCode = Left$(sCode, Len(sCode) - 1)
12320    End If

12330    NumberCode = sCode
         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       Obscure
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Word (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Obscure(ByVal Word As String) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".Obscure"

         '</EhHeader>
12100    Dim sWord      As String
12110    Dim i          As Long
12120    sWord = Word

12130    For i = 1 To Rand(20, 2, 2) 'We will give it enough chances (at least 2, up to 20) to come up with something

12140       Select Case Rand(120)

               Case 0  'Do nothing

12150          Case 1: sWord = Replace(sWord, "ate", "8")

12160          Case 2: sWord = Replace(sWord, "for", "4")

12170          Case 3: sWord = Replace(sWord, "e", "3")

12180          Case 4: sWord = Replace(sWord, "l", "1")  '1l?

12190          Case 5: sWord = Replace(sWord, "s", "z")

12200          Case 6: sWord = Replace(sWord, "o", "0")

12210          Case 7: sWord = Replace(sWord, "a", "@")

12220          Case 8: sWord = Replace(sWord, "s", "$")

12230          Case 9: sWord = Replace(sWord, "l", "|")

12240          Case 10: sWord = Replace(sWord, "ait", "8")

12250          Case 11: sWord = Replace(sWord, "a", "")

12260          Case 12: sWord = Replace(sWord, "e", "")

12270          Case 13: sWord = Replace(sWord, "ou", "u")

12280          Case 14: sWord = Replace(sWord, "cc", "x")

12290          Case 15: sWord = Replace(sWord, "oo", "ew")

12300          Case 16: sWord = Replace(sWord, "and", "&")

12310          Case 17: sWord = Replace(sWord, "are", "r")

12320          Case 18: sWord = Replace(sWord, "ks", "x")

12330          Case 19: sWord = Replace(sWord, "f", "ph")

12340          Case 20: sWord = Replace(sWord, "ph", "f")

12350          Case 21: sWord = Replace(sWord, "won", "1")

12360          Case 22: sWord = Replace(sWord, "l", "r") ' oriental accent

12370          Case 23: sWord = Replace(sWord, "ee", "eee")

12380          Case 24: sWord = Replace(sWord, "000", "k")

12390          Case 25: sWord = Replace(sWord, "er", "r")

12400          Case 26: sWord = Replace(sWord, "ex", "x")

12410          Case 27: sWord = Replace(sWord, "ecs", "x")

12420          Case 28: sWord = Replace(sWord, "m", "mm")

12430          Case 29: sWord = Replace(sWord, "cke", "x0")

12440          Case 30: sWord = Replace(sWord, "qu", "kw")

12450          Case 31: sWord = Replace(sWord, "a", "'")

12460          Case 32: sWord = Replace(sWord, "u", "'")

12470          Case 33: sWord = Replace(sWord, "ei", "ee")

12480          Case 34: sWord = Replace(sWord, "one", "own")

12490          Case 35: sWord = Replace(sWord, "oi", "oy")

12500          Case 36: sWord = Replace(sWord, "om", "um")

12510          Case 37: sWord = Replace(sWord, "a", "aa")

12520          Case 38: sWord = Replace(sWord, "ew", "u")

12530          Case 39: sWord = Replace(sWord, "us", "is")

12540          Case 40: sWord = Replace(sWord, "y", "ee")

12550          Case 41: sWord = Replace(sWord, "sh", "ch")

12560          Case 42: sWord = Replace(sWord, "to", "2")

12570          Case 43: sWord = Replace(sWord, "s", "th") 'lisp

12580          Case 44: sWord = Replace(sWord, "ck", "q")

12590          Case 45: sWord = Replace(sWord, "ci", "si")

12600          Case 46: sWord = Replace(sWord, "ie", "iye")

12610          Case 47: sWord = Replace(sWord, "tion", "shun")

12620          Case 48: sWord = Replace(sWord, "r", "w") 'Elmer Fudd

12630          Case 49: sWord = Replace(sWord, "come", "cum")

12640          Case 50: sWord = Replace(sWord, "cks", "x")

12650          Case 51: sWord = Replace(sWord, "ight", "ite")

12660          Case 52: sWord = Replace(sWord, "ing", "'n")

12670          Case 53: sWord = Replace(sWord, "th", "f")

12680          Case 54: sWord = Replace(sWord, "tion", "shun")

12690          Case 55: sWord = Replace(sWord, "too", "2")

12700          Case 56: sWord = Replace(sWord, "why", "y")

12710          Case 57: sWord = Replace(sWord, "won", "1")

12720          Case 58: sWord = Replace(sWord, "your", "yor")

12730          Case 59: sWord = Replace(sWord, "sc", "sh")

12740          Case 60: sWord = Replace(sWord, "sh", "th")

12750          Case 61: sWord = Replace(sWord, "ly", "lee")

12760          Case 62: sWord = Replace(sWord, "er", "uh") 'Evuh

12770          Case 63: sWord = Replace(sWord, "er", "a") 'Gangsta

12780          Case 64: sWord = Replace(sWord, "the", "da")

12790          Case 65: sWord = Replace(sWord, "it is", "'tis")

12800          Case 65: sWord = Replace(sWord, "you", "ya")

12810          Case 66: sWord = Replace(sWord, "l", "w")

12820          Case 67: sWord = Replace(sWord, "th", "d")

12830          Case 68: sWord = Replace(sWord, "a", "u")

12840          Case 69: sWord = Replace(sWord, "th", "'")

12850          Case 70: sWord = Replace(sWord, "your", "yer")

12860          Case 71: sWord = Replace(sWord, "ned", "nt")

12870          Case 72: sWord = Replace(sWord, "e", "_")

12880          Case 73: sWord = Replace(sWord, "t", "+")

12890          Case 74: sWord = Replace(sWord, "e", "=")

12900          Case 75: sWord = Replace(sWord, "can", "kin")

12910          Case 76: sWord = Replace(sWord, "t", "'")

12920          Case 77: sWord = Replace(sWord, "ng", "n'")

12930          Case 78: sWord = Replace(sWord, "red", "hed")

12940          Case 79: sWord = Replace(sWord, "th", "d")

12950          Case 80: sWord = Replace(sWord, "he", "eh")

12960          Case 81: sWord = Replace(sWord, "h", "")

12970          Case 82: sWord = Replace(sWord, "f", "v")

12980          Case 83: sWord = Replace(sWord, "ha", "o")

12990          Case 84: sWord = Replace(sWord, "v", "f")

13000          Case 85: sWord = Replace(sWord, "v", "b")

13010          Case 86: sWord = Replace(sWord, "N", "|\|")

13020          Case 87: sWord = Replace(sWord, "ll", "dd")

13030          Case 88: sWord = Replace(sWord, "ll", "tt")

13040          Case 89: sWord = Replace(sWord, "dd", "tt")

13050          Case 90: sWord = Replace(sWord, "h", "'")

13060          Case 91: sWord = Replace(sWord, "o", "a")

13070          Case 92: sWord = Replace(sWord, "e", "a")

13080          Case 93: sWord = Replace(sWord, "a", "uh")

13090          Case 94: sWord = Replace(sWord, "a", "u")

13100          Case 95: sWord = Replace(sWord, "oo", "u")

13110          Case 96: sWord = Replace(sWord, "i", "ih")

13120          Case 97: sWord = Replace(sWord, "a ", "ah")

13130          Case 98: sWord = Replace(sWord, "s", "ss")

13140          Case 99: sWord = Replace(sWord, "t", "tt")

13150          Case 100: sWord = Replace(sWord, "d", "dd")

13160          Case 101: sWord = Replace(sWord, "at", "@")

13170          Case 102: sWord = Replace(sWord, " ", "")

                  'Not documented yet:
13180          Case 103: sWord = Replace(sWord, "with", "w/")

13190          Case 104: sWord = Replace(sWord, "t", "d")

13200          Case 105: sWord = Replace(sWord, "t", "dd")

13210          Case 106: sWord = Replace(sWord, "d", "t")

13220          Case 107: sWord = Replace(sWord, "d", "tt")

13230          Case 108: sWord = Replace(sWord, "cks", "x")

13240          Case 109: sWord = Replace(sWord, "er", "ah")

13250          Case 110 To 120: sWord = Replace(sWord, " ", String$(Rand(3, 1), PickOne("-.></:+=\")))
                  'Case 110: sWord = Replace(sWord, "er", "oh")
13260       End Select

            'If we have done some stuff already, randomly bail out
13270       If i >= 2 And sWord <> Word And Chance(75) Then Exit For
13280    Next i

13290    Obscure = sWord
         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       PickCharacter
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Characters (String)
'                    Weight (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function PickCharacter(ByRef Characters As String, _
                              Optional Weight As Long)
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".PickCharacter"

         '</EhHeader>
12100    PickCharacter = Mid$(Characters, (Rand(Len(Characters), 1, Weight)), 1)
         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       PickOne
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       sString (String)
'                    Weight (Long = 1)
'                    Delim (String = " ")
'--------------------------------------------------------------------------------
'</CSCM>
Public Function PickOne(ByRef sString As String, _
                        Optional ByRef Weight As Long = 1, _
                        Optional ByRef Delim As String = " ") As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".PickOne"

         '</EhHeader>
12100    Dim sList()      As String
12110    Dim i            As Long
12120    Dim x            As Long
         'On Error Resume Next
12130    sList = Split(sString, Delim)

12140    If UBound(sList) Then
12150       PickOne = Trim$(sList(Rand(UBound(sList), 0, Weight)))
12160    Else
12170       PickOne = Trim$(sList(0))
12180    End If

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       PigLatin
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Words (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function PigLatin(Words As String) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".PigLatin"

         '</EhHeader>
12100    Dim strConvert      As String
12110    Dim sWords()        As String
12120    Dim i               As Long
12130    sWords = Split(Words, " ")

12140    For i = 0 To UBound(sWords)
12150       strConvert = LCase(sWords(i))

12160       If Left$(sWords(i), 1) Like "[aeiou]" Then
12170          strConvert = sWords(i) & "yay"
12180       Else
12190          strConvert = Right(sWords(i), Len(sWords(i)) - 1) & Left(sWords(i), 1) & "ay"
12200       End If

            'If word passed in was capitalized, do the same with the converted
            'word
12210       If Left$(sWords(i), 1) = UCase(Left$(sWords(i), 1)) Then
12220          Mid$(strConvert, 1, 1) = UCase(Mid$(strConvert, 1, 1))
12230       End If

12240       sWords(i) = strConvert
12250    Next i

12260    PigLatin = Join(sWords, " ")
         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       PronounceableWord
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Function PronounceableWord() As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".PronounceableWord"

         '</EhHeader>
12100    Dim sWord           As String
12110    Dim bVowelNext      As Boolean
12120    Dim iLen            As Long
12130    Dim i As Long
   
12140    bVowelNext = Rand(1)

12150    For i = 1 To Rand(5, 4)
12160       iLen = Len(sWord)

12170       If bVowelNext Then
12180          If Rand(3) = 0 And iLen > 1 Then
12190             sWord = sWord & PickOne("ing ers ance ence le ness ings ment ize ate ive ute acy ous ify ought some edness ed es ly less ment able ible les led ious ant ary iety ist ism ial ate act ure iac ice aint ent ant ure ide ify les")
12200             Exit For
12210          End If

12220          sWord = sWord & PickOne(VOWELS2, 2)
12230       Else

12240          If Rand(3) = 0 And iLen Then
12250             sWord = sWord & PickOne("cked cker tor ter ly rer tic nst lyst onic ght nge nce zer cy ly ny lic dged red ate ndle ching tching lent ged zen ted nnial lic rly stic se les")
12260             Exit For
12270          Else

12280             If Rand(3) = 0 And iLen Then
12290                sWord = sWord & PickOne(CONSONANTS3)
12300             Else
12310                sWord = sWord & PickOne(CONSONANTS2, 2)
12320             End If

12330             If Right$(sWord, 1) = "t" Then
12340                If Rand(2) = 0 And iLen > 1 Then
12350                   sWord = sWord & PickOne("ion ity ient ment ance ly less ter tor")
12360                   Exit For
12370                End If
12380             End If
12390          End If
12400       End If

12410       bVowelNext = Not bVowelNext
12420    Next i

         'Some letters shouldn't ever be doubled
12430    If InStr(sWord, "aa") Then sWord = Replace(sWord, "aa", "a")
12440    If InStr(sWord, "hh") Then sWord = Replace(sWord, "hh", "h")
12450    If InStr(sWord, "ii") Then sWord = Replace(sWord, "ii", "i")
12460    If InStr(sWord, "jj") Then sWord = Replace(sWord, "jj", "j")
12470    If InStr(sWord, "ll") Then sWord = Replace(sWord, "kk", "k")
12480    If InStr(sWord, "qq") Then sWord = Replace(sWord, "qq", "qu")
12490    If InStr(sWord, "uu") Then sWord = Replace(sWord, "uu", "u")
12500    If InStr(sWord, "vv") Then sWord = Replace(sWord, "vv", "v")
12510    If InStr(sWord, "ww") Then sWord = Replace(sWord, "ww", "w")
12520    If InStr(sWord, "xx") Then sWord = Replace(sWord, "xx", "x")
12530    If InStr(sWord, "yy") Then sWord = Replace(sWord, "yy", "y")

         'i before e except after c
12540    If InStr(sWord, "cie") Then sWord = Replace(sWord, "cie", "cei")

         'Don't start a word with double letters
12550    If Left$(sWord, 1) = Mid$(sWord, 2, 1) Then sWord = Mid$(sWord, 2, Len(sWord))
12560    PronounceableWord = sWord
         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       Rand
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Max (Long)
'                    Min (Long = 0)
'                    Weight (Long = 1)
'                    DecimalPlaces (Long = 0)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Rand(Optional ByVal Max As Long, _
                     Optional ByVal Min As Long = 0, _
                     Optional ByVal Weight As Long = 1, _
                     Optional DecimalPlaces As Long = 0) As Variant
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".Rand"

         '</EhHeader>
12100    Dim lCryptLength      As Long
12110    Dim sCryptBuffer      As String
12120    Dim i                 As Long
12130    Dim nRnd              As Single
12140    Dim nCeiling          As Single
12150    Dim lResult As Long

12160    If lHCryptprov = 0 Then InitRand
12170    If Max = 0 Then Max = 9
12180    nCeiling = Max
12190    sCryptBuffer = Chr(0)

12200    If Weight = 0 Then Weight = 1 'A weight of zero wouldn't do anything

12210    For i = 1 To Abs(Weight)

12220       If m_bUseRandAPI Then
12230          lResult = CryptGenRandom(lHCryptprov, 1, sCryptBuffer)
12240          nRnd = Asc(sCryptBuffer)
12250          nRnd = nRnd / 255
12260       Else
12270          Randomize Now
12280          nRnd = Rnd
12290       End If

12300       nCeiling = (nRnd * (nCeiling - Min)) + Min
12310    Next i

12320    If Weight > 0 Then nCeiling = Max - (nCeiling - Min)
12330    If DecimalPlaces Then
12340       Rand = Format(nCeiling, "0." & String(DecimalPlaces, "0"))
12350    Else
12360       Rand = Format(nCeiling, "0")
12370    End If

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       RandomCase
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Word (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function RandomCase(ByVal Word As String) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".RandomCase"

         '</EhHeader>
12100    Dim sLetter      As String
12110    Dim i            As Long

12120    Select Case Rand(14, 0)

            Case 0  'Do nothing

12130       Case 1  'All upper
12140          Word = StrConv(Word, vbUpperCase)

12150       Case 2  'All lower
12160          Word = StrConv(Word, vbLowerCase)

12170       Case 3  'Proper case
12180          Word = StrConv(Word, vbProperCase)

12190       Case 4   'Pick one letter and make all occurrences of that letter ucase
12200          sLetter = Mid$(Word, Rand(Len(Word), 1), 1)
12210          Word = Replace(Word, sLetter, UCase(sLetter))

12220       Case 5   'Totally random

12230          For i = 1 To Len(Word)

12240             If Rand() Then Mid$(Word, i, 1) = UCase(Mid$(Word, i, 1))
12250          Next i

12260       Case 6, 7  'Ucase one random character
12270          i = Rand(Len(Word), 1)
12280          Mid$(Word, i, 1) = UCase(Mid$(Word, i, 1))

12290       Case 8   'VOWELS UCase

12300          For i = 1 To Len(Word)

12310             If InStr(VOWELS, Mid$(Word, i, 1)) Then
12320                Mid$(Word, i, 1) = UCase(Mid$(Word, i, 1))
12330             End If

12340          Next i

12350       Case 9   'CONSONANTS UCase

12360          For i = 1 To Len(Word)

12370             If InStr(CONSONANTS, Mid$(Word, i, 1)) Then
12380                Mid$(Word, i, 1) = UCase(Mid$(Word, i, 1))
12390             End If

12400          Next i

12410       Case 10  '2 consecutive letters ucase
12420          i = Rand(Len(Word) - 1, 1)
12430          Mid$(Word, i, 1) = UCase(Mid$(Word, i, 1))
12440          Mid$(Word, i + 1, 1) = UCase(Mid$(Word, i + 1, 1))

12450       Case 11  'Last letter
12460          Mid$(Word, Len(Word), 1) = UCase(Mid(Word, Len(Word), 1))

12470       Case 12  'first and last letters
12480          Mid$(Word, 1, 1) = UCase(Mid(Word, 1, 1))
12490          Mid$(Word, Len(Word), 1) = UCase(Mid(Word, Len(Word), 1))

12500       Case 13  'first x letters
12510          i = Rand(Len(Word), 1, 2)
12520          Mid$(Word, 1, i) = UCase(Mid$(Word, 1, i))

12530       Case 14  'every other letter ucase

12540          For i = 1 To Len(Word) Step 2
12550             Mid$(Word, i, 1) = UCase(Mid$(Word, i, 1))
12560          Next i

12570    End Select

12580    RandomCase = Word
         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       RandomWord
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       FilePath (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function RandomWord(ByRef FilePath As String) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".RandomWord"

         '</EhHeader>
12100    Const BUFFER_SIZE = 128
12110    Dim sWord      As String
12120    Dim i          As Integer
12130    Dim lfile As Long
12140    Dim sBuffer As String * BUFFER_SIZE
12150    Dim sLine As String
12160    Dim lSeek As Long
12170    Dim x As Long
12180    Dim lRetryCount As Long
12190    Dim lLen As Long
12200    Dim lMaxFileSeekRetries As Long
12210    Dim iLen As Long

         'Debug.Print FilePath
12220    On Error GoTo ErrHandler
ProcStart:
12230    lfile = FreeFile()
12240    lLen = FileLen(FilePath)
12250    Open FilePath For Binary Access Read As lfile
12260    lSeek = Rand(lLen - BUFFER_SIZE, 1)

12270    Do
12280       Get lfile, lSeek, sBuffer
            '   If Err Then
            '   RandomWord = "#FILE_ERROR"
            '   Close lFile
            '   Exit Function
            '   End If
            'Find the first CrLf to find the beginning of a line
12290       x = InStr(sBuffer, vbCrLf)

12300       If x Then
12310          sLine = Mid$(sBuffer, x + 2, BUFFER_SIZE)
12320       Else
12330          lMaxFileSeekRetries = lMaxFileSeekRetries + 1

12340          If EOF(lfile) Then
12350             lSeek = Rand(lLen, 1)
12360          Else
12370             lSeek = lSeek + BUFFER_SIZE
12380          End If
12390       End If

12400       If lMaxFileSeekRetries > 5 Then
12410          Debug.Print "lMaxFileSeekRetries Exceeded"
12420          Err.Raise ERR_WORDLIST_READ, ERRSOURCE, ERR_WORDLIST_READ
12430       End If
           
12440    Loop Until x

         'Now find the next CrLf
        
12450    Dim sWords() As String
        
12460    sWords = Split(sLine, vbCrLf)
         'Debug.Print "Words: " & UBound(sWords)
12470    sLine = Trim$(sWords(1))
         'Debug.Print sLine
        
         '1155    x = InStr(sLine, vbCrLf)
         '
         '1160    If x Then
         '1165       sLine = Left$(sLine, x - 1)
         '1170    Else
         '1175       lSeek = Rand(iLen, 1)
         '
         '1180       Do Until x
         '1185          Get lFile, lSeek, sBuffer
         '
         '1190          If Err Then
         '1195             RandomWord = "#FILE_ERROR"
         '1200             Exit Function
         '
         '1205          End If
         '
         '1210          x = InStr(sBuffer, vbCrLf)
         '
         '1215          If x Then
         '1220             sLine = sLine & Left$(sBuffer, x - 1)
         '1225          Else
         '
         '1230             If EOF(lFile) Then
         '1235                sLine = sLine & sBuffer
         '1240                Exit Do
         '1245             Else
         '1250                lSeek = lSeek + 64
         '1255             End If
         '1260          End If
         '
         '1265       Loop
         '
         '1270    End If

12480    Close lfile
12490    sLine = Replace(sLine, vbLf, "")
12500    sLine = Replace(sLine, vbCr, "")

12510    If Len(sLine) = 0 Then GoTo ProcStart
         'If InStr(sLine, vbCrLf) Then Debug.Print sLine
       
         ' Debug.Print "+Word from " & FilePath & ":  " & sLine

         
12520    RandomWord = sLine
ExitHere:
12530    Close lfile
12540    Exit Function

ErrHandler:
12550    Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
12560    lRetryCount = lRetryCount + 1

12570    If lRetryCount > 2 Then
12580       Resume ExitHere
12590    Else
12600       Close lfile
12610       Resume ProcStart
12620    End If

         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       RemoveLetter
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Phrase (String)
'                    ReplaceChar (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function RemoveLetter(ByVal Phrase As String, _
                             Optional ReplaceChar As String) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".RemoveLetter"

         '</EhHeader>
12100    Dim iPos As Integer
12110    Dim iLoopCount As Integer
12120    Dim i As Integer
12130    Dim sChar As String
12140    Dim sNewPhrase As String
12150    Dim sLastChar As String
12160    Dim sLetter As String
        
12170    Do
12180       sLetter = PickCharacter(VOWELS + LETTERS)
12190       iPos = InStr(Phrase, sLetter)
12200       iLoopCount = iLoopCount + 1
12210    Loop Until iPos > 0 Or iLoopCount > 25

12220    If iPos > 0 Then
12230       sNewPhrase = Mid$(Phrase, 1, 1)

12240       For i = 2 To Len(Phrase) - 1
12250          sChar = LCase(Mid$(Phrase, i, 1))

12260          If sChar = sLetter Then
12270             If Mid$(Phrase, i - 1, 1) <> " " And Mid$(Phrase, i + 1) <> " " Then
12280                sChar = ReplaceChar
12290             End If
12300          End If

12310          sNewPhrase = sNewPhrase & sChar
12320       Next i

12330       sNewPhrase = sNewPhrase & Right(Phrase, 1)

            'Try it again if nothing happened
12340       If sNewPhrase = Phrase Then sNewPhrase = RemoveLetter(Phrase)
12350       RemoveLetter = sNewPhrase
12360    Else
12370       RemoveLetter = Phrase
12380    End If

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       ScrambleWord
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Word (String)
'                    Times (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ScrambleWord(ByVal Word As String, _
                             Optional Times As Long) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".ScrambleWord"

         '</EhHeader>
12100    Dim b()           As Byte
12110    Dim iTmp          As Byte
12120    Dim iLen          As Long
12130    Dim x1            As Long
12140    Dim x2            As Long
12150    Dim iCharLen      As Long
12160    Dim i             As Long
12170    iCharLen = LenB("A")
12180    b() = Word
12190    iLen = UBound(b)

12200    If Times = 0 Then Times = 1

12210    For i = 1 To Times
12220       x1 = Rand(iLen / iCharLen) * iCharLen
12230       x2 = Rand(iLen / iCharLen) * iCharLen
12240       iTmp = b(x1)
12250       b(x1) = b(x2)
12260       b(x2) = iTmp
12270    Next i

12280    ScrambleWord = b()
         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       SentenceCase
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Sentence (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function SentenceCase(ByVal Sentence As String) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".SentenceCase"

         '</EhHeader>
12100    Sentence = LCase(Sentence)
12110    Mid$(Sentence, 1, 1) = UCase(Mid$(Sentence, 1, 1))
12120    SentenceCase = Sentence
         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       StartTiming
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Public Sub StartTiming()
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".StartTiming"

         '</EhHeader>
12100    timeBeginPeriod 1
12110    m_lT = timeGetTime
         '<EhFooter>
ExitHere:
         Exit Sub

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Sub

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       Stutter
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       Word (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function Stutter(ByVal Word As String) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".Stutter"

         '</EhHeader>
12100    Dim iPos As Integer
12110    Dim i As Integer
12120    Dim j As Integer
12130    Dim sFirstPart As String
12140    Dim sStuttered As String
12150    Dim sSyllableMarker As String

12160    If Rand(100) > 20 Then
12170       sSyllableMarker = "aeiou"
12180    Else
12190       sSyllableMarker = "hywrtnaeiou"
12200    End If

12210    For i = 1 To Len(Word)

12220       If InStr(sSyllableMarker, LCase(Mid$(Word, i, 1))) Then
12230          sFirstPart = Left(Word, i)
12240          sStuttered = Word

12250          If Rand(100) < 5 Then sFirstPart = sFirstPart & "..."
12260          If Rand(100) < 10 Then sFirstPart = sFirstPart & " "

12270          For j = 1 To Rand(4, 1, -2)
12280             sStuttered = sFirstPart & sStuttered
12290          Next j

12300          Exit For
12310       End If

12320    Next i

12330    Stutter = sStuttered
         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       ToRoman
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       sDecNum (String)
'--------------------------------------------------------------------------------
'</CSCM>
Public Function ToRoman(ByVal sDecNum As String) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".ToRoman"

         '</EhHeader>

12100    If sDecNum <> "0" And sDecNum <> vbNullString Then
12110       Dim sNumArray()      As String

12120       If Len(sDecNum) > 3 Then ToRoman = String(Mid(sDecNum, 1, Len(sDecNum) - 3), "M")
12130       If Len(sDecNum) > 2 Then ToRoman = ToRoman & GiveLetters(Mid(sDecNum, Len(sDecNum) - 2, 1), 4)
12140       If Len(sDecNum) > 1 Then ToRoman = ToRoman & GiveLetters(Mid(sDecNum, Len(sDecNum) - 1, 1), 2)
12150       ToRoman = ToRoman & GiveLetters(Mid(sDecNum, Len(sDecNum), 1), 0)
12160       Else: ToRoman = "No Roman value For 0"
12170    End If

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       GiveLetters
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :       sInput (String)
'                    iArrStart (Long)
'--------------------------------------------------------------------------------
'</CSCM>
Private Function GiveLetters(ByVal sInput As String, _
                             ByVal iArrStart As Long) As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".GiveLetters"

         '</EhHeader>
12100    Dim sLetterArray()      As String
12110    sLetterArray() = Split("I V X L C D M", " ")

12120    Select Case sInput

            Case 4: GiveLetters = sLetterArray(iArrStart) & sLetterArray(iArrStart + 1)

12130       Case 5: GiveLetters = sLetterArray(iArrStart + 1)

12140       Case 9: GiveLetters = sLetterArray(iArrStart) & sLetterArray(iArrStart + 2)

12150       Case 6 To 8: GiveLetters = sLetterArray(iArrStart + 1) & String(sInput - 5, sLetterArray(iArrStart))

12160       Case Else: GiveLetters = GiveLetters + String(sInput, sLetterArray(iArrStart))
12170    End Select

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function

'<CSCM>
'--------------------------------------------------------------------------------
' Procedure  :       InitRand
' Description:       [type_description_here]
' Created by :       M. Burnett
' Date-Time  :       1/27/2006-23:21:30
'
' Parameters :
'--------------------------------------------------------------------------------
'</CSCM>
Private Sub InitRand()
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".InitRand"

         '</EhHeader>
12100    Dim sContainer      As String
12110    Dim lResult         As Long
12120    Dim sProvider       As String
12130    sContainer = vbNullChar
12140    sProvider = MS_DEF_PROV & vbNullChar
         'sProvider = MS_ENH_PROV & vbNullChar
         'sProvider = MS_STR_PROV & vbNullChar
         'TODO: Error handler to see if provider exists
12150    lResult = CryptAcquireContext(lHCryptprov, ByVal sContainer, ByVal sProvider, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)

12160    If lResult Then m_bUseRandAPI = True
         '<EhFooter>
ExitHere:
         Exit Sub

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, Err.Description
         Resume ExitHere
         '</EhFooter>
End Sub

Public Property Get GetWordlistDir() As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".GetWordlistDir"

         '</EhHeader>
32500    Dim sPaths(4)      As String
32510    Dim i              As Long
         Dim sDir As String

32530    sPaths(0) = GetFullPath(GetStringSetting("Pafwert", "Settings", "WordlistDir") & "\")
32540    sPaths(1) = GetFullPath(App.Path & "\wordlists" & "\")
32550    sPaths(2) = GetFullPath(GetStringSetting("Pafwert", "Settings", "LastWordlistDir") & "\")
32560    sPaths(3) = GetFullPath("C:\Program Files\Pafwert\Wordlists" & "\")
32570    sPaths(4) = GetFullPath(FileSystem.CurDir & "\")

32580    For i = 0 To 4

32590       If CheckWordlistDir(sPaths(i)) Then
32600          sDir = sPaths(i)
32610          Exit For
32620       End If

32630    Next i

32640    If Len(sDir) = 0 Then Err.Raise ERR_WORDLIST_DIR_NOT_FOUND, ERRSOURCE, ERR_WORDLIST_DIR_NOT_FOUND_DESC
32650    SaveStringSetting "Pafwert", "Settings", "WordlistDir", sDir


32670 GetWordlistDir = sDir
      '<EhFooter>
ExitHere:
      Exit Property

ErrHandler:
      Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
      Resume ExitHere
      '</EhFooter>
End Property

Public Function CheckWordlistDir(Path As String) As Boolean
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".CheckWordlistDir"

      '</EhHeader>
32500    On Error Resume Next

32510    If Len(Dir(Path)) Then
32520       If FileLen(Path & "\patterns.cfg") = 0 Or Err Then
32530          Err.Raise ERR_OPEN_PATTERNS_FILE, ERRSOURCE, ERR_OPEN_PATTERNS_FILE_DESC
32540       Else
32550          CheckWordlistDir = True
32560       End If
32570    End If

      '<EhFooter>
ExitHere:
            Exit Function

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Function


