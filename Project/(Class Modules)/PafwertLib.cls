VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PafwertLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Complexity"
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

Option Explicit

Private Const MODULE_SOURCE = "PafwertLib"
'CSEH: Pafwert Components

'TODO:
'  Change pattern property to collection
'Using UDT instead of class for better performance
Private Type Placeholder
   OriginalPattern As String
   NormalizedPattern As String
   CompletedPattern As String
   Parent As Long
   Start As Long
End Type
Private Type Modifier
   Function As String
      Parameters() As String
      Qualifier As Integer
   End Type

   Private Type Pattern
      Placeholders() As Placeholder
   End Type

   '--- Module-level Declarations
   Private mpPatterns()  As Pattern
   'Property defaults
   Const DEFAULT_MINLEN As String = 6
   Const DEFAULT_MAXLEN As String = 35
   Const DEFAULT_MINCHARSETS = 2

   'Local variables to hold property values
   Private mvarKeywords          As String
   Private mvarPattern           As String
   Private mvarPassword          As String
   Private mvarEntropy1          As Integer
   Private mvarEntropy2          As Integer
   Private mvarEntropy3          As Integer
   Private mvarWordlistDir       As String
   Private bChecked              As Boolean
   Private mvarValue             As Double
   Private bCalcKeyspace         As Boolean
   Private mvarLastPattern As String
   Private mvarTimeTaken As Long
   Private mvarComplexity As Complexity
   

Public Property Get Complexity() As Complexity
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".Complexity"

      '</EhHeader>

32500    If mvarComplexity Is Nothing Then
32510       Set mvarComplexity = New Complexity
32520    End If

32530    Set Complexity = mvarComplexity
      '<EhFooter>
ExitHere:
            Exit Property

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Property

Public Property Set Complexity(vData As Complexity)
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".Complexity"

      '</EhHeader>
32500    Set mvarComplexity = vData
      '<EhFooter>
ExitHere:
            Exit Property

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Property

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
   Set mvarComplexity = Nothing
End Sub

Public Property Get TimeTaken() As Long
Attribute TimeTaken.VB_Description = "Returns the time, in milliseconds, required to generate the last password."
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".TimeTaken"

      '</EhHeader>
32500    TimeTaken = mvarTimeTaken
      '<EhFooter>
ExitHere:
            Exit Property

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Property

'Public Function GeneratePassword(Optional ByVal Pattern As String, Optional ByVal Keywords As String, Optional ByVal MinLen As Integer = DEFAULT_MINLEN) As String
Public Function GeneratePassword(Optional ByVal Pattern As String, _
                                 Optional ByVal Keywords As String) As String
Attribute GeneratePassword.VB_Description = "Generates a password."
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".GeneratePassword"

      '</EhHeader>
32500    Dim lCount              As Long
32510    Dim pPlaceholder()      As Placeholder
32520    Dim lResult             As Long
32530    Dim i                   As Long
32540    Dim sPassword           As String
         Dim iLen As Long
         Dim MinLen As Long
MakePass:
32550    lCount = lCount + 1
32560    ReDim pPlaceholder(0)

32570    If lCount > 10 Then Err.Raise Err.Number, ERRSOURCE & " (" & Erl & ")", Err.Description

         'Get a pattern
         On Error Resume Next
32580    If Len(Pattern) = 0 Then
32590       If Len(mvarPattern) Then
32600          lResult = ParsePattern(mvarPattern, pPlaceholder())
32610       Else
32620          pPlaceholder() = GetRandomPattern
32630       End If

32640    Else
32650       lResult = ParsePattern(Pattern, pPlaceholder())
32660    End If

         'If Len(pPlaceholder()) = 0 Then Err.Raise ERR_INVALID_PATTERN, ERRSOURCE, ERR_INVALID_PATTERN_DESC
         

32670    If Err Then Err.Raise ERR_LOADING_PATTERNS, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl, ERR_LOADING_PATTERNS_DESC & ": " & Err.Description
32680
    
32690    Me.Pattern = pPlaceholder(UBound(pPlaceholder)).OriginalPattern
         'Debug.Print "--->" & Me.Pattern

32700    If UBound(pPlaceholder) = 0 Then Err.Raise ERR_INVALID_PATTERN, ERRSOURCE, ERR_INVALID_PATTERN_DESC
32710    If Err Then Err.Raise ERR_INVALID_PATTERN, ERRSOURCE, ERR_INVALID_PATTERN_DESC
   
32720    StartTiming

         'Work through each placeholder in the pattern
   

32730    For i = 1 To UBound(pPlaceholder)

32740       With pPlaceholder(i)

               'Check for back references or get a word
32750          If Left$(.NormalizedPattern, 2) = "$W" Then
32760             .CompletedPattern = pPlaceholder((Val(Mid$(.NormalizedPattern, 3, 1)))).CompletedPattern
32770          Else
32780             .CompletedPattern = GetWord(.NormalizedPattern, Keywords)
32790          End If

               'Replace placeholder in parent pattern
32800          If .Parent > 0 Then
32810             pPlaceholder(.Parent).NormalizedPattern = Replace(pPlaceholder(.Parent).NormalizedPattern, .OriginalPattern & Format(i, "00"), .CompletedPattern)
32820          End If

32830       End With

32840    Next i

         'The top-level pattern will now contain the password
32850    sPassword = pPlaceholder(i - 1).CompletedPattern
32860    iLen = Len(sPassword)

32870    If iLen < MinLen Then
32880       If lCount < 10 Then
32890          GoTo MakePass
32900       Else
               '         For i = 1 To MinLen - iLen
               '            sPassword = sPassword & PickCharacter(KEYBOARD)
               '         Next i
32910       End If
32920    End If
   
32930    bChecked = False
32940    sPassword = Trim$(sPassword)
32950    sPassword = Replace(sPassword, "#sla#", "\")
32960    sPassword = Replace(sPassword, "#pls#", "+")
32970    sPassword = Replace(sPassword, "#lbr#", "{")
32980    sPassword = Replace(sPassword, "#rbr#", "}")
32990    sPassword = Replace(sPassword, "#lba#", "[")
33000    sPassword = Replace(sPassword, "#rba#", "]")
33010    sPassword = Replace(sPassword, "#lpa#", "(")
33020    sPassword = Replace(sPassword, "#rpa#", ")")
33030    sPassword = Replace(sPassword, "#pip#", "|")
         sPassword = Replace(sPassword, "  ", " ")
33040    mvarPassword = Replace(sPassword, vbCrLf, "")
   
33050    Me.Complexity.Password = mvarPassword
33060    mvarTimeTaken = EndTiming

'33070    If InStr(LCase(mvarPassword), "{word") Then
'33080       Debug.Print "->" & Me.Pattern
'33090       Debug.Print "-   " & mvarPassword & vbCrLf
'33100    End If
   
         'Failsafe password if something went wrong somewhere
33110    If Len(mvarPassword) < 4 Then
            For i = 0 To 6
            mvarPassword = mvarPassword & PickOne(VOWELS2 & "! @ # % $ ^ & * : ' / ` ~ * - < > + =. . , , ; ; ? ? " & _
                         CONSONANTS2 & THREELETTERWORDS & "1 2 3 4 5 6 7 8 9 0")
            Next i
33130    End If

33140    mvarLastPattern = Me.Pattern
33150    Me.Pattern = ""
33160    GeneratePassword = mvarPassword

      '<EhFooter>
ExitHere:
            Exit Function

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Function
   
Private Function GetRandomPattern() As Placeholder()
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".GetRandomPattern"

         '</EhHeader>
213812   Dim sPattern            As String
213814   Dim lPos                As Long
213816   Dim sFileData           As String
213818   Dim sPatterns()         As String
213820   Dim pPlaceholder()      As Placeholder
213822   Dim x                   As Long
213824   Dim iPatternCount       As Long
213826   Dim lRandomPattern As Long
         Dim lResult As Long
213828   Dim lfile As Long
213832   Dim i As Integer
213834   Dim sPatternCheck As String
         Dim j As Long
213836   On Error Resume Next
213838   x = UBound(mpPatterns)

213840   If Err Or x = 0 Then
213842      If Len(Dir$(Me.WordlistDir & "\patterns.cfg")) = 0 Then Err.Raise ERR_OPEN_PATTERNS_FILE, ERRSOURCE, ERR_OPEN_PATTERNS_FILE_DESC
213844      lfile = FreeFile
213846      On Error Resume Next
213848      Open Me.WordlistDir & "\patterns.cfg" For Binary Access Read As #lfile
            
213850      If Err Then Err.Raise ERR_OPEN_PATTERNS_FILE, ERRSOURCE, ERR_OPEN_PATTERNS_FILE_DESC
213852      sFileData = Space$(LOF(lfile))
213854      Get #lfile, , sFileData
213856      Close lfile

213858      If Len(sFileData) = 0 Or Err Then
               Err.Raise ERR_LOADING_PATTERNS, ERRSOURCE, ERR_LOADING_PATTERNS_DESC
               GoTo ErrHandler
            End If
            
213860      sFileData = Replace(sFileData, vbCrLf & vbCrLf, vbCrLf)
213862      sPatterns = Split(sFileData, vbCrLf)

213864      If IsBounded(sPatterns) = False Then
213866         Err.Raise ERR_LOADING_PATTERNS, ERRSOURCE, ERR_LOADING_PATTERNS_DESC & Me.WordlistDir & "\patterns.cfg"
213868      End If

213870      ReDim mpPatterns(UBound(sPatterns))
            j = UBound(sPatterns)
            
213872      For i = 0 To j

               'Debug.Assert iPatternCount < 90
213874         If Left$(sPatterns(i), 1) <> "#" And Len(sPatterns(i)) Then
213876            lPos = InStr(sPatterns(i), ":")

213878            If lPos Then
213880               sPattern = (Right$(sPatterns(i), Len(sPatterns(i)) - lPos))
213882            Else
213884               sPattern = sPatterns(i)
213886            End If

213888            sPatternCheck = CheckPattern(sPattern)

213890            If Len(sPatternCheck) Then
213892               Err.Raise ERR_LOADING_PATTERNS, ERRSOURCE, sPatternCheck
213894            End If

213896            lResult = ParsePattern(sPattern, pPlaceholder)
213898            mpPatterns(iPatternCount).Placeholders = pPlaceholder
213900            iPatternCount = iPatternCount + 1
213902         Else
                  'Debug.Assert False
213904         End If

213906      Next

213908      ReDim Preserve mpPatterns(iPatternCount - 1)
213910   End If

213912   If UBound(mpPatterns) = 0 Then Err.Raise ERR_LOADING_PATTERNS, ERRSOURCE, ERR_LOADING_PATTERNS_DESC
213914   lRandomPattern = Rand(UBound(mpPatterns))
213916   mvarLastPattern = Trim$(mpPatterns(lRandomPattern).Placeholders(UBound(mpPatterns(lRandomPattern).Placeholders)).OriginalPattern)
         'Debug.Print "Pattern ---->" & mvarLastPattern & "<---"
213918   GetRandomPattern = mpPatterns(lRandomPattern).Placeholders

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function
   
Private Function ParsePattern(ByRef Pattern As String, _
                              ByRef PlaceHolderArray() As Placeholder) As Long
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".ParsePattern"

      '</EhHeader>
32500    Dim sMasterPattern         As String 'REMOVE?
32510    Dim bPattern()             As Byte
32520    Dim lCharWidth             As Long
32530    Dim lLeftCount             As Long
32540    Dim i                      As Long
32550    Dim lPlaceholderCount      As Long
32560    Dim lPosition()            As Long
32570    Dim sPatternID             As String
32580    Dim sPattern As String
         Dim j As Long
32590    Pattern = Trim$(Pattern)
32600    lCharWidth = LenB("A")
32610    sPattern = Replace(Pattern, "\\", "#sla#")
32620    sPattern = Replace(sPattern, "\+", "#pls#")
32630    sPattern = Replace(sPattern, "\{", "#lbr#")
32640    sPattern = Replace(sPattern, "\}", "#rbr#")
32650    sPattern = Replace(sPattern, "\[", "#lba#")
32660    sPattern = Replace(sPattern, "\]", "#rba#")
32670    sPattern = Replace(sPattern, "\(", "#lpa#")
32680    sPattern = Replace(sPattern, "\)", "#rpa#")
32690    sPattern = Replace(sPattern, "\|", "#pip#")
         'pad closing braces to make room for the pattern ID
32700    sPattern = Replace(sPattern, "}", "}00")
32710    bPattern() = sPattern
32720    ReDim PlaceHolderArray(1 To 1)
32730    ReDim lPosition(1 To 1)

         '--Parse out wordlist placeholders
32740    For i = 0 To UBound(bPattern) Step lCharWidth
32750       sMasterPattern = sMasterPattern & Chr$(bPattern(i))

32760       If bPattern(i) = 123 Then  '  {
32770          lLeftCount = lLeftCount + 1
32780          ReDim Preserve lPosition(1 To lLeftCount)
32790          lPosition(lLeftCount) = i
32800       End If

32810       If bPattern(i) = 125 Then  '  }
32820          If lLeftCount = 0 Then Err.Raise ERR_PATTERN_MISMATCHED_BRACES, "ParsePattern", ERR_PATTERN_MISMATCHED_BRACES_DESC & " " & Pattern
32830          lPlaceholderCount = lPlaceholderCount + 1
               'Below use bPatternID instead
32840          sPatternID = Format$(lPlaceholderCount, "00")
32850          sMasterPattern = sMasterPattern & sPatternID
32860          bPattern(i + lCharWidth) = Asc(Mid$(sPatternID, 1, 1))
32870          bPattern(i + lCharWidth + lCharWidth) = Asc(Mid$(sPatternID, 2, 1))
32880          ReDim Preserve PlaceHolderArray(1 To lPlaceholderCount)

32890          With PlaceHolderArray(lPlaceholderCount)
32900             .Start = lPosition(lLeftCount)

32910             If lLeftCount > 1 Then
32920                .Parent = lPosition(lLeftCount - 1)
32930             Else
32940                .Parent = -1
32950             End If

32960             .OriginalPattern = MidB(bPattern(), .Start + 1, i - .Start + lCharWidth)
32970             .NormalizedPattern = MidB(bPattern(), .Start + lCharWidth + 1, i - .Start - lCharWidth)
32980          End With

               'Jump ahead two characters past the pattern ID
32990          i = i + (lCharWidth * 2)
33000          lLeftCount = lLeftCount - 1
33010       End If

33020    Next i

33030    ReDim Preserve PlaceHolderArray(1 To UBound(PlaceHolderArray) + 1)
33040    PlaceHolderArray(UBound(PlaceHolderArray)).OriginalPattern = Pattern
33050    PlaceHolderArray(UBound(PlaceHolderArray)).NormalizedPattern = sMasterPattern
33060    PlaceHolderArray(UBound(PlaceHolderArray)).Start = 0
         'ShowPatterns Pattern, Placeholderarray
         'Now that we know the parent items, go through and replace them
33070    On Error GoTo 0

33080    For i = 1 To UBound(PlaceHolderArray)

33090       With PlaceHolderArray(i)

33100          For j = 1 To UBound(PlaceHolderArray) - 1

33110             If PlaceHolderArray(j).Parent = -1 Then PlaceHolderArray(j).Parent = UBound(PlaceHolderArray)
33120             If PlaceHolderArray(j).Parent = .Start Then
33130                PlaceHolderArray(j).Parent = i
33140             End If

33150          Next j

33160       End With

33170    Next i

      '<EhFooter>
ExitHere:
            Exit Function

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Function

Private Function GetWord(ByVal Placeholder As String, _
                         Optional ByRef Keywords As String) As String
         'ENHANCEMENT: Pick a random wordlist
         'ENHANCEMENT: Allow calling of external COM or DLL function or EXE
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".GetWord"

      '</EhHeader>
32500    Dim sWord                 As String
32510    Dim sFilepath             As String
32520    Dim lPos                  As Long
32530    Dim sPlaceholdername      As String
32540    Dim sParams()             As String
32550    ReDim sParams(3)
32560    Dim sModifiers()          As String
   
         'Check for modifiers
32570    lPos = InStrRev(Placeholder, "+")

32580    If lPos Then
32590       sModifiers() = Split(Mid$(Placeholder, lPos + 1, Len(Placeholder)), "+")
32600       Placeholder = Left$(Placeholder, lPos - 1)
32610    End If

         'Check for a qualifier
32620    lPos = InStr(Placeholder, "[")

32630    If lPos Then
32640       If Val(Mid$(Placeholder, lPos + 1, InStr(Placeholder, "]") - 1)) < Rand(99, 0) Then
32650          GetWord = ""
32660          Exit Function

32670       Else
32680          Placeholder = Left$(Placeholder, lPos - 1)
32690       End If
32700    End If

         'Check for parameters
32710    lPos = InStr(Placeholder, "(")

32720    If lPos Then
32730       sParams() = Split(Replace(Mid$(Placeholder, lPos + 1, InStr(Placeholder, ")") - lPos - 1), ", ", ","), ",")
32740       Placeholder = Left$(Placeholder, lPos - 1)
32750    End If

         'Check for a selection group
32760    If InStrB(Placeholder, "|") Then
32770       GetWord = PickOne(Placeholder, , "|")
32780       Exit Function

32790    End If

         'Dimension 4 parameters (0 to 4), even if they are blank
32800    ReDim Preserve sParams(3)

32810    Select Case Trim$(LCase(Placeholder))

            Case "word"

32820          If Len(Keywords) And Chance(Me.Entropy1) Then
32830             sWord = PickOne(Keywords)
32840          Else
32850             sFilepath = GetFileName(sParams(0))
32860             sWord = RandomWord(sFilepath)
32870          End If

32880       Case "sp", "space"
32890          sWord = " "

32900       Case "vowel"
32910          sWord = PickCharacter(VOWELS, Val(sParams(0)))

32920       Case "consonant"
32930          sWord = PickCharacter(CONSONANTS, Val(sParams(0)))

32940       Case "symbol"
32950          sWord = PickOne(SYMBOLS)

32960       Case "endpunctuation"
32970          sWord = PickOne(ENDPUNCTUATION)

32980       Case "sentencepunctuation"
32990          sWord = PickCharacter(SENTENCEPUNCTUATION)

33000       Case "number"
33010          sWord = Rand(Val(sParams(0)), Val(sParams(1)), Val(sParams(2)), Val(sParams(3)))

33020       Case "letter"
33030          sWord = PickCharacter(LETTERS, Val(sParams(0)))

33040       Case "smiley"
33050          sWord = PickOne(SMILEYS)

33060       Case "keyboard"
33070          sWord = PickCharacter(KEYBOARD)

33080       Case "numrow"
33090          sWord = PickCharacter(NUMROW)

33100       Case "numrowfull"
33110          sWord = PickCharacter(NUMROWFULL)

33120       Case "row1"
33130          sWord = PickCharacter(ROW1)

33140       Case "row1full"
33150          sWord = PickCharacter(ROW1FULL)

33160       Case "row2"
33170          sWord = PickCharacter(ROW2)

33180       Case "row2full"
33190          sWord = PickCharacter(ROW2FULL)

33200       Case "row3"
33210          sWord = PickCharacter(ROW3)

33220       Case "row3full"
33230          sWord = PickCharacter(ROW3FULL)

33240       Case "lefthand"
33250          sWord = PickCharacter(LEFTHAND)

33260       Case "righthand"
33270          sWord = PickCharacter(RIGHTHAND)

33280       Case "sequence"
33290          sWord = GetSequence(Val(sParams(0)))

33300       Case "ordinal"
33310          sWord = GetOrdinal(sParams(0))

33320       Case "phonetic"
33330          sWord = GetPhonetic(sParams(0), Val(sParams(1)))

33340       Case "pronounceable"
33350          sWord = PronounceableWord()

33360       Case "numberpattern"

33370          If Val(sParams(0)) = 0 Then sParams(0) = 3
33380          sWord = GetNumberPattern(sParams(0))

33390       Case "entropy1"
33400          sWord = Me.Entropy1

33410       Case "entropy2"
33420          sWord = Me.Entropy2

33430       Case "entropy3"
33440          sWord = Me.Entropy3

33450       Case "asc"
33460          sWord = Asc(sParams(0))

33470       Case "chr"

33480          If IsNumeric(Val(sParams(0))) Then
33490             sWord = Chr$(Val(sParams(0)))
33500          End If

33510       Case "now"
33520          sWord = Now()

33530       Case "anagram"

33540       Case "synonym"

33550       Case "rhyme"

33560       Case "homophone"

33570       Case "antonym"

               'Call modifier as a function
33580       Case "bracket", "num2words", " reverse", "ucase", "lcase", "propercase", "obscure", "replace", "randomcase", "scramble", "piglatin", "repeat", "right", "left", "trim", "format", "mid", "swap", "romannumeral", "hidden"

               '  sWord = ModifyWord(sParams(0), sPlaceholdername & "(" & sParams(1) & "," & sParams(2) & ")")
               '   Case "lookup"
               ' sFilepath = GetFileName(sParams(0))
               ' sWord = Lookup(sFilepath, sParams(1))
33590       Case "longmonth"
33600          sWord = PickOne(LONGMONTHS)

33610       Case "shortmonth"
33620          sWord = PickOne(SHORTMONTHS)

33630       Case "longday"
33640          sWord = PickOne(LONGDAYS)

33650       Case "shortday"
33660          sWord = PickOne(SHORTDAYS)

33670       Case "numbercode"
33680          sWord = NumberCode()

33690       Case Else
33700          sWord = Placeholder
33710    End Select

33720    If IsBounded(sModifiers) Then ModifyWord sWord, sModifiers()
         'Escape any special characters that may have shown up during the word selection
33730    sWord = Replace(sWord, "\", "#sla#")
33740    sWord = Replace(sWord, "+", "#pls#")
33750    sWord = Replace(sWord, "{", "#lbr#")
33760    sWord = Replace(sWord, "}", "#rbr#")
33770    sWord = Replace(sWord, "[", "#lba#")
33780    sWord = Replace(sWord, "]", "#rba#")
33790    sWord = Replace(sWord, "(", "#lpa#")
33800    sWord = Replace(sWord, ")", "#rpa#")
33810    sWord = Replace(sWord, "|", "#pip#")
33820    GetWord = sWord
      'Debug.Print "-" & sWord
      '<EhFooter>
ExitHere:
            Exit Function

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Function
   
Private Function ModifyWord(ByRef Word As String, _
                            ModifiersArray() As String) As String
         'ENHANCEMENT: VBScript code eval
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".ModifyWord"

         '</EhHeader>
213810   Dim j As Long
213812   Dim sTmp As String
213814   Dim sTmpArray() As String
213816   Dim sParams() As String
213818   Dim Modifier As Variant
         Dim lPos As Long
         Dim lQualifier As Long
         Dim i As Long
         Dim sWord As String
213820   ReDim sParams(3)

213822   If Len(Word) = 0 Then Exit Function

213824   For j = 0 To UBound(ModifiersArray)
213826      Modifier = ModifiersArray(j)
            'Check for a qualifier
213828      lPos = InStr(Modifier, "[")

213830      If lPos Then
213832         lQualifier = Val(Mid$(Modifier, lPos + 1, InStr(Modifier, "]") - 1))

213834         If Val(lQualifier) < Rand(99, 0) Then
213836            ModifyWord = Word
213838            Exit Function

213840         Else
213842            Modifier = Left$(Modifier, lPos - 1)
213844         End If
213846      End If

            '--Check for paramaters
213848      lPos = InStr(Modifier, "(")

213850      If lPos Then
213852         sParams() = Split(Mid$(Modifier, lPos + 1, InStr(Modifier, ")") - 1 - lPos), ",")
213854         Modifier = Left$(Modifier, lPos - 1)

213856         For i = 0 To UBound(sParams())
213858            sParams(i) = Trim$(sParams(i))

213860            If Left$(sParams(i), 1) = """" And Right$(sParams(i), 1) = """" Then
213862               sParams(i) = Mid$(sParams(i), 2, Len(sParams(i)) - 2)
213864            End If

213866         Next i

213868         If UBound(sParams()) < 3 Then
213870            ReDim Preserve sParams(3)
213872         End If
213874      End If

213876      If LCase(Modifier) = "random" Then
213878         Modifier = PickOne("bracket num2words randomcase reverse obscure piglatin scramble swap") '...
213880      End If

213882      Select Case Trim$(LCase(Modifier))

               Case "a"

213884            If InStr("aeiou", Left$(Word, 1)) Then
213886               sWord = "an " & Word
213888            Else
213890               sWord = "a " & Word
213892            End If

213894         Case "bracket"
213896            sWord = Bracket(Word, sParams(0))

213898         Case "num2word", "num2words"
213900            sWord = Trim$(SentenceCase(NumberAsText(Word)))

213902         Case "reverse"
213904            sWord = StrReverse(Word)

213906         Case "ucase", "uppercase"
213908            sWord = UCase(Word)

213910         Case "lcase", "lowercase"
213912            sWord = LCase(Word)

213914         Case "propercase"
213916            sWord = StrConv(Word, vbProperCase)

213918         Case "sentencecase"
213920            sWord = SentenceCase(Word)

213922         Case "obscure"
213924            sWord = Obscure(Word)

213926         Case "replace"
213928            sWord = Replace(Word, sParams(0), sParams(1))

213930         Case "randomcase"
213932            sWord = RandomCase(Word)

213934         Case "scramble"
213936            sWord = ScrambleWord(Word, Val(sParams(0)))

213938         Case "piglatin"
213940            sWord = PigLatin(Word)

213942         Case "repeat"

213944            If Val(sParams(0)) = 0 Then sParams(0) = 1
213946            sTmp = Word

213948            For i = 1 To Val(sParams(0))
213950               sTmp = sTmp & Word
213952            Next i

213954            sWord = sTmp

213956         Case "right"

213958            If sParams(0) = 0 Then sParams(0) = Len(Word)
213960            sWord = Right$(sWord, sParams(0))

213962         Case "left"

213964            If sParams(0) = 0 Then sParams(0) = Len(Word)
213966            sWord = Left$(Word, sParams(0))

213968         Case "trim"
213970            sWord = Trim$(sParams(0))

213972         Case "format"
213974            sWord = Format(Word, sParams(0), 1)

213976         Case "mid"

213978            If sParams(0) = 0 Then sParams(0) = 1
213980            If sParams(1) = 0 Then sParams(1) = 1
213982            sWord = Mid$(Word, sParams(0), sParams(1))

213984         Case "swap"
                  'Todo: swap by syllable if only one word
213986            sTmpArray = Split(Word, " ")

213988            If UBound(sTmpArray) Then
213990               sTmp = Mid$(sTmpArray(0), 1, 1)
213992               Mid$(sTmpArray(0), 1, 1) = Mid$(sTmpArray(1), 1, 1)
213994               Mid$(sTmpArray(1), 1, 1) = sTmp
213996            End If

213998            sWord = Join(sTmpArray, " ")

214000         Case "romannumeral"
214002            sWord = ToRoman(Val(Word))

214004         Case "hide"
214006            sWord = ""

214008         Case "quote"
214010            sWord = Chr$(34) & Word & Chr$(34)

214012         Case "stutter"
214014            sWord = Stutter(Word)

                  'Hyphenate http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=49325&lngWId=1#zip
                  'Morse Code
                  'Hex
                  'Binary
                  'Octal
               Case Else:
214016            Err.Raise ERR_INVALID_MODIFIER, Err.Source & " (" & Erl & ")", ERR_INVALID_MODIFIER_DESC
214018      End Select

214020      Word = sWord
214022   Next j

214024   ModifyWord = Word

         '<EhFooter>
ExitHere:
         Exit Function

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
         Resume ExitHere
         '</EhFooter>
End Function
   
Private Sub Class_Initialize()
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".Class_Initialize"

      '</EhHeader>
   
         'ENHANCEMENT: Get defaults from the registry
         'Roughly based on current time but clustered around 50
32500    Entropy1 = Format(45 + ((Rand(Now, 0, -3) / Now / 2) * 100), "0")
         'Roughly based on system uptime, slightly clustered around 45-65
32510    Entropy2 = Rand(Abs(sIN(timeGetTime)) * 100, 0, 3)
         'Roughly based on several system and application factors
         'Spread evenly from 1-100
32520    Entropy3 = Int(Abs(sIN(Rand(Screen.Width + Screen.Height + Screen.FontCount + App.Revision + App.ThreadID + Now)) * 100))
         'mvarMinLen = 6
         'mvarMaxLen = 50
         'mvarMinCharsets = 2


      '<EhFooter>
ExitHere:
            Exit Sub

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Sub
   
Private Function GetFileName(File As String) As String
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".GetFileName"

      '</EhHeader>
32500    Dim sFilepath      As String
         '
32510    sFilepath = File

32520    If InStr(File, "..") Then Err.Raise ERR_INVALID_FILENAME, ERRSOURCE, ERR_INVALID_FILENAME_DESC
32530    If LCase(Right$(File, 4)) <> ".txt" Then sFilepath = sFilepath & ".txt"
32540    sFilepath = GetLongFileName(GetFullPath(Me.WordlistDir & "\" & sFilepath))

32550    If Len(Dir(sFilepath)) = 0 Then Err.Raise ERR_FILE_NOT_FOUND, ERRSOURCE, ERR_FILE_NOT_FOUND_DESC
         GetFileName = sFilepath
      '<EhFooter>
ExitHere:
            Exit Function

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Function

Friend Function CountWords(File As String) As Long
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".CountWords"

      '</EhHeader>
32500    Dim sFileData        As String
32510    Dim sPatterns()      As String
32520    Dim lfile            As Long
32530    On Error Resume Next 'if the thing has an error, just return 0
32540    lfile = FreeFile(100)
32550    Open GetFileName(File) For Binary Access Read As #lfile
32560    sFileData = Space$(LOF(lfile))
32570    Get #lfile, , sFileData
32580    Close lfile
32590    sPatterns() = Split(sFileData, vbCrLf)
32600    CountWords = UBound(sPatterns)

      '<EhFooter>
ExitHere:
            Exit Function

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Function
   
Friend Function GetPatternValue(Word As String) As Long
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".GetPatternValue"

      '</EhHeader>
32500    Dim sParams()             As String
32510    Dim lPos                  As Integer
32520    Dim sPlaceholdername      As String
32530    Dim dVal                  As Double
32540    Dim sOptions()            As String
32550    Dim sOptionName           As String
32560    Dim j                     As Integer, i As Integer
   
32570    sPlaceholdername = Word
32580    dVal = 1
         'Strip off the qualifier if any
32590    lPos = InStr(sPlaceholdername, "[")

32600    If lPos Then
32610       sPlaceholdername = Left$(sPlaceholdername, lPos - 1)
32620       dVal = dVal * 2
32630    End If

         '--Check for modifiers
32640    lPos = InStrB(sPlaceholdername, "+")

32650    If lPos Then
32660       sPlaceholdername = Left$(sPlaceholdername, lPos - 1)
32670       sOptions = Split(sPlaceholdername, "+")
32680    End If

         '--Check for paramaters
32690    lPos = InStr(sPlaceholdername, "(")
32700    ReDim Preserve sParams(3)

32710    If lPos Then
32720       sParams() = Split(Mid$(sPlaceholdername, lPos + 1, InStr(sPlaceholdername, ")") - 1 - lPos), ",")
32730       sPlaceholdername = Left$(sPlaceholdername, lPos - 1)

32740       For i = 0 To UBound(sParams())
32750          sParams(i) = Trim$(sParams(i))
32760       Next i

32770       If UBound(sParams()) < 3 Then
32780          ReDim Preserve sParams(3)
32790       End If
32800    End If

32810    Select Case LCase(sPlaceholdername)

            Case "asc": dVal = dVal * 255

32820       Case "chr": dVal = dVal * 255

32830       Case "consonant": dVal = dVal * Len(CONSONANTS)

32840       Case "entropy1": dVal = dVal * 25

32850       Case "entropy2": dVal = dVal * 50

32860       Case "entropy3": dVal = dVal * 100

32870       Case "endpunctuation": dVal = dVal * Len(ENDPUNCTUATION)

32880       Case "keyboard": dVal = dVal * Len(KEYBOARD)

32890       Case "lefthand": dVal = dVal * Len(LEFTHAND)

32900       Case "letter": dVal = dVal * Len(LETTERS)

32910       Case "now": dVal = dVal * 63072000 'About two years in seconds

32920       Case "number"

32930          If Val(sParams(0)) = 0 Then sParams(0) = 10
32940          dVal = dVal * Val(sParams(0))

            Case "numberpattern":

32950          If Val(sParams(0)) = 0 Then sParams(0) = 3
32960          dVal = dVal * (10 ^ Val(sParams(0)))

32970       Case "numrow": dVal = dVal * 10

32980       Case "numrowfull": dVal = dVal * Len(NUMROWFULL)

32990       Case "phonetic": dVal = dVal * 2

33000       Case "pronounceable": dVal = dVal * 1000000000

33010       Case "righthand": dVal = dVal * Len(RIGHTHAND)

33020       Case "row1": dVal = dVal * Len(ROW1)

33030       Case "row2": dVal = dVal * Len(ROW2)

33040       Case "row3": dVal = dVal * Len(ROW3)

33050       Case "row1full": dVal = dVal * Len(ROW1FULL)

33060       Case "row2full": dVal = dVal * Len(ROW2FULL)

33070       Case "row3full": dVal = dVal * Len(ROW3FULL)

33080       Case "sequence"

33090          If Val(sParams(0)) = 0 Then sParams(0) = 3
33100          dVal = dVal * (36 ^ Val(sParams(0)))

33110       Case "smiley": dVal = dVal * 15

33120       Case "symbol": dVal = dVal * Len(SYMBOLS)

33130       Case "vowel": dVal = dVal * Len(VOWELS)

33140       Case "word":  dVal = dVal * CountWords(sParams(0))
33150    End Select

33160    If IsBounded(sOptions) Then

33170       For i = 0 To UBound(sOptions)
33180          ReDim sParams(0)
               'Strip off the qualifier if any
33190          lPos = InStr(sOptions(i), "[")

33200          If lPos Then
33210             sOptions(i) = Left$(sOptions(i), lPos - 1)
33220             dVal = dVal * 2
33230          End If

               '--Check for paramaters
33240          sOptionName = sOptions(i)
33250          lPos = InStr(sOptionName, "(")

33260          If lPos Then
33270             sParams() = Split(Mid$(sOptionName, lPos + 1, InStr(sOptionName, ")") - 1 - lPos), ",")
33280             sOptionName = Left$(sOptionName, lPos - 1)

33290             For j = 0 To UBound(sParams())
33300                sParams(j) = Trim$(sParams(j))
33310             Next j

33320             If UBound(sParams()) < 3 Then
33330                ReDim Preserve sParams(3)
33340             End If
33350          End If

33360          Select Case sOptionName

                  Case "bracket": dVal = dVal * 12

33370             Case "obscure": dVal = dVal * 100

33380             Case "random": dVal = dVal * 300 ' Rough guess

33390             Case "randomcase": dVal = dVal * 10 ' Rough guess

33400             Case "scramble": dVal = dVal * 4 'Rough guess
33410          End Select

33420       Next i

33430    End If

33440    mvarValue = mvarValue * dVal

      '<EhFooter>
ExitHere:
            Exit Function

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Function
   
Public Property Let WordlistDir(ByVal vData As String)
Attribute WordlistDir.VB_Description = "Sets or returns the directory that contains wordlists and the patterns.cfg file."
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".WordlistDir"

         '</EhHeader>

213810   If CheckWordlistDir(vData) Then
213812      mvarWordlistDir = vData
213814      SaveStringSetting "Pafwert", "Settings", "LastWordlistDir", mvarWordlistDir
213816   Else
213818      Err.Raise ERR_WORDLIST_DIR_NOT_FOUND, ERRSOURCE, ERR_WORDLIST_DIR_NOT_FOUND_DESC
213820   End If

         '<EhFooter>
ExitHere:
         Exit Property

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
         Resume ExitHere
         '</EhFooter>
End Property

Public Property Get WordlistDir() As String
         '<EhHeader>
         On Error GoTo ErrHandler
         Const PROC_SOURCE = MODULE_SOURCE & ".WordlistDir"

         '</EhHeader>
213810   Dim sPaths(4)      As String
213812   Dim i              As Long

213814   If Len(mvarWordlistDir) = 0 Then
213816      mvarWordlistDir = GetWordlistDir
213818
213820   End If
         
         WordlistDir = mvarWordlistDir
         
         
         '<EhFooter>
ExitHere:
         Exit Property

ErrHandler:
         Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
         Resume ExitHere
         '</EhFooter>
End Property

   
Public Function CheckPattern(ByVal Pattern As String) As String
Attribute CheckPattern.VB_Description = "Checks the syntax validity of a password pattern."
      '<EhHeader>
      On Error GoTo ErrHandler
      Const PROC_SOURCE = MODULE_SOURCE & ".CheckPattern"

      '</EhHeader>
32500    Dim lResult As Long
32510    Dim pPlaceholders() As Placeholder
32520    Dim i As Long
32530    Dim sTemp As String
32540    Dim iLeftBraceCount As Long
32550    Dim iRightBraceCount As Long
32560    Dim iLeftBracketCount As Long
32570    Dim iRightBracketCount As Long
32580    Dim iLeftParenCount As Long
32590    Dim iRightParenCount As Long

32600    If Len(Pattern) = 0 Then
32610       CheckPattern = "Error: Empty pattern"
32620       Exit Function

32630    End If

         'Replace escaped characters with temp placeholders
32640    Pattern = Replace(Pattern, "\\", "#sla#")
32650    Pattern = Replace(Pattern, "\+", "#pls#")
32660    Pattern = Replace(Pattern, "\{", "#lbr#")
32670    Pattern = Replace(Pattern, "\}", "#rbr#")
32680    Pattern = Replace(Pattern, "\[", "#lba#")
32690    Pattern = Replace(Pattern, "\]", "#rba#")
32700    Pattern = Replace(Pattern, "\(", "#lpa#")
32710    Pattern = Replace(Pattern, "\)", "#rpa#")
32720    Pattern = Replace(Pattern, "\|", "#pip#")

         'Check for unmatched brackets, braces, and parenthesis
32730    For i = 1 To Len(Pattern)

32740       Select Case LCase(Mid(Pattern, i, 1))

               Case "{"
32750             iLeftBraceCount = iLeftBraceCount + 1

32760          Case "}"
32770             iRightBraceCount = iRightBraceCount + 1

32780          Case "["
32790             iLeftBracketCount = iLeftBracketCount + 1

32800          Case "]"
32810             iRightBracketCount = iRightBracketCount + 1

32820          Case "("
32830             iLeftParenCount = iLeftParenCount + 1

32840          Case ")"
32850             iRightParenCount = iRightParenCount + 1
32860       End Select

32870    Next

32880    If iLeftBraceCount <> iRightBraceCount Then
32890       CheckPattern = "Error: Unmatched braces in pattern"
32900       Exit Function

32910    End If

32920    If iLeftBracketCount <> iRightBracketCount Then
32930       CheckPattern = "Error: Unmatched brackets in pattern"
32940       Exit Function

32950    End If

32960    If iLeftParenCount <> iRightParenCount Then
32970       CheckPattern = "Error: Unmatched parenthesis in pattern"
32980       Exit Function

32990    End If
   
33000    On Error Resume Next
33010    lResult = ParsePattern(Pattern, pPlaceholders)

33020    If Err Then
33030       CheckPattern = "Error: " & Err.Description
33040       Exit Function

33050    End If

33060    On Error Resume Next
33070    Err.Clear
33080    sTemp = GeneratePassword(Pattern)

33090    If Err Then
33100       CheckPattern = "Error: " & Err.Description
33110       Exit Function

33120    End If

      '<EhFooter>
ExitHere:
            Exit Function

ErrHandler:
            Err.Raise Err.Number, PROC_SOURCE & "." & MODULE_SOURCE & "." & Erl & "." & Err.Source, Err.Description
            Resume ExitHere
      '</EhFooter>
End Function

Public Property Let Entropy3(ByVal vData As Integer):  mvarEntropy3 = vData: End Property

Public Property Get Entropy3() As Integer: Entropy3 = mvarEntropy3: End Property

Public Property Let Entropy2(ByVal vData As Integer): mvarEntropy2 = vData: End Property

Public Property Get Entropy2() As Integer: Entropy2 = mvarEntropy2: End Property

Public Property Let Entropy1(ByVal vData As Integer): mvarEntropy1 = vData: End Property

Public Property Get Entropy1() As Integer: Entropy1 = mvarEntropy1: End Property

Public Property Let Password(ByVal vData As String): mvarPassword = vData: End Property

Public Property Get Password() As String: Password = mvarPassword: End Property

Public Property Let Pattern(ByVal vData As String): mvarPattern = vData: End Property

Public Property Get Pattern() As String: Pattern = mvarPattern: End Property

Public Property Let Keywords(ByVal vData As String): mvarKeywords = vData: End Property

Public Property Get Keywords() As String: Keywords = mvarKeywords: End Property

Public Property Get LastPattern() As String: LastPattern = mvarLastPattern: End Property

