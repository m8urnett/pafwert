VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Complexity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Policy"
Attribute VB_Ext_KEY = "Member1" ,"Warnings"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
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

Private mvarWarnings As Warnings
Private mvarScore As Integer
Private mvarLettersCount As Integer
Private mvarUppercaseCount As Integer
Private mvarLowercaseCount As Integer
Private mvarNumbersCount As Integer
Private mvarKeyboardSymbolsCount As Integer
Private mvarSymbolsCount As Integer
Private mvarHighUnicodeCount As Integer
Private mvarVowelsCount As Integer
Private mvarConsonentsCount As Integer
Private mvarNumbersRowCount As Integer
Private mvarFirstRowCount As Integer
Private mvarSecondRowCount As Integer
Private mvarThirdRowCount As Integer
Private mvarLeftHandCount As Integer
Private mvarRightHandCount As Integer
Private mvarRequiredCharactersCount As Integer
Private mvarExcludedCharactersCount As Integer
Private mvarCharacterDistribution(256) As Integer
Private mvarKeyspace As Double
Private mvarTimeTaken As Long
Private mvarPolicy As Policy
Private mvarNonKeyboardSymbolsCount As Integer
Private mvarNonPrintableCharsCount As Integer
Private mvarConfusingCharsCount As Integer
Private mvarCharsetCount As Integer
Private mvarPasswordLength As Integer
Private mvarPassword As String

Friend Property Let Password(ByVal vData As String)
   mvarPassword = vData
   CountCharacters
End Property

Friend Property Get Password() As String
   Password = mvarPassword
End Property

Public Property Get Policy() As Policy
   If mvarPolicy Is Nothing Then
      Set mvarPolicy = New Policy
   End If

   Set Policy = mvarPolicy
End Property

Public Property Set Policy(vData As Policy)
   Set mvarPolicy = vData
End Property

Public Property Get TimeTaken() As Long
   ' Set TimeTaken = mvarTimeTaken
End Property

Public Function CheckPassword(Optional ByVal Password As String) As Boolean
   'TODO: Password not set?
   Dim lfile As Long
   Dim sFileData As String
   Dim sWarnings As String
   Dim sErrors As String
   Dim sPass() As Byte
   Dim sPassword As String
   Dim sPasswordTrimmed As String
   Dim i As Integer
   Dim sChar As String
   Dim sTrivial() As String
   Dim sWords() As String
   Dim iLen As Integer
   Dim iHasLetters As Integer
   Dim iHasNumbers As Integer
   Dim iHasSymbols As Integer
   Dim iHasUppercase As Integer
   Dim iHasLowercase As Integer
   Dim iHasVowels As Integer
   Dim iHasConsonants As Integer
   Dim iHasNumRow As Integer
   Dim iHasRow1 As Integer
   Dim iHasRow2 As Integer
   Dim iHasRow3 As Integer
   Dim iHasLeftHand As Integer
   Dim iHasRightHand As Integer
   Dim sRequiredChars As String
   Dim iHasRequiredChars As Integer
   Dim sExcludedChars As String
   Dim iHasExcludedChars As Integer
   Dim iScore As Integer
   Dim iCharset(256) As Integer
   Dim sFirstChar As String
   Dim sLastChar As String

   'TODO:
   '      user-supplied data (city, username, mother, etc.)
   '      search online dictionary?
   '      similarity to previous pass(es)?
   '      detect sequences by looking at ascii
   
   If Len(Password) = 0 Then
      sPassword = mvarPassword
   Else
      sPassword = Password
   End If
   
   '---Initialize some vars
   sFirstChar = Left$(Password, 1)
   sLastChar = Right$(Password, 1)
   iLen = Len(sPassword)

   '---Check first and last characters
   '      If sFirstChar = " " Or sLastChar = " " Then AddWarning BEGINS_OR_ENDS_WITH_SPACE
   '      If InStr(mvarPolicy.CannotStartWithChars, sFirstChar) Then AddWarning BEGINS_WITH_RESTRICTED_CHAR
   '      If InStr(mvarPolicy.CannotEndWithChars, sLastChar) Then AddWarning ENDS_WITH_RESTRICTED_CHAR
      
   If iLen < mvarPolicy.MinimumLength Then
      AddWarning PASSWORD_TOO_SHORT
      iScore = iScore - 1
   End If

   If Me.CharsetCount = 4 Then iScore = iScore + 1
   
   '--- Less than 15 characters
   If iLen < 15 Then
      If Me.CharsetCount = 1 Then
         AddWarning NOT_ENOUGH_CHARSETS
         iScore = iScore - 1
      End If

      '--- Less than 13 characters
      If iLen < 13 Then

         For i = 1 To 255

            If iCharset(i) > iLen / 2 Then
               AddWarning TOO_MANY_SINGLE_CHARSET
               iScore = iScore - 1
               Exit For
            End If

         Next i

         '
         If iHasVowels >= (iLen * 0.8) Then
            AddWarning TOO_MANY_SINGLE_CHARSET
            iScore = iScore - 1
         End If
      
         sTrivial() = Split("qwerty pass123 passw usern letmein zzzz asas zxzx xcxc sdsd popo lklk kjkj ghgh hghg wewe qwqw pussy master admin abc123 superman mustang harley monkey ranger shadow 1111 fuckme fuckyou football baseball trustno1 xxxx aaaa l3tmein letme1n")

         For i = 0 To UBound(sTrivial())

            If InStr(sPassword, sTrivial(i)) Then
               AddWarning TRIVIAL_WORD
               iScore = iScore - 2
               Exit For
            End If

         Next i
            
         If iLen Mod 2 = 0 Then
            If Mid$(sPassword, 1, iLen / 2) = Mid$(sPassword, iLen / 2 + 1, iLen) Then
               AddWarning TRIVIAL_REPETITION
               iScore = iScore - 2
            End If

         Else

            If Mid$(sPassword, 1, (iLen - 1) / 2) = Mid$(sPassword, (iLen - 1) / 2 + 1, (iLen - 1) / 2) Then
               AddWarning TRIVIAL_REPETITION
               iScore = iScore - 2
            End If
         End If

         sTrivial() = Split("aabb bbcc ccdd qqww eerr aass ssdd zzxx xxcc 123 1969 1999 2000 1234 12345 6969 abc xyz asd qwer !@#$ uiop jkl; sdfg hjkl zxcv vbnm 2345 3456 4567 5678 6789 7890 0987 9876 8765 7654 6543 5432 4321 rewq vcxz fdsa lkjh poiu mnbv 2468 1357 1212 9090 0000 9999 2222")

         For i = 0 To UBound(sTrivial())

            If InStr(sPassword, sTrivial(i)) Then
               AddWarning TRIVIAL_SEQUENCE
               iScore = iScore - 2
               Exit For
            End If

         Next i

      End If
         
      '--- Less than 10 characters
      If iLen < 10 Then
         If IsNumeric(sPassword) Then
            AddWarning ALL_NUMERIC_PASSWORD
            iScore = iScore - 1
         End If
            
         lfile = FreeFile
         Open GetWordlistDir() & "\_dictcheck.dat" For Binary Access Read As #lfile
         sFileData = Space$(LOF(lfile))
         Get #lfile, , sFileData
         Close lfile
      
         '            If InStr(sFileData, UCase(DoubleMetaphone(sPassword))) Then
         '               sWarnings = sWarnings & "Similar to common word, name, or password" & vbCrLf
         '               iScore = iScore - 2
         '            End If
   
         'Only the first char is capitalized
         If Asc(sFirstChar) >= 65 And Asc(sFirstChar) <= 90 And mvarUppercaseCount = 1 Then
            iScore = iScore - 1
         End If
            
         'Anything but first char is capitalized
         If Asc(sFirstChar) < 65 Or Asc(sFirstChar) > 90 And mvarUppercaseCount = 1 Then
            iScore = iScore + 1
         End If

         If IsNumeric(sLastChar) And mvarNumbersCount <= 2 Then
            AddWarning COMMON_PASSWORD_PATTERN
            iScore = iScore - 1
         End If
               
         '--- Less than 9 characters
         If iLen < 9 Then
            iScore = iScore - (8 - iLen)

            If Me.CharsetCount = 1 Then
               iScore = iScore - 1
            End If
                                    
         End If
      End If
   End If

   If iLen > 15 Then iScore = iScore + 3
   If Me.CharsetCount = 4 Then iScore = iScore + 2
      
   Debug.Print "Score: " & Password, iScore
ExitHere:
   Exit Function
ErrHandler:
   Err.Raise Err.Number, ERRSOURCE & " (" & Erl & ")", Err.Description
   Resume ExitHere

End Function

Public Property Get Keyspace() As Double
   Dim lKeySpace As Integer

   If Me.UppercaseCount Then lKeySpace = lKeySpace + 26
   If Me.LowercaseCount Then lKeySpace = lKeySpace + 26
   If Me.KeyboardSymbolsCount Then lKeySpace = lKeySpace + 33
   If Me.NonKeyboardSymbolsCount Then lKeySpace = lKeySpace + 96
   If Me.NonPrintableCharsCount Then lKeySpace = lKeySpace + 37
   If Me.NumbersCount Then lKeySpace = lKeySpace + 10
   If Me.HighUnicodeCount Then lKeySpace = lKeySpace + 65000
   
   mvarKeyspace = (lKeySpace ^ Len(mvarPassword))
   
   

   Keyspace = mvarKeyspace
End Property

Public Property Get Score() As Integer
Attribute Score.VB_UserMemId = 0
   On Error GoTo ScoreGetErr
   If Score = 0 Then
      Select Case Me.Keyspace
      Case Is < 1E+15
         mvarScore = 1
      Case Is < 1E+18
         mvarScore = 2
      Case Is < 1E+21
         mvarScore = 3
      Case Is < 1E+24
         mvarScore = 4
      Case Else
         mvarScore = 5
      End Select
   End If
   Score = mvarScore
   Exit Property

ScoreGetErr:
   Call RaiseError(MyUnhandledError, "Complexity:Score Property Get")
End Property

Public Property Get Warnings() As Warnings
   On Error GoTo WarningsGetErr

   If mvarWarnings Is Nothing Then
      Set mvarWarnings = New Warnings
   End If

   Set Warnings = mvarWarnings
   Exit Property

WarningsGetErr:
   Call RaiseError(MyUnhandledError, "Complexity:Warnings Property Get")
End Property

Public Property Set Warnings(vData As Warnings)
   On Error GoTo WarningsSetErr
   Set mvarWarnings = vData
   Exit Property

WarningsSetErr:
   Call RaiseError(MyUnhandledError, "Complexity:Warnings Property Set")
End Property

Private Sub Class_Terminate()
   Set mvarPolicy = Nothing
   Set mvarWarnings = Nothing
   'Set mvarErrors = Nothing
End Sub

Private Sub CountCharacters(Optional Password As String)
       Dim i As Integer
       Dim sChar As String
       Dim sPass() As Byte

102    If Len(Password) = 0 Then Password = mvarPassword
104    sPass() = Password
106    mvarPasswordLength = UBound(sPass) / LenB("A")
   
        
108    For i = 0 To UBound(sPass) Step LenB("A")
      
110       Select Case sPass(i)

                'Lowercase
             Case 97 To 122
112             mvarLettersCount = mvarLettersCount + 1
114             mvarLowercaseCount = mvarLowercaseCount + 1
            
                'Uppercase
116          Case 65 To 90
118             mvarLettersCount = mvarLettersCount + 1
120             mvarUppercaseCount = mvarUppercaseCount + 1
         
                'Numbers
122          Case 48 To 57
124             mvarNumbersCount = mvarNumbersCount + 1
         
                'Non-printable
126          Case 1 To 31, 127, 129, 141, 143, 144, 157
128             mvarNonPrintableCharsCount = mvarNonPrintableCharsCount + 1
130             mvarSymbolsCount = mvarSymbolsCount + 1
         
                'Keyboard Symbols
132          Case 32 To 47, 91 To 96, 123 To 127
134             mvarKeyboardSymbolsCount = mvarKeyboardSymbolsCount + 1
136             mvarSymbolsCount = mvarSymbolsCount + 1
            
                'Non-keyboard symbols
138          Case 128 To 255
140             mvarNonKeyboardSymbolsCount = mvarNonKeyboardSymbolsCount + 1
142             mvarSymbolsCount = mvarSymbolsCount + 1
            
144          Case Else
146             mvarSymbolsCount = mvarSymbolsCount + 1
            
          End Select
          mvarCharacterDistribution(sPass(i)) = mvarCharacterDistribution(sPass(i)) + 1
148       sChar = Chr$(sPass(i))
          
150       If InStrB(VOWELS, sChar) Then mvarVowelsCount = mvarVowelsCount + 1
152       If InStrB(CONSONANTS, sChar) Then mvarConsonentsCount = mvarConsonentsCount + 1
154       If InStrB(NUMROWFULL, sChar) Then mvarNumbersRowCount = mvarNumbersRowCount + 1
156       If InStrB(ROW1FULL, sChar) Then mvarFirstRowCount = mvarFirstRowCount + 1
158       If InStrB(ROW2FULL, sChar) Then mvarSecondRowCount = mvarSecondRowCount + 1
160       If InStrB(ROW3FULL, sChar) Then mvarThirdRowCount = mvarThirdRowCount + 1
162       If InStrB(LEFTHAND, sChar) Then mvarLeftHandCount = mvarLeftHandCount + 1
164       If InStrB(RIGHTHAND, sChar) Then mvarRightHandCount = mvarRightHandCount + 1
166       If InStrB(CONFUSING, sChar) Then mvarConfusingCharsCount = mvarConfusingCharsCount + 1
168      ' If InStrB(mvarPolicy.RequiredChars, sChar) Then mvarRequiredCharactersCount = mvarRequiredCharactersCount + 1
170      ' If InStrB(mvarPolicy.ExcludedChars, sChar) Then mvarExcludedCharactersCount = mvarExcludedCharactersCount + 1
       Next
   
174    mvarCharsetCount = Abs(mvarLowercaseCount > 0) + Abs(mvarUppercaseCount > 0) + Abs(mvarNumbersCount > 0) + Abs(mvarSymbolsCount > 0)


End Sub

Private Sub AddWarning(ErrCode As Long)
   Dim sMessage As String

   '
   '   Select Case ErrCode
   '      Case BEGINS_OR_ENDS_WITH_SPACE
   '         sMessage = BEGINS_OR_ENDS_WITH_SPACE_DESC
   '      Case PASSWORD_TOO_SHORT
   '         sMessage = PASSWORD_TOO_SHORT_DESC
   '      Case ALL_NUMERIC_PASSWORD
   '         sMessage = ALL_NUMERIC_PASSWORD_DESC
   '      Case TRIVIAL_WORD
   '         sMessage = TRIVIAL_WORD_DESC
   '      Case TRIVIAL_SEQUENCE
   '         sMessage = TRIVIAL_SEQUENCE_DESC
   '      Case TRIVIAL_REPETITION
   '         sMessage = TRIVIAL_REPETITION_DESC
   '      Case TOO_MANY_SINGLE_CHARSET
   '         sMessage = TOO_MANY_SINGLE_CHARSET_DESC
   '      Case WORDLIST_MATCH
   '         sMessage = WORDLIST_MATCH_DESC
   '      Case PARTIAL_WORDLIST_MATCH
   '         sMessage = PARTIAL_WORDLIST_MATCH_DESC
   '      Case REVERSED_WORDLIST_MATCH
   '         sMessage = REVERSED_WORDLIST_MATCH_DESC
   '      Case NOT_ENOUGH_CHARSETS
   '         sMessage = NOT_ENOUGH_CHARSETS_DESC
   '      Case TOO_MANY_SINGLE_CHARACTER
   '         sMessage = TOO_MANY_SINGLE_CHARACTER_DESC
   '      Case SIMILAR_WORDLIST_MATCH
   '         sMessage = SIMILAR_WORDLIST_MATCH_DESC
   '      Case MISSING_REQUIRED_CHAR
   '         sMessage = MISSING_REQUIRED_CHAR_DESC
   '      Case CONTAINS_RESTRICTED_CHAR
   '         sMessage = CONTAINS_RESTRICTED_CHAR_DESC
   '      Case CONTAINS_USER_DATA
   '         sMessage = CONTAINS_USER_DATA_DESC
   '      Case COMMON_PASSWORD_PATTERN
   '         sMessage = COMMON_PASSWORD_PATTERN_DESC
   '   End Select
   
   Me.Warnings.Add ErrCode, sMessage
   
End Sub

Public Property Get PasswordLength() As Integer
   PasswordLength = Len(Me.Password)
End Property

Public Property Get CharacterDistribution() As Variant: CharacterDistribution = mvarCharacterDistribution: End Property

Public Property Get ExcludedCharactersCount() As Integer: ExcludedCharactersCount = mvarExcludedCharactersCount: End Property

Public Property Get RequiredCharactersCount() As Integer: RequiredCharactersCount = mvarRequiredCharactersCount: End Property

Public Property Get RightHandCount() As Integer: RightHandCount = mvarRightHandCount: End Property

Public Property Get LeftHandCount() As Integer: LeftHandCount = mvarLeftHandCount: End Property

Public Property Get ThirdRowCount() As Integer: ThirdRowCount = mvarThirdRowCount: End Property

Public Property Get SecondRowCount() As Integer: SecondRowCount = mvarSecondRowCount: End Property

Public Property Get FirstRowCount() As Integer: FirstRowCount = mvarFirstRowCount: End Property

Public Property Get NumbersRowCount() As Integer: NumbersRowCount = mvarNumbersRowCount: End Property

'Public Property Get ConsonentsCount() As Integer: Set ConsonentsCount = mvarConsonentsCount: End Property

Public Property Get VowelsCount() As Integer: VowelsCount = mvarVowelsCount: End Property

Public Property Get HighUnicodeCount() As Integer: HighUnicodeCount = mvarHighUnicodeCount: End Property

Public Property Get KeyboardSymbolsCount() As Integer: KeyboardSymbolsCount = mvarKeyboardSymbolsCount: End Property

Public Property Get NumbersCount() As Integer: NumbersCount = mvarNumbersCount: End Property

Public Property Get LowercaseCount() As Integer: LowercaseCount = mvarLowercaseCount: End Property

Public Property Get UppercaseCount() As Integer: UppercaseCount = mvarUppercaseCount: End Property

Public Property Get LettersCount() As Integer: LettersCount = mvarLettersCount: End Property

Public Property Get CharsetCount() As Integer: CharsetCount = mvarCharsetCount: End Property

Public Property Get SymbolsCount() As Integer: SymbolsCount = mvarSymbolsCount: End Property

Public Property Get ConfusingCharsCount() As Integer: ConfusingCharsCount = mvarConfusingCharsCount: End Property

Public Property Get NonPrintableCharsCount() As Integer: NonPrintableCharsCount = mvarNonPrintableCharsCount: End Property

Public Property Get NonKeyboardSymbolsCount() As Integer: NonKeyboardSymbolsCount = mvarNonKeyboardSymbolsCount: End Property

