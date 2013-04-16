Attribute VB_Name = "modCommandLine"
Option Explicit

Sub Main()
         '<EhHeader>
         On Error GoTo Main_Err
         '</EhHeader>
4620  Dim FSO As New FileSystemObject
4622  Dim sOut As TextStream
4624  Dim sErr As TextStream
4626  Dim oPass As New PafwertLib
4628  Dim bFailed As Boolean
4630  Dim iRetryCount As Integer
4632  Dim i As Long
4634  Dim sPatternCheck As String
4636  Dim lTotalLen As Long
4638  Dim lTotalScore As Long
4640  Dim lTotalTime As Long
4642  Dim lQTY  As Long
      Dim sPattern As String
      
4644  Set sOut = FSO.GetStandardStream(StdOut)
4646  Set sErr = FSO.GetStandardStream(StdErr)

4648  sOut.WriteLine vbCrLf & "Pafwert Pattern Tester v" & App.Major & "." & App.Minor & vbCrLf

4650  Randomize
4652  lQTY = 10
   
sPattern = Command$
sPattern = TrimOffBeginning(sPattern, Chr$(34))
sPattern = TrimOffEnd(sPattern, Chr$(34))


4654     If Len(sPattern) = 0 Then
4656     sOut.WriteLine "Usage: "
4658     sOut.WriteLine "  " & App.EXEName & " <pattern>"
4660     Exit Sub
4662     End If

4664  sPatternCheck = oPass.CheckPattern(Command$)

4666  If Len(sPatternCheck) Then
4668     sOut.WriteLine sPatternCheck

4670     Exit Sub
4672  End If

4674  sOut.WriteLine
4676  sOut.WriteLine "Sample Passwords"
4678  sOut.WriteLine "----------------"
      '-------Generate passwords
4680  For i = 0 To lQTY - 1

4682     iRetryCount = 0

4684     Do

4686        With oPass
4688           bFailed = False
4690           .GeneratePassword sPattern
               
4692           iRetryCount = iRetryCount + 1
4694        End With

4696     Loop Until (Not bFailed) Or iRetryCount > 15

4698     If iRetryCount > 15 Then
4700        sErr.WriteLine "A timeout or other error occurred generating passwords with the criteria specified. Try adjusting the complexity options and make sure the wordlist directory is available."
            GoTo ExitHere:
4702     End If

4704     lTotalLen = lTotalLen + Len(oPass.Password)
4706     lTotalScore = lTotalScore + oPass.Complexity.Score
4708     lTotalTime = lTotalTime + oPass.TimeTaken
4710     sOut.WriteLine "  " & oPass.Password

4712  Next

4714  sOut.WriteLine vbCrLf
4716  sOut.WriteLine "Statistics"
4718  sOut.WriteLine "----------"
4720  sOut.WriteLine "  Average length: " & Format(lTotalLen / lQTY, "#.0") & " characters"
4722  sOut.WriteLine "  Average time:   " & Format$(lTotalTime / lQTY, "0.0") & "ms"
4724  sOut.Write "  Average score:  "

4726  For i = 1 To Int(lTotalScore / lQTY)
4728     sOut.Write "*"
4730  Next i

4732  sOut.WriteLine vbCrLf

ExitHere:
         '<EhFooter>
         Exit Sub

Main_Err:
         sOut.WriteLine Err.Description & " in modCommandLine.Main." & Erl
         Resume ExitHere
         '</EhFooter>
End Sub



Public Function TrimOffBeginning(ByVal Data As String, _
                                 ByVal Characters As String) As String
11000    Dim X As Long
11010    X = Len(Characters)

11020    If LCase(Left$(Data, X)) = LCase(Characters) Then
11030       TrimOffBeginning = Right$(Data, Len(Data) - X)
11040    Else
11050       TrimOffBeginning = Data
11060    End If

End Function

Public Function TrimOffEnd(ByVal Data As String, _
                           ByVal Characters As String) As String
   
11000    Dim X As Long
11010    X = Len(Characters)

11020    If LCase(Right(Data, X)) = LCase(Characters) Then
11030       TrimOffEnd = Left$(Data, Len(Data) - X)
11040    Else
11050       TrimOffEnd = Data
11060    End If

End Function



