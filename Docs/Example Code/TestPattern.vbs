Dim oPaf, Pattern, sPatternCheck
Set oPaf=CreateObject("Pafwert.PafwertLib")

WScript.Echo vbCrLf & "Pafwert Pattern Tester"
WScript.Echo String(22,"-") & vbCrLf

If WScript.Arguments.Count=0 Then
	WScript.Echo "Usage: cscript.exe TestPattern.vbs <pattern> [/random]"
	WScript.Echo "       /random	Use a random pattern from patterns.cfg" & vbCrLf
	WScript.Quit
End If

Pattern=WScript.Arguments(0)

If LCase(WScript.Arguments(0))<>"/random" Then 
	sPatternCheck=oPaf.CheckPattern(Pattern)
	If Len(sPatternCheck) Then
		WScript.Echo sPatternCheck
		WScript.Quit 1
	End If
Else
	Pattern=""
End If

WScript.Echo String(78,"-")
WScript.Echo "Sample Passwords:                    Len  Charsets Upper Lower Numbers Symbols"
WScript.Echo String(78,"-")

For i=1 to 15
	With oPaf
		.GeneratePassword Pattern
		WScript.Echo Left(.Password,37) & Space(37 - Len(Left(.Password,37))) & _
						 Space(3-Len(CStr((Len(.Password))))) & Len(.Password) & _
						 Space(7-Len(.CharsetCount)) & .CharsetCount & _
						 Space(6) & Chk(.HasUpperCase) & _
						 Space(5) & Chk(.HasLowerCase) & _
						 Space(6) & Chk(.HasNumbers) & _
						 Space(7) & Chk(.HasSymbols)
		iPassLen=iPassLen + Len(.Password)
	End With
Next

WScript.Echo vbCrLf
WScript.Echo "Average length: " & Int(iPassLen/15)

Private Function Chk(Value)
   If Value = True Then
      Chk = "X"
   Else
      Chk = " "
   End If
End Function

