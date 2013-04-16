Attribute VB_Name = "modNumText"
'Module Name:   NumberAsText
'Programmer:    Frederick Rothstein
'Date Released: May 16, 1999
'Date Modified: August 20, 1999
'Copyright 1999 by Frederick N. Rothstein (All rights reserved)

Option Explicit
Dim NumberText()      As String

'Public Function NumberAsText(NumberIn As Variant, Optional _
AND_or_CHECK_or_DOLLAR As String) As String
Public Function NumberAsText(NumberIn As Variant) As String
   Dim i                   As Integer
   Dim DecimalPoint        As Integer
   Dim CardinalNumber      As Long
   Dim TestValue           As Long
   Dim CurrValue           As Currency
   Dim CentsString         As String
   Dim NumberSign          As String
   Dim WholePart           As String
   Dim BigWholePart        As String
   Dim DecimalPart         As String
   Dim UseAnd              As Boolean
   Dim UseCheck            As Boolean
   Dim UseDollars          As Boolean
   Dim CommaAdjuster       As Long
   Dim tmp                 As String
   Dim sStyle              As String
   
   '----------------------------------------
   'Begin setting conditions for formatting
   '----------------------------------------
  
   'Determine whether to apply special formatting.
   'If nothing passed, return routine result
   'converted only into its numeric equivalents,
   'with no additional format text.
   'sStyle = lcase(AND_or_CHECK_or_DOLLAR)

   'User passed "AND": "and" will be added between
   'hundreths and tens of dollars,
   'ie "Three Hundred and Fourty Two"
   'UseAnd = sStyle = "and"
   
   'User passed "DOLLAR": "dollar(s)" and "cents"
   'appended to string,
   'ie "Three Hundred and Fourty Two Dollars"
   'UseDollars = sStyle = "dollar"
   
   'User passed "CHECK" *or* "DOLLAR"
   'If "check", cent amount returned as a fraction /100
   'ie "Three Hundred Forty Two and 00/100"
   'If "dollar" was passed, "dollar(s)" and "cents"
   'appended instead.
   'UseCheck = (sStyle = "check") Or (sStyle = "dollar")
    
   '----------------------------------------
   'Check/create array. If this is the first
   'time using this routine, create the text
   'strings that will be used.
   '----------------------------------------
   If Not IsBounded(NumberText) Then
      Call BuildArray(NumberText)
   End If

   '----------------------------------------
   'Begin validating the number, and breaking
   'into constituent parts
   '----------------------------------------
  
   'prepare to check for valid value in
   NumberIn = Trim$(NumberIn)
   
   If Not IsNumeric(NumberIn) Then

      'invalid entry - abort
      NumberAsText = "Error - Number improperly formed"
      Exit Function

   Else

      'decimal check
      DecimalPoint = InStr(NumberIn, ".")

      If DecimalPoint > 0 Then

         'split the fractional and primary numbers
         DecimalPart = Mid$(NumberIn, DecimalPoint + 1)
         WholePart = Left$(NumberIn, DecimalPoint - 1)

      Else
      
         'assume the decimal is the last chr
         DecimalPoint = Len(NumberIn) + 1
         WholePart = NumberIn
         
      End If

      If InStr(NumberIn, ",,") Or _
         InStr(NumberIn, ",.") Or _
         InStr(NumberIn, ".,") Or _
         InStr(DecimalPart, ",") Then

         NumberAsText = "Error - Improper use of commas"
         Exit Function

      ElseIf InStr(NumberIn, ",") Then

         CommaAdjuster = 0
         WholePart = ""

         For i = DecimalPoint - 1 To 1 Step -1

            If Not Mid$(NumberIn, i, 1) Like "[,]" Then

               WholePart = Mid$(NumberIn, i, 1) & WholePart

            Else

               CommaAdjuster = CommaAdjuster + 1

               If (DecimalPoint - i - CommaAdjuster) Mod 3 Then

                  NumberAsText = "Error - Improper use of commas"
                  Exit Function

               End If 'If (DecimalPoint - i - CommaAdjuster)

            End If  'If Not Mid$(NumberIn, i, 1) Like
            
         Next  'For i = DecimalPoint - 1
         
      End If  'If InStr(NumberIn, ",,")
      
   End If  'If Not IsNumeric(NumberIn)
    
   If Left$(WholePart, 1) Like "[+-]" Then
      NumberSign = IIf(Left$(WholePart, 1) = "-", "Minus ", "Plus ")
      WholePart = Mid$(WholePart, 2)
   End If
   
   '----------------------------------------
   'Begin code to assure decimal portion of
   'check value is not inadvertantly rounded
   '----------------------------------------
   '   If UseCheck = True Then
   '
   '      CurrValue = CCur(Val("." & DecimalPart))
   '      DecimalPart = Mid$(Format$(CurrValue, "0.00"), 3, 2)
   '
   '      If CurrValue >= 0.995 Then
   '
   '         If WholePart = String$(Len(WholePart), "9") Then
   '
   '            WholePart = "1" & String$(Len(WholePart), "0")
   '
   '         Else
   '
   '            For i = Len(WholePart) To 1 Step -1
   '
   '              If Mid$(WholePart, i, 1) = "9" Then
   '                 Mid$(WholePart, i, 1) = "0"
   '              Else
   '                 Mid$(WholePart, i, 1) = CStr(Val(Mid$(WholePart, i, 1)) + 1)
   '                 Exit For
   '              End If
   '
   '            Next
   '
   '         End If 'If WholePart = String$(
   '      End If   'If CurrValue >= 0.995
   '   End If     'If UseCheck = True
    
   '----------------------------------------
   'Final prep step - this assures number
   'within range of formatting code below
   '----------------------------------------
   If Len(WholePart) > 9 Then
      BigWholePart = Left$(WholePart, Len(WholePart) - 9)
      WholePart = Right$(WholePart, 9)
   End If
    
   If Len(BigWholePart) > 9 Then
   
      NumberAsText = "Error - Number too large"
      Exit Function
       
   ElseIf Not WholePart Like String$(Len(WholePart), "#") Or _
      (Not BigWholePart Like String$(Len(BigWholePart), "#") _
      And BigWholePart <> "") Then
          
      NumberAsText = "Error - Number improperly formed"
      Exit Function
     
   End If

   '----------------------------------------
   'Begin creating the output string
   '----------------------------------------
    
   'Very Large values
   TestValue = Val(BigWholePart)
    
   If TestValue > 999999 Then
      CardinalNumber = TestValue \ 1000000
      tmp = HundredsTensUnits(CardinalNumber) & "Quadrillion "
      TestValue = TestValue - (CardinalNumber * 1000000)
   End If
   
   If TestValue > 999 Then
      CardinalNumber = TestValue \ 1000
      tmp = tmp & HundredsTensUnits(CardinalNumber) & "Trillion "
      TestValue = TestValue - (CardinalNumber * 1000)
   End If
   
   If TestValue > 0 Then
      tmp = tmp & HundredsTensUnits(TestValue) & "Billion "
   End If
   
   'Lesser values
   TestValue = Val(WholePart)
   
   If TestValue = 0 And BigWholePart = "" Then tmp = "Zero "
   
   If TestValue > 999999 Then
      CardinalNumber = TestValue \ 1000000
      tmp = tmp & HundredsTensUnits(CardinalNumber) & "Million "
      TestValue = TestValue - (CardinalNumber * 1000000)
   End If
    
   If TestValue > 999 Then
      CardinalNumber = TestValue \ 1000
      tmp = tmp & HundredsTensUnits(CardinalNumber) & "Thousand "
      TestValue = TestValue - (CardinalNumber * 1000)
   End If
    
   If TestValue > 0 Then
      If Val(WholePart) < 99 And BigWholePart = "" Then UseAnd = False
      tmp = tmp & HundredsTensUnits(TestValue, UseAnd)
   End If
    
   '  'If in dollar mode, assure the text is the correct plurality
   '   If UseDollars = True Then
   '
   '      CentsString = HundredsTensUnits(DecimalPart)
   '
   '      If tmp = "One " Then
   '            tmp = tmp & "Dollar"
   '      Else: tmp = tmp & "Dollars"
   '      End If
   '
   '      If CentsString <> "" Then
   '
   '         tmp = tmp & " and " & CentsString
   '
   '         If CentsString = "One " Then
   '               tmp = tmp & "Cent"
   '         Else: tmp = tmp & "Cents"
   '         End If
   '
   '      End If
      
   'ElseIf UseCheck = True Then
      
   'tmp = tmp & "and " & Left$(DecimalPart & "00", 2)
   ' tmp = tmp & "/100"
    
   'Else
    
   If DecimalPart <> "" Then
        
      tmp = tmp & "Point"
        
      For i = 1 To Len(DecimalPart)
         tmp = tmp & " " & NumberText(Mid$(DecimalPart, i, 1))
      Next
      
   End If  'If DecimalPart <> ""
   '   End If   'If UseDollars = True
    
   'done!
   NumberAsText = NumberSign & tmp
    
End Function

Private Function IsBounded(vntArray As Variant) As Boolean
 
   'note: the application in the IDE will stop
   'at this line when first run if the IDE error
   'mode is not set to "Break on Unhandled Errors"
   '(Tools/Options/General/Error Trapping)
   On Error Resume Next
   IsBounded = IsNumeric(UBound(vntArray))
   
End Function

Private Sub BuildArray(NumberText() As String)

   ReDim NumberText(0 To 27) As String
 
   NumberText(0) = "Zero"
   NumberText(1) = "One"
   NumberText(2) = "Two"
   NumberText(3) = "Three"
   NumberText(4) = "Four"
   NumberText(5) = "Five"
   NumberText(6) = "Six"
   NumberText(7) = "Seven"
   NumberText(8) = "Eight"
   NumberText(9) = "Nine"
   NumberText(10) = "Ten"
   NumberText(11) = "Eleven"
   NumberText(12) = "Twelve"
   NumberText(13) = "Thirteen"
   NumberText(14) = "Fourteen"
   NumberText(15) = "Fifteen"
   NumberText(16) = "Sixteen"
   NumberText(17) = "Seventeen"
   NumberText(18) = "Eighteen"
   NumberText(19) = "Nineteen"
   NumberText(20) = "Twenty"
   NumberText(21) = "Thirty"
   NumberText(22) = "Forty"
   NumberText(23) = "Fifty"
   NumberText(24) = "Sixty"
   NumberText(25) = "Seventy"
   NumberText(26) = "Eighty"
   NumberText(27) = "Ninety"
   
End Sub

Public Function HundredsTensUnits(ByVal TestValue As Integer, _
Optional UseAnd As Boolean) As String

   Dim CardinalNumber      As Integer
    
   If TestValue > 99 Then
      CardinalNumber = TestValue \ 100
      HundredsTensUnits = NumberText(CardinalNumber) & " Hundred "
      TestValue = TestValue - (CardinalNumber * 100)
   End If
    
   If UseAnd = True Then
      HundredsTensUnits = HundredsTensUnits & "and "
   End If
   
   If TestValue > 20 Then
      CardinalNumber = TestValue \ 10
      HundredsTensUnits = HundredsTensUnits & _
      NumberText(CardinalNumber + 18) & " "
      TestValue = TestValue - (CardinalNumber * 10)
   End If
    
   If TestValue > 0 Then
      HundredsTensUnits = HundredsTensUnits & NumberText(TestValue) & " "
   End If

End Function
