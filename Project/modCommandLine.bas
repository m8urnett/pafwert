Attribute VB_Name = "modCommandLine"
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

Sub Main()
      '<EhHeader>
      On Error GoTo Main_Err
      '</EhHeader>
3300  Dim FSO As New FileSystemObject
3310  Dim sIN As TextStream
3320  Dim sOut As TextStream
3330  Dim sErr As TextStream
3340  Dim oPass As PafwertLib
3350  Dim bFailed As Boolean
3360  Dim iRetryCount As Integer
3370  Dim i As Long
      
3380  Set sOut = FSO.GetStandardStream(StdOut)
3390  Set oPass = New PafwertLib
3400  Set sErr = FSO.GetStandardStream(StdErr)
3410  Randomize

      '-------Generate passwords
3420  For i = 0 To 11

3430     iRetryCount = 0

3440     Do

3450        With oPass
3460           bFailed = False
3470           .GeneratePassword
               .Complexity.CheckPassword
3480           iRetryCount = iRetryCount + 1
3490        End With

3500     Loop Until (Not bFailed) Or iRetryCount > 15

3510     If iRetryCount > 15 Then
3520        sErr.WriteLine "A timeout or other error occurred generating passwords with the criteria specified. Try adjusting the complexity options and make sure the wordlist directory is available."
            GoTo ExitHere:
3530     End If

3540     sOut.WriteLine oPass.Password

3550  Next

ExitHere:
      '<EhFooter>
      Exit Sub

Main_Err:
      sErr.WriteLine "Error: " & Err.Description
      Resume ExitHere
      '</EhFooter>
End Sub
