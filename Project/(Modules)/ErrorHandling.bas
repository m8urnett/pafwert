Attribute VB_Name = "modErrorHandling"
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

' Define your custom errors here.  Be sure to use numbers
' greater than 512, to avoid conflicts with OLE error numbers.
Public Const MyObjectError1 = 1000
Public Const MyObjectError2 = 1010
Public Const MyObjectErrorN = 1234
Public Const MyUnhandledError = 9999




Private Function GetErrorTextFromResource(ErrorNum As Long) _
          As String
      Dim strMsg As String
      

      ' this function will retrieve an error description from a resource
      ' file (.RES).  The ErrorNum is the index of the string
      ' in the resource file.  Called by RaiseError


      On Error GoTo GetErrorTextFromResourceError
      

      ' get the string from a resource file
      GetErrorTextFromResource = LoadResString(ErrorNum)
      

      Exit Function
      

GetErrorTextFromResourceError:
      

      If Err.Number <> 0 Then
            GetErrorTextFromResource = "An unknown error has occurred!"
      End If
      

End Function


Public Sub RaiseError(ErrorNumber As Long, Source As String)
      Dim strErrorText As String


      'there are a number of methods for retrieving the error
      'message.  The following method uses a resource file to
      'retrieve strings indexed by the error number you are
      'raising.
      strErrorText = GetErrorTextFromResource(ErrorNumber)


      'raise an error back to the client
      Err.Raise vbObjectError + ErrorNumber, Source, strErrorText


End Sub


