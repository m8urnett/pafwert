VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComplexityWarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
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

Private mvarErrorCode As Long
Private mvarMessage As String

Public Property Let Message(ByVal vData As String)
    On Error GoTo MessageLetErr

    mvarMessage = vData
    Exit Property


MessageLetErr:
    Call RaiseError(MyUnhandledError, "ComplexityError:Message Property Let")
End Property


Public Property Get Message() As String
    On Error GoTo MessageGetErr
    Message = mvarMessage
    Exit Property


MessageGetErr:
    Call RaiseError(MyUnhandledError, "ComplexityError:Message Property Get")
End Property



Public Property Let ErrorCode(ByVal vData As Long)
    On Error GoTo ErrorCodeLetErr
    mvarErrorCode = vData
    Exit Property


ErrorCodeLetErr:
    Call RaiseError(MyUnhandledError, "ComplexityError:ErrorCode Property Let")
End Property


Public Property Get ErrorCode() As Long
    On Error GoTo ErrorCodeGetErr
    
    ErrorCode = mvarErrorCode
    Exit Property


ErrorCodeGetErr:
    Call RaiseError(MyUnhandledError, "ComplexityError:ErrorCode Property Get")
End Property



