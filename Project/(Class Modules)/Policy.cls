VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Policy"
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
'TODO: set policy through registry so admins can enforce it via group policy
Private mvarMinimumLength As Integer
Private mvarMaximumLength As Integer
Private mvarMinimumCharsets As Integer
Private mvarRequiredChars As String
Private mvarExcludedChars As String
Private mvarCannotStartWithChars As String
Private mvarCannotEndWithChars As String
Private mvarRequireUppercase As Boolean
Private mvarRequireLowercase As Boolean
Private mvarRequireNumbers As Boolean
Private mvarRequireSymbols As Boolean
Private mvarMinimumScore As Integer
Private mvarMinimumEntropyBits As Integer
Private mvarAllowUserData As Boolean

Public Sub SetGroupPolicy()
End Sub

Public Sub GetGroupPolicy()
End Sub

Public Property Let AllowUserData(ByVal vData As Boolean): mvarAllowUserData = vData: End Property
Public Property Get AllowUserData() As Boolean: AllowUserData = mvarAllowUserData: End Property
Public Property Let MinimumEntropyBits(ByVal vData As Integer):  mvarMinimumEntropyBits = vData: End Property
Public Property Get MinimumEntropyBits() As Integer: MinimumEntropyBits = mvarMinimumEntropyBits: End Property
Public Property Let MinimumScore(ByVal vData As Integer):  mvarMinimumScore = vData: End Property
Public Property Get MinimumScore() As Integer: MinimumScore = mvarMinimumScore: End Property
Public Property Let RequireSymbols(ByVal vData As Boolean):  mvarRequireSymbols = vData: End Property
Public Property Get RequireSymbols() As Boolean: RequireSymbols = mvarRequireSymbols: End Property
Public Property Let RequireNumbers(ByVal vData As Boolean):  mvarRequireNumbers = vData: End Property
Public Property Get RequireNumbers() As Boolean: RequireNumbers = mvarRequireNumbers: End Property
Public Property Let RequireLowercase(ByVal vData As Boolean):  mvarRequireLowercase = vData: End Property
Public Property Get RequireLowercase() As Boolean: RequireLowercase = mvarRequireLowercase: End Property
Public Property Let RequireUppercase(ByVal vData As Boolean):  mvarRequireUppercase = vData: End Property
Public Property Get RequireUppercase() As Boolean: RequireUppercase = mvarRequireUppercase: End Property
Public Property Let CannotEndWithChars(ByVal vData As String):  mvarCannotEndWithChars = vData: End Property
Public Property Get CannotEndWithChars() As String: CannotEndWithChars = mvarCannotEndWithChars: End Property
Public Property Let CannotStartWithChars(ByVal vData As String):  mvarCannotStartWithChars = vData: End Property
Public Property Get CannotStartWithChars() As String: CannotStartWithChars = mvarCannotStartWithChars: End Property
Public Property Let ExcludedChars(ByVal vData As String):  mvarExcludedChars = vData: End Property
Public Property Get ExcludedChars() As String: ExcludedChars = mvarExcludedChars: End Property
Public Property Let RequiredChars(ByVal vData As String):  mvarRequiredChars = vData: End Property
Public Property Get RequiredChars() As String: RequiredChars = mvarRequiredChars: End Property
Public Property Let MinimumCharsets(ByVal vData As Integer):  mvarMinimumCharsets = vData: End Property
Public Property Get MinimumCharsets() As Integer: MinimumCharsets = mvarMinimumCharsets: End Property
Public Property Let MaximumLength(ByVal vData As Integer):  mvarMaximumLength = vData: End Property
Public Property Get MaximumLength() As Integer: MaximumLength = mvarMaximumLength: End Property
Public Property Let MinimumLength(ByVal vData As Integer):  mvarMinimumLength = vData: End Property
Public Property Get MinimumLength() As Integer: MinimumLength = mvarMinimumLength: End Property
