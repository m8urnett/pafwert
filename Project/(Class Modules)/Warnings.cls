VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Warnings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"ComplexityError"
Attribute VB_Ext_KEY = "Member0" ,"ComplexityWarning"
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

'local variable to hold collection
Private mCol As Collection
Private mvarComplexityWarning As ComplexityWarning


Public Property Get ComplexityWarning() As ComplexityWarning
    If mvarComplexityWarning Is Nothing Then
        Set mvarComplexityWarning = New ComplexityWarning
    End If


    Set ComplexityWarning = mvarComplexityWarning
End Property


Public Property Set ComplexityWarning(vData As ComplexityWarning)
    Set mvarComplexityWarning = vData
End Property


Public Function Add(ErrorCode As Long, Message As String, Optional sKey As String) As ComplexityWarning
    'create a new object
'    Dim objNewMember As ComplexityError
'    Set objNewMember = New ComplexityError
'
'
'    'set the properties passed into the method
'    objNewMember.ErrorCode = ErrorCode
'    objNewMember.Message = Message
'    If Len(sKey) = 0 Then
'        mCol.Add objNewMember
'    Else
'        mCol.Add objNewMember, sKey
'    End If
'
'
'    'return the object created
'    Set Add = objNewMember
'    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As ComplexityWarning
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
  Set mvarComplexityWarning = Nothing
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

