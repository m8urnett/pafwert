VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Options"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5460
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ButtonPlusCtl.ButtonPlus ButtonPlus1 
      Height          =   315
      Left            =   4200
      TabIndex        =   22
      Top             =   5640
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      BackStyle       =   0
      Picture         =   "frmOptions.frx":2B8A
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   180
      TabIndex        =   21
      Text            =   "c:\program files\pafwert\wordlists"
      Top             =   5640
      Width           =   3915
   End
   Begin VB.ComboBox cboScore 
      Height          =   315
      ItemData        =   "frmOptions.frx":305E
      Left            =   3420
      List            =   "frmOptions.frx":3071
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1160
      Width           =   675
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Generate any password"
      Height          =   255
      Index           =   3
      Left            =   180
      TabIndex        =   17
      Top             =   600
      Value           =   -1  'True
      Width           =   3555
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4140
      TabIndex        =   12
      Top             =   600
      Width           =   1155
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4140
      TabIndex        =   11
      Top             =   60
      Width           =   1155
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Custom requirements"
      Height          =   255
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   1500
      Width           =   3555
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Must have a complexity score of at least"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   1200
      Width           =   3555
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Must meet Windows complexity requirements"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   900
      Width           =   3555
   End
   Begin FramePlusCtl.FramePlus fraCustom 
      Height          =   3315
      Left            =   480
      TabIndex        =   13
      Top             =   1740
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      Begin VB.TextBox txtMinLen 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1860
         TabIndex        =   3
         Text            =   "8"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtMaxLen 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1860
         TabIndex        =   4
         Text            =   "30"
         Top             =   600
         Width           =   375
      End
      Begin VB.CheckBox chkSymbols 
         Caption         =   "Symbols and Punctuation"
         Enabled         =   0   'False
         Height          =   255
         Left            =   660
         TabIndex        =   10
         Top             =   2880
         Width           =   2295
      End
      Begin VB.CheckBox chkNumbers 
         Caption         =   "Numbers"
         Enabled         =   0   'False
         Height          =   255
         Left            =   660
         TabIndex        =   9
         Top             =   2580
         Width           =   1095
      End
      Begin VB.CheckBox chkLower 
         Caption         =   "Lowercase letters"
         Enabled         =   0   'False
         Height          =   255
         Left            =   660
         TabIndex        =   8
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CheckBox chkUpper 
         Caption         =   "Uppercase letters"
         Enabled         =   0   'False
         Height          =   255
         Left            =   660
         TabIndex        =   7
         Top             =   1980
         Width           =   1815
      End
      Begin VB.CheckBox chkRequire 
         Caption         =   "All passwords must contain:"
         Height          =   255
         Left            =   300
         TabIndex        =   6
         Top             =   1680
         Width           =   2715
      End
      Begin VB.TextBox txtMinCharsets 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1860
         TabIndex        =   5
         Text            =   "2"
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Minimum Length:"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   16
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M&aximum Length:"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   15
         Top             =   660
         Width           =   1245
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum number of character sets:"
         Height          =   495
         Left            =   300
         TabIndex        =   14
         Top             =   1140
         Width           =   1635
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wordlist Directory:"
      Height          =   195
      Left            =   180
      TabIndex        =   20
      Top             =   5400
      Width           =   1290
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Note that complexity requirements may slow down password generation."
      Height          =   435
      Left            =   180
      TabIndex        =   19
      Top             =   60
      Width           =   3615
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private Sub chkRequire_Click()
   If chkRequire.Value = vbChecked Then
      chkUpper.Enabled = True
      chkLower.Enabled = True
      chkNumbers.Enabled = True
      chkSymbols.Enabled = True
   Else
      chkUpper.Enabled = False
      chkLower.Enabled = False
      chkNumbers.Enabled = False
      chkSymbols.Enabled = False
   End If
   
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub Option1_Click(Index As Integer)
   If Index = 2 Then
      fraCustom.Enabled = True
   Else
      fraCustom.Enabled = False
   End If
   
   If Index = 1 Then
      cboScore.Enabled = True
   Else
      cboScore.Enabled = False
   End If
End Sub
