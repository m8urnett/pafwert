VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pafwert"
   ClientHeight    =   4905
   ClientLeft      =   2310
   ClientTop       =   2625
   ClientWidth     =   7395
   FillColor       =   &H00FF8080&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   3400
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin FramePlusCtl.FramePlus FramePlus7 
      Align           =   1  'Align Top
      Height          =   510
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   900
      BackgroundPicture=   "frmMain.frx":2B8A
      BackgroundPictureAlignment=   10
      Style           =   0
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
      Begin ButtonPlusCtl.ButtonPlus cmdHelp 
         Height          =   220
         Left            =   9900
         TabIndex        =   19
         ToolTipText     =   "Help"
         Top             =   135
         Width           =   220
         _ExtentX        =   397
         _ExtentY        =   397
         BackStyle       =   0
         FocusStyle      =   0
         Picture         =   "frmMain.frx":5C34
         BorderStyle     =   0
         UseVisualStyles =   0   'False
         BackColor       =   0
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
      Begin ButtonPlusCtl.ButtonPlus cmdSuggest 
         Height          =   495
         Left            =   240
         TabIndex        =   0
         Top             =   10
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         BackStyle       =   0
         ThemeColor      =   4484202
         WordWrap        =   0   'False
         Picture         =   "frmMain.frx":5F86
         BorderStyle     =   0
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Suggest..."
      End
      Begin ButtonPlusCtl.ButtonPlus cmdAbout 
         Height          =   225
         Left            =   10200
         TabIndex        =   20
         ToolTipText     =   "Help"
         Top             =   135
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   397
         BackStyle       =   0
         FocusStyle      =   0
         Picture         =   "frmMain.frx":9030
         BorderStyle     =   0
         UseVisualStyles =   0   'False
         BackColor       =   0
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
      Begin ButtonPlusCtl.ButtonPlus cmdOptions 
         Default         =   -1  'True
         Height          =   495
         Left            =   1500
         TabIndex        =   22
         Top             =   10
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         BackStyle       =   0
         ThemeColor      =   4484202
         WordWrap        =   0   'False
         Picture         =   "frmMain.frx":9382
         BorderStyle     =   0
         ForeColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Options..."
      End
   End
   Begin FramePlusCtl.FramePlus FramePlus6 
      Align           =   1  'Align Top
      Height          =   1200
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   2117
      BackgroundPicture=   "frmMain.frx":C42C
      BackgroundPictureAlignment=   10
      BackColor       =   12506837
      Style           =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Desktop Edition v2.1"
      Begin VB.Label lblCopyright 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright ©2006 M. Burnett"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   5400
         TabIndex        =   6
         Top             =   960
         Width           =   2070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Smarter Passwords"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lblVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desktop Edition v2.1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   5640
         TabIndex        =   4
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblPafwert 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pafwert"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   120
         Width           =   1725
      End
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   11
      Left            =   3660
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   3400
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   9
      Left            =   3660
      TabIndex        =   10
      Top             =   2160
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   10
      Left            =   3660
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   6
      Left            =   3660
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Visible         =   0   'False
      Width           =   3400
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   7
      Left            =   3660
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   3400
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   3400
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   3400
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin ButtonPlusCtl.ButtonPlus Password 
      Height          =   315
      Index           =   8
      Left            =   3660
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   556
      BackStyle       =   0
      HotEffects      =   -1  'True
      FocusStyle      =   0
      WordWrap        =   0   'False
      CaptionAlignment=   3
      BorderStyle     =   0
      Appearance      =   0
      UseVisualStyles =   0   'False
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
   End
   Begin VB.Label lblDesktopEdition 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   1485
   End
   Begin VB.Menu mnuPassword 
      Caption         =   "Password Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy &Password"
      End
      Begin VB.Menu mnuCopySelected 
         Caption         =   "Copy &Selected"
      End
      Begin VB.Menu mnuCopyAll 
         Caption         =   "Copy &All"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Option Explicit
Private Declare Function HtmlHelp _
                Lib "hhctrl.ocx" _
                Alias "HtmlHelpA" (ByVal hwndCaller As Long, _
                                   ByVal pszFile As String, _
                                   ByVal uCommand As Long, _
                                   ByVal dwData As Long) As Long

Const HH_DISPLAY_TOPIC = &H0
Const HH_SET_WIN_TYPE = &H4
Const HH_GET_WIN_TYPE = &H5
Const HH_GET_WIN_HANDLE = &H6
Const HH_DISPLAY_TEXT_POPUP = &HE   ' Display string resource ID or text in a pop-up window.
Const HH_HELP_CONTEXT = &HF         ' Display mapped numeric value in dwData.
Const HH_TP_HELP_CONTEXTMENU = &H10 ' Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU.
Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help, similar to WinHelp's HELP_WM_HELP.

Private Const MODULE_NAME As String = "UI"
Private Const MAX_RETRIES As Long = 25
Private Const BORDER_WIDTH = 300



Private Sub cmdHelp_Click()
   '<EhHeader>
   On Error GoTo cmdHelp_Click_Err
   '</EhHeader>
   Dim hwndHH As Long
   hwndHH = HtmlHelp(0, App.Path & "\" & "pafwert.chm", HH_DISPLAY_TOPIC, 0)
   '<EhFooter>
   Exit Sub

cmdHelp_Click_Err:
   MsgBox Err.Description & vbCrLf & "in Pafwert.frmMain.cmdHelp_Click " & "at line " & Erl
   Resume Next
   '</EhFooter>
End Sub

Private Sub cmdOptions_Click()
   frmOptions.Show vbModal, Me
   
End Sub

Private Sub cmdSuggest_Click()
      '<EhHeader>
      On Error GoTo cmdSuggest_Click_Err
      '</EhHeader>
      Dim i As Long
      Dim sKeywords As String
      Dim sPattern As String
      Dim oPass As PafwertLib
      Dim bFailed As Boolean
      Dim iRetryCount As Integer

      Const PROCNAME As String = "cmdSuggest_Click"
      'On Error GoTo ErrHandler
      '---Set UI stuff
100   Set oPass = New PafwertLib

102   Screen.MousePointer = vbArrowHourglass
104   sKeywords = ""

106   Randomize

      Dim jj As Integer

      'For jj = 1 To Int(Rnd * Rnd * Rnd * 7) + 1
      '-------Generate passwords
108   For i = 0 To 11

         '         DoEvents
110      iRetryCount = 0

         Do

112         With oPass
114            bFailed = False
116            .GeneratePassword
               '.GeneratePassword sPattern, sKeywords
               
118            iRetryCount = iRetryCount + 1
            End With

120      Loop Until (Not bFailed) Or iRetryCount > MAX_RETRIES

122      If iRetryCount > MAX_RETRIES Then
124         MsgBox "A timeout or other error occurred generating passwords with the criteria specified. Try adjusting the complexity options and make sure the wordlist directory is available.", vbInformation
            GoTo ExitHere:
         End If

126      With Password(i)
128         .Caption = Trim$(oPass.Password)
130         .Width = Me.TextWidth(oPass.Password & "________")
            '.ToolTipText = oPass.LastPattern
132         .Visible = True
         End With

      Next

      'Next
 
ExitHere:
134   Screen.MousePointer = vbNormal

136   cmdSuggest.SetFocus

      Exit Sub


cmdSuggest_Click_Err:
      MsgBox "An error occurred while generating your passwords: " & vbCrLf & Err.Number & ": " & Err.Description & vbCrLf & "Location: Suggest." & Erl & "." & Err.Source
      Resume ExitHere
      '</EhFooter>
End Sub
'

Private Sub Form_Load()
      '<EhHeader>
      On Error GoTo Form_Load_Err
      '</EhHeader>
      Dim i As Integer
100   lblVer.Caption = EDITION & " v" & App.Major & "." & App.Minor & "." & App.Revision

102   For i = 0 To 11

104      With Password(i)
106         .DropDownItems.Add "Copy", "Copy Password"
108         .DropDownItems.Add "CopyAll", "Copy All Passwords"
110         .DropDownItems.Add "ClearList", "Clear List"
         End With

112   Next i


      #If Pro Then
         cmdOptions.Enabled = True
         cmdOptions.Visible = True
      #End If
      
      
      '<EhFooter>
      Exit Sub

Form_Load_Err:
      MsgBox Err.Description & vbCrLf & "in Pafwert.frmMain.Form_Load " & "at line " & Erl
      Resume Next
      '</EhFooter>
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
'   Call HtmlHelp(Me.hWnd, "", HH_CLOSE_ALL, 0)
End Sub

Private Sub Form_Resize()
   '<EhHeader>
   On Error Resume Next
   '</EhHeader>
   Dim i As Integer
   Dim iWidth As Integer
   On Error Resume Next
   lblCopyright.Left = Me.Width - lblCopyright.Width - BORDER_WIDTH
   lblVer.Left = Me.Width - lblVer.Width - BORDER_WIDTH
   cmdHelp.Left = Me.Width - cmdHelp.Width - BORDER_WIDTH
   cmdAbout.Left = cmdHelp.Left - cmdAbout.Width - (BORDER_WIDTH / 2)

   ' cmdCopy.Left = Me.Width - BORDER_WIDTH * 5 - cmdHelp.Width - BORDER_WIDTH
   
   iWidth = Int((Me.ScaleWidth / 2) - (BORDER_WIDTH * 3))

   For i = 0 To 5
      Password(i).Left = BORDER_WIDTH
      Password(i).Width = Me.TextWidth(Password(i).Caption & "___________") ' iWidth
   Next

   For i = 6 To 11
      Password(i).Left = iWidth + (BORDER_WIDTH * 3)
      Password(i).Width = Me.TextWidth(Password(i).Caption & "___________")  'iWidth
   Next
   
End Sub

'Public Property Get WordlistDir() As String
'1000    Dim sPaths(6)      As String
'1005    Dim i              As Long
'1010    Dim Dir    As String
'        'If Len(mvarWordlistDir) = 0 Then
'1015    sPaths(0) = GetFullPath(GetStringSetting("Pafwert", "Settings", "WordlistDir") & "\")
'1020    sPaths(1) = GetFullPath(App.Path & "\wordlists\")
'1025    sPaths(2) = GetFullPath(GetStringSetting("Pafwert", "Settings", "LastWordlistDir") & "\")
'1030    sPaths(3) = GetFullPath("C:\Program Files\Pafwert\Wordlists" & "\")
'1035    sPaths(4) = GetFullPath(CurDir & "\")
'1040    sPaths(5) = GetFullPath(App.Path & "\..\wordlists\")
'1045    sPaths(6) = GetFullPath(Command$ & "\")
'
'1050    For i = 0 To 6
'1055       Debug.Print i, sPaths(i)
'
'1060       If CheckWordlistDir(sPaths(i)) Then
'1065          Dir = sPaths(i)
'1070          Exit For
'1075       End If
'
'1080    Next
'
'        ' TODO: Prompt for wordlist dir
'        ' TODO: Have a command-line param for wordlist dir
'
'1085    If Len(Dir) = 0 Then Err.Raise ERR_WORDLIST_DIR_NOT_FOUND, ERRSOURCE, ERR_WORDLIST_DIR_NOT_FOUND_DESC
'        'SaveStringSetting "Pafwert", "Settings", "WordlistDir", mvarWordlistDir
'        'End If
'1090    WordlistDir = Dir
'End Property

'Private Function CheckWordlistDir(Path As String) As Boolean
'1000    On Error Resume Next
'
'1005    If Len(Dir(Path)) Then
'1010       If FileLen(Path & "\patterns.cfg") = 0 Or Err Then
'1015          Err.Raise ERR_OPEN_PATTERNS_FILE, ERRSOURCE, ERR_OPEN_PATTERNS_FILE_DESC
'1020       Else
'1025          CheckWordlistDir = True
'1030       End If
'1035    End If
'
'End Function

Private Sub Password_DropdownItemClick(Index As Integer, _
                                       ByVal Item As ButtonPlusCtl.DropDownItem)
      '<EhHeader>
      On Error GoTo Password_DropdownItemClick_Err
      '</EhHeader>
      Dim i As Integer
      Dim sPasswords As String

100   Select Case Item.Key

         Case "Copy"
102         Screen.MousePointer = vbArrowHourglass
104         Clipboard.Clear
106         Clipboard.SetText Password(Index).Caption
108         Screen.MousePointer = vbNormal
      
110      Case "CopyAll"
         
112         Screen.MousePointer = vbArrowHourglass

114         For i = 0 To 11
116            sPasswords = sPasswords & Password(i).Caption & vbCrLf
            Next

118         Clipboard.Clear
120         Clipboard.SetText sPasswords
122         Screen.MousePointer = vbNormal
      
124      Case "ClearList"

126         For i = 0 To 11
128            Password(i).Caption = ""
130            Password(i).Visible = False
            Next

      End Select

      '<EhFooter>
      Exit Sub

Password_DropdownItemClick_Err:
      MsgBox Err.Description & vbCrLf & "in Pafwert.frmMain.Password_DropdownItemClick " & "at line " & Erl
      Resume Next
      '</EhFooter>
End Sub

Private Sub cmdAbout_Click()
      '<EhHeader>
      On Error GoTo cmdAbout_Click_Err
      '</EhHeader>
100   frmAbout.Show vbModal, Me
      '<EhFooter>
      Exit Sub

cmdAbout_Click_Err:
      MsgBox Err.Description & vbCrLf & "in Pafwert.frmMain.cmdAbout_Click " & "at line " & Erl
      Resume Next
      '</EhFooter>
End Sub

Private Sub Password_MouseEnter(Index As Integer)
      '<EhHeader>
      On Error GoTo Password_MouseEnter_Err
      '</EhHeader>
100   Password(Index).Style = bpStyleDropdownRight

      '<EhFooter>
      Exit Sub

Password_MouseEnter_Err:
      MsgBox Err.Description & vbCrLf & "in Pafwert.frmMain.Password_MouseEnter " & "at line " & Erl
      Resume Next
      '</EhFooter>
End Sub

Private Sub Password_MouseExit(Index As Integer)
      '<EhHeader>
      On Error GoTo Password_MouseExit_Err
      '</EhHeader>
100   Password(Index).Style = bpStyleStandard
      '<EhFooter>
      Exit Sub

Password_MouseExit_Err:
      MsgBox Err.Description & vbCrLf & "in Pafwert.frmMain.Password_MouseExit " & "at line " & Erl
      Resume Next
      '</EhFooter>
End Sub

