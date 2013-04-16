VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5625
   ClientLeft      =   6210
   ClientTop       =   3930
   ClientWidth     =   5685
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3882.475
   ScaleMode       =   0  'User
   ScaleWidth      =   5338.509
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5100
      Width           =   1260
   End
   Begin VB.Label lblSmarterPasswords 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XATO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5040
      TabIndex        =   9
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "See the Help file for more information."
      Height          =   255
      Index           =   1
      Left            =   1740
      TabIndex        =   8
      Top             =   3900
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Do you like Pafwert? Read more about the theories behind these patterns in Perfect Passwords."
      Height          =   735
      Index           =   0
      Left            =   1740
      TabIndex        =   7
      Top             =   3120
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   2250
      Left            =   60
      Picture         =   "frmAbout.frx":000C
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label lblPafwert 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pafwert"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Smarter Passwords"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   2025
   End
   Begin VB.Label lblLicense 
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed for non-commercial use."
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1620
      Width           =   3870
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1D6F
      ForeColor       =   &H00000000&
      Height          =   675
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   2520
      Width           =   5430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5224.884
      Y1              =   1408.044
      Y2              =   1408.044
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   4785
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©2006 Mark Burnett, All Rights Reserved."
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   2160
      Width           =   4830
   End
   Begin VB.Image Image1 
      Height          =   5730
      Left            =   0
      Picture         =   "frmAbout.frx":1DF6
      Top             =   -120
      Width           =   5715
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   ' Me.Caption = "About " & "Pafwert"
    lblVersion.Caption = EDITION & " v" & App.Major & "." & App.Minor & "." & App.Revision
   #If Pro Then
      lblLicense.Caption = "Registered for commercial use."
   #Else
      lblLicense.Caption = "Licensed for personal, non-commercial use."
   #End If
      
End Sub

Private Sub Image1_Click()

End Sub
