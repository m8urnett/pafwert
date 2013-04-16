VERSION 5.00
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Begin VB.Form frmPassDetails 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Password Details"
   ClientHeight    =   5070
   ClientLeft      =   2130
   ClientTop       =   1035
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   4577
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Password Details"
      Begin VB.Label lblPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Left            =   300
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmPassDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

