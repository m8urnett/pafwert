VERSION 5.00
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Object = "{81D621F1-9E9D-4240-9A81-DD63C0382C3D}#5.0#0"; "RMChart.ocx"
Begin VB.Form frmBuilder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PatternBuilder"
   ClientHeight    =   8580
   ClientLeft      =   7725
   ClientTop       =   2370
   ClientWidth     =   8730
   Icon            =   "frmBuilder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   8730
   Begin ButtonPlusCtl.ButtonPlus ButtonPlus1 
      Height          =   375
      Left            =   6720
      TabIndex        =   29
      Top             =   300
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Test"
   End
   Begin VB.TextBox txtPattern 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   360
      Width           =   6555
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7260
      TabIndex        =   22
      Text            =   "20"
      Top             =   780
      Width           =   555
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   60
      TabIndex        =   21
      Top             =   2580
      Width           =   4155
   End
   Begin FramePlusCtl.FramePlus FramePlus2 
      Height          =   1935
      Left            =   60
      TabIndex        =   0
      Top             =   6540
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3413
      BorderStyle     =   5
      BackColorGradient=   14737632
      Style           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Password Statistics"
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Length Distribution"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   48
         Top             =   180
         Width           =   1635
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   14
         Left            =   2820
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "15+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   2820
         TabIndex        =   47
         Top             =   660
         Width           =   255
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   13
         Left            =   2640
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   2640
         TabIndex        =   46
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   12
         Left            =   2460
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   2460
         TabIndex        =   45
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   11
         Left            =   2280
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   2280
         TabIndex        =   44
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   10
         Left            =   2100
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   2100
         TabIndex        =   43
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   9
         Left            =   1920
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   1920
         TabIndex        =   42
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   8
         Left            =   1740
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1740
         TabIndex        =   41
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   7
         Left            =   1560
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1560
         TabIndex        =   40
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   6
         Left            =   1380
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   1380
         TabIndex        =   39
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   5
         Left            =   1200
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   38
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   1020
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1020
         TabIndex        =   37
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   840
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   36
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   660
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   660
         TabIndex        =   35
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   480
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   34
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpLen 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   300
         Top             =   480
         Width           =   135
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   33
         Top             =   660
         Width           =   135
      End
      Begin VB.Shape shpRating 
         BorderColor     =   &H00808080&
         FillColor       =   &H0080FFFF&
         Height          =   135
         Index           =   4
         Left            =   3000
         Top             =   1125
         Width           =   135
      End
      Begin VB.Shape shpRating 
         BorderColor     =   &H00808080&
         FillColor       =   &H0080FFFF&
         Height          =   135
         Index           =   3
         Left            =   2820
         Top             =   1125
         Width           =   135
      End
      Begin VB.Shape shpRating 
         BorderColor     =   &H00808080&
         FillColor       =   &H0080FFFF&
         Height          =   135
         Index           =   2
         Left            =   2640
         Top             =   1125
         Width           =   135
      End
      Begin VB.Shape shpRating 
         BorderColor     =   &H00808080&
         FillColor       =   &H0080FFFF&
         Height          =   135
         Index           =   1
         Left            =   2460
         Top             =   1125
         Width           =   135
      End
      Begin VB.Shape shpRating 
         BorderColor     =   &H00808080&
         FillColor       =   &H0080FFFF&
         Height          =   135
         Index           =   0
         Left            =   2280
         Top             =   1125
         Width           =   135
      End
      Begin VB.Label lblComplexity 
         BackStyle       =   0  'Transparent
         Caption         =   "Average complexity rating:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Average generation time:"
         Height          =   315
         Left            =   240
         TabIndex        =   27
         Top             =   1380
         Width           =   3135
      End
   End
   Begin FramePlusCtl.FramePlus FramePlus1 
      Height          =   1935
      Left            =   4380
      TabIndex        =   1
      Top             =   6540
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3413
      BorderStyle     =   5
      BackColorGradient=   14737632
      Style           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Character Distribution"
      Begin VB.PictureBox picGraph 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   120
         ScaleHeight     =   885
         ScaleWidth      =   3765
         TabIndex        =   11
         Top             =   360
         Width           =   3795
         Begin VB.Line lnGraph 
            BorderColor     =   &H00E0E0E0&
            Index           =   7
            X1              =   0
            X2              =   3795
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line lnGraph 
            BorderColor     =   &H00E0E0E0&
            Index           =   6
            X1              =   0
            X2              =   3795
            Y1              =   840
            Y2              =   840
         End
         Begin VB.Line lnGraph 
            BorderColor     =   &H00E0E0E0&
            Index           =   5
            X1              =   0
            X2              =   3795
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line lnGraph 
            BorderColor     =   &H00E0E0E0&
            Index           =   4
            X1              =   0
            X2              =   3795
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line lnGraph 
            BorderColor     =   &H00E0E0E0&
            Index           =   3
            X1              =   0
            X2              =   3795
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Line lnGraph 
            BorderColor     =   &H00E0E0E0&
            Index           =   2
            X1              =   0
            X2              =   3795
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line lnGraph 
            BorderColor     =   &H00E0E0E0&
            Index           =   1
            X1              =   0
            X2              =   3795
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line lnGraph 
            BorderColor     =   &H00E0E0E0&
            Index           =   0
            X1              =   0
            X2              =   3795
            Y1              =   120
            Y2              =   120
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   2340
         TabIndex        =   2
         Top             =   1380
         Width           =   2340
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E17B59&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   5
            Left            =   1140
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   6
            Top             =   40
            Width           =   135
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H001CDFDF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   4
            Left            =   60
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   5
            Top             =   40
            Width           =   135
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H0054CF66&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   3
            Left            =   60
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   4
            Top             =   270
            Width           =   135
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C74982&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Index           =   0
            Left            =   1140
            ScaleHeight     =   135
            ScaleWidth      =   135
            TabIndex        =   3
            Top             =   270
            Width           =   135
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uppercase"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   10
            Top             =   15
            Width           =   780
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Lowercase"
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   9
            Top             =   15
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Numbers"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Symbols"
            Height          =   195
            Index           =   3
            Left            =   1320
            TabIndex        =   7
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Charcter Distribution"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   53
         Top             =   60
         Width           =   1635
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00BD889B&
         BorderWidth     =   2
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   135
         Y1              =   0
         Y2              =   975
      End
   End
   Begin FramePlusCtl.FramePlus FramePlus3 
      Height          =   3735
      Left            =   4380
      TabIndex        =   12
      Top             =   2640
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6588
      BorderStyle     =   5
      BackColorGradient=   14737632
      Style           =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Character Sets"
      Begin RMChart.RMChartX RMChartX1 
         Height          =   1455
         Left            =   2520
         TabIndex        =   49
         Top             =   180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         RMCBackColor    =   -1
         RMCHeight       =   97
         RMCWidth        =   97
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RMChart.RMChartX RMChartX2 
         Height          =   1455
         Left            =   2520
         TabIndex        =   50
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   2566
         RMCBackColor    =   -1
         RMCHeight       =   97
         RMCWidth        =   97
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Character Set Distribution"
         Height          =   195
         Left            =   180
         TabIndex        =   52
         Top             =   1980
         Width           =   1800
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Character Sets"
         Height          =   195
         Left            =   120
         TabIndex        =   51
         Top             =   60
         Width           =   1050
      End
      Begin VB.Label lbl1Charset 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% use 1 character set"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   2220
         Width           =   1530
      End
      Begin VB.Label lbl4Charsets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% use 4 character sets"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2940
         Width           =   1605
      End
      Begin VB.Label lbl3Charsets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% use 3 character sets"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2700
         Width           =   1605
      End
      Begin VB.Label lbl2Charsets 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% use 2 character sets"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2460
         Width           =   1605
      End
      Begin VB.Label lblUppercase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% contain uppercase letters"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   540
         Width           =   1950
      End
      Begin VB.Label lblNumbers 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% contain numbers"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   780
         Width           =   1335
      End
      Begin VB.Label lblSymbols 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% contain symbols"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1020
         Width           =   1290
      End
      Begin VB.Label lblLowercase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "% contain lowercase letters"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   300
         Width           =   1920
      End
   End
   Begin ButtonPlusCtl.ButtonPlus cmdModifier 
      Height          =   375
      Left            =   2580
      TabIndex        =   30
      Top             =   1620
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackStyle       =   0
      Style           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Modifier"
   End
   Begin ButtonPlusCtl.ButtonPlus cmdWordlist 
      Height          =   375
      Left            =   60
      TabIndex        =   31
      Top             =   1620
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackStyle       =   0
      Style           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Wordlist"
   End
   Begin ButtonPlusCtl.ButtonPlus cmdFunction 
      Height          =   375
      Left            =   1320
      TabIndex        =   32
      Top             =   1620
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BackStyle       =   0
      Style           =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Function"
   End
   Begin VB.Label Label1 
      Caption         =   "Pattern:"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   26
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Passwords:"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   25
      Top             =   2280
      Width           =   1380
   End
   Begin VB.Label Label3 
      Caption         =   "Qty:"
      Height          =   315
      Left            =   6780
      TabIndex        =   24
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "frmBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'
'



Private Sub ButtonPlus1_Click()
       '<EhHeader>
       On Error GoTo ButtonPlus1_Click_Err
       '</EhHeader>
24620  Dim oPass As PafwertLib
24622  Dim i As Long
24624  Dim sPass As String
24626  Dim sKeywords As String
24628  Dim sPattern As String
24630  Dim sPasswords As String
24632  Dim iQty As Long
24634  Dim lErrNumber   As Long
24636  Dim sErrSource   As String
24638  Dim sErrDescription As String
       'Dim gItem As GroupItem
24640  Dim sPatternCheck As String
24642  Dim bFailed As Boolean
24644  Dim iRetryCount As Integer
24646  Dim vLenData() As Variant
24648  Dim lPassLen As Long
24650  Dim lTotalLen As Long
24652  Dim lMinLen As Long
24654  Dim lMaxLen As Long
24656  Dim ChartXunit As Single
24658  Dim ChartYUnit As Single
24660  Dim J As Integer
24662  Dim iChar As Integer
24664  Dim lMaxHeight As Long
24666  Dim lChCount(256) As Long
24668  Dim lTotalLower As Long
24670  Dim lTotalUpper As Long
24672  Dim lTotalNumbers As Long
24674  Dim lTotalSymbols As Long
24676  Dim iCharsetCount(4) As Long
24678  Dim lTotalTime As Single
24680  Dim lTotalScore As Integer
24682  Dim lAvgScore As Integer
   
24684  Set oPass = New PafwertLib
24686  Screen.MousePointer = vbArrowHourglass
24688  List1.Clear
      
24690  iQty = Val(txtQty)
24692  lMinLen = 9999
24694  ReDim vLenData(15)

24696  For i = 0 To 14
24698     shpLen(i).FillColor = vbWhite
24700     shpLen(i).BorderColor = &H808080
24702     shpLen(i).BorderWidth = 1
24704  Next i

24706  Me.Refresh
24708  sPatternCheck = oPass.CheckPattern(txtPattern)

24710  If Len(sPatternCheck) Then
24712     MsgBox sPatternCheck, vbExclamation, "Pattern Error"
24714     Screen.MousePointer = vbNormal
24716     Exit Sub
24718  End If
   
24720  For i = 1 To iQty
24722     iRetryCount = 0

24724     With oPass
24726        .GeneratePassword txtPattern, sKeywords
         
24728        With oPass.Complexity
24730           lPassLen = .PasswordLength
            
24732           If .PasswordLength < lMinLen Then lMinLen = lPassLen
24734           If .PasswordLength > lMaxLen Then lMaxLen = lPassLen
24736           lTotalLen = lTotalLen + lPassLen

24738           For J = 1 To lPassLen
24740              iChar = Asc(Mid$(.Password, J, 1))
24742              lChCount(iChar) = lChCount(iChar) + .CharacterDistribution(iChar)

24744              If lChCount(iChar) > lMaxHeight Then
24746                 lMaxHeight = lChCount(iChar)
24748              End If

24750           Next J

24752           lTotalLower = lTotalLower + Abs(.LowercaseCount > 0)
24754           lTotalUpper = lTotalUpper + Abs(.UppercaseCount > 0)
24756           lTotalNumbers = lTotalNumbers + Abs(.NumbersCount > 0)
24758           lTotalSymbols = lTotalSymbols + Abs(.SymbolsCount > 0)
24760           iCharsetCount(.CharsetCount) = iCharsetCount(.CharsetCount) + 1
24762           lTotalScore = lTotalScore + .Score
24764           lTotalTime = lTotalTime + oPass.TimeTaken
24766        End With

24768     End With
      
24770     If Len(oPass.Password) < 15 Then
24772        vLenData(Len(oPass.Password)) = vLenData(Len(oPass.Password)) + 1
24774     Else
24776        vLenData(15) = vLenData(15) + 1
24778     End If

24780     List1.AddItem oPass.Password

24782  Next
   
24784  For i = 2 To 14
24786     vLenData(i) = (vLenData(i - 1) + vLenData(i) + vLenData(i + 1)) / 3
24788  Next i
   
24790  For i = 1 To 15

24792     Select Case Val(vLenData(i))

             Case 0
24794           shpLen(i - 1).FillColor = vbWhite

24796        Case iQty / iQty
24798           shpLen(i - 1).FillColor = &H80FFFF

24800        Case (iQty / iQty) To ((iQty / 8))
24802           shpLen(i - 1).FillColor = vbYellow
            
24804        Case (iQty / 8) To (iQty / 6)
24806           shpLen(i - 1).FillColor = &H80FF& 'Orange
         
24808        Case (iQty / 6) To (iQty / 5)
24810           shpLen(i - 1).FillColor = &H40C0&

24812        Case (iQty / 5) To (iQty / 4)
24814           shpLen(i - 1).FillColor = vbRed

24816        Case (iQty / 4) To iQty
24818           shpLen(i - 1).FillColor = &HC0&
24820     End Select
      
24822  Next i

24824  If (lTotalLen / iQty) >= 15 Then
24826     shpLen(14).BorderColor = vbBlack
24828  Else
24830     shpLen(Int(lTotalLen / iQty)).BorderColor = vbBlack
24832     shpLen(Int(lTotalLen / iQty)).BorderWidth = 2
24834  End If
   
       'lblShortest.Caption = "Shortest password: " & Str$(lMinLen) & " characters."
       'lblLongest.Caption = "Longest password: " & Str$(lMaxLen) & " characters."
       'lblAverage.Caption = "Average length: " & Format$(lTotalLen / iQty, "0.0") & " characters."
24836  lblLowercase.Caption = Format$(lTotalLower / iQty, "0%") & " contain lowercase letters"
24838  lblUppercase.Caption = Format$(lTotalUpper / iQty, "0%") & " contain uppercase letters"
24840  lblNumbers.Caption = Format$(lTotalNumbers / iQty, "0%") & " contain numbers"
24842  lblSymbols.Caption = Format$(lTotalSymbols / iQty, "0%") & " contain symbols"
   
24844  DoChart RMChartX1, lTotalLower / iQty * 100, lTotalUpper / iQty * 100, lTotalNumbers / iQty * 100, lTotalSymbols / iQty * 100
   
24846  lbl1Charset.Caption = Format$(iCharsetCount(1) / iQty, "0%") & " use 1 character set"
24848  lbl2Charsets.Caption = Format$(iCharsetCount(2) / iQty, "0%") & " use 2 character sets"
24850  lbl3Charsets.Caption = Format$(iCharsetCount(3) / iQty, "0%") & " use 3 character sets"
24852  lbl4Charsets.Caption = Format$(iCharsetCount(4) / iQty, "0%") & " use 4 character sets"
24854  lblTime.Caption = "Average generation time: " & Format$(lTotalTime / iQty, "0.0") & "ms"
   
24856  DoChart RMChartX2, iCharsetCount(1) / iQty * 100, iCharsetCount(2) / iQty * 100, iCharsetCount(3) / iQty * 100, iCharsetCount(4) / iQty * 100
   
24858  lAvgScore = Int(lTotalScore / iQty)
   
24860  For i = 0 To 4

24862     If lAvgScore >= i + 1 Then
24864        shpRating(i).FillStyle = 0
24866     Else
24868        shpRating(i).FillStyle = 1
24870     End If

24872  Next
   
24874  On Error Resume Next
24876  ChartXunit = (picGraph.Width / 257)
24878  ChartYUnit = picGraph.Height / lMaxHeight
   
24880  For i = 1 To 128
24882     Load Line1(i)

24884     With Line1(i)
24886        Set .Container = picGraph
24888        .x1 = i * ChartXunit
24890        .x2 = i * ChartXunit
24892        .Y1 = (picGraph.Height - (lChCount(i)) * ChartYUnit) '+ 50
24894        .Y2 = picGraph.Height - 3
   
24896        Select Case i
   
                Case 48 To 57
24898              .BorderColor = &H54CF66
   
24900           Case 64 To 90
24902              .BorderColor = &H1CDFDF
   
24904           Case 95 To 122
24906              .BorderColor = &HE17B59
   
24908           Case Else
24910              .BorderColor = &HC74982
24912        End Select
   
24914        .ZOrder 0
24916        .Visible = True
24918     End With

24920  Next

24922  Screen.MousePointer = vbNormal
   
       '<EhFooter>
       Exit Sub

ButtonPlus1_Click_Err:
       MsgBox Err.Description & vbCrLf & "in Pafwert.Suggest." & Erl & "." & Err.Source
       Resume Next
       '</EhFooter>
End Sub

Private Sub cmdFunction_DropdownItemClick(ByVal Item As ButtonPlusCtl.DropDownItem)
          '<EhHeader>
          On Error GoTo cmdFunction_DropdownItemClick_Err
          '</EhHeader>
24620    txtPattern.SelText = "{" & Item.Caption & "()" & "}"
          '<EhFooter>
          Exit Sub

cmdFunction_DropdownItemClick_Err:
          MsgBox Err.Description & vbCrLf & _
                 "in frmBuilder.cmdFunction_DropdownItemClick." & Erl
          Resume Next
          '</EhFooter>
End Sub

Private Sub cmdModifier_DropdownItemClick(ByVal Item As ButtonPlusCtl.DropDownItem)
          '<EhHeader>
          On Error GoTo cmdModifier_DropdownItemClick_Err
          '</EhHeader>
24620 Dim lPos As Long
24622 lPos = InStr(txtPattern.SelStart, txtPattern.Text, "}") - 1
24624 If lPos <= 0 Then
24626    lPos = InStrRev(txtPattern.Text, "}") - 1
24628 End If
24630 txtPattern.SelStart = lPos

24632    txtPattern.SelText = "+" & Item.Caption
          '<EhFooter>
          Exit Sub

cmdModifier_DropdownItemClick_Err:
          MsgBox Err.Description & vbCrLf & _
                 "in frmBuilder.cmdModifier_DropdownItemClick." & Erl
          Resume Next
          '</EhFooter>
End Sub

Private Sub cmdWordlist_DropdownItemClick(ByVal Item As ButtonPlusCtl.DropDownItem)
          '<EhHeader>
          On Error GoTo cmdWordlist_DropdownItemClick_Err
          '</EhHeader>
24620    txtPattern.SelText = "{Word(" & Item.Caption & ")}"
   
          '<EhFooter>
          Exit Sub

cmdWordlist_DropdownItemClick_Err:
          MsgBox Err.Description & vbCrLf & _
                 "in frmBuilder.cmdWordlist_DropdownItemClick." & Erl
          Resume Next
          '</EhFooter>
End Sub

Private Sub Form_Load()
          '<EhHeader>
          On Error GoTo Form_Load_Err
          '</EhHeader>
24620 Dim oTmp As PafwertLib
24622 Dim sPath As String
24624 Dim sFilename As String

24626 Set oTmp = New PafwertLib
24628    sPath = oTmp.WordlistDir

24630 With cmdWordlist.DropDownItems

   
24632 sFilename = Dir(sPath & "*.txt")
24634 Do While sFilename <> ""
   
24636    If sFilename <> "." And sFilename <> ".." And Len(sFilename) <> 0 Then
24638       .Add , Left$(sFilename, Len(sFilename) - 4)
24640    End If
24642    sFilename = Dir()
24644 Loop

24646 End With


24648 With cmdModifier.DropDownItems
24650    .Add , "A"
24652    .Add , "Bracket"
24654    .Add , "Format"
24656    .Add , "LCase"
24658    .Add , "Left"
24660    .Add , "Hide"
24662    .Add , "Mid"
24664    .Add , "Num2word"
24666    .Add , "Obscure"
24668    .Add , "PigLatin"
24670    .Add , "ProperCase"
24672    .Add , "Quote"
24674    .Add , "RandomCase"
24676    .Add , "Repeat"
24678    .Add , "Replace"
24680    .Add , "Reverse"
24682    .Add , "Right"
24684    .Add , "RomanNumeral"
24686    .Add , "Scramble"
24688    .Add , "SentenceCase"
24690    .Add , "Stutter"
24692    .Add , "Swap"
24694    .Add , "Trim"
24696    .Add , "UCase"
24698 End With

24700 With cmdFunction.DropDownItems
24702    .Add , "Asc"
24704    .Add , "Chr"
24706    .Add , "Consonant"
24708    .Add , "Entropy1"
24710    .Add , "Entropy2"
24712    .Add , "Entropy3"
24714    .Add , "EndPunctuation"
24716    .Add , "Keyboard"
24718    .Add , "LeftHand"
24720    .Add , "Letter"
24722    .Add , "LongDay"
24724    .Add , "LongMonth"
24726    .Add , "Now"
24728    .Add , "Number"
24730    .Add , "NumberCode"
24732    .Add , "NumberPattern"
24734    .Add , "NumRow"
24736    .Add , "NumRowFull"
24738    .Add , "Ordinal"
24740    .Add , "Phoenetic"
24742    .Add , "Pronounceable"
24744    .Add , "RightHand"
24746    .Add , "Row1"
24748    .Add , "Row1Full"
24750    .Add , "Row2"
24752    .Add , "Row2Full"
24754    .Add , "Row3"
24756    .Add , "Row3Full"
24758    .Add , "Sequence"
24760    .Add , "ShortDay"
24762    .Add , "ShortMonth"
24764    .Add , "Smiley"
24766    .Add , "Space"
24768    .Add , "Symbol"
24770    .Add , "Vowel"
24772 End With




          '<EhFooter>
          Exit Sub

Form_Load_Err:
          MsgBox Err.Description & vbCrLf & _
                 "in frmBuilder.Form_Load." & Erl
          Resume Next
          '</EhFooter>
End Sub

Private Sub txtPattern_KeyPress(KeyAscii As Integer)
          '<EhHeader>
          On Error GoTo txtPattern_KeyPress_Err
          '</EhHeader>
24620    If KeyAscii = 1 Then
24622       txtPattern.SelStart = 0
24624       txtPattern.SelLength = 32767
24626    End If
          '<EhFooter>
          Exit Sub

txtPattern_KeyPress_Err:
          MsgBox Err.Description & vbCrLf & _
                 "in frmBuilder.txtPattern_KeyPress." & Erl
          Resume Next
          '</EhFooter>
End Sub


Sub DoChart(Control As RMChartX, First As Long, Second As Long, Third As Long, Fourth As Long)
          '<EhHeader>
          On Error GoTo DoChart_Err
          '</EhHeader>

24620     Dim nRetval As Long
24622     Dim sTemp As String
 
24624     With Control
24626         .Reset
24628         .RMCBackColor = Default
24630         .RMCStyle = RMC_CTRLSTYLEFLAT
24632         .RMCWidth = 100
24634         .RMCHeight = 100
24636         .RMCBgImage = ""
24638         .Font = "Tahoma"
              '************** Add Region 1 *****************************
24640         .AddRegion
24642         With .Region(1)
24644             .Left = 5
24646             .Top = 5
24648             .Width = -5
24650             .Height = -5
24652             .Footer = ""
                  '************** Add Series 1 to region 1 *******************************
24654             .AddGridlessSeries
24656             With .GridlessSeries
24658                 .SeriesStyle = RMC_PIE_3D
24660                 .Alignment = RMC_FULL
24662                 .ExplodeMode = RMC_EXPLODE_NONE
24664                 .Lucent = True
24666                 .ValueLabelOn = RMC_VLABEL_NONE
24668                 .HatchMode = RMC_HATCHBRUSH_OFF
24670                 .StartAngle = 0
                      '****** Set color values ******
24672                 .SetColorValue 1, RoyalBlue
24674                 .SetColorValue 2, Yellow
24676                 .SetColorValue 3, SpringGreen
24678                 .SetColorValue 4, DarkViolet
                      '****** Set data values ******
24680                 sTemp = First & "*" & Second & "*" & Third & "*" & Fourth
24682                 .DataString = sTemp
24684             End With 'GridLessSeries
24686         End With 'Region(1)
24688         nRetval = .Draw(True)
24690     End With 'RMChartX1
          '<EhFooter>
          Exit Sub

DoChart_Err:
          MsgBox Err.Description & vbCrLf & _
                 "in frmBuilder.DoChart." & Erl
          Resume Next
          '</EhFooter>
End Sub
