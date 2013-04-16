VERSION 5.00
Object = "{78F87B0F-C7C3-4096-9287-F5A4383498E1}#1.31#0"; "ExpBar1.ocx"
Object = "{A19332D7-D707-4A30-9F38-796D120AF5B3}#1.2#0"; "BtnPlus1.ocx"
Object = "{2BF72F7D-D367-4712-9146-5533EF3E691A}#1.2#0"; "FraPlus1.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pafwert"
   ClientHeight    =   10275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9885
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin ExplorerBarCtl.ExplorerBar expPasses 
      Align           =   3  'Align Left
      Height          =   9060
      Left            =   4485
      TabIndex        =   5
      Top             =   1215
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   15981
      EmptyGroupHeight=   50
      HighlightBackColor=   33023
      BackgroundStyle =   3
      BackColorAlt    =   16777215
      HeaderBackColor =   12632256
      GroupBackColor  =   16777215
      GroupBackColorAlt=   14737632
      BorderStyle     =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GroupsCount     =   2
      Group1Caption   =   "Suggested Passwords"
      Group1Id        =   "{78BCC245-D7BC-474A-A3B4-978D8FE0D69B}"
      Group1Key       =   "passwords"
      Group1HeaderBackgroundStyle=   1
      Group1BackgroundStyle=   2
      Group1EmptyText =   "Click on the Suggest button to generate passwords."
      Group2Caption   =   "Password Statistics"
      Group2Id        =   "{3C91B162-435C-4752-9D52-83F91F575B2A}"
      Group2Key       =   "passwordstatistics"
      Group2Visible   =   0   'False
      ImageWidth      =   16
      ImageHeight     =   16
      ImageItemCount  =   1
      ImageItem1Key   =   "bullet"
      ImageItem1Tag   =   "C:\Program Files\Innovasys\ExplorerBar\Demonstration\Icons\bullet.bmp"
      ImageItem1Id    =   "{E9893275-1CFB-4567-AC8B-286CC24E182C}"
      ImageItem1Picture=   "frmMain.frx":17002
      HotStyleMousePointer=   99
      LicenceData     =   "Unlicensed"
      Begin VB.Frame fraPassStats 
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   4875
         Begin FramePlusCtl.FramePlus FramePlus1 
            Height          =   1995
            Left            =   0
            TabIndex        =   7
            Top             =   1260
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   3519
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Character Frequency"
            Begin VB.PictureBox picGraph 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   1090
               Left            =   120
               ScaleHeight     =   1065
               ScaleWidth      =   3765
               TabIndex        =   17
               Top             =   240
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
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   495
               Left            =   960
               ScaleHeight     =   465
               ScaleWidth      =   2205
               TabIndex        =   8
               Top             =   1380
               Width           =   2235
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
                  TabIndex        =   12
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
                  TabIndex        =   11
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
                  TabIndex        =   10
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
                  TabIndex        =   9
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
                  TabIndex        =   16
                  Top             =   15
                  Width           =   780
               End
               Begin VB.Label Label4 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lowercase"
                  Height          =   255
                  Index           =   2
                  Left            =   1320
                  TabIndex        =   15
                  Top             =   10
                  Width           =   1275
               End
               Begin VB.Label Label4 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Numbers"
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   14
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
                  TabIndex        =   13
                  Top             =   240
                  Width           =   585
               End
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
         Begin FramePlusCtl.FramePlus FramePlus2 
            Height          =   1155
            Left            =   480
            TabIndex        =   18
            Top             =   0
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   2037
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Password Lengths"
            Begin VB.Label lblAverage 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Average length:"
               Height          =   195
               Left            =   180
               TabIndex        =   23
               Top             =   720
               Width           =   1125
            End
            Begin VB.Label lblLongest 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Longest password:"
               Height          =   195
               Left            =   180
               TabIndex        =   22
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label lblShortest 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Shortest password:"
               Height          =   195
               Left            =   180
               TabIndex        =   21
               Top             =   240
               Width           =   1350
            End
         End
         Begin FramePlusCtl.FramePlus FramePlus3 
            Height          =   2355
            Left            =   0
            TabIndex        =   19
            Top             =   3300
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   4154
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
            Begin VB.Label lbl1Charset 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% of the passwords use 1 character set."
               Height          =   195
               Left            =   180
               TabIndex        =   31
               Top             =   1320
               Width           =   2820
            End
            Begin VB.Label lbl4Charsets 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% of the passwords use 4 character sets."
               Height          =   195
               Left            =   180
               TabIndex        =   30
               Top             =   2040
               Width           =   2895
            End
            Begin VB.Label lbl3Charsets 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% of the passwords use 3 character sets."
               Height          =   195
               Left            =   180
               TabIndex        =   29
               Top             =   1800
               Width           =   2895
            End
            Begin VB.Label lbl2Charsets 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% of the passwords use 2 character sets."
               Height          =   195
               Left            =   180
               TabIndex        =   28
               Top             =   1560
               Width           =   2895
            End
            Begin VB.Label lblUppercase 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% of the passwords contain uppercase letters."
               Height          =   195
               Left            =   180
               TabIndex        =   27
               Top             =   480
               Width           =   3240
            End
            Begin VB.Label lblNumbers 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% of the passwords contain numbers."
               Height          =   195
               Left            =   180
               TabIndex        =   26
               Top             =   720
               Width           =   2625
            End
            Begin VB.Label lblSymbols 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% of the passwords contain symbols."
               Height          =   195
               Left            =   180
               TabIndex        =   25
               Top             =   960
               Width           =   2580
            End
            Begin VB.Label lblLowercase 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "% of the passwords contain lowercase letters."
               Height          =   195
               Left            =   180
               TabIndex        =   24
               Top             =   240
               Width           =   3210
            End
         End
         Begin FramePlusCtl.FramePlus FramePlus4 
            Height          =   975
            Left            =   0
            TabIndex        =   20
            Top             =   5760
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   1720
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Other Stats"
            Begin VB.Shape shpRating 
               BorderColor     =   &H00C0C0C0&
               Height          =   135
               Index           =   4
               Left            =   3000
               Top             =   590
               Width           =   135
            End
            Begin VB.Shape shpRating 
               BorderColor     =   &H00C0C0C0&
               Height          =   135
               Index           =   3
               Left            =   2820
               Top             =   590
               Width           =   135
            End
            Begin VB.Shape shpRating 
               BorderColor     =   &H00C0C0C0&
               Height          =   135
               Index           =   2
               Left            =   2640
               Top             =   590
               Width           =   135
            End
            Begin VB.Shape shpRating 
               BorderColor     =   &H00C0C0C0&
               Height          =   135
               Index           =   1
               Left            =   2460
               Top             =   590
               Width           =   135
            End
            Begin VB.Shape shpRating 
               BorderColor     =   &H00C0C0C0&
               FillColor       =   &H0080FFFF&
               FillStyle       =   0  'Solid
               Height          =   135
               Index           =   0
               Left            =   2280
               Top             =   590
               Width           =   135
            End
            Begin VB.Label lblComplexity 
               BackStyle       =   0  'Transparent
               Caption         =   "Average complexity rating:"
               Height          =   255
               Left            =   240
               TabIndex        =   33
               Top             =   540
               Width           =   2055
            End
            Begin VB.Label lblTime 
               BackStyle       =   0  'Transparent
               Caption         =   "Average time:"
               Height          =   315
               Left            =   240
               TabIndex        =   32
               Top             =   240
               Width           =   3135
            End
         End
      End
   End
   Begin VB.PictureBox picCover 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4815
      TabIndex        =   4
      Top             =   0
      Width           =   4815
   End
   Begin VB.PictureBox picCover 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   0
      Left            =   3300
      ScaleHeight     =   15
      ScaleWidth      =   4815
      TabIndex        =   3
      Top             =   1320
      Width           =   4815
   End
   Begin ExplorerBarCtl.ExplorerBar ExplorerBar1 
      Align           =   3  'Align Left
      Height          =   9060
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   2
      Top             =   1215
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   15981
      EmptyGroupHeight=   60
      HighlightStyle  =   1
      BackColorAlt    =   16777215
      ItemGap         =   0
      PlaySounds      =   0   'False
      HeaderBackColorAlt=   16777215
      HeaderBackColor =   12506837
      GroupBackColor  =   12506837
      GroupBackColorAlt=   16777215
      BackColor       =   4484202
      ScrollbarStyle  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Margin          =   10
      InnerMargin     =   6
      Gap             =   6
      HeaderHeight    =   18
      GroupsCount     =   6
      Group1Caption   =   "Pafwert"
      Group1Id        =   "{5CFBEBBC-7C71-4CA8-98F8-FF2A2DA286C6}"
      Group1Key       =   "pafwert"
      Group1HeaderStyle=   0
      Group2Caption   =   "--Based on a pattern"
      Group2Expanded  =   0   'False
      Group2Id        =   "{EF7ABB48-35DE-4771-A9AF-DF222F5A6EEE}"
      Group2HeaderImage=   "SymbolConfiguration2"
      Group2Key       =   "basedonapattern"
      Group2Visible   =   0   'False
      Group2HeaderBackgroundStyle=   3
      Group3Caption   =   "Policy requirements"
      Group3Expanded  =   0   'False
      Group3Id        =   "{07346D21-BDB1-46CE-BBF8-4BF31BF83642}"
      Group3HeaderImage=   "SecurityEnabled"
      Group3Key       =   "basedoncomplexityconstraints"
      Group3HeaderBackgroundStyle=   3
      Group3BackgroundStyle=   3
      Group4Caption   =   "Modify randomness"
      Group4Expanded  =   0   'False
      Group4Id        =   "{1E3BD5FC-98A0-40EC-9EB5-EC1126DD1EC5}"
      Group4HeaderImage=   "Dice"
      Group4Key       =   "basedonrandomness"
      Group4HeaderBackgroundStyle=   3
      Group4BackgroundStyle=   3
      Group5Caption   =   "Other Options"
      Group5Expanded  =   0   'False
      Group5Id        =   "{BC9DC205-8533-42BB-8260-73D8F090217C}"
      Group5HeaderImage=   "SymbolConfiguration2"
      Group5Key       =   "otheroptions"
      Group5HeaderBackgroundStyle=   3
      Group5BackgroundStyle=   3
      Group6Caption   =   "Password Tasks"
      Group6Id        =   "{516EAE9D-179D-484B-A85F-7E34124D0386}"
      Group6HeaderImage=   "Menu2"
      Group6HeaderBackgroundStyle=   3
      Group6BackgroundStyle=   3
      Group6GroupItemsCount=   3
      Group6Item1Caption=   "Export to file..."
      Group6Item1Id   =   "{AA2DC790-8AA6-4608-AF59-AEF8BD36DD57}"
      Group6Item1Indent=   4
      Group6Item1Key  =   "exporttofile"
      Group6Item2Caption=   "Copy to clipboard"
      Group6Item2Id   =   "{07C6BC9E-BAAB-484D-875A-08245C41B39B}"
      Group6Item2Indent=   4
      Group6Item2Key  =   "copytoclipboard"
      Group6Item3Caption=   "Clear list"
      Group6Item3Id   =   "{A5EED8DE-AF30-49BA-8C24-B1F7096B7955}"
      Group6Item3Indent=   4
      Group6Item3Key  =   "clearlist"
      HeaderImageWidth=   24
      HeaderImageHeight=   24
      HeaderImageItemCount=   11
      HeaderImageItem1Key=   "SymbolRefresh3"
      HeaderImageItem1Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Symbol Refresh 3.bmp"
      HeaderImageItem1Id=   "{83EADB09-6A17-4306-92C8-10B4D899A033}"
      HeaderImageItem1Picture=   "frmMain.frx":17354
      HeaderImageItem2Key=   "Copy"
      HeaderImageItem2Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Copy.bmp"
      HeaderImageItem2Id=   "{C4C035AB-F950-4E71-A333-A83EB442A8D7}"
      HeaderImageItem2Picture=   "frmMain.frx":17A66
      HeaderImageItem3Key=   "Folder2Configuration"
      HeaderImageItem3Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Folder 2 Configuration.bmp"
      HeaderImageItem3Id=   "{D5A8CC52-B01C-439F-88E6-764264F1E636}"
      HeaderImageItem3Picture=   "frmMain.frx":18178
      HeaderImageItem4Key=   "Help"
      HeaderImageItem4Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Help.bmp"
      HeaderImageItem4Id=   "{CEDDA140-5EB4-4920-BE14-D4690A94B648}"
      HeaderImageItem4Picture=   "frmMain.frx":1888A
      HeaderImageItem5Key=   "PieChart"
      HeaderImageItem5Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Pie Chart.bmp"
      HeaderImageItem5Id=   "{7A63A269-44A0-4C3E-A626-7F55DCF4FA2A}"
      HeaderImageItem5Picture=   "frmMain.frx":18F9C
      HeaderImageItem6Key=   "BarChart5"
      HeaderImageItem6Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Bar Chart 5.bmp"
      HeaderImageItem6Id=   "{4289FB8D-4F6B-4670-B719-B49D68550BD3}"
      HeaderImageItem6Picture=   "frmMain.frx":196AE
      HeaderImageItem7Key=   "SaveCheck"
      HeaderImageItem7Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Save Check.bmp"
      HeaderImageItem7Id=   "{0BA42EB9-74EA-4F41-B856-182648163837}"
      HeaderImageItem7Picture=   "frmMain.frx":19DC0
      HeaderImageItem8Key=   "Dice"
      HeaderImageItem8Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Dice.bmp"
      HeaderImageItem8Id=   "{7B5F3207-011D-4A3D-B0EF-733AA4507806}"
      HeaderImageItem8Picture=   "frmMain.frx":1A4D2
      HeaderImageItem9Key=   "SecurityEnabled"
      HeaderImageItem9Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Security Enabled.bmp"
      HeaderImageItem9Id=   "{66A1F067-DDC0-4AA7-8EFD-B6C89AC364D3}"
      HeaderImageItem9Picture=   "frmMain.frx":1ABE4
      HeaderImageItem10Key=   "SymbolConfiguration2"
      HeaderImageItem10Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Symbol Configuration 2.bmp"
      HeaderImageItem10Id=   "{B753320C-2009-4F7B-A3FA-8B12D6AD7C9B}"
      HeaderImageItem10Picture=   "frmMain.frx":1B2F6
      HeaderImageItem11Key=   "Menu2"
      HeaderImageItem11Tag=   "K:\Users\Mark\Development\Pafwert 2.0\Graphics\Menu 2.bmp"
      HeaderImageItem11Id=   "{7498661A-1425-4D8D-8F13-F2270BE6E809}"
      HeaderImageItem11Picture=   "frmMain.frx":1BA08
      HotStyleMousePointer=   99
      LicenceData     =   "Unlicensed"
      Begin FramePlusCtl.FramePlus fraOpt 
         Height          =   3135
         Index           =   3
         Left            =   -2460
         TabIndex        =   52
         Tag             =   "4"
         Top             =   4800
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   5530
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
         Caption         =   "FramePlus8"
         Begin FramePlusCtl.FramePlus fraEntropy 
            Height          =   2295
            Left            =   0
            TabIndex        =   53
            Top             =   300
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   4048
            ShowCheckbox    =   -1  'True
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
            Caption         =   "Use these seed values"
            Begin VB.TextBox txtKeywords 
               Enabled         =   0   'False
               Height          =   735
               Left            =   420
               MultiLine       =   -1  'True
               TabIndex        =   54
               Top             =   1440
               Width           =   2115
            End
            Begin Pafwert.mbSlider mbSlider1 
               Height          =   255
               Left            =   1260
               TabIndex        =   55
               Top             =   210
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               Value           =   0
            End
            Begin Pafwert.mbSlider mbSlider2 
               Height          =   255
               Left            =   1260
               TabIndex        =   56
               Top             =   510
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               Value           =   0
            End
            Begin Pafwert.mbSlider mbSlider3 
               Height          =   255
               Left            =   1260
               TabIndex        =   57
               Top             =   810
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   450
               Value           =   0
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Entropy 3:"
               Enabled         =   0   'False
               Height          =   255
               Index           =   2
               Left            =   420
               TabIndex        =   61
               Top             =   840
               Width           =   1395
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Entropy 2:"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   420
               TabIndex        =   60
               Top             =   540
               Width           =   1395
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Entropy 1:"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   420
               TabIndex        =   59
               Top             =   240
               Width           =   1395
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Seed keywords:"
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   420
               TabIndex        =   58
               Top             =   1200
               Width           =   1815
            End
         End
      End
      Begin FramePlusCtl.FramePlus fraOpt 
         Height          =   3315
         Index           =   2
         Left            =   1440
         TabIndex        =   42
         Tag             =   "2"
         Top             =   4140
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   5847
         BackStyle       =   0
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
         Caption         =   "FramePlus7"
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FF00FF&
            Caption         =   "Enforce length"
            Height          =   255
            Left            =   60
            TabIndex        =   66
            Top             =   60
            Width           =   2715
         End
         Begin VB.TextBox txtMinLen 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   720
            TabIndex        =   63
            Text            =   "8"
            Top             =   300
            Width           =   495
         End
         Begin VB.TextBox txtMaxLen 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1860
            TabIndex        =   62
            Text            =   "128"
            Top             =   300
            Width           =   495
         End
         Begin FramePlusCtl.FramePlus fraCharSets 
            Height          =   2235
            Left            =   60
            TabIndex        =   43
            Top             =   960
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   3942
            ShowCheckbox    =   -1  'True
            BackStyle       =   0
            ThemeColor      =   4484202
            BackColor       =   16711935
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
            Caption         =   "Character sets"
            Begin VB.CheckBox chkSymbols 
               Caption         =   "Symbols"
               Enabled         =   0   'False
               Height          =   255
               Left            =   600
               TabIndex        =   51
               Top             =   1920
               Width           =   1875
            End
            Begin VB.CheckBox chkNumbers 
               Caption         =   "Numbers"
               Enabled         =   0   'False
               Height          =   255
               Left            =   600
               TabIndex        =   50
               Top             =   1680
               Width           =   1875
            End
            Begin VB.CheckBox chkUpper 
               Caption         =   "Uppercase"
               Enabled         =   0   'False
               Height          =   255
               Left            =   600
               TabIndex        =   49
               Top             =   1440
               Width           =   1875
            End
            Begin VB.CheckBox chkLower 
               Caption         =   "Lowercase"
               Enabled         =   0   'False
               Height          =   255
               Left            =   600
               MaskColor       =   &H00FF00FF&
               TabIndex        =   48
               Top             =   1200
               Width           =   1875
            End
            Begin VB.OptionButton optComplexity 
               Caption         =   "These character sets:"
               Enabled         =   0   'False
               Height          =   255
               Index           =   3
               Left            =   300
               TabIndex        =   47
               Top             =   960
               Width           =   2175
            End
            Begin VB.OptionButton optComplexity 
               Caption         =   "All 4 character sets"
               Enabled         =   0   'False
               Height          =   255
               Index           =   2
               Left            =   300
               TabIndex        =   46
               Top             =   720
               Width           =   2175
            End
            Begin VB.OptionButton optComplexity 
               Caption         =   "At least 3 character sets"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   300
               TabIndex        =   45
               Top             =   480
               Width           =   2175
            End
            Begin VB.OptionButton optComplexity 
               Caption         =   "At least 2 character sets"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   300
               TabIndex        =   44
               Top             =   240
               Value           =   -1  'True
               Width           =   2175
            End
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H00FF00FF&
            Caption         =   "Min:"
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   65
            Top             =   360
            Width           =   435
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H00FF00FF&
            Caption         =   "Max:"
            Height          =   195
            Index           =   2
            Left            =   1440
            TabIndex        =   64
            Top             =   360
            Width           =   555
         End
      End
      Begin FramePlusCtl.FramePlus fraOpt 
         Height          =   495
         Index           =   4
         Left            =   1140
         TabIndex        =   40
         Tag             =   "4"
         Top             =   3420
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         BackStyle       =   0
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
         Caption         =   "FramePlus6"
         Begin VB.CheckBox chkShowStats 
            BackColor       =   &H00FF00FF&
            Caption         =   "Show statistics"
            Height          =   315
            Left            =   120
            TabIndex        =   41
            Top             =   60
            Width           =   2055
         End
      End
      Begin FramePlusCtl.FramePlus fraOpt 
         Height          =   1275
         Index           =   0
         Left            =   1560
         TabIndex        =   34
         Tag             =   "0"
         Top             =   6540
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   2249
         BackStyle       =   0
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
         Caption         =   "FramePlus5"
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   900
            TabIndex        =   35
            Text            =   "15"
            Top             =   120
            Width           =   435
         End
         Begin ButtonPlusCtl.ButtonPlus cmdSuggest 
            Height          =   375
            Left            =   1500
            TabIndex        =   36
            Top             =   780
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            BackStyle       =   0
            HotEffects      =   -1  'True
            WordWrap        =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Suggest"
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Suggest "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   39
            Top             =   180
            Width           =   1020
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "based on the criteria below."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   38
            Top             =   440
            Width           =   2520
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "passwords"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   37
            Top             =   180
            Width           =   1200
         End
      End
   End
   Begin VB.CommandButton cmdCopy 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copy to Clipboard"
      Enabled         =   0   'False
      Height          =   135
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4740
      Width           =   15
   End
   Begin VB.PictureBox picBanner 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   0
      Picture         =   "frmMain.frx":1C11A
      ScaleHeight     =   1215
      ScaleWidth      =   9885
      TabIndex        =   0
      Top             =   0
      Width           =   9885
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
Option Explicit

Private Const MODULE_NAME As String = "Pafwert"

Private mPatterns() As String

Private Const BAR_WIDTH = 3475

Private Sub chkShowStats_Click()

   If chkShowStats.Value = vbChecked Then
      expPasses.Groups("passwordstatistics").Visible = True
      fraPassStats.Visible = False
   Else
      expPasses.Groups("passwordstatistics").Visible = False
      fraPassStats.Visible = False
   End If

End Sub

Private Sub cmdSuggest_Click()
      Dim i As Long
      Dim sPass As String
      Dim sKeywords As String
      Dim sPattern As String
      Dim sPasswords As String
      Dim iQty As Long
      Dim oPass As PafwertLib
      Dim lErrNumber   As Long
      Dim sErrSource   As String
      Dim sErrDescription As String
      Dim gItem As GroupItem
      Dim sPatternCheck As String
      Dim bFailed As Boolean
      Dim iRetryCount As Integer
      Dim vLenData() As Variant
      Dim lPassLen As Long
      Dim lTotalLen As Long
      Dim lMinLen As Long
      Dim lMaxLen As Long
      Dim ChartXunit As Single
      Dim ChartYUnit As Single
      Dim j As Integer
      Dim iChar As Integer
      Dim lMaxHeight As Long
      Dim lChCount(256) As Long
      Dim lTotalLower As Long
      Dim lTotalUpper As Long
      Dim lTotalNumbers As Long
      Dim lTotalSymbols As Long
      Dim iCharsetCount(4) As Long
      Dim lTotalTime As Single
      Const PROCNAME As String = "cmdSuggest_Click"
10    'On Error GoTo ErrHandler
      '---Set UI stuff
20    Set oPass = New PafwertLib
30    cmdSuggest.Enabled = False
40    Screen.MousePointer = vbArrowHourglass
      expPasses.Groups(1).GroupItems.Clear
      expPasses.Groups(1).Caption = "Generating random passwords..."
      expPasses.Redraw = False

      '-------Set Pattern
50    If optPattern(1).Value = True Then
60       sPattern = mPatterns(cboPattern.ItemData(cboPattern.ListIndex))
70    End If

      If optPattern(2).Value Then
         sPattern = txtPattern
         sPatternCheck = oPass.CheckPattern(sPattern)

         If Len(sPatternCheck) Then
            MsgBox sPatternCheck
            Exit Sub

         End If
      End If

      '-------Set entropy, keywords and quantity
      If fraEntropy.Checked Then
         oPass.Entropy1 = mbSlider1.Value
         oPass.Entropy2 = mbSlider2.Value
         oPass.Entropy3 = mbSlider3.Value
80       sKeywords = Trim$(txtKeywords)
      End If

      '---Set password policy
      With oPass.Complexity.Policy

         If fraPassLengths.Checked Then
            oPass.Complexity.Policy.MinimumLength = Val(txtMinLen)
            oPass.Complexity.Policy.MaximumLength = Val(txtMaxLen)
         End If

         If fraCharSets.Checked Then
            If optComplexity(0).Value = True Then .MinimumCharsets = 2
            If optComplexity(1).Value = True Then .MinimumCharsets = 3
            If optComplexity(2).Value = True Then .MinimumCharsets = 4
            If optComplexity(3).Value = True Then
               If chkLower.Value = vbChecked Then .RequireLowercase = True
               If chkUpper.Value = vbChecked Then .RequireUppercase = True
               If chkNumbers.Value = vbChecked Then .RequireNumbers = True
               If chkSymbols.Value = vbChecked Then .RequireSymbols = True
            End If
         End If

      End With

90    iQty = Val(txtQty)
      lMinLen = 9999
      ReDim vLenData(iQty, 1)

      '-------Generate passwords
100   For i = 0 To iQty - 1
110      'DoEvents
         iRetryCount = 0

120      Do

            With oPass
               bFailed = False
130            .GeneratePassword sPattern, sKeywords
               iRetryCount = iRetryCount + 1

               If iRetryCount > 5 Then Debug.Print iRetryCount, .Password
            End With

140      Loop Until (Not bFailed) Or iRetryCount > 50

         If iRetryCount > 25 Then
            MsgBox "Timeout occurred generating passwords with the criteria specified. Try adjusting the complexity options.", vbInformation
            GoTo ExitHere:
         End If

         '---Tally the stats
         If chkShowStats.Value = vbChecked Then

            With oPass.Complexity

               If .PasswordLength < lMinLen Then lMinLen = .PasswordLength
               If .PasswordLength > lMaxLen Then lMaxLen = .PasswordLength
               lTotalLen = lTotalLen + lPassLen

               For j = 1 To Len(.Password)
                  lChCount(iChar) = lChCount(iChar) + .CharacterDistribution(j)

                  If lChCount(iChar) > lMaxHeight Then
                     lMaxHeight = lChCount(iChar)
                  End If

               Next j

               lTotalLower = lTotalLower + Abs(.LowercaseCount > 0)
               lTotalUpper = lTotalUpper + Abs(.UppercaseCount > 0)
               lTotalNumbers = lTotalNumbers + Abs(.NumbersCount > 0)
               lTotalSymbols = lTotalSymbols + Abs(.SymbolsCount > 0)
               iCharsetCount(.CharsetCount) = iCharsetCount(.CharsetCount) + .CharsetCount
               lTotalTime = lTotalTime + oPass.TimeTaken
            End With

         End If

         vLenData(i, 1) = Len(oPass.Password)
150      Set gItem = expPasses.Groups(1).GroupItems.Add("p" & i)
         gItem.Caption = oPass.Password
         gItem.ForeColor = vbBlack
         gItem.Indent = 4
         gItem.Image = "bullet"
         

         
160   Next i

      cmdCopy.Enabled = True

      '------Configure Charts
      If chkShowStats.Value = vbChecked Then
         '            If fraPassStats.Visible = False Then
         '               expPasses.Groups("passwordstatistics").Visible = True
         '
         '               fraPassStats.Visible = True
         '            End If
         fraPassStats.Visible = True
         lblShortest.Caption = "Shortest password: " & Str$(lMinLen) & " characters."
         lblLongest.Caption = "Longest password: " & Str$(lMaxLen) & " characters."
         lblAverage.Caption = "Average length: " & Format$(lTotalLen / iQty, "0.0") & " characters."
         lblLowercase.Caption = Format$(lTotalLower / iQty, "0%") & " of the passwords contain lowercase letters."
         lblUppercase.Caption = Format$(lTotalUpper / iQty, "0%") & " of the passwords contain uppercase letters."
         lblNumbers.Caption = Format$(lTotalNumbers / iQty, "0%") & " of the passwords contain numbers."
         lblSymbols.Caption = Format$(lTotalSymbols / iQty, "0%") & " of the passwords contain symbols."
         lbl1Charset.Caption = Format$(iCharsetCount(1) / iQty, "0%") & " of the passwords use 1 character set."
         lbl2Charsets.Caption = Format$(iCharsetCount(2) / iQty, "0%") & " of the passwords use 2 character sets."
         lbl3Charsets.Caption = Format$(iCharsetCount(3) / iQty, "0%") & " of the passwords use 3 character sets."
         lbl4Charsets.Caption = Format$(iCharsetCount(4) / iQty, "0%") & " of the passwords use 4 character sets."
         lblTime.Caption = "Average time: " & Format$(lTotalTime / iQty, "0.0") & "ms"
         On Error Resume Next
         ChartXunit = (picGraph.Width / 257)
         ChartYUnit = picGraph.Height / lMaxHeight

         For i = 1 To 254
            Load Line1(i)

            With Line1(i)
               Set .Container = picGraph
               .x1 = i * ChartXunit
               .x2 = i * ChartXunit
               .Y1 = (picGraph.Height - (lChCount(i)) * ChartYUnit) + 50
               .Y2 = picGraph.Height - 3

               Select Case i

                  Case 48 To 57
                     .BorderColor = &H54CF66

                  Case 64 To 90
                     .BorderColor = &H1CDFDF

                  Case 95 To 122
                     .BorderColor = &HE17B59

                  Case Else
                     .BorderColor = &HC74982
               End Select

               .ZOrder 0
               .Visible = True
            End With

         Next i

      End If

ExitHere:
      On Error Resume Next
      expPasses.Groups(1).Caption = "Suggested Passwords"
      expPasses.Redraw = True
200   Screen.MousePointer = vbNormal
210   cmdSuggest.Enabled = True
      cmdSuggest.SetFocus
220   On Error Resume Next

230   If lErrNumber <> 0 Then
240      ShowUnexpectedError lErrNumber, sErrDescription, sErrSource
250   End If

260   Exit Sub

ErrHandler:
270   lErrNumber = Err.Number
280   sErrDescription = Err.Description
290   sErrSource = FormatErrorSource(Err.Source, MODULE_NAME, PROCNAME & " (" & Erl & ")")
300   Resume ExitHere
End Sub
   
Private Sub Form_Load()
   Dim i As Long
   Dim Index               As Long
   ExplorerBar1.Width = BAR_WIDTH
   ' ExplorerBar2.Width = BAR_WIDTH
   'match the colors
   SynchronizeColors Controls, ExplorerBar1.GroupBackColor

   For i = 0 To 3
      optComplexity(i).BackColor = ExplorerBar1.GroupBackColor
   Next i

   chkLower.BackColor = ExplorerBar1.GroupBackColor
   chkUpper.BackColor = ExplorerBar1.GroupBackColor
   chkNumbers.BackColor = ExplorerBar1.GroupBackColor
   chkSymbols.BackColor = ExplorerBar1.GroupBackColor

   'attach the clients
   For Index = 0 To fraOpt.UBound

      With ExplorerBar1.Groups(Index + 1).GroupItems.Add()
         .ClientResizeStyle = ebcrsHorizontal
         Set .Client = fraOpt(Index)
      End With

   Next

   With expPasses.Groups("passwordstatistics").GroupItems.Add()
      .ClientResizeStyle = ebcrsHorizontal
      Set .Client = fraPassStats
   End With

End Sub

Private Sub Form_Resize()
   On Error Resume Next
   expPasses.Width = Me.Width - expPasses.Left - (Screen.TwipsPerPixelX * 8)
End Sub

Private Sub fraCharSets_AfterCheck(ByVal Checked As Boolean)

   If Checked = True Then
      If optComplexity(3).Value = True Then
         chkLower.Enabled = True
         chkUpper.Enabled = True
         chkNumbers.Enabled = True
         chkSymbols.Enabled = True
      Else
         chkLower.Enabled = False
         chkUpper.Enabled = False
         chkNumbers.Enabled = False
         chkSymbols.Enabled = False
      End If
   End If

End Sub

Private Sub optComplexity_Click(Index As Integer)

   Select Case Index

      Case 0
         chkLower.Enabled = False
         chkUpper.Enabled = False
         chkNumbers.Enabled = False
         chkSymbols.Enabled = False

      Case 1
         chkLower.Enabled = False
         chkUpper.Enabled = False
         chkNumbers.Enabled = False
         chkSymbols.Enabled = False

      Case 2
         chkLower.Enabled = False
         chkUpper.Enabled = False
         chkNumbers.Enabled = False
         chkSymbols.Enabled = False

      Case 3
         chkLower.Enabled = True
         chkUpper.Enabled = True
         chkNumbers.Enabled = True
         chkSymbols.Enabled = True
   End Select

End Sub

Private Sub optPattern_Click(Index As Integer)

   Select Case Index

      Case 0
         cboPattern.Enabled = False
         txtPattern.Enabled = False

      Case 1
         cboPattern.Enabled = True
         txtPattern.Enabled = False
         LoadPatterns

      Case 2
         cboPattern.Enabled = False
         txtPattern.Enabled = True
   End Select

End Sub

   '*******************************************************************************
   ' FormatErrorSource (FUNCTION)
   '
   ' PARAMETERS:
   ' (In) - sErrSource - String - The current Err.Source
   ' (In) - sModule - String - The module where it occurred
' (In) - sFunction  - String - The function where it occurred
   '
   ' RETURN VALUE:
   ' String - A formatted string to return as the error source
   '
   ' DESCRIPTION:
' Takes the error source and appends module and function information so it can
   ' be used to trace the stack.
   '*******************************************************************************
Public Function FormatErrorSource(ByVal sErrSource As String, _
                                  ByVal sModule As String, _
                                  ByVal sFunction As String) As String
      Static s_sDefaultErrorSource As String
10    On Error Resume Next

20    If LenB(s_sDefaultErrorSource) = 0 Then
30       Err.Raise vbObjectError
40       s_sDefaultErrorSource = Err.Source
50       Err.Clear
60    End If

70    If sErrSource = s_sDefaultErrorSource Then

80       FormatErrorSource = App.ProductName & "." & sModule & "." & sFunction
90    Else

100      FormatErrorSource = sErrSource & vbCrLf & App.ProductName & "." & sModule & "." & sFunction
110   End If

End Function
   
   '*******************************************************************************
   ' ShowUnexpectedError (SUB)
   '
   ' PARAMETERS:
   ' (In) - lNumber   - Long   - Error number
   ' (In) - sDescription - String - Error description
   ' (In) - sLocation - String - Error source
   '
   ' DESCRIPTION:
   ' If an unexpected error is found, this routine gets called to display the error
   ' to the user.
   '*******************************************************************************
Public Sub ShowUnexpectedError(ByVal lNumber As Long, _
                               ByVal sDescription As String, _
                               ByVal sLocation As String)
      Dim sMessage As String
10    On Error Resume Next
20    sMessage = "An error occurred in the application: " & lNumber & vbCrLf & sDescription & vbCrLf & sLocation
30    Debug.Print sMessage
40    MsgBox sMessage, vbCritical, App.ProductName
End Sub
   
Public Sub SynchronizeColors(Controls As Object, _
                             BackColor As OLE_COLOR)
   Dim Control             As Object
   'match the colors
   On Error Resume Next

   For Each Control In Controls

      If TypeOf Control.Container Is Frame Then

         Select Case LCase$(TypeName(Control))

            Case "textbox", "combobox"

               'skip
            Case Else
               'apply backcolor
               Control.BackColor = BackColor
         End Select

      ElseIf TypeOf Control Is Frame Then
         'apply backcolor
         Control.BackColor = BackColor
      End If

   Next

End Sub '(Public) Sub SynchronizeColors ()

Public Sub LoadPatterns()
   Dim lResult             As Long
   Dim sPattern            As String
   Dim lPos                As Long
   Dim sFileData           As String
   Dim sPatternNames()     As String
   Dim lFile               As Long
   Dim i                   As Integer
   Dim sName               As String

   'On Error GoTo ErrHandler
   If cboPattern.ListCount = 0 Then
      If Len(Dir(WordlistDir & "\patterns.cfg")) = 0 Then Err.Raise ERR_OPEN_PATTERNS_FILE, ERRSOURCE, ERR_OPEN_PATTERNS_FILE_DESC
      lFile = FreeFile
      On Error Resume Next
      Open Me.WordlistDir & "\patterns.cfg" For Binary Access Read As #lFile

      If Err Then Err.Raise ERR_OPEN_PATTERNS_FILE, ERRSOURCE, ERR_OPEN_PATTERNS_FILE_DESC
      sFileData = Space$(LOF(lFile))
      Get #lFile, , sFileData
      Close lFile

      If Len(sFileData) = 0 Or Err Then Err.Raise ERR_LOADING_PATTERNS, ERRSOURCE, ERR_LOADING_PATTERNS_DESC
      'On Error GoTo ErrHandler
      sFileData = Replace(sFileData, vbCrLf & vbCrLf, vbCrLf)
      sPatternNames = Split(sFileData, vbCrLf)

      If IsBounded(sPatternNames) = False Then
         Err.Raise ERR_LOADING_PATTERNS, ERRSOURCE, ERR_LOADING_PATTERNS_DESC & Me.WordlistDir & "\patterns.cfg"
      End If

      ReDim mPatterns(UBound(sPatternNames))

      For i = 0 To UBound(sPatternNames)

         If Left$(sPatternNames(i), 1) <> "#" And Len(sPatternNames(i)) Then
            lPos = InStr(sPatternNames(i), ":")

            If lPos Then
               sName = (Left$(sPatternNames(i), lPos - 1))
               sPattern = (Right$(sPatternNames(i), Len(sPatternNames(i)) - lPos))
            Else
               sName = "(No name)"
               sPattern = sPatternNames(i)
            End If

            cboPattern.AddItem sName
            cboPattern.ItemData(cboPattern.NewIndex) = i
            mPatterns(i) = sPattern
            Debug.Print i, cboPattern.NewIndex, sName, sPattern
         End If

      Next i

   End If

End Sub

Public Property Get WordlistDir() As String
   Dim sPaths(6)      As String
   Dim i              As Long
   Dim Dir    As String
   'If Len(mvarWordlistDir) = 0 Then
   sPaths(0) = GetFullPath(GetStringSetting("Pafwert", "Settings", "WordlistDir") & "\")
   sPaths(1) = GetFullPath(App.Path & "\wordlists\")
   sPaths(2) = GetFullPath(GetStringSetting("Pafwert", "Settings", "LastWordlistDir") & "\")
   sPaths(3) = GetFullPath("C:\Program Files\Pafwert\Wordlists" & "\")
   sPaths(4) = GetFullPath(CurDir & "\")
   sPaths(5) = GetFullPath(App.Path & "\..\wordlists\")
   sPaths(6) = GetFullPath(Command$ & "\")

   For i = 0 To 6
      Debug.Print i, sPaths(i)
      If CheckWordlistDir(sPaths(i)) Then
         Dir = sPaths(i)
         Exit For
      End If

   Next i

' TODO: Prompt for wordlist dir
' TODO: Have a command-line param for wordlist dir



   If Len(Dir) = 0 Then Err.Raise ERR_WORDLIST_DIR_NOT_FOUND, ERRSOURCE, ERR_WORDLIST_DIR_NOT_FOUND_DESC
   'SaveStringSetting "Pafwert", "Settings", "WordlistDir", mvarWordlistDir
