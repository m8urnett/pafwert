VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Generator"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstPasswords 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   60
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   60
      Width           =   3675
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Top             =   3180
      Width           =   1095
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

Private Declare Function SendMessage Lib "user32" Alias _
   "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Private Const LB_GETSELITEMS = &H191
Private Const LB_GETITEMRECT = &H198
Private Const LB_ERR = (-1)

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Sub cmdGenerate_Click()
   Dim i As Integer
   Dim oPaf As New PafwertLib
   
   lstPasswords.Clear
   For i = 1 To 10
      lstPasswords.AddItem oPaf.GeneratePassword
   Next i
End Sub

Private Sub Form_Load()
   cmdGenerate_Click
End Sub

Private Sub lstPasswords_DblClick()
Dim i As Integer
   For i = 0 To lstPasswords.ListCount - 1
      lstPasswords.Selected(i) = True
   Next i
End Sub

Private Sub lstPasswords_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iItem As Integer
Dim lXcoord As Long, lYcoord As Long
Dim lRet As Long
Dim CurRect As RECT
Dim i As Integer

   If Button = vbRightButton Then
      lXcoord = X \ Screen.TwipsPerPixelX
      lYcoord = Y \ Screen.TwipsPerPixelY
 
      For i = 0 To lstPasswords.ListCount - 1
         lRet = SendMessage(lstPasswords.hwnd, LB_GETITEMRECT, i, CurRect)
         If (lXcoord >= CurRect.Left) And (lXcoord <= CurRect.Right) And (lYcoord >= CurRect.Top) And (lYcoord <= CurRect.Bottom) Then
            iItem = i
            Exit For
         End If
      Next i
      
      If iItem <> -1 Then
         If lstPasswords.SelCount <= 1 Then
            For i = 0 To lstPasswords.ListCount - 1
               lstPasswords.Selected(i) = False
            Next i
         End If
         lstPasswords.ListIndex = iItem
         lstPasswords.Selected(iItem) = True
         PopupMenu mnuPassword
      End If
   End If
   
End Sub

Private Sub mnuCopy_Click()
   Clipboard.SetText lstPasswords.Text
End Sub

Private Sub mnuCopyAll_Click()
Dim i As Integer
Dim sPasswords As String
   
   For i = 0 To lstPasswords.ListCount - 1
      sPasswords = sPasswords & lstPasswords.List(i) & vbCrLf
   Next i
   Clipboard.SetText sPasswords
   
End Sub

Private Sub mnuCopySelected_Click()
    Dim SelItems() As Long
    Dim SelectedCount As Long
    Dim i As Integer
    Dim sPasswords As String
    
    SelectedCount = lstPasswords.SelCount
    If SelectedCount > 0 Then
        ReDim SelItems(SelectedCount - 1)
        SendMessage lstPasswords.hwnd, LB_GETSELITEMS, ByVal SelectedCount, SelItems(0)
    End If

    If SelectedCount > 0 Then
        For i = 0 To SelectedCount - 1
            sPasswords = sPasswords & lstPasswords.List(SelItems(i)) & vbCrLf
        Next i
        Clipboard.SetText sPasswords
    End If
End Sub

Private Sub mnuPassword_Click()
   If lstPasswords.SelCount > 1 Then
      mnuCopy.Enabled = False
   Else
      mnuCopy.Enabled = True
   End If
End Sub
