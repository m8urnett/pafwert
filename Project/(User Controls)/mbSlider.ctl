VERSION 5.00
Begin VB.UserControl mbSlider 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6495
   KeyPreview      =   -1  'True
   MaskColor       =   &H00FFFFFF&
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   433
   ToolboxBitmap   =   "mbSlider.ctx":0000
   Begin VB.PictureBox picThumbDisabled 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   195
      Left            =   4020
      Picture         =   "mbSlider.ctx":0312
      ScaleHeight     =   195
      ScaleWidth      =   90
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox picThumbEnabled 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      HasDC           =   0   'False
      Height          =   255
      Left            =   4020
      Picture         =   "mbSlider.ctx":0458
      ScaleHeight     =   255
      ScaleWidth      =   150
      TabIndex        =   3
      Top             =   660
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picLineDisabled 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      HasDC           =   0   'False
      Height          =   90
      Left            =   1200
      Picture         =   "mbSlider.ctx":059E
      ScaleHeight     =   90
      ScaleWidth      =   150
      TabIndex        =   2
      Top             =   1380
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picLineEnabled 
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      HasDC           =   0   'False
      Height          =   90
      Left            =   1140
      Picture         =   "mbSlider.ctx":0608
      ScaleHeight     =   90
      ScaleWidth      =   150
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox picThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   465
      Picture         =   "mbSlider.ctx":0672
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   0
      Top             =   495
      Width           =   90
   End
   Begin VB.Image imgMid 
      Height          =   30
      Left            =   1080
      Picture         =   "mbSlider.ctx":07B8
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1770
   End
End
Attribute VB_Name = "mbSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mMin As Long
Private mMax As Long
Private mValue As Long
Private mDragging As Boolean
Private Const clrWhite As Long = 16777215
Private Const clrBlack As Long = 0
Private Const clrLtGrey As Long = 13160660
Private Const clrDkGrey As Long = 8421504
Private mLastX As Long
Private mLastY As Long
Private mEnabled As Boolean

Public Event Change(ByVal NewValue As Long)

Private Sub imgMid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Show what the position would be if the user were to click
    Dim Max As Long
    Max = ScaleWidth - picThumb.Width
    imgMid.ToolTipText = Int((ScaleX(X, vbTwips, vbPixels) * (mMax - mMin)) / Max) + mMin
End Sub


Private Sub imgMid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If mEnabled = False Then Exit Sub
    ' Set value via click position on bar
    Dim Max As Long
    Max = ScaleWidth - picThumb.Width
    Value = Int((ScaleX(X, vbTwips, vbPixels) * (mMax - mMin)) / Max) + mMin
End Sub

Private Sub picThumb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If mEnabled = False Then Exit Sub
    ' Start dragging the thumb
    mDragging = True
    mLastX = X
End Sub


Private Sub picThumb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If mEnabled = False Then Exit Sub
    ' Drag the thumb horizontally
    Dim NewX As Long
    Dim Max As Long
    Max = ScaleWidth - picThumb.Width
    If mDragging = True Then
        If X <> mLastX Then
            mLastX = X
            NewX = picThumb.Left + X
            If NewX < 0 Then
                NewX = 0
            ElseIf NewX > Max Then
                NewX = Max
            End If
            picThumb.Left = NewX
            SetValueFromThumb
        End If
    End If
End Sub


Private Sub picThumb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If mEnabled = False Then Exit Sub
    ' Stop dragging the thumb
    mDragging = False
End Sub


Private Sub SetValueFromThumb()
    ' Set value from thumb position
    Dim Max As Long
'    If mOrientation = sldVertical Then
'        Max = ScaleHeight - picThumbVertical.Height
'        Value = (picThumbVertical.Top * (mMax - mMin)) / Max + mMin
'    Else
        Max = ScaleWidth - picThumb.Width
        Value = (picThumb.Left * (mMax - mMin)) / Max + mMin
   ' End If
End Sub


Private Sub SetThumbFromValue()
    ' Set thumb position from value
    Dim Max As Long
    Dim X As Long
    Dim Y As Long
'
'    ' Set vertical thumb
'    Max = ScaleHeight - picThumbVertical.Height
'    picThumbVertical.Top = (mValue - mMin) / (mMax - mMin) * Max
'    picThumbVertical.ToolTipText = mValue
'
'   ' Set horizontal thumb
'    Max = ScaleWidth - picThumb.Width
'    picThumb.Left = (mValue - mMin) / (mMax - mMin) * Max
'    picThumb.ToolTipText = mValue
End Sub


Private Sub UserControl_InitProperties()
    ' Set default properties (executed once when control is added to form)
    MinValue = 0
    MaxValue = 100
    Value = 50
End Sub


Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   If mEnabled = False Then Exit Sub
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Then
        Value = Value - 1
    ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyDown Then
        Value = Value + 1
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' Load the properties
    MinValue = PropBag.ReadProperty("MinValue", 0)
    MaxValue = PropBag.ReadProperty("MaxValue", 100)
    Value = PropBag.ReadProperty("Value", 50)
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    ' Save the properties
    PropBag.WriteProperty "MinValue", mMin, 0
    PropBag.WriteProperty "MaxValue", mMax, 100
    PropBag.WriteProperty "Value", mValue, 50
End Sub


Private Sub UserControl_Resize()
    ' Resize the slider bar
    Dim X As Long
    Dim Y As Long
    Dim w As Long
    Dim h As Long
    Dim pos As Long
    
    ' Draw the horizontal bar
    X = 0
    Y = Int(ScaleHeight / 2)
    w = ScaleWidth
    h = imgMid.Height
    ' Draw the line
    imgMid.Move X, Y, w, h
    'imgLeft.Move X, Y
    'imgRight.Move w - 2, Y
    ' Draw the thumb
    pos = mValue * Int((w - picThumb.Width)) / mMax
    picThumb.Move pos - Int(picThumb.Width / 2) + 2, Y - picThumb.Height / 2 + 2
    
'    ' Draw the vertical bar
'    X = Int(ScaleWidth / 2)
'    Y = 0
'    w = imgMidVertical.Width
'    h = ScaleHeight
'    ' Draw the line
'    imgMidVertical.Move X, Y, w, h
'    imgTop.Move X, Y
'    imgBottom.Move w, Y - 2
'    ' Draw the thumb
'    pos = mValue * Int((h - picThumbVertical.Height)) / mMax
'    picThumbVertical.Move X - picThumbVertical.Width / 2 + 2, pos - Int(picThumbVertical.Height / 2) + 2
    
    'SetThumbFromValue
End Sub


Public Property Let MinValue(ByVal Value As Long)
    ' Set the minimum value (no negative numbers)
    If Value < 0 Then
        Value = 0
    End If
    mMin = Value
    PropertyChanged "MinValue"
    If mValue < mMin Then
        mValue = mMin
        PropertyChanged "Value"
    End If
    If mMax < mMin + 1 Then
        mMax = mMin + 1
        PropertyChanged "MaxValue"
    End If
End Property


Public Property Get MinValue() As Long
    ' Get the minimum value
    MinValue = mMin
End Property


Public Property Let MaxValue(ByVal Value As Long)
    ' Set the maximum value (must be greater than min)
    If Value < mMin + 1 Then
        Value = mMin + 1
    End If
    mMax = Value
    PropertyChanged "MaxValue"
    If mValue > mMax Then
        mValue = mMax
        PropertyChanged "Value"
    End If
End Property


Public Property Get MaxValue() As Long
    ' Get the maximum value
    MaxValue = mMax
End Property


Public Property Let Value(ByVal NewValue As Long)
    ' Set the value (must be >= min and <= max)
    If NewValue = mValue Then
        Exit Property
    End If
    If NewValue < mMin Then
        NewValue = mMin
    ElseIf NewValue > mMax Then
        NewValue = mMax
    End If
    mValue = NewValue
    SetThumbFromValue
    PropertyChanged "Value"
    RaiseEvent Change(NewValue)
End Property


Public Property Get Value() As Long
    ' Get the value
    Value = mValue
End Property

Public Property Get Enabled() As Boolean
   Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
   mEnabled = Value
   If mEnabled = True Then
      imgMid.Picture = picLineEnabled.Picture
      picThumb.Picture = picThumbEnabled.Picture
   Else
      imgMid.Picture = picLineDisabled.Picture
      picThumb.Picture = picThumbDisabled.Picture
   End If
   UserControl.Refresh
   PropertyChanged "Value"
   RaiseEvent Change(Value)
End Property


'Public Property Let Orientation(ByVal Value As sldOrientation)
'    ' Set the orientation, horizontal or vertical
'    mOrientation = Value
'
'    picThumbVertical.Visible = False
'    imgTop.Visible = False
'    imgBottom.Visible = False
'    imgMidVertical.Visible = False
'
'    picThumb.Visible = False
'    imgLeft.Visible = False
'    imgRight.Visible = False
'    imgMid.Visible = False
'
'    If mOrientation = sldVertical Then
'        picThumbVertical.Visible = True
'        imgTop.Visible = True
'        imgBottom.Visible = True
'        imgMidVertical.Visible = True
'    Else
'        picThumb.Visible = True
'        imgLeft.Visible = True
'        imgRight.Visible = True
'        imgMid.Visible = True
'    End If
'
'    PropertyChanged "Orientation"
'End Property
'
'
'Public Property Get Orientation() As sldOrientation
'    ' Get the orientation, horizontal or vertical
'    Orientation = mOrientation
'End Property
