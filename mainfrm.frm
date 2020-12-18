VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Pixel Traverser"
   ClientHeight    =   3750
   ClientLeft      =   345
   ClientTop       =   6885
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6015
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   240
      Left            =   3120
      TabIndex        =   23
      Top             =   3240
      Width           =   990
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   240
      Left            =   2160
      TabIndex        =   22
      Top             =   3240
      Width           =   975
   End
   Begin VB.CheckBox chkTopMost 
      Caption         =   "Top Most"
      Height          =   195
      Left            =   4560
      TabIndex        =   21
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame FraParameters 
      Caption         =   " Parameters"
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   5775
      Begin VB.CheckBox chkRandom 
         Caption         =   "Random motion"
         Height          =   255
         Left            =   3360
         TabIndex        =   20
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton OptClickEach 
         Caption         =   "Click each point"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton OptHoldClick 
         Caption         =   "Hold left click"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.TextBox txtDelayMs 
         Height          =   285
         Left            =   4080
         TabIndex        =   17
         Text            =   "1"
         ToolTipText     =   "Interval between mouse moves in milliseconds"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtDy 
         Height          =   285
         Left            =   2160
         TabIndex        =   15
         Text            =   "5"
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtDx 
         Height          =   285
         Left            =   600
         TabIndex        =   13
         Text            =   "5"
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblDelay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delay"
         Height          =   195
         Left            =   3360
         TabIndex        =   16
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblDy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dy"
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   360
         Width           =   180
      End
      Begin VB.Label lblDx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dx"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   180
      End
   End
   Begin VB.Timer hotkeyListenerTimer 
      Interval        =   50
      Left            =   3000
      Top             =   3480
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   240
      Left            =   1080
      TabIndex        =   10
      ToolTipText     =   "Ctrl+F3"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Timer cursorMoveTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3600
      Top             =   3480
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Move"
      Height          =   240
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Ctrl+F2"
      Top             =   3240
      Width           =   975
   End
   Begin VB.Timer cursorpostimer 
      Interval        =   20
      Left            =   4440
      Top             =   3480
   End
   Begin VB.Timer regionselectiontimer 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5280
      Top             =   3480
   End
   Begin VB.Frame FraSelectRegion 
      Caption         =   "Select region"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtBottomRightCoord 
         Height          =   285
         Left            =   3240
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtTopLeftCoord 
         Height          =   285
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdBottomRight 
         Caption         =   "Select"
         Height          =   300
         Left            =   4800
         TabIndex        =   4
         ToolTipText     =   "Click this button and move your cursor to the desired point. Wait for 2s"
         Top             =   840
         Width           =   800
      End
      Begin VB.CommandButton cmdTopLeft 
         Caption         =   "Select"
         Height          =   300
         Left            =   4800
         TabIndex        =   3
         ToolTipText     =   "Click this button and move your cursor to the desired point. Wait for 2s"
         Top             =   360
         Width           =   800
      End
      Begin VB.TextBox txtBottomright 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "bottomright"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtTopleft 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Text            =   "topleft"
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblBottomRight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bottom right"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lblTopLeft 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Top left"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Declare Function SetCursorPos _
                Lib "user32" (ByVal x As Long, _
                              ByVal y As Long) As Long

Private Declare Sub mouse_event _
                Lib "user32" (ByVal dwFlags As Long, _
                              ByVal dx As Long, _
                              ByVal dy As Long, _
                              ByVal cButtons As Long, _
                              ByVal dwExtraInfo As Long)

Private Const SWP_NOSIZE = &H1

Private Const SWP_SHOWWINDOW = &H40

Private Const SWP_NOMOVE = &H2

Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal x As Long, _
                              ByVal y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1

Private Const HWND_NOTOPMOST = -2

Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down

Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up

Private Type POINTAPI

    x As Long
    y As Long

End Type

Private Const VK_LBUTTON As Long = &H1

Private Const VK_CONTROL = &H11

Private Const VK_F2 = &H71

Private Const VK_F3 = &H72

Private topLeftSelect     As Boolean

Private bottomRightSelect As Boolean

Private topLeftPoint      As POINTAPI '//this iscontinuously changed during traversal

Private bottomRightPoint  As POINTAPI '//this is continuously changed during traversal

Private Sub chkRandom_Click()

    If chkRandom.Value = 0 Then
        '//random motion disabled
        txtDx.Enabled = True
        txtDy.Enabled = True
        OptClickEach.Enabled = True
        OptHoldClick.Enabled = True
    ElseIf chkRandom.Value = 1 Then
        '//random motion enabled
        txtDx.Enabled = False
        txtDy.Enabled = False
        OptClickEach.Enabled = False
        OptHoldClick.Enabled = False
    End If
    
End Sub

Private Sub chkTopMost_Click()

    If chkTopMost.Value = 1 Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
    ElseIf chkTopMost.Value = 0 Then
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_SHOWWINDOW
    End If
        
End Sub

Private Sub cmdAbout_Click()
    MsgBox "Author: s0ft" + vbNewLine + "Blog: c0dew0rth.blogspot.com" + vbNewLine + "Contact: yciloplabolg@gmail.com", vbOKOnly, ":D"
End Sub

Private Sub cmdBottomRight_Click()
    cmdBottomRight.Enabled = False
    bottomRightSelect = True
    topLeftSelect = False
    regionselectiontimer.Enabled = True
End Sub

Private Sub cmdHelp_Click()
    MsgBox "1. Select the top left and the bottom right points of the required rectangular region." + vbNewLine + "2. Specify the pixel distance in the horizontal (dx) and in the vertical (dy) directions for rectangular motion pattern. No such parameter is required for random motion." + vbNewLine + "3. Specify the interval between consecutive cursor translations and whether left click should be held or clicked between the points." + vbNewLine + "4. Press Ctrl+F2 to start. Press Ctrl+F3 to stop.", vbInformation, "Help"
End Sub

Private Sub cmdStart_Click()
    '//save the specified boundary points to global variables. these will be used and modified in the traversal
    topLeftPoint.x = Val(Split(txtTopLeftCoord.Text, ",")(0))
    topLeftPoint.y = Val(Split(txtTopLeftCoord.Text, ",")(1))
    bottomRightPoint.x = Val(Split(txtBottomRightCoord.Text, ",")(0))
    bottomRightPoint.y = Val(Split(txtBottomRightCoord.Text, ",")(1))
    
    cursorMoveTimer.Interval = Val(txtDelayMs.Text)
    cursorMoveTimer.Enabled = True
End Sub

Private Sub cmdStop_Click()
    cursorMoveTimer.Enabled = False
End Sub

Private Sub cmdTopLeft_Click()
    cmdTopLeft.Enabled = False
    bottomRightSelect = False
    topLeftSelect = True
    regionselectiontimer.Enabled = True
End Sub

Private Sub cursorMoveTimer_Timer()
    
    If chkRandom.Value = 0 Then '//if random motion is not checked
        HorizontalRowMotion
    ElseIf chkRandom.Value = 1 Then '//if random motion is checked
        RandomMotion
    End If
    
End Sub

Private Sub RandomMotion()

    Dim randomX, randomY As Integer

    randomX = topLeftPoint.x + Rnd() * (bottomRightPoint.x - topLeftPoint.x)
    randomY = topLeftPoint.y + Rnd() * (bottomRightPoint.y - topLeftPoint.y)
    
    SetCursorPos randomX, randomY
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub

Private Sub HorizontalRowMotion()

    Dim topleftx, toplefty, bottomrightx, bottomrighty As Integer

    topleftx = topLeftPoint.x
    toplefty = topLeftPoint.y
    bottomrightx = bottomRightPoint.x
    bottomrighty = bottomRightPoint.y

    SetCursorPos topleftx, toplefty

    If OptHoldClick.Value Then '//if left click is to be held down while moving cursor
        If topleftx = Val(Split(txtTopLeftCoord.Text, ",")(0)) And toplefty = Val(Split(txtTopLeftCoord.Text, ",")(1)) Then '//if this is the first point in the selected region
            mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        End If

    ElseIf OptClickEach.Value Then '//if instead left click should occur for each pixel moved to
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    End If

    If topleftx >= bottomrightx Then
        If toplefty >= bottomrighty Then
            cursorMoveTimer.Enabled = False '//reached the final pixel of the selected region

            Exit Sub

        Else
            toplefty = toplefty + Val(txtDy.Text) '//reached the end of the first row of the selected region. goto next row
            topleftx = Val(Split(txtTopLeftCoord.Text, ",")(0)) '//reset the column to the first
        End If

    Else
        topleftx = topleftx + Val(txtDx.Text) '//move onto the next column of the current row of the selected region
    End If

    '//update the first pixel of the traversal region
    topLeftPoint.x = topleftx
    topLeftPoint.y = toplefty

End Sub

Private Sub cursorpostimer_Timer()

    Dim position As POINTAPI

    If GetCursorPos(position) Then
        
        Dim coord As String

        coord = Trim$(Str$(position.x)) + "," + Trim$(Str$(position.y))
        txtTopleft.Text = coord
        txtBottomright.Text = coord
           
    End If
    
End Sub

Private Sub Form_Resize()
Me.Width = 6255
Me.Height = 4335
End Sub

Private Sub hotkeyListenerTimer_Timer()

    Dim keystateCtrl, keystateF2, keystateF3 As Integer

    keystateCtrl = GetAsyncKeyState(VK_CONTROL)

    If keystateCtrl = -32768 Then '//CONTROL key is being held down
        keystateF2 = GetAsyncKeyState(VK_F2)
        keystateF3 = GetAsyncKeyState(VK_F3)

        If keystateF2 <> 0 Then '//F2 pressed
            Call cmdStart_Click
        ElseIf keystateF3 <> 0 Then '//F3 pressed
            Call cmdStop_Click
        End If
        
    End If
    
End Sub

Private Sub regionselectiontimer_Timer()
    
    If (topLeftSelect) Then
        txtTopLeftCoord.Text = txtTopleft.Text
        topLeftSelect = False
        cmdTopLeft.Enabled = True
    ElseIf (bottomRightSelect) Then
        txtBottomRightCoord.Text = txtBottomright.Text
        bottomRightSelect = False
        cmdBottomRight.Enabled = True
    End If
    
    regionselectiontimer.Enabled = False
End Sub
