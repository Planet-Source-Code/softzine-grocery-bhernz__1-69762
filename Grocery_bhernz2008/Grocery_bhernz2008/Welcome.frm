VERSION 5.00
Begin VB.Form welcome_bhernz 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3225
   ClientLeft      =   5520
   ClientTop       =   2400
   ClientWidth     =   3765
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Welcome.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   215
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   251
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   0
      Picture         =   "Welcome.frx":24A2
      ScaleHeight     =   3195
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press ESC to exit me"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label lbl_day 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date Today"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   2
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Label lbl_time 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   12
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   960
         Width           =   2895
      End
   End
End
Attribute VB_Name = "welcome_bhernz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Const HWND_TOPMOST = -1
'Const HWND_NOTOPMOST = -2
'Const SWP_NOSIZE = &H1
'Const SWP_NOMOVE = &H2
'Const SWP_NOACTIVATE = &H10
'Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim i As Integer
Private Sub popup()
On Error Resume Next
    Picture1.Visible = True
    i = Me.Height
    Me.Height = 0
    While Me.Height < i
        Me.Height = Me.Height + 2
        Me.Top = Me.Top - 2
        DoEvents
    Wend
End Sub
Private Sub popdown()
On Error Resume Next
    i = Me.Height
    While Me.Height > 500
        Me.Height = Me.Height - 2
        Me.Top = Me.Top + 2
        DoEvents
    Wend
End Sub
Private Sub Form_Activate()
On Error Resume Next
    Grocery.Enabled = False
    lbl_time.Caption = "Login at:" & Format$(Now, "hh:mm:ss AM/PM")
    lbl_day.Caption = "Today:" & Format$(Date, "dd-MMM-yy")
    Call popup
    'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    'Sleep welcometime 'Wait for 1 Seconds
    'Call popdown
Grocery.Enabled = True

End Sub
Private Sub Form_Load()
On Error Resume Next
    Me.Left = Screen.Width - (Me.Width + 50)
    Me.Top = Screen.Height - 450 '450 assumed height for taskbar
    Picture1.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub
