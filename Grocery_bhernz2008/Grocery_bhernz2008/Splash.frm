VERSION 5.00
Object = "{8770CE95-D0D2-4A5F-BD93-E531C279B841}#1.7#0"; "VCBUTT~1.OCX"
Begin VB.Form Splash 
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   3180
   ScaleMode       =   0  'User
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin vcButtonCTL.vcButton vcSI 
      Height          =   315
      Left            =   4320
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "System Info"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      ForeColor       =   255
      PictureLeftUp   =   "Splash.frx":0000
      PictureRightUp  =   "Splash.frx":02F2
      PictureMiddleUp =   "Splash.frx":05E4
      PictureLeftDown =   "Splash.frx":082E
      PictureRightDown=   "Splash.frx":0B20
      PictureMiddleDown=   "Splash.frx":0DBE
      HoverPictureLeft=   "Splash.frx":10B0
      HoverPictureRight=   "Splash.frx":132A
      HoverPictureMiddle=   "Splash.frx":1600
      MouseIcon       =   "Splash.frx":198E
   End
   Begin vcButtonCTL.vcButton vcOk 
      Height          =   315
      Left            =   5760
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   -1  'True
      ForeColor       =   255
      PictureLeftUp   =   "Splash.frx":1CA8
      PictureRightUp  =   "Splash.frx":1F9A
      PictureMiddleUp =   "Splash.frx":228C
      PictureLeftDown =   "Splash.frx":24D6
      PictureRightDown=   "Splash.frx":27C8
      PictureMiddleDown=   "Splash.frx":2A66
      HoverPictureLeft=   "Splash.frx":2D58
      HoverPictureRight=   "Splash.frx":2FD2
      HoverPictureMiddle=   "Splash.frx":32A8
      MouseIcon       =   "Splash.frx":3636
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   5760
      Picture         =   "Splash.frx":3950
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   1680
      Picture         =   "Splash.frx":1A9DA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   5640
      Picture         =   "Splash.frx":2FB3C
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Cell: +639184614971"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   120
      Picture         =   "Splash.frx":3097E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ALL RIGHTS RESERVE 2008"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "S Y S T E M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   135
      Left            =   2760
      TabIndex        =   3
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEMS"
      BeginProperty Font 
         Name            =   "Imprint MT Shadow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bernie O. Aquino        Philippines 2008"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   600
      X2              =   720
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Feed Back @ Web_struck24m@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   8010
      Left            =   -360
      Picture         =   "Splash.frx":36C08
      Top             =   -2400
      Width           =   8010
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub Image2_Click()
Unload Me

End Sub

Private Sub vcOk_Click()
Unload Me
Grocery.Show

End Sub

Private Sub vcSI_Click()
  Call StartSysInfo

End Sub
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Load()
  Me.Top = 3000
  Me.Left = (Screen.Width - Me.Width) / 2
 End Sub



