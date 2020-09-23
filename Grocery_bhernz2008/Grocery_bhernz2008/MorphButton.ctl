VERSION 5.00
Begin VB.UserControl MorphButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2340
   ScaleHeight     =   71
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   156
End
Attribute VB_Name = "MorphButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Enum TRACKMOUSEEVENT_FLAGS
   TME_HOVER = &H1&
   TME_LEAVE = &H2&
   TME_QUERY = &H40000000
   TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
   cbSize                             As Long
   dwFlags                            As TRACKMOUSEEVENT_FLAGS
   hwndTrack                          As Long
   dwHoverTime                        As Long
End Type

Private bTrack                        As Boolean
Private bTrackUser32                  As Boolean
Private bInCtrl                       As Boolean
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'==================================================================================================
'Subclasser declarations.

' windows messages to be intercepted by subclassing.
Private Const WM_MOUSEMOVE            As Long = &H200
Private Const WM_MOUSELEAVE           As Long = &H2A3
Private Const WM_SETFOCUS             As Long = &H7
Private Const WM_KILLFOCUS            As Long = &H8

Private Enum eMsgWhen                                                       'When to callback
  MSG_BEFORE = 1                                                            'Callback before the original WndProc
  MSG_AFTER = 2                                                             'Callback after the original WndProc
  MSG_BEFORE_AFTER = MSG_BEFORE Or MSG_AFTER                                'Callback before and after the original WndProc
End Enum

Private Const ALL_MESSAGES  As Long = -1                                    'All messages callback
Private Const MSG_ENTRIES   As Long = 32                                    'Number of msg table entries
Private Const CODE_LEN      As Long = 240                                   'Thunk length in bytes
Private Const WNDPROC_OFF   As Long = &H30                                  'WndProc execution offset
Private Const MEM_LEN       As Long = CODE_LEN + (8 * (MSG_ENTRIES + 1))    'Bytes to allocate per thunk, data + code + msg tables
Private Const PAGE_RWX      As Long = &H40                                  'Allocate executable memory
Private Const MEM_COMMIT    As Long = &H1000                                'Commit allocated memory
Private Const GWL_WNDPROC   As Long = -4                                    'SetWindowsLong WndProc index
Private Const IDX_SHUTDOWN  As Long = 1                                     'Shutdown flag data index
Private Const IDX_HWND      As Long = 2                                     'hWnd data index
Private Const IDX_EBMODE    As Long = 3                                     'EbMode data index
Private Const IDX_CWP       As Long = 4                                     'CallWindowProc data index
Private Const IDX_SWL       As Long = 5                                     'SetWindowsLong data index
Private Const IDX_FREE      As Long = 6                                     'VirtualFree data index
Private Const IDX_ME        As Long = 7                                     'Owner data index
Private Const IDX_WNDPROC   As Long = 8                                     'Original WndProc data index
Private Const IDX_CALLBACK  As Long = 9                                     'zWndProc data index
Private Const IDX_BTABLE    As Long = 10                                    'Before table data index
Private Const IDX_ATABLE    As Long = 11                                    'After table data index
Private Const IDX_EBX       As Long = 14                                    'Data code index

Private z_Code(29)          As Currency                                     'Thunk machine-code initialised here
' not too familiar w/Paul's new stuff but zdata() ubound needs to be modified on per-control basis?
' ubound represents # of procedures in control.  This has 106 procedures, I padded it for future expansion.
Private z_Data(120)         As Long                                        'Array whose data pointer is re-mapped to arbitary memory addresses
Private z_DataDataPtr       As Long                                         'Address of z_Data()'s SafeArray data pointer
Private z_DataOrigData      As Long                                         'Address of z_Data()'s original data
Private z_hWnds             As Collection                                   'hWnd/thunk-address collection

Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

'==================================================================================================

'  declares for Unicode support.
Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Const VER_PLATFORM_WIN32_NT = 2
Private Type OSVERSIONINFO
   dwOSVersionInfoSize                As Long
   dwMajorVersion                     As Long
   dwMinorVersion                     As Long
   dwBuildNumber                      As Long
   dwPlatformId                       As Long
   szCSDVersion                       As String * 128        '  Maintenance string for PSS usage
End Type
Private mWindowsNT                    As Boolean
Private Const DT_CALCRECT             As Long = &H400        ' if used, DrawText API just calculates rectangle.
Private Const DT_SINGLELINE           As Long = &H20         ' strip cr/lf from string before draw.
Private Const DT_NOPREFIX             As Long = &H800        ' ignore access key ampersand.
Private Const DT_LEFT                 As Long = &H0          ' draw from left edge of rectangle.
Private Const DT_NOCLIP               As Long = &H100        ' ignores right edge of rectangle when drawing.

'  graphics api declares.
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveTo Lib "gdi32" Alias "MoveToEx" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'  declares for Carles P.V.'s gradient paint routine.
Private Type BITMAPINFOHEADER
   biSize                             As Long
   biWidth                            As Long
   biHeight                           As Long
   biPlanes                           As Integer
   biBitCount                         As Integer
   biCompression                      As Long
   biSizeImage                        As Long
   biXPelsPerMeter                    As Long
   biYPelsPerMeter                    As Long
   biClrUsed                          As Long
   biClrImportant                     As Long
End Type
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Const DIB_RGB_COLORS          As Long = 0
Private Const PI                      As Single = 3.14159265358979
Private Const TO_DEG                  As Single = 180 / PI
Private Const TO_RAD                  As Single = PI / 180
Private Const INT_ROT                 As Long = 1000

'  enum for determining if the control is a checkbox or option button.
Public Enum ControlTypeOptions
   [CheckBox]                                                ' control is a checkbox.
   [OptionButton]                                            ' control is an option button.
End Enum

Public Enum CBAlignmentOptions
   [Align Left]                                              ' checkbox is displayed in left side of control.
   [Align Center]
   [Align Right]                                             ' checkbox is displayed in right side of control.
End Enum

Public Enum MouseOverOptions
   [None]                                                    ' no graphics changes on mouseover.
   [Border]                                                  ' border color is changed on mouseover.
End Enum

'  for use in determining if mouse is in control.
Private Type POINTAPI
   x As Long                                                 ' horizontal pixel position.
   y As Long                                                 ' vertical pixel position.
End Type
Private MousePos As POINTAPI

'  rectangle structure for API drawing of text onto control.
Private Type RECT
   Left                               As Long
   Top                                As Long
   Right                              As Long
   Bottom                             As Long
End Type

' property default constants.
Private Const m_def_CaptionHover = vbBlack
Private Const m_def_MouseOverActions = 1                     ' no recoloring when mouseover.
Private Const m_def_MOverBorderColor = 0                     ' black control border when mouseover.
Private Const m_def_CaptionAlignment = 0                    ' checkbox aligned to left.
Private Const m_def_ControlType = 0                          ' checkbox control.
Private Const m_def_BackAngle = 90                           ' horizontal gradient.
Private Const m_def_BorderCurvature = 0                      ' no border curvature.
Private Const m_def_BorderWidth = 1                          ' 1-pixel border width.
Private Const m_def_BorderColor = 0                          ' black border.
Private Const m_def_BackMiddleOut = True                     ' background gradient is middle-out.
Private Const m_def_CaptionColor = 0                         ' black caption color.
Private Const m_def_BackColor1 = &H4040&                     ' gold gradient background.
Private Const m_def_BackColor2 = &HC0FFFF                    ' gold gradient background.
Private Const m_def_Value = 0                                ' option not selected/checked.
Private Const m_def_Enabled = True                           ' control is enabled.
Private Const m_def_Caption = "MorphButton"                  ' caption.
Private Const m_def_FocusRectColor = &H0                     ' black custom focus rectangle.
Private Const m_def_FocusEnabled = False
Private Const m_def_CaptionClick = vbBlack

'  these variables allow the control to switch between enabled/disabled appearances.
Private ActiveBackColor1              As OLE_COLOR           ' current first background gradient color.
Private ActiveBackColor2              As OLE_COLOR           ' current second background gradient color.
Private ActiveBorderColor             As OLE_COLOR           ' current control border color.
Private ActiveCaptionColor            As OLE_COLOR           ' current caption text color.

' property variables.
Private m_CaptionHover                As OLE_COLOR
Private m_MouseOverActions            As MouseOverOptions    ' what gets colored when mouse is over control?
Private m_MOverBorderColor            As OLE_COLOR           ' control border color when mouse is over control.
Private m_CaptionAlignment           As CBAlignmentOptions  ' left or right checkbox alignment enum.
Private m_ControlType                 As ControlTypeOptions  ' allows selection of the type of control to display.
Private m_BackAngle                   As Single              ' angle of control background gradient.
Private m_BorderCurvature             As Integer             ' amount of curvature the control's corners are to have.
Private m_Font                        As Font                ' the font used to draw the caption.
Private m_CheckBorderColor            As OLE_COLOR           ' the color of the checkbox border.
Private m_BorderWidth                 As Long                ' the width of the control's border.
Private m_BorderColor                 As OLE_COLOR           ' the color of the control's border.
Private m_BackMiddleOut               As Boolean             ' if true, control background gradient is middle-out.
Private m_CaptionColor                As OLE_COLOR           ' the color of the caption text.
Private m_BackColor1                  As OLE_COLOR           ' the first color of the control's background gradient.
Private m_BackColor2                  As OLE_COLOR           ' the second color of the control's background gradient.
Private m_Value                       As Integer             ' the selected status of the optionbutton/checkbox.
Private m_Enabled                     As Boolean             ' control's enabled status.
Private m_Caption                     As String              ' the text to display alongside the checkbox.
Private m_FocusRectColor              As OLE_COLOR           ' the color of the custom focus rectangle.
Private m_FocusEnabled                As Boolean
Private m_CaptionClick                As OLE_COLOR

' event declarations.
Public Event Click()
Public Event MouseEnter()
Public Event MouseLeave()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Private HasFocus As Boolean                                  ' indicates if MorphOptionButton has focus.
Private CheckBoxX As Long                                    ' x coordinate of left edge of checkbox.
Private KeyIsDown As Boolean                                 ' indicates if a key is being pressed.
Private MouseIsDown As Boolean                               ' indicates if the left mouse mutton is down.
Private SaveBorderColor As Long                              ' original control border color.
Dim r           As RECT
Dim TextColor1 As Long
Dim CheckColor As Long

Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)

'*************************************************************************
'* track the mouse leaving the indicated window.                         *
'*************************************************************************

   Dim tme As TRACKMOUSEEVENT_STRUCT

   If bTrack Then
      With tme
         .cbSize = Len(tme)
         .dwFlags = TME_LEAVE
         .hwndTrack = lng_hWnd
       End With
       If bTrackUser32 Then
         Call TrackMouseEvent(tme)
       Else
         Call TrackMouseEventComCtl(tme)
       End If
   End If

End Sub



'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<< Event-Handling Routines >>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* not used in the control, but passed along to the rest of the project. *
'*************************************************************************

   If m_Enabled Then
      KeyIsDown = True
      RaiseEvent KeyDown(KeyCode, Shift)
   End If

End Sub

Private Sub UserControl_Initialize()

'*************************************************************************
'* the first event in the control's life cycle.                          *
'*************************************************************************

   Dim OS As OSVERSIONINFO

   OS.dwOSVersionInfoSize = Len(OS)
   Call GetVersionEx(OS)
   mWindowsNT = ((OS.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

'*************************************************************************
'* allows the space bar to "click" the control like vb equivalents.      *
'*************************************************************************

   KeyIsDown = False
'  if mouse button is down, ignore.
   If MouseIsDown Then
      Exit Sub
   End If

'  only "click" control if the control is enabled and key pressed is space bar.
   If m_Enabled Then
      If KeyCode = 32 Then
            m_Value = IIf(m_Value = vbChecked, vbUnchecked, vbChecked)
            RedrawControl
      End If
     
'     pass along the KeyUp and KeyPress events to project regardless of key pressed.
      RaiseEvent KeyUp(KeyCode, Shift)
      RaiseEvent KeyPress(KeyCode)
   End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'*************************************************************************
'* not used in the control, but passed along to the rest of the project. *
'*************************************************************************
   BorderWidth = 1
   If m_Enabled Then
      MouseIsDown = True
      RaiseEvent MouseDown(Button, Shift, x, y)
   End If
  If Button = vbLeftButton Then
     BorderWidth = BorderWidth + 1
     CaptionColor = CaptionClick
     RaiseEvent Click
  End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

'*************************************************************************
'* sets MorphOptionButton's value to True (only on left button click).   *
'*************************************************************************

   MouseIsDown = False
'  if a key is down, ignore.
   If KeyIsDown Then
      Exit Sub
   End If
   If BorderWidth < 2 Then BorderWidth = 2
'  if mouse has left control, ignore.  This allows you to "back out" of a click, so
'  to speak, by holding down mouse button, dragging the mouse out, and releasing.
   GetCursorPos MousePos
   If WindowFromPoint(MousePos.x, MousePos.y) <> UserControl.hWnd Then
      Exit Sub   ' don't send event.
   End If

'  only bother if the control is enabled.
   If m_Enabled Then
         If Button = vbLeftButton Then
            BorderWidth = BorderWidth - 1
            m_Value = IIf(m_Value = vbChecked, vbUnchecked, vbChecked)
            RedrawControl
         End If
'     pass along both the MouseUp and Click events (Click on left button only).
      RaiseEvent MouseUp(Button, Shift, x, y)
   End If

End Sub

Private Sub UserControl_Resize()
   RedrawControl
End Sub

Private Sub UserControl_Show()
   GetEnabledDisplayProperties
   RedrawControl
End Sub

Private Sub RedrawControl()

'*************************************************************************
'* the master routine for displaying textbox and its contents.           *
'*************************************************************************
   
   SetBackGround      ' display the background gradient.
   SetBorder          ' display the control's border, if defined.
   SetText            ' display the caption.

End Sub

Private Sub SetBackGround()

'*************************************************************************
'* displays the control's background gradient.                           *
'*************************************************************************

   PaintGradient hDC, 0, 0, ScaleWidth, ScaleHeight, TranslateColor(ActiveBackColor1), _
                 TranslateColor(ActiveBackColor2), m_BackAngle, m_BackMiddleOut

End Sub

Private Sub SetBorder()

'*************************************************************************
'* draws the border around the control, using appropriate curvature.     *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim hBrush As Long    ' the brush pattern used to 'paint' the border.
   Dim hRgn1  As Long    ' the outer boundary of the border region.
   Dim hRgn2  As Long    ' the inner boundary of the border region.

   hRgn1 = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, m_BorderCurvature, m_BorderCurvature)
   hRgn2 = CreateRoundRectRgn(m_BorderWidth, m_BorderWidth, ScaleWidth - m_BorderWidth, _
                              ScaleHeight - m_BorderWidth, m_BorderCurvature, m_BorderCurvature)
   CombineRgn hRgn2, hRgn1, hRgn2, 3

'  create and apply the color brush.
   hBrush = CreateSolidBrush(TranslateColor(ActiveBorderColor))
   FillRgn hDC, hRgn2, hBrush

'  set the control region.
   SetWindowRgn hWnd, hRgn1, True

'  free the memory.
   DeleteObject hRgn2
   DeleteObject hBrush
   DeleteObject hRgn1

ErrHandler:
   Exit Sub

End Sub


Private Sub SetText()

'*************************************************************************
'* displays the caption text.  Selected text is displayed using the      *
'* SelTextColor value.                                                   *
'*************************************************************************

   If Not m_Font Is Nothing Then

     ' Dim r           As RECT      ' the rectangle that defines the text draw area.
      Dim tHeight     As Long      ' the height of the text.
      Dim tWidth      As Long      ' the width of the text.
      Dim Clearance   As Long      ' (to end of checkbox) + 1 letter width = clearance.


'     get the height and width of the text based on the selected font.
      tHeight = TextHeight(m_Caption)
      tWidth = TextWidthU(hDC, m_Caption)

'     make the left clearance one letter width.  Also account for up to right edge of checkbox.
      If m_CaptionAlignment = [Align Left] Then Clearance = TextWidthU(hDC, "n")
      If m_CaptionAlignment = [Align Center] Then Clearance = (ScaleWidth / 2) - (tWidth / 2)
      If m_CaptionAlignment = [Align Right] Then Clearance = ScaleWidth - tWidth - TextWidthU(hDC, "n")
      
'     set the text color.
      UserControl.ForeColor = TextColor1  ' CaptionColor 'TranslateColor(ActiveCaptionColor)

'     define the text drawing area rectangle size.
      If MouseIsDown = False Then
      With r
         .Left = Clearance
         .Top = (ScaleHeight - tHeight) / 2
         .Bottom = r.Top + tHeight
         .Right = .Left + tWidth
      End With
      Else
      With r
         .Left = Clearance + 1
         .Top = ((ScaleHeight - tHeight) / 2) + 1
         .Bottom = r.Top + tHeight
         .Right = .Left + tWidth
      End With
      End If
'     display the text using DrawText API.
      DrawText UserControl.hDC, m_Caption, -1, r, DT_LEFT Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP

'     if the MorphOptionButton has the focus, draw a focus rectangle
'     around the text, adding a one-pixel clearance around the text.
   
      If HasFocus And FocusEnabled Then
         With r
            .Left = .Left - 1
            .Top = .Top - 1
            .Bottom = .Bottom + 1
            .Right = .Right + 1
           DrawRectangle .Left, .Top, .Right, .Bottom, m_FocusRectColor
         End With
      End If
   End If

End Sub

Private Sub DrawRectangle(x1 As Long, y1 As Long, x2 As Long, y2 As Long, lColor As Long)

'*************************************************************************
'* draws the checkbox and focus rectangles.                              *
'*************************************************************************

   On Error GoTo ErrHandler

   Dim hBrush As Long    ' the brush pattern used to 'paint' the border.
   Dim hRgn1  As Long    ' the outer boundary of the rectangle region.
   Dim hRgn2  As Long    ' the inner boundary of the rectangle region.

'  create the outer region.
   hRgn1 = CreateRoundRectRgn(x1, y1, x2, y2, 0, 0)
'  create the inner region.
   hRgn2 = CreateRoundRectRgn(x1 + 1, y1 + 1, x2 - 1, y2 - 1, 0, 0)
   
'  combine the regions into one border region.
   CombineRgn hRgn2, hRgn1, hRgn2, 3

'  create and apply the color brush.
   hBrush = CreateSolidBrush(TranslateColor(lColor))
   FillRgn hDC, hRgn2, hBrush

'  free the memory.
   DeleteObject hRgn2
   DeleteObject hBrush
   DeleteObject hRgn1

ErrHandler:
   Exit Sub

End Sub

Private Function TextWidthU(ByVal hDC As Long, sString As String) As Long

'*************************************************************************
'* a better alternative to the VB method .TextWidth.  Thanks LaVolpe!    *
'*************************************************************************

   Dim Flags    As Long
   Dim TextRect As RECT

   SetRect TextRect, 0, 0, 0, 0
   Flags = DT_CALCRECT Or DT_SINGLELINE Or DT_NOPREFIX Or DT_LEFT
   DrawText hDC, sString, -1, TextRect, Flags
   TextWidthU = TextRect.Right + 1

End Function

Private Sub DrawText(ByVal hDC As Long, ByVal lpString As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long)

'*************************************************************************
'* draws the text with Unicode support based on OS version.              *
'* Thanks to Richard Mewett.                                             *
'*************************************************************************

   If mWindowsNT Then
      DrawTextW hDC, StrPtr(lpString), nCount, lpRect, wFormat
   Else
      DrawTextA hDC, lpString, nCount, lpRect, wFormat
   End If

End Sub

Private Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long

'*************************************************************************
'* translates ole color into COLORREF long for drawing purposes.         *
'*************************************************************************

   If OleTranslateColor(oClr, hPal, TranslateColor) Then
      TranslateColor = -1
   End If

End Function

Private Sub PaintGradient(ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, _
                          ByVal Height As Long, ByVal Color1 As Long, ByVal Color2 As Long, _
                          ByVal Angle As Single, ByVal bMOut As Boolean)

'*************************************************************************
'* Carles P.V.'s routine, modified by Matthew R. Usner for middle-out    *
'* gradient capability.  Original submission at PSC, txtCodeID=60580.    *
'*************************************************************************

   Dim uBIH      As BITMAPINFOHEADER
   Dim lBits()   As Long
   Dim lGrad()   As Long, lGrad2() As Long

   Dim lClr      As Long
   Dim R1        As Long, G1 As Long, b1 As Long
   Dim R2        As Long, G2 As Long, b2 As Long
   Dim dR        As Long, dG As Long, dB As Long

   Dim Scan      As Long
   Dim i         As Long, j As Long, k As Long
   Dim jIn       As Long
   Dim iEnd      As Long, jEnd As Long
   Dim Offset    As Long

   Dim lQuad     As Long
   Dim AngleDiag As Single
   Dim AngleComp As Single

   Dim g         As Long
   Dim luSin     As Long, luCos As Long
 
   If (Width > 0 And Height > 0) Then

'     Matthew R. Usner - solves weird problem of when angle is
'     >= 91 and <= 270, the colors invert in MiddleOut mode.
      If bMOut And Angle >= 91 And Angle <= 270 Then
         g = Color1
         Color1 = Color2
         Color2 = g
      End If

'     -- Right-hand [+] (ox=0º)
      Angle = -Angle + 90

'     -- Normalize to [0º;360º]
      Angle = Angle Mod 360
      If (Angle < 0) Then
         Angle = 360 + Angle
      End If

'     -- Get quadrant (0 - 3)
      lQuad = Angle \ 90

'     -- Normalize to [0º;90º]
        Angle = Angle Mod 90

'     -- Calc. gradient length ('distance')
      If (lQuad Mod 2 = 0) Then
         AngleDiag = Atn(Width / Height) * TO_DEG
      Else
         AngleDiag = Atn(Height / Width) * TO_DEG
      End If
      AngleComp = (90 - Abs(Angle - AngleDiag)) * TO_RAD
      Angle = Angle * TO_RAD
      g = Sqr(Width * Width + Height * Height) * Sin(AngleComp) 'Sinus theorem

'     -- Decompose colors
      If (lQuad > 1) Then
         lClr = Color1
         Color1 = Color2
         Color2 = lClr
      End If
      R1 = (Color1 And &HFF&)
      G1 = (Color1 And &HFF00&) \ 256
      b1 = (Color1 And &HFF0000) \ 65536
      R2 = (Color2 And &HFF&)
      G2 = (Color2 And &HFF00&) \ 256
      b2 = (Color2 And &HFF0000) \ 65536

'     -- Get color distances
      dR = R2 - R1
      dG = G2 - G1
      dB = b2 - b1

'     -- Size gradient-colors array
      ReDim lGrad(0 To g - 1)
      ReDim lGrad2(0 To g - 1)

'     -- Calculate gradient-colors
      iEnd = g - 1
      If (iEnd = 0) Then
'        -- Special case (1-pixel wide gradient)
         lGrad2(0) = (b1 \ 2 + b2 \ 2) + 256 * (G1 \ 2 + G2 \ 2) + 65536 * (R1 \ 2 + R2 \ 2)
      Else
         For i = 0 To iEnd
            lGrad2(i) = b1 + (dB * i) \ iEnd + 256 * (G1 + (dG * i) \ iEnd) + 65536 * (R1 + (dR * i) \ iEnd)
         Next i
      End If

'     'if block' added by Matthew R. Usner - accounts for possible MiddleOut gradient draw.
      If bMOut Then
         k = 0
         For i = 0 To iEnd Step 2
            lGrad(k) = lGrad2(i)
            k = k + 1
         Next i
         For i = iEnd - 1 To 1 Step -2
            lGrad(k) = lGrad2(i)
            k = k + 1
         Next i
      Else
         For i = 0 To iEnd
            lGrad(i) = lGrad2(i)
         Next i
      End If

'     -- Size DIB array
      ReDim lBits(Width * Height - 1) As Long
      iEnd = Width - 1
      jEnd = Height - 1
      Scan = Width

'     -- Render gradient DIB
      Select Case lQuad

         Case 0, 2
            luSin = Sin(Angle) * INT_ROT
            luCos = Cos(Angle) * INT_ROT
            Offset = 0
            jIn = 0
            For j = 0 To jEnd
               For i = 0 To iEnd
                  lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
               Next i
               jIn = jIn + luCos
               Offset = Offset + Scan
            Next j

         Case 1, 3
            luSin = Sin(90 * TO_RAD - Angle) * INT_ROT
            luCos = Cos(90 * TO_RAD - Angle) * INT_ROT
            Offset = jEnd * Scan
            jIn = 0
            For j = 0 To jEnd
               For i = 0 To iEnd
                  lBits(i + Offset) = lGrad((i * luSin + jIn) \ INT_ROT)
               Next i
               jIn = jIn + luCos
               Offset = Offset - Scan
            Next j

      End Select

'     -- Define DIB header
      With uBIH
         .biSize = 40
         .biPlanes = 1
         .biBitCount = 32
         .biWidth = Width
         .biHeight = Height
      End With

'     -- Paint it!
      If MouseIsDown = True Then
      Call StretchDIBits(hDC, x + 1, y + 1, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
      Else
       Call StretchDIBits(hDC, x, y, Width, Height, 0, 0, Width, Height, lBits(0), uBIH, DIB_RGB_COLORS, vbSrcCopy)
      End If
    End If

End Sub

' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< Property Routines  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
' <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub UserControl_InitProperties()

'*************************************************************************
'* initialize properties to default values.                              *
'*************************************************************************

   Set m_Font = Ambient.Font
   m_CaptionHover = m_def_CaptionHover
   m_Enabled = m_def_Enabled
   m_Caption = m_def_Caption
   m_Value = m_def_Value
   m_BackColor1 = m_def_BackColor1
   m_BackColor2 = m_def_BackColor2
   m_CaptionColor = m_def_CaptionColor
   m_BackMiddleOut = m_def_BackMiddleOut
   m_BorderColor = m_def_BorderColor
   m_BorderWidth = m_def_BorderWidth
   m_BorderCurvature = m_def_BorderCurvature
   m_BackAngle = m_def_BackAngle
   m_ControlType = m_def_ControlType
   m_CaptionAlignment = m_def_CaptionAlignment
   m_MOverBorderColor = m_def_MOverBorderColor
   m_MouseOverActions = m_def_MouseOverActions
   m_FocusEnabled = m_def_FocusEnabled
   m_CaptionClick = m_def_CaptionClick
   
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'*************************************************************************
'* load property values from storage.                                    *
'*************************************************************************

   With PropBag
      Set m_Font = .ReadProperty("Font", Ambient.Font)
      Set UserControl.Font = m_Font
      m_Enabled = .ReadProperty("Enabled", m_def_Enabled)
      m_Caption = .ReadProperty("Caption", m_def_Caption)
      m_Value = .ReadProperty("Value", m_def_Value)
      m_BackColor1 = .ReadProperty("BackColor1", m_def_BackColor1)
      m_BackColor2 = .ReadProperty("BackColor2", m_def_BackColor2)
      m_CaptionColor = .ReadProperty("CaptionColor", m_def_CaptionColor)
      m_BackMiddleOut = .ReadProperty("BackMiddleOut", m_def_BackMiddleOut)
      m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
      m_BorderWidth = .ReadProperty("BorderWidth", m_def_BorderWidth)
      m_BorderCurvature = .ReadProperty("BorderCurvature", m_def_BorderCurvature)
      m_BackAngle = .ReadProperty("BackAngle", m_def_BackAngle)
      m_ControlType = .ReadProperty("ControlType", m_def_ControlType)
      m_CaptionAlignment = .ReadProperty("CaptionAlignment", m_def_CaptionAlignment)
      m_MOverBorderColor = .ReadProperty("MOverBorderColor", m_def_MOverBorderColor)
      m_MouseOverActions = .ReadProperty("MouseOverActions", m_def_MouseOverActions)
      m_FocusRectColor = .ReadProperty("FocusRectColor", m_def_FocusRectColor)
      m_CaptionHover = .ReadProperty("CaptionHover", m_def_CaptionHover)
      m_CaptionClick = .ReadProperty("CaptionClick", m_def_CaptionClick)
      m_FocusEnabled = .ReadProperty("FocusEnabled", m_def_FocusEnabled)
      
   End With

'  save the default border and checkbox border colors so that when the
'  mouse cursor leaves the control we can restore the original color(s).
   SaveBorderColor = m_BorderColor
   TextColor1 = m_CaptionColor
   CheckColor = m_CaptionColor
   FocusEnabled = m_FocusEnabled
   
  If Ambient.UserMode Then
     bTrack = True
     With UserControl
        sc_Subclass .hWnd                            ' subclass a window handle... or three
        sc_AddMsg .hWnd, WM_MOUSEMOVE, MSG_AFTER     ' for mouse enter detect.
        sc_AddMsg .hWnd, WM_MOUSELEAVE, MSG_AFTER    ' for mouse leave detect.
        sc_AddMsg .hWnd, WM_SETFOCUS, MSG_AFTER      ' for got focus detect.
        sc_AddMsg .hWnd, WM_KILLFOCUS, MSG_AFTER     ' for lost focus detect.
     End With
  End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'*************************************************************************
'* write property values to storage.                                     *
'*************************************************************************

   With PropBag
      .WriteProperty "Enabled", m_Enabled, m_def_Enabled
      .WriteProperty "Caption", m_Caption, m_def_Caption
      .WriteProperty "Value", m_Value, m_def_Value
      .WriteProperty "BackColor1", m_BackColor1, m_def_BackColor1
      .WriteProperty "BackColor2", m_BackColor2, m_def_BackColor2
      .WriteProperty "CaptionColor", m_CaptionColor, m_def_CaptionColor
      .WriteProperty "BackMiddleOut", m_BackMiddleOut, m_def_BackMiddleOut
      .WriteProperty "BorderColor", m_BorderColor, m_def_BorderColor
      .WriteProperty "BorderWidth", m_BorderWidth, m_def_BorderWidth
      .WriteProperty "Font", m_Font, Ambient.Font
      .WriteProperty "BorderCurvature", m_BorderCurvature, m_def_BorderCurvature
      .WriteProperty "BackAngle", m_BackAngle, m_def_BackAngle
      .WriteProperty "ControlType", m_ControlType, m_def_ControlType
      .WriteProperty "CaptionAlignment", m_CaptionAlignment, m_def_CaptionAlignment
      .WriteProperty "MOverBorderColor", m_MOverBorderColor, m_def_MOverBorderColor
      .WriteProperty "MouseOverActions", m_MouseOverActions, m_def_MouseOverActions
      .WriteProperty "FocusRectColor", m_FocusRectColor, m_def_FocusRectColor
      .WriteProperty "CaptionHover", m_CaptionHover, m_def_CaptionHover
      .WriteProperty "CaptionClick", m_CaptionClick, m_def_CaptionClick
      .WriteProperty "FocusEnabled", m_FocusEnabled, m_def_FocusEnabled
      
   End With

End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Sets/returns whether or not the control can be accessed."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Misc"
   Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
   m_Enabled = New_Enabled
   If m_Enabled Then GetEnabledDisplayProperties
   PropertyChanged "Enabled"
   RedrawControl
End Property

Public Property Get FocusEnabled() As Boolean
   FocusEnabled = m_FocusEnabled
End Property

Public Property Let FocusEnabled(ByVal New_FocusEnabled As Boolean)
   m_FocusEnabled = New_FocusEnabled
   PropertyChanged "FocusEnabled"
   RedrawControl
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "The text to display on the MorphOptionCheck control."
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Text"
   Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   RedrawControl
End Property

Public Property Get CaptionClick() As OLE_COLOR
   CaptionClick = m_CaptionClick
End Property

Public Property Let CaptionClick(ByVal New_CaptionClick As OLE_COLOR)
   m_CaptionClick = New_CaptionClick
   PropertyChanged "CaptionClick"
   RedrawControl
End Property

Public Property Get CaptionHover() As OLE_COLOR
   CaptionHover = m_CaptionHover
End Property

Public Property Let CaptionHover(ByVal New_CaptionHover As OLE_COLOR)
   m_CaptionHover = New_CaptionHover
   PropertyChanged "CaptionHover"
   RedrawControl
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Text"
Attribute Font.VB_UserMemId = -512
   Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As StdFont)
   Set m_Font = New_Font
   Set UserControl.Font = m_Font
   PropertyChanged "Font"
   RedrawControl
End Property

Public Property Get Value() As Integer
Attribute Value.VB_Description = "MorphOptionCheck status (Checked, Unchecked, True, False)."
   Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
   m_Value = New_Value
   PropertyChanged "Value"
   RedrawControl
End Property

Public Property Get BackColor1() As OLE_COLOR
Attribute BackColor1.VB_Description = "The first gradient color for the MorphCheckBox background."
Attribute BackColor1.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BackColor1 = m_BackColor1
End Property

Public Property Let BackColor1(ByVal New_BackColor1 As OLE_COLOR)
   m_BackColor1 = New_BackColor1
   PropertyChanged "BackColor1"
   RedrawControl
End Property

Public Property Get BackColor2() As OLE_COLOR
Attribute BackColor2.VB_Description = "The second gradient color for the MorphCheckBox background."
Attribute BackColor2.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BackColor2 = m_BackColor2
End Property

Public Property Let BackColor2(ByVal New_BackColor2 As OLE_COLOR)
   m_BackColor2 = New_BackColor2
   PropertyChanged "BackColor2"
   RedrawControl
End Property

Public Property Get CaptionColor() As OLE_COLOR
Attribute CaptionColor.VB_Description = "The color of the MorphOptionCheck text."
Attribute CaptionColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   CaptionColor = m_CaptionColor
End Property

Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
   m_CaptionColor = New_CaptionColor
   TextColor1 = m_CaptionColor
   PropertyChanged "CaptionColor"
   RedrawControl
End Property

Public Property Get BackMiddleOut() As Boolean
Attribute BackMiddleOut.VB_Description = "When True, the background gradient is displayed in middle-out fashion."
Attribute BackMiddleOut.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BackMiddleOut = m_BackMiddleOut
End Property

Public Property Let BackMiddleOut(ByVal New_BackMiddleOut As Boolean)
   m_BackMiddleOut = New_BackMiddleOut
   PropertyChanged "BackMiddleOut"
   RedrawControl
End Property

Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "The color of the control's border."
Attribute BorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
   m_BorderColor = New_BorderColor
   PropertyChanged "BorderColor"
   RedrawControl
End Property

Public Property Get BorderWidth() As Long
Attribute BorderWidth.VB_Description = "The width, in pixels, of the MorphCheckBox border."
Attribute BorderWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Long)
   m_BorderWidth = New_BorderWidth
   PropertyChanged "BorderWidth"
   RedrawControl
End Property

Public Property Get BorderCurvature() As Integer
Attribute BorderCurvature.VB_Description = "The amount of curvature that each corner of the control is to have."
Attribute BorderCurvature.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BorderCurvature = m_BorderCurvature
End Property

Public Property Let BorderCurvature(ByVal New_BorderCurvature As Integer)
   m_BorderCurvature = New_BorderCurvature
   PropertyChanged "BorderCurvature"
   RedrawControl
End Property

Public Property Get BackAngle() As Single
Attribute BackAngle.VB_Description = "The gradient angle of the background."
Attribute BackAngle.VB_ProcData.VB_Invoke_Property = ";Appearance"
   BackAngle = m_BackAngle
End Property

Public Property Let BackAngle(ByVal New_BackAngle As Single)
   m_BackAngle = New_BackAngle
   PropertyChanged "BackAngle"
   RedrawControl
End Property

Public Property Get hWnd() As Long
   hWnd = UserControl.hWnd
End Property

Public Property Get hDC() As Long
   hDC = UserControl.hDC
End Property

Public Property Get CaptionAlignment() As CBAlignmentOptions
Attribute CaptionAlignment.VB_Description = "Sets the checkbox to left or right side of the control."
Attribute CaptionAlignment.VB_ProcData.VB_Invoke_Property = ";Behavior"
   CaptionAlignment = m_CaptionAlignment
End Property

Public Property Let CaptionAlignment(ByVal New_CaptionAlignment As CBAlignmentOptions)
   m_CaptionAlignment = New_CaptionAlignment
   PropertyChanged "CaptionAlignment"
   RedrawControl
End Property

Public Property Get MOverBorderColor() As OLE_COLOR
Attribute MOverBorderColor.VB_Description = "The border color when the mouse pointer is over the control."
Attribute MOverBorderColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
   MOverBorderColor = m_MOverBorderColor
End Property

Public Property Let MOverBorderColor(ByVal New_MOverBorderColor As OLE_COLOR)
   m_MOverBorderColor = New_MOverBorderColor
   PropertyChanged "MOverBorderColor"
End Property

Public Property Get FocusRectColor() As OLE_COLOR
   FocusRectColor = m_FocusRectColor
End Property

Public Property Let FocusRectColor(ByVal New_FocusRectColor As OLE_COLOR)
   m_FocusRectColor = New_FocusRectColor
   PropertyChanged "FocusRectColor"
End Property

Private Sub UserControl_Terminate()
   On Error GoTo Catch
   sc_Terminate
Catch:
End Sub

Private Sub GetEnabledDisplayProperties()

'*************************************************************************
'* applies enabled graphics properties to the active display properties. *
'*************************************************************************

   ActiveBackColor1 = m_BackColor1
   ActiveBackColor2 = m_BackColor2
   ActiveBorderColor = m_BorderColor
   ActiveCaptionColor = m_CaptionColor
End Sub

'-uSelfSub code-----------------------------------------------------------------------------------
Private Function sc_Subclass(ByVal lng_hWnd As Long) As Boolean             'Subclass the specified window handle
  Dim nAddr As Long
  
  If IsWindow(lng_hWnd) = 0 Then                                            'Ensure the window handle is valid
    zError "sc_Subclass", "Invalid window handle"
  End If
  
  If z_hWnds Is Nothing Then
    RtlMoveMemory VarPtr(z_DataDataPtr), VarPtrArray(z_Data), 4             'Get the address of z_Data()'s SafeArray header
    z_DataDataPtr = z_DataDataPtr + 12                                      'Bump the address to point to the pvData data pointer
    RtlMoveMemory VarPtr(z_DataOrigData), z_DataDataPtr, 4                  'Get the value of z_Data()'s SafeArray pvData data pointer
  
    nAddr = zGetCallback                                                    'Get the address of this UserControl's zWndProc callback routine
    
    'Initialise the machine-code thunk
    z_Code(6) = -490736517001394.5807@: z_Code(7) = 484417356483292.94@: z_Code(8) = -171798741966746.6996@: z_Code(9) = 843649688964536.7412@: z_Code(10) = -330085705188364.0817@: z_Code(11) = 41621208.9739@: z_Code(12) = -900372920033759.9903@: z_Code(13) = 291516653989344.1016@: z_Code(14) = -621553923181.6984@: z_Code(15) = 291551690021556.6453@: z_Code(16) = 28798458374890.8543@: z_Code(17) = 86444073845629.4399@: z_Code(18) = 636540268579660.4789@: z_Code(19) = 60911183420250.2143@: z_Code(20) = 846934495644380.8767@: z_Code(21) = 14073829823.4668@: z_Code(22) = 501055845239149.5051@: z_Code(23) = 175724720056981.1236@: z_Code(24) = 75457451135513.7931@: z_Code(25) = -576850389355798.3357@: z_Code(26) = 146298060653075.5445@: z_Code(27) = 850256350680294.7583@: z_Code(28) = -4888724176660.092@: z_Code(29) = 21456079546.6867@
    
    zMap VarPtr(z_Code(0))                                                  'Map the address of z_Code()'s first element to the z_Data() array
    z_Data(IDX_EBMODE) = zFnAddr("vba6", "EbMode")                          'Store the EbMode function address in the thunk data
    z_Data(IDX_CWP) = zFnAddr("user32", "CallWindowProcA")                  'Store CallWindowProc function address in the thunk data
    z_Data(IDX_SWL) = zFnAddr("user32", "SetWindowLongA")                   'Store the SetWindowLong function address in the thunk data
    z_Data(IDX_FREE) = zFnAddr("kernel32", "VirtualFree")                   'Store the VirtualFree function address in the thunk data
    z_Data(IDX_ME) = ObjPtr(Me)                                             'Store my object address in the thunk data
    z_Data(IDX_CALLBACK) = nAddr                                            'Store the zWndProc address in the thunk data
    zMap z_DataOrigData                                                     'Restore z_Data()'s original data pointer
    
    Set z_hWnds = New Collection                                            'Create the window-handle/thunk-memory-address collection
  End If

  nAddr = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX)                    'Allocate executable memory
  RtlMoveMemory nAddr, VarPtr(z_Code(0)), CODE_LEN                          'Copy the machine-code to the allocated memory

  On Error GoTo Catch                                                       'Catch double subclassing
    z_hWnds.Add nAddr, "h" & lng_hWnd                                       'Add the hWnd/thunk-address to the collection
  On Error GoTo 0

  zMap nAddr                                                                'Map z_Data() to the subclass thunk machine-code
  z_Data(IDX_EBX) = nAddr                                                   'Patch the data address
  z_Data(IDX_HWND) = lng_hWnd                                               'Store the window handle in the thunk data
  z_Data(IDX_BTABLE) = nAddr + CODE_LEN                                     'Store the address of the before table in the thunk data
  z_Data(IDX_ATABLE) = z_Data(IDX_BTABLE) + ((MSG_ENTRIES + 1) * 4)         'Store the address of the after table in the thunk data
  nAddr = nAddr + WNDPROC_OFF                                               'Execution address of the thunk's WndProc
  z_Data(IDX_WNDPROC) = SetWindowLongA(lng_hWnd, GWL_WNDPROC, nAddr)        'Set the new WndProc and store the original WndProc in the thunk data
  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
  sc_Subclass = True                                                        'Indicate success
  Exit Function                                                             'Exit

Catch:
  zError "sc_Subclass", "Window handle is already subclassed"
End Function

'Terminate all subclassing
Private Sub sc_Terminate()
  Dim i     As Long
  Dim nAddr As Long

  If z_hWnds Is Nothing Then                                                'Ensure that subclassing has been started
  Else
    With z_hWnds
      For i = .Count To 1 Step -1                                           'Loop through the collection of window handles in reverse order
        nAddr = .Item(i)                                                    'Map z_Data() to the hWnd thunk address
        If IsBadCodePtr(nAddr) = 0 Then                                     'Ensure that the thunk hasn't already freed itself
          zMap nAddr                                                        'Map the thunk memory to the z_Data() array
          sc_UnSubclass z_Data(IDX_HWND)                                    'UnSubclass
        End If
      Next i                                                                'Next member of the collection
    End With
    
    Set z_hWnds = Nothing                                                   'Destroy the window-handle/thunk-address collection
  End If
End Sub

'UnSubclass the specified window handle
Public Sub sc_UnSubclass(ByVal lng_hWnd As Long)
  If z_hWnds Is Nothing Then                                                'Ensure that subclassing has been started
    zError "UnSubclass", "Subclassing hasn't been started", False
  Else
    zDelMsg lng_hWnd, ALL_MESSAGES, IDX_BTABLE                              'Delete all before messages
    zDelMsg lng_hWnd, ALL_MESSAGES, IDX_ATABLE                              'Delete all after messages
    zMap_hWnd lng_hWnd                                                      'Map the thunk memory to the z_Data() array
    z_Data(IDX_SHUTDOWN) = -1                                               'Set the shutdown indicator
    zMap z_DataOrigData                                                     'Restore z_Data()'s original data pointer
    z_hWnds.Remove "h" & lng_hWnd                                           'Remove the specified window handle from the collection
  End If
End Sub

'Add the message value to the window handle's specified callback table
Private Sub sc_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If When And MSG_BEFORE Then                                               'If the message is to be added to the before original WndProc table...
    zAddMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Add the message to the before table
  End If

  If When And MSG_AFTER Then                                                'If message is to be added to the after original WndProc table...
    zAddMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Add the message to the after table
  End If

  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Sub

'Delete the message value from the window handle's specified callback table
Private Sub sc_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = eMsgWhen.MSG_AFTER)
  If When And MSG_BEFORE Then                                               'If the message is to be deleted from the before original WndProc table...
    zDelMsg lng_hWnd, uMsg, IDX_BTABLE                                      'Delete the message from the before table
  End If

  If When And MSG_AFTER Then                                                'If the message is to be deleted from the after original WndProc table...
    zDelMsg lng_hWnd, uMsg, IDX_ATABLE                                      'Delete the message from the after table
  End If

  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Sub

'Call the original WndProc
Private Function sc_CallOrigWndProc(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  sc_CallOrigWndProc = CallWindowProcA(z_Data(IDX_WNDPROC), lng_hWnd, uMsg, _
                                                            wParam, lParam) 'Call the original WndProc of the passed window handle parameter
  zMap z_DataOrigData                                                       'Restore z_Data()'s original data pointer
End Function

'Add the message to the specified table of the window handle
Private Sub zAddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim i      As Long                                                        'Loop index

  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  zMap z_Data(nTable)                                                       'Map z_Data() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being added to the table...
    nCount = ALL_MESSAGES                                                   'Set the table entry count to ALL_MESSAGES
  Else
    nCount = z_Data(0)                                                      'Get the current table entry count

    If nCount >= MSG_ENTRIES Then                                           'Check for message table overflow
      zError "zAddMsg", "Message table overflow. Either increase the value of Const MSG_ENTRIES or use ALL_MESSAGES instead of specific message values", False
      Exit Sub
    End If

    For i = 1 To nCount                                                     'Loop through the table entries
      If z_Data(i) = 0 Then                                                 'If the element is free...
        z_Data(i) = uMsg                                                    'Use this element
        Exit Sub                                                            'Bail
      ElseIf z_Data(i) = uMsg Then                                          'If the message is already in the table...
        Exit Sub                                                            'Bail
      End If
    Next i                                                                  'Next message table entry

    nCount = i                                                              'On drop through: i = nCount + 1, the new table entry count
    z_Data(nCount) = uMsg                                                   'Store the message in the appended table entry
  End If

  z_Data(0) = nCount                                                        'Store the new table entry count
End Sub

'Delete the message from the specified table of the window handle
Private Sub zDelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal nTable As Long)
  Dim nCount As Long                                                        'Table entry count
  Dim i      As Long                                                        'Loop index

  zMap_hWnd lng_hWnd                                                        'Map z_Data() to the thunk of the specified window handle
  zMap z_Data(nTable)                                                       'Map z_Data() to the table address

  If uMsg = ALL_MESSAGES Then                                               'If ALL_MESSAGES are being deleted from the table...
    z_Data(0) = 0                                                           'Zero the table entry count
  Else
    nCount = z_Data(0)                                                      'Get the table entry count
    
    For i = 1 To nCount                                                     'Loop through the table entries
      If z_Data(i) = uMsg Then                                              'If the message is found...
        z_Data(i) = 0                                                       'Null the msg value -- also frees the element for re-use
        Exit Sub                                                            'Exit
      End If
    Next i                                                                  'Next message table entry
    
    zError "zDelMsg", "Message &H" & Hex$(uMsg) & " not found in table", False
  End If
End Sub

'Error handler
Private Sub zError(ByVal sRoutine As String, ByVal sMsg As String, Optional ByVal bEnd As Boolean = True)
  App.LogEvent TypeName(Me) & "." & sRoutine & ": " & sMsg, vbLogEventTypeError
  
  MsgBox sMsg & ".", IIf(bEnd, vbCritical, vbExclamation) + vbApplicationModal, "Error in " & TypeName(Me) & "." & sRoutine
  
  If bEnd Then
    End
  End If
End Sub

'Return the address of the specified DLL/procedure
Private Function zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
  zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)                   'Get the specified procedure address
  Debug.Assert zFnAddr                                                      'In the IDE, validate that the procedure address was located
End Function

'Map z_Data() to the specified address
Private Sub zMap(ByVal nAddr As Long)
  RtlMoveMemory z_DataDataPtr, VarPtr(nAddr), 4                             'Set z_Data()'s SafeArray data pointer to the specified address
End Sub

'Map z_Data() to the thunk address for the specified window handle
Private Function zMap_hWnd(ByVal lng_hWnd As Long) As Long
  If z_hWnds Is Nothing Then                                                'Ensure that subclassing has been started
    zError "zMap_hWnd", "Subclassing hasn't been started", True
  Else
    On Error GoTo Catch                                                     'Catch unsubclassed window handles
    zMap_hWnd = z_hWnds("h" & lng_hWnd)                                     'Get the thunk address
    zMap zMap_hWnd                                                          'Map z_Data() to the thunk address
  End If
  
  Exit Function                                                             'Exit returning the thunk address

Catch:
  zError "zMap_hWnd", "Window handle isn't subclassed"
End Function

'Determine the address of the final private method, zWndProc
Private Function zGetCallback() As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte                                                         'Value pointed at by the vTable entry
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Upper bound of z_Data()
  Dim k     As Long                                                         'vTable entry value
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(Me), 4                                'Get the address of my vTable
  zMap nAddr + &H7A4                                                        'Map z_Data() to the first possible vTable entry for a UserControl

  j = UBound(z_Data())                                                      'Get the upper bound of z_Data()
  For i = 0 To j                                                            'Loop through the vTable looking for the first method entry
    k = z_Data(i)                                                           'Get the vTable entry
    
    If k <> 0 Then                                                          'Skip implemented interface entries
      RtlMoveMemory VarPtr(bVal), k, 1                                      'Get the first byte pointed to by this vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'If a method (pcode or native)
        bSub = bVal                                                         'Store which of the method markers was found (pcode or native)
        Exit For                                                            'Method found, quit loop and scan methods
      End If
    End If
  Next i
  
  For i = i To j                                                            'Loop through the remaining vTable entries
    k = z_Data(i)                                                           'Get the vTable entry
    
    If IsBadCodePtr(k) Then                                                 'Is the vTable entry an invalid code address...
      Exit For                                                              'Bad code pointer, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), k, 1                                        'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      Exit For                                                              'Bad method signature, quit loop
    End If
  Next i
  
  If i > j Then                                                             'Loop completed without finding the last method
    zError "zGetCallback", "z_Data() overflow. Increase the number of elements in the z_Data() array"
  End If
 
  zGetCallback = z_Data(i - 1)                                              'Return the last good vTable entry address
End Function

'-Subclass callback: must be private and the last method in the source file-----------------------
Private Sub zWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

   Select Case uMsg

      Case WM_MOUSEMOVE
'        detect when mouse has entered the control.
         If m_Enabled And Not bInCtrl Then
            bInCtrl = True
            Call TrackMouseLeave(lng_hWnd)
            RaiseEvent MouseEnter
'           repaint control based on selected mouseover actions.
            'm_MouseOverActions = Border
            ActiveBorderColor = m_MOverBorderColor
            TextColor1 = CaptionHover
            DrawText UserControl.hDC, m_Caption, -1, r, DT_LEFT Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP
            RedrawControl
         End If
       
'     detect when mouse has left the control.
      Case WM_MOUSELEAVE
         bInCtrl = False
         RaiseEvent MouseLeave
'        restore default control appearance if any mouseover actions were specified.
         
            ActiveBorderColor = SaveBorderColor
            TextColor1 = CheckColor
            DrawText UserControl.hDC, m_Caption, -1, r, DT_LEFT Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP
            RedrawControl
        
'     detect when control has gained the focus.
      Case WM_SETFOCUS
         If m_Enabled Then
            HasFocus = True
            RedrawControl
         End If

'     detect when control has lost the focus.
      Case WM_KILLFOCUS
         HasFocus = False
         KeyIsDown = False
         MouseIsDown = False
         RedrawControl
   End Select

End Sub
