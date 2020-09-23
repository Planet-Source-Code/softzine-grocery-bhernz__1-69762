Attribute VB_Name = "Module1"
'Author£ºWXJ_Lake
'Email: webmaster@archtide.com
'Homepage£ºwww.archtide.com
Option Explicit

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Const RGN_OR = 2
Public welcome As Boolean
Public welcometime As Integer

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type RECT
      Left As Long
      Top As Long
      Right As Long
      Bottom As Long
   End Type

   Private Type CharRange
     cpMin As Long     ' First character of range (0 for start of doc)
     cpMax As Long     ' Last character of range (-1 for end of doc)
   End Type

   Private Type FormatRange
     hDC As Long       ' Actual DC to draw on
     hdcTarget As Long ' Target DC for determining text formatting
     rc As RECT        ' Region of the DC to draw to (in twips)
     rcPage As RECT    ' Region of the entire DC (page size) (in twips)
     chrg As CharRange ' Range of text to draw (see above declaration)
   End Type

   Public Const WM_USER As Long = &H400
   Private Const EM_FORMATRANGE As Long = WM_USER + 57
   Public Const EM_SETTARGETDEVICE As Long = WM_USER + 72
   Private Const PHYSICALOFFSETX As Long = 112
   Private Const PHYSICALOFFSETY As Long = 113
   
   Private Declare Function GetDeviceCaps Lib "gdi32" ( _
      ByVal hDC As Long, ByVal nIndex As Long) As Long
   Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" _
      (ByVal lpDriverName As String, ByVal lpDeviceName As String, _
      ByVal lpOutput As Long, ByVal lpInitData As Long) As Long
      
'declare for moving the form
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

'for translucent effect in win2k, remove this if run in win9x or NT
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1

Public Sub SetAutoRgn(hForm As Form, Optional transColor As Long = vbNull)
  Dim X As Long, Y As Long
  Dim Rgn1 As Long, Rgn2 As Long
  Dim SPos As Long, EPos As Long
  Dim wID As Long, Hgt As Long
  Dim xoff As Long, yoff As Long
  Dim DIB As New cDIBSection
  Dim bDib() As Byte
  Dim tSA As SAFEARRAY2D
  
  
    'get the picture size of the form
  DIB.CreateFromPicture hForm.Picture
  wID = DIB.Width
  Hgt = DIB.Height
  
  With hForm
    .ScaleMode = vbPixels
    'compute the title bar's offset
    xoff = (.ScaleX(.Width, vbTwips, vbPixels) - .ScaleWidth) / 2
    yoff = .ScaleY(.Height, vbTwips, vbPixels) - .ScaleHeight - xoff
    'change the form size
    .Width = (wID + xoff * 2) * Screen.TwipsPerPixelX
    .Height = (Hgt + xoff + yoff) * Screen.TwipsPerPixelY
  End With
  
  ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = DIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = DIB.BytesPerScanLine
        .pvData = DIB.DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4
        
      
' if there is no transColor specified, use the first pixel as the transparent color
  If transColor = vbNull Then transColor = RGB(bDib(0, 0), bDib(1, 0), bDib(2, 0))
  
  Rgn1 = CreateRectRgn(0, 0, 0, 0)
  
  For Y = 0 To Hgt - 1 'line scan
    X = -3
    Do
     X = X + 3
     
     While RGB(bDib(X, Y), bDib(X + 1, Y), bDib(X + 2, Y)) = transColor And (X < wID * 3 - 3)
       X = X + 3 'skip the transparent point
     Wend
     SPos = X / 3
     While RGB(bDib(X, Y), bDib(X + 1, Y), bDib(X + 2, Y)) <> transColor And (X < wID * 3 - 3)
       X = X + 3 'skip the nontransparent point
     Wend
     EPos = X / 3
     
     'combine the region
     If SPos <= EPos Then
         Rgn2 = CreateRectRgn(SPos + xoff, Hgt - Y + yoff, EPos + xoff, Hgt - 1 - Y + yoff)
         CombineRgn Rgn1, Rgn1, Rgn2, RGN_OR
         DeleteObject Rgn2
     End If
    Loop Until X >= wID * 3 - 3
  Next Y
  
  SetWindowRgn hForm.hWnd, Rgn1, True  'set the final shap region
  DeleteObject Rgn1
 
End Sub

'miscellaneous stuff  =================================
Public Function FileExists(strPath As String) As Boolean
   FileExists = LenB(Dir$(strPath))
End Function

