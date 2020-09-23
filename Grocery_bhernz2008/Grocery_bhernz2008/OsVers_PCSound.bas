Attribute VB_Name = "OsVers_PCSound"
Option Explicit

Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Public Enum OsType
    Nt2000
    Win9xMe
    OsUnknown
End Enum

'following lines are for Win9xMe platforms
'For these systems, the file WIN95IO.DLL must be copied
'to the Windows/System folder.
'WIN95IO.DLL is available from http://www.softcircuits.com
Declare Sub vbOut Lib "WIN95IO.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
Declare Function vbInp Lib "WIN95IO.DLL" (ByVal nPort As Integer) As Integer

'This line is for NT2000 platforms
Public Declare Function NtBeep Lib "kernel32" Alias "Beep" (ByVal FreqHz As Long, ByVal DurationMs As Long) As Long

Public Function GetVersion() As String
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer

   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)

   With osinfo
   Select Case .dwPlatformId
      Case 1
         If .dwMinorVersion = 0 Then
            GetVersion = "Windows 95"
         ElseIf .dwMinorVersion = 10 Then
            GetVersion = "Windows 98"
         ElseIf .dwMinorVersion = 90 Then
            GetVersion = "Windows Me"
         End If
      Case 2
         If .dwMajorVersion = 3 Then
            GetVersion = "Windows NT 3.51"
         ElseIf .dwMajorVersion = 4 Then
            GetVersion = "Windows NT 4.0"
         ElseIf .dwMajorVersion = 5 Then
            GetVersion = "Windows 2000"
         End If
      Case Else
         GetVersion = "Failed"
   End Select
   End With
End Function

'for speaker beep function, only the platform type is relevant
Public Function GetPlatform() As OsType
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer

   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)
   
   Select Case osinfo.dwPlatformId
      Case 1
        GetPlatform = Win9xMe
      Case 2
        GetPlatform = Nt2000
      Case Else
        GetPlatform = OsUnknown
   End Select
   
End Function

'This is where the beep method depends on the operating system
Public Sub PcSpeakerBeep(ByVal FreqHz As Integer, ByVal LengthMs As Single)
    
    Select Case GetPlatform
        Case Win9xMe
            Call Win9xBeep(FreqHz, LengthMs)
        Case Nt2000
            Call NtBeep(CLng(FreqHz), CLng(LengthMs))
        Case OsUnknown
            Beep    'use the default beep routine, probably the sound card
    End Select
            
End Sub

'following routine largely by Jorge Loubet
Private Sub Win9xBeep(ByVal Freq As Integer, ByVal Length As Single)

    Dim LoByte As Integer
    Dim HiByte As Integer
    Dim Clicks As Integer
    Dim SpkrOn As Integer
    Dim SpkrOff As Integer
    Dim TimeEnd As Single
    
    TimeEnd = Timer + Length / 1000
    
    'Ports 66, 67, and 97 control timer and speaker
    '
    'Divide clock frequency by sound frequency
    'to get number of "clicks" clock must produce.
        Clicks = CInt(1193280 / Freq)
        LoByte = Clicks And &HFF
        HiByte = Clicks \ 256
    'Tell timer that data is coming
        vbOut 67, 182
    'Send count to timer
        vbOut 66, LoByte
        vbOut 66, HiByte
    'Turn speaker on by setting bits 0 and 1 of PPI chip.
        SpkrOn = vbInp(97) Or &H3
        vbOut 97, SpkrOn
    
    'Leave speaker on (while timer runs)
        Do While Timer < TimeEnd
            'Let processor do other tasks
            DoEvents
        Loop
    'Turn speaker off.
        SpkrOff = vbInp(97) And &HFC
        vbOut 97, SpkrOff
End Sub

Public Sub Warble(ByVal FreqHz As Integer, ByVal DurationMs As Single)
    Dim EndTime As Single
    EndTime = Timer + DurationMs / 1000
    
    If FreqHz < 100 Then FreqHz = 100
    Do While EndTime > Timer
        Call PcSpeakerBeep(FreqHz, 10)
        Call PcSpeakerBeep(FreqHz / 1.1, 10)
        Call PcSpeakerBeep(FreqHz / 1.2, 80)
        Call PcSpeakerBeep(FreqHz / 1.3, 10)
        Call PcSpeakerBeep(FreqHz / 1.2, 80)
        Call PcSpeakerBeep(FreqHz / 1.1, 10)
    Loop
End Sub

