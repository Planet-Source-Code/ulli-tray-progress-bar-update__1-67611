VERSION 5.00
Begin VB.Form fProgress 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FFFF&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   330
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   360
   ControlBox      =   0   'False
   DrawMode        =   10  'Stift maskieren
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   24
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "fProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const English = True

#If English Then '******************************************************

Private Const sTR       As String = "Estimated Time Remaining" & vbCrLf
Private Const Mask      As String = "h\ \h\r\s  m\ \m\i\n\s  s\ \s\e\c\s"

#Else 'german **********************************************************

Rem Mark off silent
Private Const sTR       As String = "Verbleibenden Zeit etwa" & vbCrLf
Private Const Mask      As String = "h\ \S\t\d  m\ \M\i\n  s\ \S\e\k"
Rem Mark on

#End If '***************************************************************

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private CPUFreq         As Currency
Private CPUCount        As Currency

Private Type POINT
    X                   As Long
    Y                   As Long
End Type

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Public BalloonText      As String
Public BarText          As String
Public ShowRemaining    As Boolean

Private WindowRect      As RECT
Private CursorPos       As POINT
Private Systray         As clsSystray
Private hWndTray        As Long
Private PrintY          As Long
Private FColor          As Long
Private BColor          As Long
Private PrevCheckAt     As Currency
Private PercentSoFar    As Double
Private PercentText     As String
Private PrevBalloonText As String

Public Property Let Progress(Percent As Double)

  Dim Elapsed   As Single

    Select Case Percent
      Case Is < 0 'done
        If hWndTray Then 'it is still showing
            hWndTray = 0
            BalloonText = vbNullString
            With Systray
                .HideBalloon
                .RemoveIconFromTray
            End With 'SYSTRAY
            Set Systray = Nothing
        End If
        Unload Me
      Case Is <= 100 'a progress % value
        If hWndTray = 0 Then
            PrevCheckAt = 0
            If BalloonText = vbNullString Then
                If ShowRemaining Then
                    BalloonText = sTR & " ..."
                  Else 'SHOWREMAINING = FALSE/0
                    BalloonText = "Progress"
                End If
              Else 'NOT BALLOONTEXT...
                BalloonText = Left$(BalloonText, 32)
            End If
            PrevBalloonText = BalloonText
            Set Systray = New clsSystray
            With Systray
                .SetOwner Me
                .AddIconToTray Icon.Handle, , True
                .ShowBalloon BalloonText, , SoundOff
            End With 'SYSTRAY
            hWndTray = FindWindow("Shell_TrayWnd", "") 'find tray
            GetWindowRect hWndTray, WindowRect
            With WindowRect
                Width = (.Right - .Left - 2) * 15 'adjust my size
                Height = (.Bottom - .Top - 2) * 15
            End With 'WINDOWRECT
            SetParent hWnd, hWndTray 'tray is my parent
            ScaleWidth = 1000 'percent * 10
            PrintY = (ScaleHeight - TextHeight("A")) / 2 'vertical print pos
            FColor = ForeColor 'colors...
            If FColor < 0 Then
                FColor = GetSysColor(FColor And &H7FFFFFFF)
            End If
            BColor = BackColor
            If BColor < 0 Then
                BColor = GetSysColor(BColor And &H7FFFFFFF)
            End If
            FColor = Not (FColor Xor BColor)
            Show
            QueryPerformanceFrequency CPUFreq
            QueryPerformanceCounter CPUCount
            PrevCheckAt = CPUCount
            PercentSoFar = Percent
        End If
        If ShowRemaining And Percent - PercentSoFar > 0 Then
            QueryPerformanceCounter CPUCount
            Elapsed = (CPUCount - PrevCheckAt) / CPUFreq
            If Elapsed >= 1 Then
                BalloonText = sTR & Format$(((100 - Percent) * Elapsed / (Percent - PercentSoFar)) / 86400, Mask)
                PrevCheckAt = CPUCount
                PercentSoFar = Percent
            End If
        End If
        If PrevBalloonText <> BalloonText Then
            PrevBalloonText = BalloonText
            Systray.ShowBalloon BalloonText, , SoundOff
        End If
        PercentText = Int(Percent) & "%"
        Cls
        CurrentY = PrintY
        CurrentX = 5
        Print Left$(BarText, 32);
        CurrentX = 500 - TextWidth(PercentText) / 2
        Print PercentText
        Line (0, 3)-(Percent * 10, ScaleHeight - 4), FColor, BF
    End Select

End Property

':) Ulli's VB Code Formatter V2.21.8 (2007-Jan-19 14:14)  Decl: 54  Code: 87  Total: 141 Lines
':) CommentOnly: 2 (1,4%)  Commented: 17 (12,1%)  Empty: 15 (10,6%)  Max Logic Depth: 5
