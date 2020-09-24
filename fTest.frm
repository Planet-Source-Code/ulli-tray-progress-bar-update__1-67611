VERSION 5.00
Begin VB.Form fTest 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Test"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btDone 
      Appearance      =   0  '2D
      Cancel          =   -1  'True
      Caption         =   "Done"
      Height          =   495
      Left            =   1515
      TabIndex        =   1
      Top             =   510
      Width           =   1080
   End
   Begin VB.CommandButton btTest 
      Appearance      =   0  '2D
      Caption         =   "Test it"
      Default         =   -1  'True
      Height          =   495
      Left            =   390
      TabIndex        =   0
      Top             =   510
      Width           =   1080
   End
   Begin VB.Timer Ticker 
      Enabled         =   0   'False
      Left            =   45
      Top             =   1185
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Percent As Double

Private Sub btDone_Click()

    Unload Me

End Sub

Private Sub btTest_Click()

    btTest.Enabled = False

    With fProgress

        'optionally set colors - default is red on yellow
        .BackColor = vbInfoBackground
        .ForeColor = vbInfoText

        'optionally set balloon text
        .BalloonText = "This may take a while"

        'alternatively or additionally to balloon text
        .ShowRemaining = True

        'optionally set bar text - this may be modified while it is running
        .BarText = "Testing progress bar..."

    End With 'FPROGRESS

    Percent = 0 'reset percent

    With Ticker 'prepare ticker
        .Interval = 20
        .Enabled = True
    End With 'TICKER

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If btTest.Enabled = False Then 'it's still active so we have to stop it
        fProgress.Progress = -1
    End If

End Sub

Private Sub Ticker_Timer()

    With fProgress
        If Percent > 100 Then
            Ticker.Enabled = False 'switch off ticker

            .Progress = -1 'hide progress bar (by any negative value)

            btTest.Enabled = True

          Else 'NOT PERCENT...

            .Progress = Percent '.Progress is a Property_Let of fProgress

            'simulate variable progress speed for testing time to go estimation
            Select Case Percent

              Case Is < 30
                Percent = Percent + 0.14

              Case Is < 31
                Percent = Percent + 0.001
                .BarText = "Dead slow now..."

              Case Is < 50
                Percent = Percent + 0.02
                .BarText = "Still rather slow..."

              Case Else
                Percent = Percent + 0.06
                .BarText = "Faster again..."
            End Select

        End If
    End With 'FPROGRESS

End Sub

':) Ulli's VB Code Formatter V2.21.8 (2007-Jan-19 14:14)  Decl: 3  Code: 85  Total: 88 Lines
':) CommentOnly: 5 (5,7%)  Commented: 10 (11,4%)  Empty: 31 (35,2%)  Max Logic Depth: 4
