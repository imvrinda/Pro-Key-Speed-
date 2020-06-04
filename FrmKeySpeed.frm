VERSION 5.00
Begin VB.Form frmKeySpeed 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form2"
   ClientHeight    =   11115
   ClientLeft      =   765
   ClientTop       =   870
   ClientWidth     =   15240
   Icon            =   "FrmKeySpeed.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Timer time_sec 
      Interval        =   1000
      Left            =   4560
      Top             =   6600
   End
   Begin VB.Timer Timer_F 
      Interval        =   1000
      Left            =   3840
      Top             =   6600
   End
   Begin VB.Timer Timer_E 
      Interval        =   1000
      Left            =   3120
      Top             =   6600
   End
   Begin VB.Timer Timer_D 
      Interval        =   1000
      Left            =   2400
      Top             =   6600
   End
   Begin VB.Timer Timer_C 
      Interval        =   1000
      Left            =   1680
      Top             =   6600
   End
   Begin VB.Timer Timer_B 
      Interval        =   1000
      Left            =   960
      Top             =   6600
   End
   Begin VB.Timer Timer_A 
      Interval        =   1000
      Left            =   240
      Top             =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      X1              =   12600
      X2              =   12600
      Y1              =   240
      Y2              =   11160
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   12840
      TabIndex        =   11
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Of Drop Counts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   12720
      TabIndex        =   10
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   12960
      TabIndex        =   9
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Of Counts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   12960
      TabIndex        =   8
      Top             =   360
      Width           =   1890
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   13200
      TabIndex        =   7
      Top             =   3840
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   12840
      TabIndex        =   6
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label_F 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   54
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1500
      Left            =   11160
      TabIndex        =   5
      Top             =   5400
      Width           =   660
   End
   Begin VB.Label Label_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   54
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1500
      Left            =   9840
      TabIndex        =   4
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label Label_D 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   54
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1380
      Left            =   7200
      TabIndex        =   3
      Top             =   4320
      Width           =   780
   End
   Begin VB.Label Label_C 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   54
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1500
      Left            =   5760
      TabIndex        =   2
      Top             =   360
      Width           =   675
   End
   Begin VB.Label Label_B 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   54
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1500
      Left            =   3600
      TabIndex        =   1
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label Label_A 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1350
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   720
   End
End
Attribute VB_Name = "frmKeySpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer
Dim cnt As Integer
Dim sec As Integer
Dim cnt1 As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)

Dim ctl As Control

  For Each ctl In Me.Controls
  
    If TypeOf ctl Is VB.Label Then
     
      If ctl.Caption = Chr(KeyAscii) Then
       cnt = cnt + 1
       Label4.Caption = cnt
       
        ctl.Top = Me.Height
        If KeyAscii < 90 Then
            ctl.Caption = Chr(KeyAscii + 1)
        Else
            KeyAscii = 65
            
             ctl.Caption = Chr(65)
        End If
        
        Exit Sub
        
       End If
    End If
    
  Next
End Sub

Private Sub Form_Load()
Timer_A.Interval = myInt
Timer_B.Interval = myInt
Timer_C.Interval = myInt
Timer_D.Interval = myInt
Timer_E.Interval = myInt
Timer_F.Interval = myInt

End Sub


Private Sub time_sec_Timer()
 sec = sec + 1
    Label2.Caption = sec
End Sub

Private Sub Timer_A_Timer()
    If Label_A.Top > Me.Height Then
        Label_A.Top = 0 - Label_A.Height
        cnt1 = cnt1 + 1
        Label6.Caption = cnt1
    Else
        Label_A.Top = Label_A.Top + 600
    End If
    If Label_A.ForeColor = vbRed Then
        Label_A.ForeColor = vbBlue
    Else
        Label_A.ForeColor = vbRed
    End If
    
End Sub

Private Sub Timer_B_Timer()
If Label_B.Top > Me.Height Then
Label_B.Top = 0 - Label_B.Height
cnt1 = cnt1 + 1
        Label6.Caption = cnt1
Else
Label_B.Top = Label_B.Top + 600
End If
If Label_B.ForeColor = vbRed Then
Label_B.ForeColor = vbBlue
Else
Label_B.ForeColor = vbRed
End If

End Sub


Private Sub Timer_C_Timer()
If Label_C.Top > Me.Height Then
Label_C.Top = 0 - Label_C.Height
cnt1 = cnt1 + 1
        Label6.Caption = cnt1
Else
Label_C.Top = Label_C.Top + 600
End If
If Label_C.ForeColor = vbRed Then
Label_C.ForeColor = vbBlue
Else
Label_C.ForeColor = vbRed
End If

End Sub

Private Sub Timer_D_Timer()
If Label_D.Top > Me.Height Then
Label_D.Top = 0 - Label_D.Height
cnt1 = cnt1 + 1
        Label6.Caption = cnt1
Else
Label_D.Top = Label_D.Top + 600
End If
If Label_D.ForeColor = vbRed Then
Label_D.ForeColor = vbBlue
Else
Label_D.ForeColor = vbRed
End If

End Sub

Private Sub Timer_E_Timer()
    If Label_E.Top > Me.Height Then
        Label_E.Top = 0 - Label_E.Height
        cnt1 = cnt1 + 1
        Label6.Caption = cnt1
    Else
    Label_E.Top = Label_E.Top + 600
    End If
    If Label_E.ForeColor = vbRed Then
    Label_E.ForeColor = vbBlue
    Else
    Label_E.ForeColor = vbRed
    End If
   
End Sub

Private Sub Timer_F_Timer()
If Label_F.Top > Me.Height Then
Label_F.Top = 0 - Label_F.Height
cnt1 = cnt1 + 1
        Label6.Caption = cnt1

Else
Label_F.Top = Label_F.Top + 500
End If
If Label_F.ForeColor = vbRed Then
Label_F.ForeColor = vbBlue
Else
Label_F.ForeColor = vbRed
End If

End Sub
