VERSION 5.00
Begin VB.Form frmSetting 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   4605
   ClientTop       =   2265
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   7245
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2040
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   3
      Top             =   4200
      Width           =   1695
   End
   Begin VB.OptionButton optCap 
      Caption         =   "Capital"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2040
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.OptionButton optMix 
      Caption         =   "Mixed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.OptionButton optSmall 
      Caption         =   "Small"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGo_Click()
If Combo1.Text = "Level1" Then
myInt = 800
End If
If Combo1.Text = "Level2" Then
myInt = 600
End If
If Combo1.Text = "Level3" Then
myInt = 400
End If
If Combo1.Text = "Level4" Then
myInt = 300
End If
If Combo1.Text = "Level5" Then
myInt = 200
End If
frmKeySpeed.Show
Unload Me

End Sub

Private Sub Form_Load()
Combo1.AddItem ("Level1")
Combo1.AddItem ("Level2")
Combo1.AddItem ("Level3")
Combo1.AddItem ("Level4")
Combo1.AddItem ("Level5")


Combo1.Text = "Level1"
End Sub
