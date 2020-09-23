VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1943
      TabIndex        =   30
      Top             =   3360
      Width           =   795
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3773
      TabIndex        =   29
      Top             =   2880
      Width           =   555
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   353
      TabIndex        =   28
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blah pgyblah blah"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   773
      TabIndex        =   27
      Top             =   2880
      Width           =   2715
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   3773
      TabIndex        =   26
      Top             =   2580
      Width           =   555
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   353
      TabIndex        =   25
      Top             =   2580
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blah pgyblah blah"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   773
      TabIndex        =   24
      Top             =   2580
      Width           =   2715
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   3773
      TabIndex        =   23
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   353
      TabIndex        =   22
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blah pgyblah blah"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   773
      TabIndex        =   21
      Top             =   2280
      Width           =   2715
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3773
      TabIndex        =   20
      Top             =   1980
      Width           =   555
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   353
      TabIndex        =   19
      Top             =   1980
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blah pgyblah blah"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   773
      TabIndex        =   18
      Top             =   1980
      Width           =   2715
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3773
      TabIndex        =   17
      Top             =   1680
      Width           =   555
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   353
      TabIndex        =   16
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blah pgyblah blah"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   773
      TabIndex        =   15
      Top             =   1680
      Width           =   2715
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3773
      TabIndex        =   14
      Top             =   1380
      Width           =   555
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   353
      TabIndex        =   13
      Top             =   1380
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blah pgyblah blah"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   773
      TabIndex        =   12
      Top             =   1380
      Width           =   2715
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3773
      TabIndex        =   11
      Top             =   1080
      Width           =   555
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   353
      TabIndex        =   10
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blah pgyblah blah"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   773
      TabIndex        =   9
      Top             =   1080
      Width           =   2715
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3773
      TabIndex        =   8
      Top             =   780
      Width           =   555
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   353
      TabIndex        =   7
      Top             =   780
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blah pgyblah blah"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   773
      TabIndex        =   6
      Top             =   780
      Width           =   2715
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3773
      TabIndex        =   5
      Top             =   480
      Width           =   555
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   353
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blah pgyblah blah"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   773
      TabIndex        =   3
      Top             =   480
      Width           =   2715
   End
   Begin VB.Label lblScore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3773
      TabIndex        =   2
      Top             =   180
      Width           =   555
   End
   Begin VB.Label lblPos 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "10."
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   353
      TabIndex        =   1
      Top             =   180
      Width           =   375
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Blah pgyblah blah"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   773
      TabIndex        =   0
      Top             =   180
      Width           =   2715
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = App.Title & " - Hall of fame!"
    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
Dim x As Integer
    For x = 1 To 10
        lblPos(x - 1) = x & "."
        lblName(x - 1) = GetSetting(App.Title, "HiScores", "Name" & Format(x, "00"))
        lblScore(x - 1) = CStr(GetSetting(App.Title, "HiScores", "Score" & Format(x, "00")))
    Next x
End Sub

