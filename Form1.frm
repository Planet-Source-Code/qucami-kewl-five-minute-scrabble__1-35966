VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4740
      Left            =   7020
      Sorted          =   -1  'True
      TabIndex        =   94
      Top             =   240
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hi Scores"
      Height          =   375
      Left            =   1380
      TabIndex        =   93
      Top             =   4740
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "New Game"
      Height          =   375
      Left            =   180
      TabIndex        =   92
      Top             =   4740
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "I can't do any more!"
      Height          =   375
      Left            =   3120
      TabIndex        =   91
      Top             =   4740
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check my words"
      Height          =   375
      Left            =   4800
      TabIndex        =   89
      Top             =   4740
      Width           =   1395
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   4140
      TabIndex        =   7
      Text            =   "WWWWWWW"
      Top             =   3180
      Width           =   1335
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   4140
      TabIndex        =   8
      Text            =   "WWWWWWW"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4140
      TabIndex        =   6
      Text            =   "WWWWWWW"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4140
      TabIndex        =   5
      Text            =   "WWWWWWW"
      Top             =   1620
      Width           =   1335
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   4140
      TabIndex        =   4
      Text            =   "WWWWWWW"
      Top             =   840
      Width           =   1335
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   5595
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10742
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTotalScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1395
      TabIndex        =   90
      Top             =   5280
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Score"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   5640
      TabIndex        =   88
      Top             =   600
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Your word"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   4140
      TabIndex        =   87
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   4
      Left            =   5640
      TabIndex        =   86
      Top             =   3960
      Width           =   555
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   85
      Top             =   3180
      Width           =   555
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   84
      Top             =   2400
      Width           =   555
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   83
      Top             =   1620
      Width           =   555
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "9999"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   82
      Top             =   840
      Width           =   555
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   3480
      TabIndex        =   81
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   34
      Left            =   3720
      TabIndex        =   80
      Top             =   4080
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   2940
      TabIndex        =   79
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   33
      Left            =   3180
      TabIndex        =   78
      Top             =   4080
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   2400
      TabIndex        =   77
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   32
      Left            =   2640
      TabIndex        =   76
      Top             =   4080
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   1860
      TabIndex        =   75
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   2100
      TabIndex        =   74
      Top             =   4080
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   1320
      TabIndex        =   73
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   1560
      TabIndex        =   72
      Top             =   4080
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   780
      TabIndex        =   71
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   1020
      TabIndex        =   70
      Top             =   4080
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   240
      TabIndex        =   69
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   480
      TabIndex        =   68
      Top             =   4080
      Width           =   195
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   3720
      TabIndex        =   66
      Top             =   3300
      Width           =   195
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   3180
      TabIndex        =   64
      Top             =   3300
      Width           =   195
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   2640
      TabIndex        =   62
      Top             =   3300
      Width           =   195
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   2100
      TabIndex        =   60
      Top             =   3300
      Width           =   195
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   1560
      TabIndex        =   58
      Top             =   3300
      Width           =   195
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   1020
      TabIndex        =   56
      Top             =   3300
      Width           =   195
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   480
      TabIndex        =   54
      Top             =   3300
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   3480
      TabIndex        =   53
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   3720
      TabIndex        =   52
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   2940
      TabIndex        =   51
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   3180
      TabIndex        =   50
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   2400
      TabIndex        =   49
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   2640
      TabIndex        =   48
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   1860
      TabIndex        =   47
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   2100
      TabIndex        =   46
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   1320
      TabIndex        =   45
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   1560
      TabIndex        =   44
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   780
      TabIndex        =   43
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   1020
      TabIndex        =   42
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   41
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   480
      TabIndex        =   40
      Top             =   2520
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   3480
      TabIndex        =   39
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   3720
      TabIndex        =   38
      Top             =   1740
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   2940
      TabIndex        =   37
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   3180
      TabIndex        =   36
      Top             =   1740
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   2400
      TabIndex        =   35
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   2640
      TabIndex        =   34
      Top             =   1740
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   1860
      TabIndex        =   33
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   2100
      TabIndex        =   32
      Top             =   1740
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1320
      TabIndex        =   31
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   1560
      TabIndex        =   30
      Top             =   1740
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   780
      TabIndex        =   29
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   1020
      TabIndex        =   28
      Top             =   1740
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   27
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   26
      Top             =   1740
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   25
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   3720
      TabIndex        =   24
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2940
      TabIndex        =   23
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   3180
      TabIndex        =   22
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   21
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   2640
      TabIndex        =   20
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1860
      TabIndex        =   19
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   2100
      TabIndex        =   18
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   17
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1560
      TabIndex        =   16
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   780
      TabIndex        =   15
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1020
      TabIndex        =   14
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Triple word score"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   4
      Left            =   1080
      TabIndex        =   13
      Top             =   4380
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Triple word score"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   1080
      TabIndex        =   12
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Triple word score"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   11
      Top             =   2820
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Triple word score"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   10
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Triple word score"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   1260
      Width           =   2775
   End
   Begin VB.Label lblLetterScore 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Width           =   195
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"Form1.frx":0E42
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   135
      TabIndex        =   1
      Top             =   60
      Width           =   6135
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   6
      Left            =   3420
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   5
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   4
      Left            =   2340
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   3
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   2
      Left            =   1260
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   0
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   1
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   13
      Left            =   3420
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   12
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   11
      Left            =   2340
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   10
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   9
      Left            =   1260
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   8
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   7
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   20
      Left            =   3420
      Shape           =   4  'Rounded Rectangle
      Top             =   2340
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   19
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   2340
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   18
      Left            =   2340
      Shape           =   4  'Rounded Rectangle
      Top             =   2340
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   17
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   2340
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   16
      Left            =   1260
      Shape           =   4  'Rounded Rectangle
      Top             =   2340
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   15
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   2340
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   14
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   2340
      Width           =   495
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   3480
      TabIndex        =   67
      Top             =   3180
      Width           =   255
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   2940
      TabIndex        =   65
      Top             =   3180
      Width           =   255
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   2400
      TabIndex        =   63
      Top             =   3180
      Width           =   255
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   1860
      TabIndex        =   61
      Top             =   3180
      Width           =   255
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   1320
      TabIndex        =   59
      Top             =   3180
      Width           =   255
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   780
      TabIndex        =   57
      Top             =   3180
      Width           =   255
   End
   Begin VB.Label lblLetter 
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "BMW Helvetica Light"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   240
      TabIndex        =   55
      Top             =   3180
      Width           =   255
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   27
      Left            =   3420
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   26
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   25
      Left            =   2340
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   24
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   23
      Left            =   1260
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   22
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   21
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   34
      Left            =   3420
      Shape           =   4  'Rounded Rectangle
      Top             =   3900
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   33
      Left            =   2880
      Shape           =   4  'Rounded Rectangle
      Top             =   3900
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   32
      Left            =   2340
      Shape           =   4  'Rounded Rectangle
      Top             =   3900
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   31
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   3900
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   30
      Left            =   1260
      Shape           =   4  'Rounded Rectangle
      Top             =   3900
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   29
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   3900
      Width           =   495
   End
   Begin VB.Shape shp 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   435
      Index           =   28
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   3900
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sWords() As String
Dim iTileScores(27) As Integer
Dim sScoreMultiplier(5) As String
Dim sCurrentWords(5) As String
Function IsThereAHiScore(lIn As String) As Boolean
    If lIn > GetSetting(App.Title, "HiScores", "Score10") Then
        IsThereAHiScore = True
    Else
        IsThereAHiScore = False
    End If
End Function
Private Sub Command1_Click()
Dim lScore As Long
Dim lTotalScore As Long
Dim X As Integer
Dim y As Long
    SB.Panels(1).Text = ""
    For X = 0 To 4
        txtAnswer(X) = UCase(txtAnswer(X))
        If txtAnswer(X) > "" Then
            If Not IsWordInScrabbleWord(txtAnswer(X), X) Then
                SB.Panels(1).Text = "'" & txtAnswer(X) & "' cannot be formed from '" & sCurrentWords(X) & "'"
                Exit Sub
            End If
            If InStr(txtAnswer(X), " ") > 0 Then
                SB.Panels(1).Text = "'" & txtAnswer(X) & "' contains a space"
                Exit Sub
            End If
            If Not IsWordInDictionary(txtAnswer(X)) Then
                SB.Panels(1).Text = "'" & txtAnswer(X) & "' is not in my dictionary"
                Exit Sub
            End If
        End If
    Next X

    lScore = 0
    lTotalScore = 0
    For X = 0 To 4
        lScore = 0
        If txtAnswer(X) > "" Then
            For y = 1 To Len(txtAnswer(X))
                If Mid(sScoreMultiplier(X), 2, 1) = "L" And Mid(sScoreMultiplier(X), 3, 1) = CStr(y) Then
                    If Left(sScoreMultiplier(X), 1) = "D" Then
                        lScore = lScore + (iTileScores(Asc(Mid(txtAnswer(X), y, 1)) - 65) * 2)
                    ElseIf Left(sScoreMultiplier(X), 1) = "T" Then
                        lScore = lScore + (iTileScores(Asc(Mid(txtAnswer(X), y, 1)) - 65) * 3)
                    Else
                        lScore = lScore + iTileScores(Asc(Mid(txtAnswer(X), y, 1)) - 65)
                    End If
                Else
                    lScore = lScore + iTileScores(Asc(Mid(txtAnswer(X), y, 1)) - 65)
                End If
            Next y
            If Mid(sScoreMultiplier(X), 2, 1) = "W" Then
                If Left(sScoreMultiplier(X), 1) = "D" Then
                    lblScore(X) = CStr(lScore * 2)
                ElseIf Left(sScoreMultiplier(X), 1) = "T" Then
                    lblScore(X) = CStr(lScore * 3)
                Else
                    lblScore(X) = CStr(lScore)
                End If
            Else
                lblScore(X) = CStr(lScore)
            End If
            If Len(txtAnswer(X)) = 7 Then
                lScore = lScore + 50
                lblScore(X) = lScore
            End If
            lTotalScore = lTotalScore + lScore
        Else
            lblScore(X) = "0"
        End If
    Next X
    lblTotalScore = "Final Score: " & CStr(lTotalScore)
End Sub
Function IsWordInScrabbleWord(sIn As String, iLine As Integer) As Boolean
Dim X As Integer
Dim p As Integer
Dim sT As String
    IsWordInScrabbleWord = True
    sT = UCase(sCurrentWords(iLine))
    For X = 1 To Len(sIn)
        p = InStr(sT, Mid(sIn, X, 1))
        If p = 0 Then
            IsWordInScrabbleWord = False
            Exit Function
        Else
            Mid(sT, p, 1) = " "
        End If
    Next X
    
End Function
Function IsWordInDictionary(sIn As String) As Boolean
IsWordInDictionary = False
Dim X As Long
    For X = 0 To UBound(sWords)
        If sIn = UCase(sWords(X)) Then
            IsWordInDictionary = True
            Exit Function
        End If
    Next X
End Function

Private Sub Command2_Click()
Dim sName As String
    If IsThereAHiScore(Mid(lblTotalScore, InStr(lblTotalScore, ":") + 2)) Then
        sName = InputBox("Congratulations, you made the hall of fame", App.Title, Environ("username"))
        If Trim(sName) = "" Then
            sName = Environ("username")
        End If
        SaveScore sName
    End If
    Command3_Click
End Sub
Sub SaveScore(sIn As String)
Dim X As Integer
    List1.Clear
    For X = 1 To 10
        List1.AddItem Format(GetSetting(App.Title, "HiScores", "Score" & Format(X, "00")), "0000") & GetSetting(App.Title, "HiScores", "Name" & Format(X, "00"))
    Next X
    List1.AddItem Format(Mid(lblTotalScore, InStr(lblTotalScore, ":") + 2), "0000") & sIn
    For X = List1.ListCount - 1 To List1.ListCount - 10 Step -1
        SaveSetting App.Title, "HiScores", "Name" & Format(11 - X, "00"), Mid(List1.List(X), 5)
        SaveSetting App.Title, "HiScores", "Score" & Format(11 - X, "00"), CInt(Left(List1.List(X), 4))
    Next X
    
    Form2.Show 1
    
End Sub




Private Sub Command3_Click()
    ResetTable
    SetTileScores
    SetScoreMultiplier
    SelectWords
    displaywordsandscores
    SetScoreMultiplier
    DisplayMultiplier
End Sub

Private Sub Command4_Click()
    Form2.Show 1
    
End Sub

Private Sub Form_Load()
    Randomize Timer ^ Format(Now, "SS")
    LoadDict
    SB.Panels(1).Text = "Dictionary loaded. " & UBound(sWords) & " words in memory."

    Me.Left = (Screen.Width / 2) - (Me.Width / 2)
    Me.Top = (Screen.Height / 2) - (Me.Height / 2)
    SortTheHighScores
    Me.Show
    DoEvents
    Command3_Click

End Sub
Sub SortTheHighScores()
    If GetSetting(App.Title, "HiScores", "Name01") > "" Then
        Exit Sub
    End If
Dim X As Integer
    For X = 1 To 10
        SaveSetting App.Title, "HiScores", "Name" & Format(X, "00"), "Guest"
        SaveSetting App.Title, "HiScores", "Score" & Format(X, "00"), 30
    Next X
End Sub
Sub DisplayMultiplier()
Dim X As Integer
    For X = 0 To 4
        If Left(sScoreMultiplier(X), 1) = "D" Then
            lblInfo(X) = "Double "
        ElseIf Left(sScoreMultiplier(X), 1) = "T" Then
            lblInfo(X) = "Tripple "
        End If
        If Mid(sScoreMultiplier(X), 2, 1) = "L" Then
            lblInfo(X) = lblInfo(X) & "letter score on the " & WordPlace(Mid(sScoreMultiplier(X), 3, 1)) & " letter"
        ElseIf Mid(sScoreMultiplier(X), 2, 1) = "W" Then
            lblInfo(X) = lblInfo(X) & "word score"
        End If
    Next X
End Sub
Function WordPlace(sIn As String) As String
    Select Case sIn
        Case "1"
            WordPlace = "first"
        Case "2"
            WordPlace = "second"
        Case "3"
            WordPlace = "third"
        Case "4"
            WordPlace = "fourth"
        Case "5"
            WordPlace = "fifth"
        Case "6"
            WordPlace = "sixth"
        Case "7"
            WordPlace = "seventh"
    End Select
End Function
Sub SetScoreMultiplier()
Dim X As Integer
Dim seed As Long
    'AAB where AA is the multiplier and B is the tile (zero offset)
    '--- =standard score
    'DL9 =double letter
    'TL9 =triple letter
    'DW- =double word
    'TW- =triple word
    For X = 0 To 4
        seed = Int(Rnd * 1000)
        Select Case seed 'we can set the loading on what
            Case Is < 400
                sScoreMultiplier(X) = "---"
            Case Is < 600
                seed = Int(Rnd * 7)
                sScoreMultiplier(X) = "DL" & seed
            Case Is < 700
                seed = Int(Rnd * 7)
                sScoreMultiplier(X) = "TL" & seed
            Case Is < 900
                sScoreMultiplier(X) = "DW-"
            Case Is < 1000
                sScoreMultiplier(X) = "TW-"
        End Select
    Next X
    
End Sub
Sub displaywordsandscores()
Dim X As Integer
Dim sT As String
    sT = ""
    For X = 0 To 4
        sT = sT & sCurrentWords(X)
    Next X
    For X = 1 To Len(sT)
        lblLetter(X - 1) = Mid(sT, X, 1)
        lblLetterScore(X - 1) = iTileScores(Asc(Mid(sT, X, 1)) - 65)
    Next X
End Sub
Sub SelectWords()
Dim X As Integer
Dim seed As Long
    For X = 0 To 4
        Do
            seed = Int((UBound(sWords) + 1) * Rnd)
        Loop Until Len(sWords(seed)) = 7
        sCurrentWords(X) = ShuffleWord(UCase(sWords(seed)))
    Next X
End Sub
Function ShuffleWord(sIn As String)
Dim sFrom As String
Dim sTo As String
Dim X As Integer
    sFrom = sIn
    sTo = ""
    Do
        X = Int(Rnd * Len(sIn)) + 1
        If Mid(sFrom, X, 1) <> "¬" Then
            sTo = sTo & Mid(sFrom, X, 1)
            Mid(sFrom, X, 1) = "¬"
        End If
    Loop Until Len(sTo) = Len(sIn)
    ShuffleWord = sTo
End Function
Sub LoadDict()
Dim sIn As String
    App.Title = "5 Minute Scrabble"
    Me.Caption = App.Title
    
    sIn = FileText(App.Path & "\words.dat")
    sWords = Split(sIn, vbCrLf)
    ReDim Preserve sWords(UBound(sWords) - 1)
    sIn = ""

End Sub
Sub ResetTable()
Dim X As Integer
    For X = 0 To shp.Count - 1
        shp(X).FillColor = &HE0E0E0
        shp(X).BorderColor = &H808080
        lblLetter(X) = ""
        lblLetterScore(X) = ""
    Next X
    For X = 0 To txtAnswer.Count - 1
        txtAnswer(X) = ""
        lblInfo(X) = ""
        lblScore(X) = "0"
    Next X
    lblTotalScore = "Final Score: 0"
End Sub
Sub SetTileScores()
Dim sScores As String
Dim X As Integer
    sScores = "010303020104020401080501030101031001010101040408041000"
    For X = 1 To 27
        iTileScores(X - 1) = CInt(Mid(sScores, (X * 2) - 1, 2))
    Next X
End Sub
Function FileText(ByVal filename As String) As String
    Dim handle As Integer
    
    If Len(Dir$(filename)) = 0 Then
        Err.Raise 53
    End If
    
    handle = FreeFile
    Open filename$ For Binary As #handle
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
End Function

