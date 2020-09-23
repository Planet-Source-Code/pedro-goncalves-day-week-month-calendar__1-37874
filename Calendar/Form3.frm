VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   LinkTopic       =   "Form3"
   ScaleHeight     =   5640
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   10920
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Semana"
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   5160
      Width           =   855
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      MonthBackColor  =   16777215
      ShowToday       =   0   'False
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   22740993
      CurrentDate     =   37477
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   147
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saturday"
      Height          =   255
      Index           =   6
      Left            =   10080
      TabIndex        =   139
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Friday"
      Height          =   255
      Index           =   5
      Left            =   9000
      TabIndex        =   138
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thursday"
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   137
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wednesday"
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   136
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tuesday"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   135
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monday"
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   134
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sunday"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   133
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   7
      Left            =   10080
      TabIndex        =   146
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   6
      Left            =   9000
      TabIndex        =   145
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   144
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   143
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   142
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   141
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   140
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   36
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   10080
      TabIndex        =   132
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   10080
      TabIndex        =   131
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   10080
      TabIndex        =   130
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   10080
      TabIndex        =   129
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   10080
      TabIndex        =   128
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   10080
      TabIndex        =   127
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   10080
      TabIndex        =   126
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   10080
      TabIndex        =   125
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   10080
      TabIndex        =   124
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   10080
      TabIndex        =   123
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   10080
      TabIndex        =   122
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   10080
      TabIndex        =   121
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   10080
      TabIndex        =   120
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   10080
      TabIndex        =   119
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   10080
      TabIndex        =   118
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   9000
      TabIndex        =   117
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   9000
      TabIndex        =   116
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   9000
      TabIndex        =   115
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   9000
      TabIndex        =   114
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   9000
      TabIndex        =   113
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   9000
      TabIndex        =   112
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   9000
      TabIndex        =   111
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   9000
      TabIndex        =   110
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   9000
      TabIndex        =   109
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   9000
      TabIndex        =   108
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   9000
      TabIndex        =   107
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   9000
      TabIndex        =   106
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   9000
      TabIndex        =   105
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   9000
      TabIndex        =   104
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   9000
      TabIndex        =   103
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   7920
      TabIndex        =   102
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   7920
      TabIndex        =   101
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   7920
      TabIndex        =   100
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   7920
      TabIndex        =   99
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   7920
      TabIndex        =   98
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   7920
      TabIndex        =   97
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   7920
      TabIndex        =   96
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   7920
      TabIndex        =   95
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   7920
      TabIndex        =   94
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   7920
      TabIndex        =   93
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   7920
      TabIndex        =   92
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   7920
      TabIndex        =   91
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   7920
      TabIndex        =   90
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7920
      TabIndex        =   89
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   88
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   6840
      TabIndex        =   87
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   6840
      TabIndex        =   86
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   6840
      TabIndex        =   85
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   6840
      TabIndex        =   84
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   6840
      TabIndex        =   83
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   6840
      TabIndex        =   82
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   6840
      TabIndex        =   81
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   6840
      TabIndex        =   80
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   6840
      TabIndex        =   79
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   6840
      TabIndex        =   78
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   6840
      TabIndex        =   77
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   6840
      TabIndex        =   76
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   6840
      TabIndex        =   75
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   6840
      TabIndex        =   74
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   73
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   5760
      TabIndex        =   72
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   5760
      TabIndex        =   71
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   5760
      TabIndex        =   70
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   5760
      TabIndex        =   69
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   5760
      TabIndex        =   68
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   5760
      TabIndex        =   67
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   5760
      TabIndex        =   66
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   5760
      TabIndex        =   65
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   5760
      TabIndex        =   64
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   63
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   62
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   61
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   60
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   59
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   58
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblSaturday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   10080
      TabIndex        =   57
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblFriday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   9000
      TabIndex        =   56
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblThursday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   55
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblWednesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6840
      TabIndex        =   54
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5760
      TabIndex        =   53
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   4680
      TabIndex        =   52
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   4680
      TabIndex        =   51
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   4680
      TabIndex        =   50
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   4680
      TabIndex        =   49
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   48
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   4680
      TabIndex        =   47
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   4680
      TabIndex        =   46
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   4680
      TabIndex        =   45
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   4680
      TabIndex        =   44
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   4680
      TabIndex        =   43
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   42
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   41
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4680
      TabIndex        =   40
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4680
      TabIndex        =   39
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   38
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblMonday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4680
      TabIndex        =   37
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   35
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   34
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   33
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   32
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   31
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   30
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   29
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   28
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   27
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   3600
      TabIndex        =   26
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   3600
      TabIndex        =   25
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   3600
      TabIndex        =   24
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   3600
      TabIndex        =   23
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   3600
      TabIndex        =   22
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lblSunday 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   3600
      TabIndex        =   21
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "08:00"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   20
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   19
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "09:00"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   18
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   17
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10:00"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   16
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   15
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11:00"
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   14
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   13
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12:00"
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   12
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   11
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "13:00"
      Height          =   255
      Index           =   10
      Left            =   3000
      TabIndex        =   10
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   9
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14:00"
      Height          =   255
      Index           =   12
      Left            =   3000
      TabIndex        =   8
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   13
      Left            =   3000
      TabIndex        =   7
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15:00"
      Height          =   255
      Index           =   14
      Left            =   3000
      TabIndex        =   6
      Top             =   4560
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   15
      Left            =   3000
      TabIndex        =   5
      Top             =   4800
      Width           =   495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim datSelectedDate As Date
Dim intSelectedHour As Integer
Dim intSelectdDay As Integer

Private Sub Command1_Click()
    datSelectedDate = datSelectedDate + 7
    Call GetWeek(datSelectedDate)
End Sub

Private Sub Command2_Click()
    datSelectedDate = datSelectedDate - 7
    Call GetWeek(datSelectedDate)
End Sub

Private Sub Command3_Click()
    datSelectedDate = Now
    Call GetWeek(datSelectedDate)
End Sub

Private Sub Form_Load()
    datSelectedDate = Now
    Call GetWeek(datSelectedDate)
End Sub

Private Sub GetWeek(Data As Date)
    Dim i As Integer
    i = Weekday(Data)
    
    Select Case i
        Case 1
            Label11(1).Caption = Format(Data, "DD")
            Label11(2).Caption = Format(Data + 1, "DD")
            Label11(3).Caption = Format(Data + 2, "DD")
            Label11(4).Caption = Format(Data + 3, "DD")
            Label11(5).Caption = Format(Data + 4, "DD")
            Label11(6).Caption = Format(Data + 5, "DD")
            Label11(7).Caption = Format(Data + 6, "DD")
            Label3.Caption = "Semana de " & Label11(1).Caption & " de " & Format(Data, "MMMM") & " a " & Label11(7).Caption & " de " & Format(Data + 6, "MMMM")
        Case 2
            Label11(1).Caption = Format(Data - 1, "DD")
            Label11(2).Caption = Format(Data, "DD")
            Label11(3).Caption = Format(Data + 1, "DD")
            Label11(4).Caption = Format(Data + 2, "DD")
            Label11(5).Caption = Format(Data + 3, "DD")
            Label11(6).Caption = Format(Data + 4, "DD")
            Label11(7).Caption = Format(Data + 5, "DD")
            Label3.Caption = "Semana de " & Label11(1).Caption & " de " & Format(Data - 1, "MMMM") & " a " & Label11(7).Caption & " de " & Format(Data + 5, "MMMM")
        Case 3
            Label11(1).Caption = Format(Data - 2, "DD")
            Label11(2).Caption = Format(Data - 1, "DD")
            Label11(3).Caption = Format(Data, "DD")
            Label11(4).Caption = Format(Data + 1, "DD")
            Label11(5).Caption = Format(Data + 2, "DD")
            Label11(6).Caption = Format(Data + 3, "DD")
            Label11(7).Caption = Format(Data + 4, "DD")
            Label3.Caption = "Semana de " & Label11(1).Caption & " de " & Format(Data - 2, "MMMM") & " a " & Label11(7).Caption & " de " & Format(Data + 4, "MMMM")
        Case 4
            Label11(1).Caption = Format(Data - 3, "DD")
            Label11(2).Caption = Format(Data - 2, "DD")
            Label11(3).Caption = Format(Data - 1, "DD")
            Label11(4).Caption = Format(Data, "DD")
            Label11(5).Caption = Format(Data + 1, "DD")
            Label11(6).Caption = Format(Data + 2, "DD")
            Label11(7).Caption = Format(Data + 3, "DD")
            Label3.Caption = "Semana de " & Label11(1).Caption & " de " & Format(Data - 3, "MMMM") & " a " & Label11(7).Caption & " de " & Format(Data + 3, "MMMM")
        Case 5
            Label11(1).Caption = Format(Data - 4, "DD")
            Label11(2).Caption = Format(Data - 3, "DD")
            Label11(3).Caption = Format(Data - 2, "DD")
            Label11(4).Caption = Format(Data - 1, "DD")
            Label11(5).Caption = Format(Data, "DD")
            Label11(6).Caption = Format(Data + 1, "DD")
            Label11(7).Caption = Format(Data + 2, "DD")
            Label3.Caption = "Semana de " & Label11(1).Caption & " de " & Format(Data - 4, "MMMM") & " a " & Label11(7).Caption & " de " & Format(Data + 2, "MMMM")
        Case 6
            Label11(1).Caption = Format(Data - 5, "DD")
            Label11(2).Caption = Format(Data - 4, "DD")
            Label11(3).Caption = Format(Data - 3, "DD")
            Label11(4).Caption = Format(Data - 2, "DD")
            Label11(5).Caption = Format(Data - 1, "DD")
            Label11(6).Caption = Format(Data, "DD")
            Label11(7).Caption = Format(Data + 1, "DD")
            Label3.Caption = "Semana de " & Label11(1).Caption & " de " & Format(Data - 5, "MMMM") & " a " & Label11(7).Caption & " de " & Format(Data + 1, "MMMM")
        Case 7
            Label11(1).Caption = Format(Data - 6, "DD")
            Label11(2).Caption = Format(Data - 5, "DD")
            Label11(3).Caption = Format(Data - 4, "DD")
            Label11(4).Caption = Format(Data - 3, "DD")
            Label11(5).Caption = Format(Data - 2, "DD")
            Label11(6).Caption = Format(Data - 1, "DD")
            Label11(7).Caption = Format(Data, "DD")
            Label3.Caption = "Semana de " & Label11(1).Caption & " de " & Format(Data - 6, "MMMM") & " a " & Label11(7).Caption & " de " & Format(Data, "MMMM")
    End Select
End Sub

Private Sub lblFriday_Click(Index As Integer)
    intSelectedHour = Index
    intSelectdDay = 6
    Text1.Top = lblFriday(intSelectedHour).Top
    Text1.Left = lblFriday(intSelectedHour).Left
    Text1.Width = lblFriday(intSelectedHour).Width
    Text1.Height = lblFriday(intSelectedHour).Height
    Text1.BackColor = lblFriday(intSelectedHour).BackColor
    Text1.Text = lblFriday(intSelectedHour).Caption
    Text1.Visible = True
    Text1.SetFocus
End Sub

Private Sub lblMonday_Click(Index As Integer)
    intSelectedHour = Index
    intSelectdDay = 2
    Text1.Top = lblMonday(intSelectedHour).Top
    Text1.Left = lblMonday(intSelectedHour).Left
    Text1.Width = lblMonday(intSelectedHour).Width
    Text1.Height = lblMonday(intSelectedHour).Height
    Text1.BackColor = lblMonday(intSelectedHour).BackColor
    Text1.Text = lblMonday(intSelectedHour).Caption
    Text1.Visible = True
    Text1.SetFocus
End Sub

Private Sub lblSaturday_Click(Index As Integer)
    intSelectedHour = Index
    intSelectdDay = 7
    Text1.Top = lblSaturday(intSelectedHour).Top
    Text1.Left = lblSaturday(intSelectedHour).Left
    Text1.Width = lblSaturday(intSelectedHour).Width
    Text1.Height = lblSaturday(intSelectedHour).Height
    Text1.BackColor = lblSaturday(intSelectedHour).BackColor
    Text1.Text = lblSaturday(intSelectedHour).Caption
    Text1.Visible = True
    Text1.SetFocus
End Sub

Private Sub lblSunday_Click(Index As Integer)
    intSelectedHour = Index
    intSelectdDay = 1
    Text1.Top = lblSunday(intSelectedHour).Top
    Text1.Left = lblSunday(intSelectedHour).Left
    Text1.Width = lblSunday(intSelectedHour).Width
    Text1.Height = lblSunday(intSelectedHour).Height
    Text1.BackColor = lblSunday(intSelectedHour).BackColor
    Text1.Text = lblSunday(intSelectedHour).Caption
    Text1.Visible = True
    Text1.SetFocus
End Sub

Private Sub lblThursday_Click(Index As Integer)
    intSelectedHour = Index
    intSelectdDay = 5
    Text1.Top = lblThursday(intSelectedHour).Top
    Text1.Left = lblThursday(intSelectedHour).Left
    Text1.Width = lblThursday(intSelectedHour).Width
    Text1.Height = lblThursday(intSelectedHour).Height
    Text1.BackColor = lblThursday(intSelectedHour).BackColor
    Text1.Text = lblThursday(intSelectedHour).Caption
    Text1.Visible = True
    Text1.SetFocus
End Sub

Private Sub lblTuesday_Click(Index As Integer)
    intSelectedHour = Index
    intSelectdDay = 3
    Text1.Top = lblTuesday(intSelectedHour).Top
    Text1.Left = lblTuesday(intSelectedHour).Left
    Text1.Width = lblTuesday(intSelectedHour).Width
    Text1.Height = lblTuesday(intSelectedHour).Height
    Text1.BackColor = lblTuesday(intSelectedHour).BackColor
    Text1.Text = lblTuesday(intSelectedHour).Caption
    Text1.Visible = True
    Text1.SetFocus
End Sub

Private Sub lblWednesday_Click(Index As Integer)
    intSelectedHour = Index
    intSelectdDay = 4
    Text1.Top = lblWednesday(intSelectedHour).Top
    Text1.Left = lblWednesday(intSelectedHour).Left
    Text1.Width = lblWednesday(intSelectedHour).Width
    Text1.Height = lblWednesday(intSelectedHour).Height
    Text1.BackColor = lblWednesday(intSelectedHour).BackColor
    Text1.Text = lblWednesday(intSelectedHour).Caption
    Text1.Visible = True
    Text1.SetFocus
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Select Case intSelectdDay
            Case 1
                lblSunday(intSelectedHour).Caption = Text1.Text
                Text1.Visible = False
                'Aqui será o procedimento de gravação
                If Label2(intSelectedHour).Caption = ":30" Then
                    MsgBox Mid$(Label2(intSelectedHour - 1).Caption, 1, 2) & Label2(intSelectedHour).Caption & " " & lblSunday(intSelectedHour).Caption
                Else
                    MsgBox Label2(intSelectedHour).Caption & " " & lblSunday(intSelectedHour).Caption
                End If
                Exit Sub

            Case 2
                lblMonday(intSelectedHour).Caption = Text1.Text
                Text1.Visible = False
                'Aqui será o procedimento de gravação
                If Label2(intSelectedHour).Caption = ":30" Then
                    MsgBox Mid$(Label2(intSelectedHour - 1).Caption, 1, 2) & Label2(intSelectedHour).Caption & " " & lblMonday(intSelectedHour).Caption
                Else
                    MsgBox Label2(intSelectedHour).Caption & " " & lblMonday(intSelectedHour).Caption
                End If
                Exit Sub

            Case 3
                lblTuesday(intSelectedHour).Caption = Text1.Text
                Text1.Visible = False
                'Aqui será o procedimento de gravação
                If Label2(intSelectedHour).Caption = ":30" Then
                    MsgBox Mid$(Label2(intSelectedHour - 1).Caption, 1, 2) & Label2(intSelectedHour).Caption & " " & lblTuesday(intSelectedHour).Caption
                Else
                    MsgBox Label2(intSelectedHour).Caption & " " & lblTuesday(intSelectedHour).Caption
                End If
                Exit Sub

            Case 4
                lblWednesday(intSelectedHour).Caption = Text1.Text
                Text1.Visible = False
                'Aqui será o procedimento de gravação
                If Label2(intSelectedHour).Caption = ":30" Then
                    MsgBox Mid$(Label2(intSelectedHour - 1).Caption, 1, 2) & Label2(intSelectedHour).Caption & " " & lblWednesday(intSelectedHour).Caption
                Else
                    MsgBox Label2(intSelectedHour).Caption & " " & lblWednesday(intSelectedHour).Caption
                End If
                Exit Sub

            Case 5
                lblThursday(intSelectedHour).Caption = Text1.Text
                Text1.Visible = False
                'Aqui será o procedimento de gravação
                If Label2(intSelectedHour).Caption = ":30" Then
                    MsgBox Mid$(Label2(intSelectedHour - 1).Caption, 1, 2) & Label2(intSelectedHour).Caption & " " & lblThursday(intSelectedHour).Caption
                Else
                    MsgBox Label2(intSelectedHour).Caption & " " & lblThursday(intSelectedHour).Caption
                End If
                Exit Sub

            Case 6
                lblFriday(intSelectedHour).Caption = Text1.Text
                Text1.Visible = False
                'Aqui será o procedimento de gravação
                If Label2(intSelectedHour).Caption = ":30" Then
                    MsgBox Mid$(Label2(intSelectedHour - 1).Caption, 1, 2) & Label2(intSelectedHour).Caption & " " & lblFriday(intSelectedHour).Caption
                Else
                    MsgBox Label2(intSelectedHour).Caption & " " & lblFriday(intSelectedHour).Caption
                End If
                Exit Sub

            Case 7
                lblSaturday(intSelectedHour).Caption = Text1.Text
                Text1.Visible = False
                'Aqui será o procedimento de gravação
                If Label2(intSelectedHour).Caption = ":30" Then
                    MsgBox Mid$(Label2(intSelectedHour - 1).Caption, 1, 2) & Label2(intSelectedHour).Caption & " " & lblSaturday(intSelectedHour).Caption
                Else
                    MsgBox Label2(intSelectedHour).Caption & " " & lblSaturday(intSelectedHour).Caption
                End If
                Exit Sub

        End Select
    End If
End Sub
