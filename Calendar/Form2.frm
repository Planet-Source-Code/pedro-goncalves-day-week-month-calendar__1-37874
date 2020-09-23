VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form2"
   ScaleHeight     =   6570
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   120
      TabIndex        =   37
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
   Begin VB.CommandButton Command3 
      Caption         =   "Hoje"
      Height          =   375
      Left            =   7080
      TabIndex        =   36
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<"
      Height          =   375
      Left            =   3000
      TabIndex        =   35
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">"
      Height          =   375
      Left            =   6720
      TabIndex        =   34
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   0
      Top             =   4680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3240
      TabIndex        =   33
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   15
      Left            =   3000
      TabIndex        =   32
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15:00"
      Height          =   255
      Index           =   14
      Left            =   3000
      TabIndex        =   31
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   13
      Left            =   3000
      TabIndex        =   30
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "14:00"
      Height          =   255
      Index           =   12
      Left            =   3000
      TabIndex        =   29
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   11
      Left            =   3000
      TabIndex        =   28
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "13:00"
      Height          =   255
      Index           =   10
      Left            =   3000
      TabIndex        =   27
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   26
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "12:00"
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   25
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   24
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11:00"
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   23
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   22
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10:00"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   21
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   20
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "09:00"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   19
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   ":30"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   18
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "08:00"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   17
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   3600
      TabIndex        =   16
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   3600
      TabIndex        =   15
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   3600
      TabIndex        =   14
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   3600
      TabIndex        =   13
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   3600
      TabIndex        =   12
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   3600
      TabIndex        =   11
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   10
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   9
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   3600
      TabIndex        =   8
      Top             =   2280
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   7
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   6
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   5
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   3
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   1
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intSelectedHour As String
Dim datSelectedDate As Date

Private Sub Command1_Click()
    datSelectedDate = datSelectedDate + 1
    Label3.Caption = Format(datSelectedDate, "Long Date")
End Sub

Private Sub Command2_Click()
    datSelectedDate = datSelectedDate - 1
    Label3.Caption = Format(datSelectedDate, "Long Date")
End Sub

Private Sub Command3_Click()
    datSelectedDate = Now
    Label3.Caption = Format(datSelectedDate, "Long Date")
End Sub

Private Sub Form_Load()
    datSelectedDate = Now
    Label3.Caption = Format(datSelectedDate, "Long Date")
End Sub

Private Sub Label1_Click(Index As Integer)
    intSelectedHour = Index
    Text1.Top = Label1(intSelectedHour).Top
    Text1.Left = Label1(intSelectedHour).Left
    Text1.Width = Label1(intSelectedHour).Width
    Text1.Height = Label1(intSelectedHour).Height
    Text1.BackColor = Label1(intSelectedHour).BackColor
    Text1.Text = Label1(intSelectedHour).Caption
    Text1.Visible = True
    Text1.SetFocus
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    datSelectedDate = DateClicked
    Label3.Caption = Format(datSelectedDate, "Long Date")
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Label1(intSelectedHour).Caption = Text1.Text
        Text1.Visible = False
        'Aqui será o procedimento de gravação
        If Label2(intSelectedHour).Caption = ":30" Then
            MsgBox Mid$(Label2(intSelectedHour - 1).Caption, 1, 2) & Label2(intSelectedHour).Caption & " " & Label1(intSelectedHour).Caption
        Else
            MsgBox Label2(intSelectedHour).Caption & " " & Label1(intSelectedHour).Caption
        End If
        Exit Sub
    End If
End Sub
