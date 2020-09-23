VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12795
   LinkTopic       =   "Form4"
   ScaleHeight     =   8730
   ScaleWidth      =   12795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   375
      Left            =   9120
      TabIndex        =   136
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   6
      Left            =   8040
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   26
      Top             =   1080
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   6
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   5
      Left            =   6720
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   23
      Top             =   1080
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   5
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   4
      Left            =   5400
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   20
      Top             =   1080
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   4
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   3
      Left            =   4080
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   17
      Top             =   1080
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   3
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   2
      Left            =   2760
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   2
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   1
      Left            =   1440
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   11
      Top             =   1080
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   1
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   0
      Left            =   120
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   0
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   13
      Left            =   8040
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   47
      Top             =   2310
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   13
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   48
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   12
      Left            =   6720
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   44
      Top             =   2310
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   12
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   45
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   11
      Left            =   5400
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   41
      Top             =   2310
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   11
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   10
      Left            =   4080
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   38
      Top             =   2310
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   10
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   9
      Left            =   2760
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   35
      Top             =   2310
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   9
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   8
      Left            =   1440
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   32
      Top             =   2310
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   8
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   7
      Left            =   120
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   29
      Top             =   2310
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   7
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   20
      Left            =   8040
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   68
      Top             =   3540
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   20
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   69
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   20
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   19
      Left            =   6720
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   65
      Top             =   3540
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   19
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   66
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   19
         Left            =   0
         TabIndex        =   67
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   18
      Left            =   5400
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   62
      Top             =   3540
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   18
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   18
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   17
      Left            =   4080
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   59
      Top             =   3540
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   17
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   60
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   16
      Left            =   2760
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   56
      Top             =   3540
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   16
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   16
         Left            =   0
         TabIndex        =   58
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   15
      Left            =   1440
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   53
      Top             =   3540
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   15
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   14
      Left            =   120
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   50
      Top             =   3540
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   14
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   51
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   27
      Left            =   8040
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   89
      Top             =   4770
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   27
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   90
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   27
         Left            =   0
         TabIndex        =   91
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   26
      Left            =   6720
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   86
      Top             =   4770
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   26
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   87
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   26
         Left            =   0
         TabIndex        =   88
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   25
      Left            =   5400
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   83
      Top             =   4770
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   25
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   84
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   25
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   24
      Left            =   4080
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   80
      Top             =   4770
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   24
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   81
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   24
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   23
      Left            =   2760
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   77
      Top             =   4770
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   23
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   78
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   23
         Left            =   0
         TabIndex        =   79
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   22
      Left            =   1440
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   74
      Top             =   4770
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   22
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   75
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   22
         Left            =   0
         TabIndex        =   76
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   21
      Left            =   120
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   71
      Top             =   4770
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   21
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   72
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   21
         Left            =   0
         TabIndex        =   73
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   34
      Left            =   8040
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   110
      Top             =   6000
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   34
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   111
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   34
         Left            =   0
         TabIndex        =   112
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   33
      Left            =   6720
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   107
      Top             =   6000
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   33
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   108
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   33
         Left            =   0
         TabIndex        =   109
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   32
      Left            =   5400
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   104
      Top             =   6000
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   32
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   105
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   32
         Left            =   0
         TabIndex        =   106
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   31
      Left            =   4080
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   101
      Top             =   6000
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   31
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   102
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   31
         Left            =   0
         TabIndex        =   103
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   30
      Left            =   2760
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   98
      Top             =   6000
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   30
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   99
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   30
         Left            =   0
         TabIndex        =   100
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   29
      Left            =   1440
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   95
      Top             =   6000
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   29
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   96
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   29
         Left            =   0
         TabIndex        =   97
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   28
      Left            =   120
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   92
      Top             =   6000
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   28
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   93
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   28
         Left            =   0
         TabIndex        =   94
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   41
      Left            =   8040
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   114
      Top             =   7230
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   41
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   115
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   41
         Left            =   0
         TabIndex        =   116
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   40
      Left            =   6720
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   117
      Top             =   7230
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   40
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   118
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   40
         Left            =   0
         TabIndex        =   119
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   39
      Left            =   5400
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   120
      Top             =   7230
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   39
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   121
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   39
         Left            =   0
         TabIndex        =   122
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   38
      Left            =   4080
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   123
      Top             =   7230
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   38
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   124
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   38
         Left            =   0
         TabIndex        =   125
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   37
      Left            =   2760
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   126
      Top             =   7230
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   37
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   127
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   37
         Left            =   0
         TabIndex        =   128
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   36
      Left            =   1440
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   129
      Top             =   7230
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   36
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   130
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   36
         Left            =   0
         TabIndex        =   131
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1245
      Index           =   35
      Left            =   120
      ScaleHeight     =   1245
      ScaleWidth      =   1335
      TabIndex        =   132
      Top             =   7230
      Width           =   1335
      Begin VB.TextBox txtDay 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   35
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   133
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblDay 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   35
         Left            =   0
         TabIndex        =   134
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   375
      Left            =   120
      TabIndex        =   113
      Top             =   120
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10440
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   360
      TabIndex        =   135
      Top             =   120
      Width           =   8775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saturday"
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Friday"
      Height          =   255
      Left            =   6720
      TabIndex        =   6
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Thursday"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Wednesday"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tuesday"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Monday"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblSunday 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sunday"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim datSelectedMonth As Date

Private Sub Command1_Click()
    datSelectedMonth = datSelectedMonth - 30
    Call GetMonth(datSelectedMonth)
End Sub

Private Sub Command2_Click()
    datSelectedMonth = datSelectedMonth + 30
    Call GetMonth(datSelectedMonth)
End Sub

Private Sub Form_Load()
    datSelectedMonth = Now
    Call GetMonth(datSelectedMonth)
End Sub

Private Sub GetMonth(Data As Date)
    'Get the first day of the month
    Dim datFirstDay As Date
    Dim intNumberOfDays As Integer
    Dim i As Integer
    Dim x As Integer
    
    'Displays the selected month
    Label7.Caption = Format(Data, "MMMM") & " of " & Format(Data, "YYYY")
    
    'Clean the txtDay / lblDay
    For i = 0 To 41
        lblDay(i).Caption = ""
        txtDay(i).Text = ""
        picDay(i).Visible = True
    Next i
    
    Select Case Format(Data, "M")
        Case 1, 3, 5, 7, 8, 10, 12
            intNumberOfDays = 31
        Case 4, 6, 9, 11
            intNumberOfDays = 30
        Case 2
            intNumberOfDays = 28
    End Select
    
    datFirstDay = "1/" & Format(Data, "M") & "/" & Format(Data, "YYYY")
    'picDay(0) -> Sunday
    'picDay(1) -> Monday
    'picDay(2) -> Tuesday
    'picDay(3) -> Wednesday
    'picDay(4) -> Thursday
    'picDay(5) -> Friday
    'picDay(6) -> Saturday
    'lblDay(Weekday(datFirstDay) - 1).Caption = "1"
    x = 1
    For i = Weekday(datFirstDay) - 1 To intNumberOfDays + Weekday(datFirstDay) - 2
        lblDay(i).Caption = x
        x = x + 1
    Next i
    
    'Make the picDays that aren't from the month invisible
    For i = 0 To 41
        If lblDay(i).Caption = "" Then picDay(i).Visible = False
    Next i
    
End Sub

Private Sub txtDay_Change(Index As Integer)
    MsgBox "clo"
End Sub
