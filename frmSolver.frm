VERSION 5.00
Begin VB.Form frmSolver 
   Caption         =   "WR Solver"
   ClientHeight    =   6570
   ClientLeft      =   465
   ClientTop       =   345
   ClientWidth     =   3660
   Icon            =   "frmSolver.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   3660
   Begin VB.ListBox lstWords 
      Columns         =   3
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   8
      IntegralHeight  =   0   'False
      ItemData        =   "frmSolver.frx":0742
      Left            =   0
      List            =   "frmSolver.frx":0749
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.ListBox lstWords 
      Columns         =   4
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   7
      IntegralHeight  =   0   'False
      ItemData        =   "frmSolver.frx":0757
      Left            =   0
      List            =   "frmSolver.frx":075E
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5160
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.PictureBox picSetup 
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   3600
      TabIndex        =   6
      Top             =   0
      Width           =   3660
      Begin VB.TextBox txtRound 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   121
         Text            =   "1"
         Top             =   0
         Width           =   255
      End
      Begin VB.TextBox txtQuick 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   120
         Top             =   0
         Width           =   3255
      End
      Begin VB.OptionButton optShow 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   3300
         Width           =   200
      End
      Begin VB.OptionButton optShow 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   3300
         Width           =   200
      End
      Begin VB.OptionButton optShow 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   3300
         Width           =   200
      End
      Begin VB.OptionButton optShow 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   3300
         Width           =   200
      End
      Begin VB.OptionButton optShow 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   3300
         Width           =   200
      End
      Begin VB.OptionButton optShow 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   3300
         Width           =   200
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   0
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdSolve 
         Caption         =   "Solve"
         Height          =   375
         Left            =   1200
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2400
         TabIndex        =   109
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtMin 
         Height          =   285
         Left            =   600
         TabIndex        =   108
         Text            =   "3"
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtMax 
         Height          =   285
         Left            =   1680
         TabIndex        =   107
         Text            =   "5"
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   106
         Text            =   "a"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   105
         Text            =   "b"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   720
         TabIndex        =   104
         Text            =   "c"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   103
         Text            =   "d"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   102
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   101
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   100
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   7
         Left            =   2520
         TabIndex        =   99
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   8
         Left            =   2880
         TabIndex        =   98
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   9
         Left            =   3240
         TabIndex        =   97
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   10
         Left            =   0
         TabIndex        =   96
         Text            =   "e"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   11
         Left            =   360
         TabIndex        =   95
         Text            =   "f"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   12
         Left            =   720
         TabIndex        =   94
         Text            =   "g"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   13
         Left            =   1080
         TabIndex        =   93
         Text            =   "h"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   14
         Left            =   1440
         TabIndex        =   92
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   15
         Left            =   1800
         TabIndex        =   91
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   16
         Left            =   2160
         TabIndex        =   90
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   17
         Left            =   2520
         TabIndex        =   89
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   18
         Left            =   2880
         TabIndex        =   88
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   19
         Left            =   3240
         TabIndex        =   87
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   20
         Left            =   0
         TabIndex        =   86
         Text            =   "i"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   21
         Left            =   360
         TabIndex        =   85
         Text            =   "j"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   22
         Left            =   720
         TabIndex        =   84
         Text            =   "k"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   23
         Left            =   1080
         TabIndex        =   83
         Text            =   "l"
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   24
         Left            =   1440
         TabIndex        =   82
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   25
         Left            =   1800
         TabIndex        =   81
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   26
         Left            =   2160
         TabIndex        =   80
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   27
         Left            =   2520
         TabIndex        =   79
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   28
         Left            =   2880
         TabIndex        =   78
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   29
         Left            =   3240
         TabIndex        =   77
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   30
         Left            =   0
         TabIndex        =   76
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   31
         Left            =   360
         TabIndex        =   75
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   32
         Left            =   720
         TabIndex        =   74
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   33
         Left            =   1080
         TabIndex        =   73
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   34
         Left            =   1440
         TabIndex        =   72
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   35
         Left            =   1800
         TabIndex        =   71
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   36
         Left            =   2160
         TabIndex        =   70
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   37
         Left            =   2520
         TabIndex        =   69
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   38
         Left            =   2880
         TabIndex        =   68
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   39
         Left            =   3240
         TabIndex        =   67
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   40
         Left            =   0
         TabIndex        =   66
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   41
         Left            =   360
         TabIndex        =   65
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   42
         Left            =   720
         TabIndex        =   64
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   43
         Left            =   1080
         TabIndex        =   63
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   44
         Left            =   1440
         TabIndex        =   62
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   45
         Left            =   1800
         TabIndex        =   61
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   46
         Left            =   2160
         TabIndex        =   60
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   47
         Left            =   2520
         TabIndex        =   59
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   48
         Left            =   2880
         TabIndex        =   58
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   49
         Left            =   3240
         TabIndex        =   57
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   50
         Left            =   0
         TabIndex        =   56
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   51
         Left            =   360
         TabIndex        =   55
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   52
         Left            =   720
         TabIndex        =   54
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   53
         Left            =   1080
         TabIndex        =   53
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   54
         Left            =   1440
         TabIndex        =   52
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   55
         Left            =   1800
         TabIndex        =   51
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   56
         Left            =   2160
         TabIndex        =   50
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   57
         Left            =   2520
         TabIndex        =   49
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   58
         Left            =   2880
         TabIndex        =   48
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   59
         Left            =   3240
         TabIndex        =   47
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   60
         Left            =   0
         TabIndex        =   46
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   61
         Left            =   360
         TabIndex        =   45
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   62
         Left            =   720
         TabIndex        =   44
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   63
         Left            =   1080
         TabIndex        =   43
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   64
         Left            =   1440
         TabIndex        =   42
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   65
         Left            =   1800
         TabIndex        =   41
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   66
         Left            =   2160
         TabIndex        =   40
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   67
         Left            =   2520
         TabIndex        =   39
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   68
         Left            =   2880
         TabIndex        =   38
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   69
         Left            =   3240
         TabIndex        =   37
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   70
         Left            =   0
         TabIndex        =   36
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   71
         Left            =   360
         TabIndex        =   35
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   72
         Left            =   720
         TabIndex        =   34
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   73
         Left            =   1080
         TabIndex        =   33
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   74
         Left            =   1440
         TabIndex        =   32
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   75
         Left            =   1800
         TabIndex        =   31
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   76
         Left            =   2160
         TabIndex        =   30
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   77
         Left            =   2520
         TabIndex        =   29
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   78
         Left            =   2880
         TabIndex        =   28
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   79
         Left            =   3240
         TabIndex        =   27
         Top             =   2040
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   80
         Left            =   0
         TabIndex        =   26
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   81
         Left            =   360
         TabIndex        =   25
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   82
         Left            =   720
         TabIndex        =   24
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   83
         Left            =   1080
         TabIndex        =   23
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   84
         Left            =   1440
         TabIndex        =   22
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   85
         Left            =   1800
         TabIndex        =   21
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   86
         Left            =   2160
         TabIndex        =   20
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   87
         Left            =   2520
         TabIndex        =   19
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   88
         Left            =   2880
         TabIndex        =   18
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   89
         Left            =   3240
         TabIndex        =   17
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   90
         Left            =   0
         TabIndex        =   16
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   91
         Left            =   360
         TabIndex        =   15
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   92
         Left            =   720
         TabIndex        =   14
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   93
         Left            =   1080
         TabIndex        =   13
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   94
         Left            =   1440
         TabIndex        =   12
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   95
         Left            =   1800
         TabIndex        =   11
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   96
         Left            =   2160
         TabIndex        =   10
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   97
         Left            =   2520
         TabIndex        =   9
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   98
         Left            =   2880
         TabIndex        =   8
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtL 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   99
         Left            =   3240
         TabIndex        =   7
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MIN"
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   113
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MAX"
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   112
         Top             =   3240
         Width           =   615
      End
   End
   Begin VB.ListBox lstWords 
      Columns         =   4
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      IntegralHeight  =   0   'False
      ItemData        =   "frmSolver.frx":076B
      Left            =   0
      List            =   "frmSolver.frx":0772
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.ListBox lstWords 
      Columns         =   5
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   5
      IntegralHeight  =   0   'False
      ItemData        =   "frmSolver.frx":077E
      Left            =   0
      List            =   "frmSolver.frx":0785
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4320
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.ListBox lstWords 
      Columns         =   5
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   4
      IntegralHeight  =   0   'False
      ItemData        =   "frmSolver.frx":0790
      Left            =   0
      List            =   "frmSolver.frx":0797
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.ListBox lstWords 
      Columns         =   6
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   3
      IntegralHeight  =   0   'False
      ItemData        =   "frmSolver.frx":07A1
      Left            =   0
      List            =   "frmSolver.frx":07A8
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   3660
   End
End
Attribute VB_Name = "frmSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False














Sub EnterRoundText(Round As Integer, Letters As String)
Dim i As Integer
Select Case Round
    Case 1
        txtL(0).Text = Mid(Letters, 1, 1)
        txtL(1).Text = Mid(Letters, 2, 1)
        txtL(2).Text = Mid(Letters, 3, 1)
        txtL(3).Text = Mid(Letters, 4, 1)
        
        txtL(10).Text = Mid(Letters, 5, 1)
        txtL(11).Text = Mid(Letters, 6, 1)
        txtL(12).Text = Mid(Letters, 7, 1)
        txtL(13).Text = Mid(Letters, 8, 1)
        
        txtL(20).Text = Mid(Letters, 9, 1)
        txtL(21).Text = Mid(Letters, 10, 1)
        txtL(22).Text = Mid(Letters, 11, 1)
        txtL(23).Text = Mid(Letters, 12, 1)
        
        txtL(30).Text = Mid(Letters, 13, 1)
        txtL(31).Text = Mid(Letters, 14, 1)
        txtL(32).Text = Mid(Letters, 15, 1)
        txtL(33).Text = Mid(Letters, 16, 1)
        
    Case 2
        txtL(2).Text = Mid(Letters, 1, 1)
        txtL(3).Text = Mid(Letters, 2, 1)
        
        txtL(11).Text = Mid(Letters, 3, 1)
        txtL(12).Text = Mid(Letters, 4, 1)
        txtL(13).Text = Mid(Letters, 5, 1)
        txtL(14).Text = Mid(Letters, 6, 1)
        
        txtL(20).Text = Mid(Letters, 7, 1)
        txtL(21).Text = Mid(Letters, 8, 1)
        txtL(22).Text = Mid(Letters, 9, 1)
        txtL(23).Text = Mid(Letters, 10, 1)
        txtL(24).Text = Mid(Letters, 11, 1)
        txtL(25).Text = Mid(Letters, 12, 1)
        
        txtL(30).Text = Mid(Letters, 13, 1)
        txtL(31).Text = Mid(Letters, 14, 1)
        txtL(32).Text = Mid(Letters, 15, 1)
        txtL(33).Text = Mid(Letters, 16, 1)
        txtL(34).Text = Mid(Letters, 17, 1)
        txtL(35).Text = Mid(Letters, 18, 1)
        
        txtL(41).Text = Mid(Letters, 19, 1)
        txtL(42).Text = Mid(Letters, 20, 1)
        txtL(43).Text = Mid(Letters, 21, 1)
        txtL(44).Text = Mid(Letters, 22, 1)
        
        txtL(52).Text = Mid(Letters, 23, 1)
        txtL(53).Text = Mid(Letters, 24, 1)
    
    Case 3
        txtL(0).Text = Mid(Letters, 1, 1)
        txtL(1).Text = Mid(Letters, 2, 1)
        txtL(2).Text = Mid(Letters, 3, 1)
        txtL(3).Text = Mid(Letters, 4, 1)
        
        txtL(10).Text = Mid(Letters, 5, 1)
        txtL(11).Text = Mid(Letters, 6, 1)
        txtL(12).Text = Mid(Letters, 7, 1)
        txtL(13).Text = Mid(Letters, 8, 1)
        
        txtL(20).Text = Mid(Letters, 9, 1)
        txtL(21).Text = Mid(Letters, 10, 1)
        txtL(22).Text = Mid(Letters, 11, 1)
        txtL(23).Text = Mid(Letters, 12, 1)
        txtL(24).Text = Mid(Letters, 13, 1)
        txtL(25).Text = Mid(Letters, 14, 1)
        
        txtL(30).Text = Mid(Letters, 15, 1)
        txtL(31).Text = Mid(Letters, 16, 1)
        txtL(32).Text = Mid(Letters, 17, 1)
        txtL(33).Text = Mid(Letters, 18, 1)
        txtL(34).Text = Mid(Letters, 19, 1)
        txtL(35).Text = Mid(Letters, 20, 1)
        
        txtL(42).Text = Mid(Letters, 21, 1)
        txtL(43).Text = Mid(Letters, 22, 1)
        txtL(44).Text = Mid(Letters, 23, 1)
        txtL(45).Text = Mid(Letters, 24, 1)
        
        txtL(52).Text = Mid(Letters, 25, 1)
        txtL(53).Text = Mid(Letters, 26, 1)
        txtL(54).Text = Mid(Letters, 27, 1)
        txtL(55).Text = Mid(Letters, 28, 1)
        
    Case 4
        txtL(0).Text = Mid(Letters, 1, 1)
        txtL(1).Text = Mid(Letters, 2, 1)
        txtL(2).Text = Mid(Letters, 3, 1)
        txtL(3).Text = Mid(Letters, 4, 1)
        txtL(4).Text = Mid(Letters, 5, 1)
        txtL(5).Text = Mid(Letters, 6, 1)
        
        txtL(10).Text = Mid(Letters, 7, 1)
        txtL(11).Text = Mid(Letters, 8, 1)
        txtL(12).Text = Mid(Letters, 9, 1)
        txtL(13).Text = Mid(Letters, 10, 1)
        txtL(14).Text = Mid(Letters, 11, 1)
        txtL(15).Text = Mid(Letters, 12, 1)
        
        txtL(20).Text = Mid(Letters, 13, 1)
        txtL(21).Text = Mid(Letters, 14, 1)
        
        txtL(24).Text = Mid(Letters, 15, 1)
        txtL(25).Text = Mid(Letters, 16, 1)
        
        txtL(30).Text = Mid(Letters, 17, 1)
        txtL(31).Text = Mid(Letters, 18, 1)
        
        txtL(34).Text = Mid(Letters, 19, 1)
        txtL(35).Text = Mid(Letters, 20, 1)
        
        txtL(40).Text = Mid(Letters, 21, 1)
        txtL(41).Text = Mid(Letters, 22, 1)
        txtL(42).Text = Mid(Letters, 23, 1)
        txtL(43).Text = Mid(Letters, 24, 1)
        txtL(44).Text = Mid(Letters, 25, 1)
        txtL(45).Text = Mid(Letters, 26, 1)
        
        txtL(50).Text = Mid(Letters, 27, 1)
        txtL(51).Text = Mid(Letters, 28, 1)
        txtL(52).Text = Mid(Letters, 29, 1)
        txtL(53).Text = Mid(Letters, 30, 1)
        txtL(54).Text = Mid(Letters, 31, 1)
        txtL(55).Text = Mid(Letters, 32, 1)
End Select

For i = 0 To 99
    If txtL(i).Text = "q" Then txtL(i).Text = "qu"
Next i
End Sub

Sub GetWordList()
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim P As Integer
Dim MyChar As String
Dim TX As Double
'Get our boundaries
Erase Lookup()
Erase XGrid()
Erase MyWords()
ReDim Lookup(255)
ReDim XGrid(9, 9)
x1 = 100
y1 = 100
x2 = -1
y2 = -1
k = 48
P = 0
For j = 0 To 9
    For i = 0 To 9
        MyChar = txtL(P).Text
        If MyChar <> "" Then
            If i < x1 Then x1 = i
            If i > x2 Then x2 = i
            If j < y1 Then y1 = j
            If j > y2 Then y2 = j
            Lookup(k) = MyChar
            XGrid(i, j) = Chr(k)
            k = k + 1
        End If
        P = P + 1
    Next i
Next j
'now we start to generate words
TX = Timer
SearchWord
TX = Timer - TX
'sHCreateThread AddressOf SearchWord, ByVal 0&, CTF_INSIST, ByVal 0&
k = 0
For i = 3 To 8
    k = k + lstWords(i).ListCount
Next i
Me.Caption = k & " words in " & CInt(TX) & " secs."
End Sub







Private Sub cmdCancel_Click()
AllStop = True
End Sub

Private Sub cmdClear_Click()
Dim i As Integer
AllStop = True
For i = 0 To 99
    txtL(i).Text = ""
Next i
For i = 3 To 8
    lstWords(i).Clear
Next i
End Sub

Private Sub cmdSolve_Click()
AllStop = False
For i = 3 To 8
    lstWords(i).Clear
Next i
GetWordList
End Sub

Private Sub Form_Load()
LoadDict
MIN = 3
MAX = 5
optShow(MIN).Value = True
DoEvents
InitWindow Me
Call Form_Resize
End Sub




Private Sub Form_Resize()
Dim i As Integer
For i = 3 To 8
    lstWords(i).Top = 3600
    lstWords(i).Height = Me.Height - picSetup.Height - 400
Next i
End Sub


Private Sub Form_Unload(Cancel As Integer)
AllStop = True
End
End Sub



Private Sub optShow_Click(Index As Integer)
Dim i As Integer
For i = 3 To 8
    lstWords(i).Visible = False
Next i
lstWords(Index).Visible = True
    
End Sub

Private Sub txtL_GotFocus(Index As Integer)
txtL(Index).SelStart = 0
txtL(Index).SelLength = 100
End Sub

Private Sub txtL_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Local Error Resume Next
Select Case KeyCode
    Case vbKeyUp
        txtL(Index - 10).SetFocus
    Case vbKeyDown
        txtL(Index + 10).SetFocus
    Case vbKeyRight
        If Right(Index, 1) < 9 Then
            txtL(Index + 1).SetFocus
        End If
    Case vbKeyLeft
        If Index Mod 10 <> 0 Then
            txtL(Index - 1).SetFocus
        End If
End Select
End Sub

Private Sub txtMax_Change()
If Not IsNumeric(txtMax.Text) Then txtMax.Text = 8
If Val(txtMax.Text) > 8 Then txtMax.Text = 8
If Val(txtMax.Text) < Val(txtMin.Text) Then txtMax.Text = Val(txtMin.Text)

MAX = Val(txtMax.Text)
optShow(MAX).Value = True
End Sub

Private Sub txtMin_Change()
If Not IsNumeric(txtMin.Text) Then txtMin.Text = 3
If Val(txtMin.Text) < 3 Then txtMin.Text = 3
If Val(txtMin.Text) > Val(txtMax.Text) Then txtMin.Text = Val(txtMax.Text)

MIN = Val(txtMin.Text)
optShow(MIN).Value = True
End Sub


Private Sub txtQuick_Change()
EnterRoundText Val(txtRound.Text), txtQuick.Text
End Sub

Private Sub txtRound_Change()
For i = 0 To 99
    txtL(i).Text = ""
Next i
End Sub


