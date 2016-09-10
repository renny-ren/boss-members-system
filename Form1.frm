VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "老板一卡通会员系统"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15990
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   15990
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "控制台"
      Height          =   2055
      Left            =   6000
      TabIndex        =   138
      Top             =   8280
      Width           =   2175
      Begin VB.CommandButton Command1 
         Caption         =   "保存"
         Height          =   615
         Left            =   480
         TabIndex        =   140
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "清空记录"
         Height          =   615
         Left            =   480
         TabIndex        =   139
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Index           =   0
      ItemData        =   "Form1.frx":70CA
      Left            =   13680
      List            =   "Form1.frx":70D1
      TabIndex        =   106
      Top             =   240
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Caption         =   "今日商品"
      Height          =   9615
      Left            =   9840
      TabIndex        =   53
      Top             =   720
      Width           =   6015
      Begin VB.CommandButton Command22 
         Height          =   375
         Left            =   120
         TabIndex        =   149
         Top             =   8520
         Width           =   375
      End
      Begin VB.CommandButton Command21 
         Height          =   375
         Left            =   120
         TabIndex        =   147
         Top             =   8040
         Width           =   375
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   137
         Top             =   9000
         Width           =   1455
      End
      Begin VB.CommandButton Command20 
         Height          =   375
         Left            =   120
         TabIndex        =   136
         Top             =   9000
         Width           =   375
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   135
         Text            =   "自定义"
         Top             =   9000
         Width           =   1815
      End
      Begin VB.CommandButton Command19 
         Height          =   375
         Left            =   120
         TabIndex        =   134
         Top             =   7560
         Width           =   375
      End
      Begin VB.CommandButton Command18 
         Height          =   375
         Left            =   120
         TabIndex        =   133
         Top             =   7080
         Width           =   375
      End
      Begin VB.CommandButton Command17 
         Height          =   375
         Left            =   120
         TabIndex        =   132
         Top             =   6600
         Width           =   375
      End
      Begin VB.CommandButton Command16 
         Height          =   375
         Left            =   120
         TabIndex        =   131
         Top             =   6120
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         Height          =   375
         Left            =   120
         TabIndex        =   130
         Top             =   5640
         Width           =   375
      End
      Begin VB.CommandButton Command14 
         Height          =   375
         Left            =   120
         TabIndex        =   129
         Top             =   5160
         Width           =   375
      End
      Begin VB.CommandButton Command13 
         Height          =   375
         Left            =   120
         TabIndex        =   128
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         Height          =   375
         Left            =   120
         TabIndex        =   127
         Top             =   4200
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         Height          =   375
         Left            =   120
         TabIndex        =   126
         Top             =   3720
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Height          =   375
         Left            =   120
         TabIndex        =   125
         Top             =   3240
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Height          =   375
         Left            =   120
         TabIndex        =   124
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Height          =   375
         Left            =   120
         TabIndex        =   123
         Top             =   2280
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Height          =   375
         Left            =   120
         TabIndex        =   105
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Height          =   375
         Left            =   120
         TabIndex        =   104
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Left            =   120
         TabIndex        =   103
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   5160
         TabIndex        =   153
         Top             =   8520
         Width           =   495
      End
      Begin VB.Label Label81 
         Alignment       =   2  'Center
         Caption         =   "6.2"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   152
         Top             =   8520
         Width           =   495
      End
      Begin VB.Label Label80 
         Alignment       =   2  'Center
         Caption         =   "6.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   151
         Top             =   8520
         Width           =   495
      End
      Begin VB.Label Label79 
         Caption         =   "百醇"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   150
         Top             =   8520
         Width           =   2295
      End
      Begin VB.Label Label78 
         Caption         =   "八宝粥"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   600
         TabIndex        =   148
         Top             =   8040
         Width           =   2295
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   5160
         TabIndex        =   146
         Top             =   8040
         Width           =   495
      End
      Begin VB.Label Label77 
         Alignment       =   2  'Center
         Caption         =   "3.4"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   145
         Top             =   8040
         Width           =   495
      End
      Begin VB.Label Label76 
         Alignment       =   2  'Center
         Caption         =   "3.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   144
         Top             =   8040
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   5160
         TabIndex        =   122
         Top             =   7560
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   5160
         TabIndex        =   121
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   5160
         TabIndex        =   120
         Top             =   6600
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   5160
         TabIndex        =   119
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   5160
         TabIndex        =   118
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   5160
         TabIndex        =   117
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   5160
         TabIndex        =   116
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   5160
         TabIndex        =   115
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   5160
         TabIndex        =   114
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   5160
         TabIndex        =   113
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5160
         TabIndex        =   112
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   5160
         TabIndex        =   111
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5160
         TabIndex        =   110
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   109
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label74 
         Alignment       =   2  'Center
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5160
         TabIndex        =   108
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label73 
         Alignment       =   2  'Center
         Caption         =   "2.8"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   102
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label Label72 
         Alignment       =   2  'Center
         Caption         =   "2.1"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   101
         Top             =   7560
         Width           =   495
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "2.3"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   100
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Label70 
         Alignment       =   2  'Center
         Caption         =   "1.7"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   99
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label69 
         Alignment       =   2  'Center
         Caption         =   "1.4"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   98
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         Caption         =   "3.3"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   97
         Top             =   6600
         Width           =   495
      End
      Begin VB.Label Label67 
         Alignment       =   2  'Center
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   96
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label Label66 
         Alignment       =   2  'Center
         Caption         =   "2.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   95
         Top             =   7560
         Width           =   495
      End
      Begin VB.Label Label65 
         Alignment       =   2  'Center
         Caption         =   "3.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   94
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label64 
         Alignment       =   2  'Center
         Caption         =   "2.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   93
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label Label63 
         Alignment       =   2  'Center
         Caption         =   "2.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   92
         Top             =   4680
         Width           =   495
      End
      Begin VB.Label Label62 
         Alignment       =   2  'Center
         Caption         =   "2.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   91
         Top             =   5160
         Width           =   495
      End
      Begin VB.Label Label61 
         Alignment       =   2  'Center
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   90
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label Label60 
         Alignment       =   2  'Center
         Caption         =   "1.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   89
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label Label59 
         Alignment       =   2  'Center
         Caption         =   "3.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   88
         Top             =   6600
         Width           =   495
      End
      Begin VB.Label Label58 
         Alignment       =   2  'Center
         Caption         =   "3.8"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   87
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label57 
         Alignment       =   2  'Center
         Caption         =   "0.8"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   86
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label56 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   85
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label55 
         Alignment       =   2  'Center
         Caption         =   "1.9"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   84
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label54 
         Alignment       =   2  'Center
         Caption         =   "2.3"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   83
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label53 
         Alignment       =   2  'Center
         Caption         =   "3.8"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   82
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         Caption         =   "3.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   81
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label51 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   80
         Top             =   3720
         Width           =   495
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "2.5"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   79
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "2.8"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   78
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   77
         Top             =   2760
         Width           =   495
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   76
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   75
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   74
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   73
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label43 
         Caption         =   "听装可乐"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   72
         Top             =   7560
         Width           =   2295
      End
      Begin VB.Label Label42 
         Caption         =   "浪味仙蔬菜"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   375
         Left            =   600
         TabIndex        =   71
         Top             =   7080
         Width           =   2295
      End
      Begin VB.Label Label41 
         Caption         =   "旺旺小小酥"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   600
         TabIndex        =   70
         Top             =   6600
         Width           =   2295
      End
      Begin VB.Label Label40 
         Caption         =   "好丽友巧克力派"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   600
         TabIndex        =   69
         Top             =   6120
         Width           =   2295
      End
      Begin VB.Label Label39 
         Caption         =   "奇多干杯脆"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   600
         TabIndex        =   68
         Top             =   5640
         Width           =   2295
      End
      Begin VB.Label Label38 
         Caption         =   "小米锅巴孜然味"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   600
         TabIndex        =   67
         Top             =   5160
         Width           =   2295
      End
      Begin VB.Label Label37 
         Caption         =   "卫龙大面筋128g"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   66
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label Label36 
         Caption         =   "乐事薯片烤肉味"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   600
         TabIndex        =   65
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label35 
         Caption         =   "乐事薯片"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   600
         TabIndex        =   64
         Top             =   3720
         Width           =   2295
      End
      Begin VB.Label Label34 
         Caption         =   "烧烤素鸡"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   63
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label Label33 
         Caption         =   "土老帽火腿香干"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   62
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label Label32 
         Caption         =   "劲仔小鱼麻辣味"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   61
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label31 
         Caption         =   "金锣火腿肠"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   600
         TabIndex        =   60
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label30 
         Caption         =   "方便面"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   600
         TabIndex        =   59
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label29 
         Caption         =   "瓶装可口可乐500ML"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   600
         TabIndex        =   58
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label28 
         Caption         =   "库存"
         BeginProperty Font 
            Name            =   "恬夕 梦巴黎"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   57
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "会员价"
         BeginProperty Font 
            Name            =   "恬夕 梦巴黎"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   56
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "普通价"
         BeginProperty Font 
            Name            =   "恬夕 梦巴黎"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   55
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "名称"
         BeginProperty Font 
            Name            =   "恬夕 梦巴黎"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   54
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " 信息"
      Height          =   7455
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   9495
      Begin VB.Label Label14 
         Caption         =   "12"
         Height          =   375
         Index           =   12
         Left            =   3000
         TabIndex        =   143
         Top             =   6720
         Width           =   6000
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   12
         Left            =   1800
         TabIndex        =   142
         Top             =   6720
         Width           =   855
      End
      Begin VB.Label Label75 
         Caption         =   "013"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   141
         Top             =   6720
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "11"
         Height          =   375
         Index           =   11
         Left            =   3000
         TabIndex        =   52
         Top             =   6240
         Width           =   6000
      End
      Begin VB.Label Label14 
         Caption         =   "10"
         Height          =   375
         Index           =   10
         Left            =   3000
         TabIndex        =   51
         Top             =   5760
         Width           =   6000
      End
      Begin VB.Label Label14 
         Caption         =   "9"
         Height          =   375
         Index           =   9
         Left            =   3000
         TabIndex        =   50
         Top             =   5280
         Width           =   6000
      End
      Begin VB.Label Label14 
         Caption         =   "8"
         Height          =   375
         Index           =   8
         Left            =   3000
         TabIndex        =   49
         Top             =   4800
         Width           =   6000
      End
      Begin VB.Label Label14 
         Caption         =   "7"
         Height          =   375
         Index           =   7
         Left            =   3000
         TabIndex        =   48
         Top             =   4320
         Width           =   6000
      End
      Begin VB.Label Label14 
         Caption         =   "6"
         Height          =   375
         Index           =   6
         Left            =   3000
         TabIndex        =   47
         Top             =   3840
         Width           =   6000
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "11"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   11
         Left            =   1800
         TabIndex        =   46
         Top             =   6195
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   10
         Left            =   1800
         TabIndex        =   45
         Top             =   5700
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   9
         Left            =   1800
         TabIndex        =   44
         Top             =   5205
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   8
         Left            =   1800
         TabIndex        =   43
         Top             =   4695
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   7
         Left            =   1800
         TabIndex        =   42
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   6
         Left            =   1800
         TabIndex        =   41
         Top             =   3705
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "012"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   6300
         Width           =   495
      End
      Begin VB.Label Label23 
         Caption         =   "011"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   5800
         Width           =   495
      End
      Begin VB.Label Label19 
         Caption         =   "010"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   5300
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "009"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   4800
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "008"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   4300
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "007006"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   3800
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   32
         Top             =   840
         Width           =   6000
      End
      Begin VB.Label Label14 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   31
         Top             =   1320
         Width           =   6000
      End
      Begin VB.Label Label14 
         Caption         =   "2"
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   30
         Top             =   1800
         Width           =   6000
      End
      Begin VB.Label Label14 
         Caption         =   "3"
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   29
         Top             =   2280
         Width           =   6000
      End
      Begin VB.Label Label14 
         Caption         =   "4"
         Height          =   375
         Index           =   4
         Left            =   3000
         TabIndex        =   28
         Top             =   2760
         Width           =   6000
      End
      Begin VB.Label Label14 
         Caption         =   "5"
         Height          =   375
         Index           =   5
         Left            =   3000
         TabIndex        =   27
         Top             =   3240
         Width           =   6000
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   26
         Top             =   705
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   25
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   24
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   23
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   4
         Left            =   1800
         TabIndex        =   22
         Top             =   2700
         Width           =   855
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Smudger LET"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   5
         Left            =   1800
         TabIndex        =   21
         Top             =   3195
         Width           =   855
      End
      Begin VB.Label Label20 
         Caption         =   "谢鑫"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   20
         Top             =   795
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "001"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   800
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "002"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1300
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "003"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label10 
         Caption         =   "004"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2300
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "005"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2800
         Width           =   495
      End
      Begin VB.Label Label12 
         Caption         =   "006"
         BeginProperty Font 
            Name            =   "Square721 BT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3300
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "记录"
         BeginProperty Font 
            Name            =   "恬夕 梦巴黎"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "余额"
         BeginProperty Font 
            Name            =   "恬夕 梦巴黎"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "卡号"
         BeginProperty Font 
            Name            =   "恬夕 梦巴黎"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "恬夕 梦巴黎"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "新增"
      Height          =   2055
      Left            =   3120
      TabIndex        =   5
      Top             =   8280
      Width           =   2175
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "添加"
         Default         =   -1  'True
         Height          =   615
         Left            =   480
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "姓名"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "充值"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   8280
      Width           =   2055
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         ItemData        =   "Form1.frx":70DB
         Left            =   720
         List            =   "Form1.frx":70E2
         TabIndex        =   33
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "确定"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "迷你霹雳体"
            Size            =   20.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "金额"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "姓名"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Label Label4 
      Caption         =   "买家"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13080
      TabIndex        =   107
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "老板一卡通会员计费系统"
      BeginProperty Font 
         Name            =   "钟齐翰墨毛笔"
         Size            =   30
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   3960
      TabIndex        =   34
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'本程序共计用时8小时完成  2016年3月7日14:32:01
'第一次使用前需要在注册表的HKEY_CURRENT_USER\Software\VB and VBA Program Settings键值下新建余额1~余额12的数据并初始化为0值，库存1~库存x的数据并初始化，否则会出错



Private Sub Command1_Click()  '保存按钮
Dim j, k As Integer
For j = 0 To 12   '卡号数-1
SaveSetting App.Title, "Set", "余额" & j + 1, Label22(j).Caption  '保存各个余额数据到注册表
SaveSetting App.Title, "Set", "记录" & j + 1, Label14(j).Caption  '保存各个记录数据到注册表
Next
For k = 0 To 16 ' 商品数-1
SaveSetting App.Title, "Set", "库存" & k + 1, Label74(k).Caption  '保存各个库存数据到注册表
Next
End Sub




Private Sub Command2_Click()  '充值按钮
Dim a As String

Select Case Combo1(1).Text
Case Label20(0).Caption: Label22(0).Caption = Val(Label22(0).Caption) + Text2.Text  '加对应余额
a = Label20(0).Caption                                                              'a取值为充值的名字，后面显示提示信息用
Set b = Label22(0)                                                             'b取值为余额，后面显示提示信息用
Case Label20(1).Caption: Label22(1).Caption = Val(Label22(1).Caption) + Text2.Text
a = Label20(1).Caption
Set b = Label22(1)
Case Label20(2).Caption: Label22(2).Caption = Val(Label22(2).Caption) + Text2.Text
a = Label20(2).Caption
Set b = Label22(2)
Case Label20(3).Caption: Label22(3).Caption = Val(Label22(3).Caption) + Text2.Text
a = Label20(3).Caption
Set b = Label22(3)
Case Label20(4).Caption: Label22(4).Caption = Val(Label22(4).Caption) + Text2.Text
a = Label20(4).Caption
Set b = Label22(4)
Case Label20(5).Caption: Label22(5).Caption = Val(Label22(5).Caption) + Text2.Text
a = Label20(5).Caption
Set b = Label22(5)
Case Label20(6).Caption: Label22(6).Caption = Val(Label22(6).Caption) + Text2.Text
a = Label20(6).Caption
Set b = Label22(6)
Case Label20(7).Caption: Label22(7).Caption = Val(Label22(7).Caption) + Text2.Text
a = Label20(7).Caption
Set b = Label22(7)
Case Label20(8).Caption: Label22(8).Caption = Val(Label22(8).Caption) + Text2.Text
a = Label20(8).Caption
Set b = Label22(8)
Case Label20(9).Caption: Label22(9).Caption = Val(Label22(9).Caption) + Text2.Text
a = Label20(9).Caption
Set b = Label22(9)
Case Label20(10).Caption: Label22(10).Caption = Val(Label22(10).Caption) + Text2.Text
a = Label20(10).Caption
Set b = Label22(10)
Case Label20(11).Caption: Label22(11).Caption = Val(Label22(11).Caption) + Text2.Text
a = Label20(11).Caption
Set b = Label22(11)
Case Label20(12).Caption: Label22(12).Caption = Val(Label22(12).Caption) + Text2.Text
a = Label20(12).Caption
Set b = Label22(12)

End Select

If Text2.Text < 20 Then
MsgBox a & "  " & Text2.Text & "元充值成功！" & vbCrLf & "没有奖励" & vbCrLf & "卡上余额：" & b.Caption & "元", , 提示
  ElseIf Text2.Text >= 20 And Text2.Text < 50 Then
  b.Caption = b.Caption + 1
  MsgBox a & "  " & Text2.Text & "元充值成功！" & vbCrLf & "奖励1元！" & vbCrLf & "卡上余额：" & b.Caption & "元", , 提示
  
 ElseIf Text2.Text >= 50 And Text2.Text < 100 Then
 b.Caption = b.Caption + 2
 MsgBox a & "  " & Text2.Text & "元充值成功！" & vbCrLf & "奖励2元！" & vbCrLf & "卡上余额：" & b.Caption & "元", , 提示
 
   ElseIf Text2.Text >= 100 And Text2.Text < 200 Then
   b.Caption = b.Caption + 5
   MsgBox a & "  " & Text2.Text & "元充值成功！" & vbCrLf & "奖励5元！" & vbCrLf & "卡上余额：" & b.Caption & "元", , 提示
   
    ElseIf Text2.Text >= 200 Then
    b.Caption = b.Caption + 15
   MsgBox a & "  " & Text2.Text & "元充值成功！" & vbCrLf & "奖励15元！" & vbCrLf & "卡上余额：" & b.Caption & "元", , 提示
   End If


End Sub

Private Sub Command20_Click() '自定义
Select Case Combo1(0).Text

Case Label20(0).Caption: Label22(0).Caption = Label22(0).Caption - Text4.Text
Label14(0).Caption = Label14(0).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - Text4.Text
Label14(1).Caption = Label14(1).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - Text4.Text
Label14(2).Caption = Label14(2).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - Text4.Text
Label14(3).Caption = Label14(3).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - Text4.Text
Label14(4).Caption = Label14(4).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - Text4.Text
Label14(5).Caption = Label14(5).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - Text4.Text
Label14(6).Caption = Label14(6).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - Text4.Text
Label14(7).Caption = Label14(7).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - Text4.Text
Label14(8).Caption = Label14(8).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - Text4.Text
Label14(9).Caption = Label14(9).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - Text4.Text
Label14(10).Caption = Label14(10).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - Text4.Text
Label14(11).Caption = Label14(11).Caption + Text3.Text & " -" & Text4.Text & "；"

Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - Text4.Text
Label14(12).Caption = Label14(12).Caption + Text3.Text & " -" & Text4.Text & "；"

End Select

End Sub





Private Sub Command3_Click()  '瓶装可乐按钮
Select Case Combo1(0).Text

Case Label20(0).Caption: Label22(0).Caption = Label22(0).Caption - 2.8   '2.8为该商品价格
Label14(0).Caption = Label14(0).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.8
Label14(1).Caption = Label14(1).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.8
Label14(2).Caption = Label14(2).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.8
Label14(3).Caption = Label14(3).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.8
Label14(4).Caption = Label14(4).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.8
Label14(5).Caption = Label14(5).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.8
Label14(6).Caption = Label14(6).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.8
Label14(7).Caption = Label14(7).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.8
Label14(8).Caption = Label14(8).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.8
Label14(9).Caption = Label14(9).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.8
Label14(10).Caption = Label14(10).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.8
Label14(11).Caption = Label14(11).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.8
Label14(12).Caption = Label14(12).Caption + "瓶装可乐 -2.8；"
Label74(0).Caption = Label74(0).Caption - 1

End Select

End Sub

Private Sub Command4_Click()  '方便面按钮
Select Case Combo1(0).Text

Case Label20(0).Caption: Label22(0).Caption = Label22(0).Caption - 3.8
Label14(0).Caption = Label14(0).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 3.8
Label14(1).Caption = Label14(1).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 3.8
Label14(2).Caption = Label14(2).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 3.8
Label14(3).Caption = Label14(3).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 3.8
Label14(4).Caption = Label14(4).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 3.8
Label14(5).Caption = Label14(5).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 3.8
Label14(6).Caption = Label14(6).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 3.8
Label14(7).Caption = Label14(7).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 3.8
Label14(8).Caption = Label14(8).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 3.8
Label14(9).Caption = Label14(9).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 3.8
Label14(10).Caption = Label14(10).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 3.8
Label14(11).Caption = Label14(11).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 3.8
Label14(12).Caption = Label14(12).Caption + "方便面 -3.8；"
Label74(1).Caption = Label74(1).Caption - 1
End Select

End Sub

Private Sub Command5_Click() '火腿肠按钮
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 0.8  '扣除金额
Label14(0).Caption = Label14(0).Caption + "火腿肠 -0.8；"                '添加记录
Label74(2).Caption = Label74(2).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 0.8
Label14(1).Caption = Label14(1).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 0.8
Label14(2).Caption = Label14(2).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 0.8
Label14(3).Caption = Label14(3).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 0.8
Label14(4).Caption = Label14(4).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 0.8
Label14(5).Caption = Label14(5).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 0.8
Label14(6).Caption = Label14(6).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 0.8
Label14(7).Caption = Label14(7).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 0.8
Label14(8).Caption = Label14(8).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 0.8
Label14(9).Caption = Label14(9).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 0.8
Label14(10).Caption = Label14(10).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 0.8
Label14(11).Caption = Label14(11).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 0.8
Label14(12).Caption = Label14(12).Caption + "火腿肠 -0.8；"
Label74(2).Caption = Label74(2).Caption - 1
End Select
End Sub
Private Sub Command8_Click()  '劲仔
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 1  '扣除金额
Label14(0).Caption = Label14(0).Caption + "劲仔-1；"                '添加记录
Label74(3).Caption = Label74(3).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 1
Label14(1).Caption = Label14(1).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 1
Label14(2).Caption = Label14(2).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 1
Label14(3).Caption = Label14(3).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 1
Label14(4).Caption = Label14(4).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 1
Label14(5).Caption = Label14(5).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 1
Label14(6).Caption = Label14(6).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 1
Label14(7).Caption = Label14(7).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 1
Label14(8).Caption = Label14(8).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 1
Label14(9).Caption = Label14(9).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 1
Label14(10).Caption = Label14(10).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 1
Label14(11).Caption = Label14(11).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 1
Label14(12).Caption = Label14(12).Caption + "劲仔-1；"
Label74(3).Caption = Label74(3).Caption - 1
End Select

End Sub
Private Sub Command9_Click()  '土老帽火腿香干
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 1.9  '扣除金额
Label14(0).Caption = Label14(0).Caption + "土老帽-1.9；"                '添加记录
Label74(4).Caption = Label74(4).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 1.9
Label14(1).Caption = Label14(1).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 1.9
Label14(2).Caption = Label14(2).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 1.9
Label14(3).Caption = Label14(3).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 1.9
Label14(4).Caption = Label14(4).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 1.9
Label14(5).Caption = Label14(5).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 1.9
Label14(6).Caption = Label14(6).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 1.9
Label14(7).Caption = Label14(7).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 1.9
Label14(8).Caption = Label14(8).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 1.9
Label14(9).Caption = Label14(9).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 1.9
Label14(10).Caption = Label14(10).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 1.9
Label14(11).Caption = Label14(11).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 1.9
Label14(12).Caption = Label14(12).Caption + "土老帽-1.9；"
Label74(4).Caption = Label74(4).Caption - 1
End Select

End Sub
Private Sub Command10_Click()  '烧烤素鸡
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 2.3  '扣除金额
Label14(0).Caption = Label14(0).Caption + "素鸡-2.3；"                '添加记录
Label74(5).Caption = Label74(5).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.3
Label14(1).Caption = Label14(1).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.3
Label14(2).Caption = Label14(2).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.3
Label14(3).Caption = Label14(3).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.3
Label14(4).Caption = Label14(4).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.3
Label14(5).Caption = Label14(5).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.3
Label14(6).Caption = Label14(6).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.3
Label14(7).Caption = Label14(7).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.3
Label14(8).Caption = Label14(8).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.3
Label14(9).Caption = Label14(9).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.3
Label14(10).Caption = Label14(10).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.3
Label14(11).Caption = Label14(11).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.3
Label14(12).Caption = Label14(12).Caption + "素鸡-2.3；"
Label74(5).Caption = Label74(5).Caption - 1
End Select

End Sub
Private Sub Command11_Click()  '乐事薯片
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 3.8  '扣除金额
Label14(0).Caption = Label14(0).Caption + "薯片-3.8；"                '添加记录
Label74(6).Caption = Label74(6).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 3.8
Label14(1).Caption = Label14(1).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 3.8
Label14(2).Caption = Label14(2).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 3.8
Label14(3).Caption = Label14(3).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 3.8
Label14(4).Caption = Label14(4).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 3.8
Label14(5).Caption = Label14(5).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 3.8
Label14(6).Caption = Label14(6).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 3.8
Label14(7).Caption = Label14(7).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 3.8
Label14(8).Caption = Label14(8).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 3.8
Label14(9).Caption = Label14(9).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 3.8
Label14(10).Caption = Label14(10).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 3.8
Label14(11).Caption = Label14(11).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 3.8
Label14(12).Caption = Label14(12).Caption + "薯片-3.8；"
Label74(6).Caption = Label74(6).Caption - 1
End Select

End Sub
Private Sub Command12_Click()  '乐事薯片烤肉味
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 3.5  '扣除金额
Label14(0).Caption = Label14(0).Caption + "薯片扁的-3.5；"                '添加记录
Label74(7).Caption = Label74(7).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 3.5
Label14(1).Caption = Label14(1).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 3.5
Label14(2).Caption = Label14(2).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 3.5
Label14(3).Caption = Label14(3).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 3.5
Label14(4).Caption = Label14(4).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 3.5
Label14(5).Caption = Label14(5).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 3.5
Label14(6).Caption = Label14(6).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 3.5
Label14(7).Caption = Label14(7).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 3.5
Label14(8).Caption = Label14(8).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 3.5
Label14(9).Caption = Label14(9).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 3.5
Label14(10).Caption = Label14(10).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 3.5
Label14(11).Caption = Label14(11).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 3.5
Label14(12).Caption = Label14(12).Caption + "薯片扁的-3.5；"
Label74(7).Caption = Label74(7).Caption - 1
End Select

End Sub
Private Sub Command13_Click()  '卫龙大面筋
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 2.5  '扣除金额
Label14(0).Caption = Label14(0).Caption + "大面筋-2.5；"                '添加记录
Label74(8).Caption = Label74(8).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.5
Label14(1).Caption = Label14(1).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.5
Label14(2).Caption = Label14(2).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.5
Label14(3).Caption = Label14(3).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.5
Label14(4).Caption = Label14(4).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.5
Label14(5).Caption = Label14(5).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.5
Label14(6).Caption = Label14(6).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.5
Label14(7).Caption = Label14(7).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.5
Label14(8).Caption = Label14(8).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.5
Label14(9).Caption = Label14(9).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.5
Label14(10).Caption = Label14(10).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.5
Label14(11).Caption = Label14(11).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.5
Label14(12).Caption = Label14(12).Caption + "大面筋-2.5；"
Label74(8).Caption = Label74(8).Caption - 1
End Select

End Sub
Private Sub Command14_Click()  '小米锅巴
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 2.3  '扣除金额
Label14(0).Caption = Label14(0).Caption + "小米锅巴-2.3；"                '添加记录
Label74(9).Caption = Label74(9).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.3
Label14(1).Caption = Label14(1).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.3
Label14(2).Caption = Label14(2).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.3
Label14(3).Caption = Label14(3).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.3
Label14(4).Caption = Label14(4).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.3
Label14(5).Caption = Label14(5).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.3
Label14(6).Caption = Label14(6).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.3
Label14(7).Caption = Label14(7).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.3
Label14(8).Caption = Label14(8).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.3
Label14(9).Caption = Label14(9).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.3
Label14(10).Caption = Label14(10).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.3
Label14(11).Caption = Label14(11).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.3
Label14(12).Caption = Label14(12).Caption + "小米锅巴-2.3；"
Label74(9).Caption = Label74(9).Caption - 1
End Select

End Sub
Private Sub Command15_Click()  '奇多干杯脆
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 1.7  '扣除金额
Label14(0).Caption = Label14(0).Caption + "干杯脆-1.7；"                '添加记录
Label74(10).Caption = Label74(10).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 1.7
Label14(1).Caption = Label14(1).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 1.7
Label14(2).Caption = Label14(2).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 1.7
Label14(3).Caption = Label14(3).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 1.7
Label14(4).Caption = Label14(4).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 1.7
Label14(5).Caption = Label14(5).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 1.7
Label14(6).Caption = Label14(6).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 1.7
Label14(7).Caption = Label14(7).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 1.7
Label14(8).Caption = Label14(8).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 1.7
Label14(9).Caption = Label14(9).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 1.7
Label14(10).Caption = Label14(10).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 1.7
Label14(11).Caption = Label14(11).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 1.7
Label14(12).Caption = Label14(12).Caption + "干杯脆-1.7；"
Label74(10).Caption = Label74(10).Caption - 1
End Select

End Sub
Private Sub Command16_Click()  '好丽友
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 1.4  '扣除金额
Label14(0).Caption = Label14(0).Caption + "好丽友-1.4；"                '添加记录
Label74(11).Caption = Label74(11).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 1.4
Label14(1).Caption = Label14(1).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 1.4
Label14(2).Caption = Label14(2).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 1.4
Label14(3).Caption = Label14(3).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 1.4
Label14(4).Caption = Label14(4).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 1.4
Label14(5).Caption = Label14(5).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 1.4
Label14(6).Caption = Label14(6).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 1.4
Label14(7).Caption = Label14(7).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 1.4
Label14(8).Caption = Label14(8).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 1.4
Label14(9).Caption = Label14(9).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 1.4
Label14(10).Caption = Label14(10).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 1.4
Label14(11).Caption = Label14(11).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 1.4
Label14(12).Caption = Label14(12).Caption + "好丽友-1.4；"
Label74(11).Caption = Label74(11).Caption - 1
End Select

End Sub

Private Sub Command17_Click() '旺旺小小酥
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 3.3  '扣除金额
Label14(0).Caption = Label14(0).Caption + "小小酥-3.3；"                '添加记录
Label74(12).Caption = Label74(12).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 3.3
Label14(1).Caption = Label14(1).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 3.3
Label14(2).Caption = Label14(2).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 3.3
Label14(3).Caption = Label14(3).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 3.3
Label14(4).Caption = Label14(4).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 3.3
Label14(5).Caption = Label14(5).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 3.3
Label14(6).Caption = Label14(6).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 3.3
Label14(7).Caption = Label14(7).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 3.3
Label14(8).Caption = Label14(8).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 3.3
Label14(9).Caption = Label14(9).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 3.3
Label14(10).Caption = Label14(10).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 3.3
Label14(11).Caption = Label14(11).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 3.3
Label14(12).Caption = Label14(12).Caption + "小小酥-3.3；"
Label74(12).Caption = Label74(12).Caption - 1
End Select

End Sub
Private Sub Command18_Click()  '浪味仙
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 2.8  '扣除金额
Label14(0).Caption = Label14(0).Caption + "浪味仙-2.8；"                '添加记录
Label74(13).Caption = Label74(13).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.8
Label14(1).Caption = Label14(1).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.8
Label14(2).Caption = Label14(2).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.8
Label14(3).Caption = Label14(3).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.8
Label14(4).Caption = Label14(4).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.8
Label14(5).Caption = Label14(5).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.8
Label14(6).Caption = Label14(6).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.8
Label14(7).Caption = Label14(7).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.8
Label14(8).Caption = Label14(8).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.8
Label14(9).Caption = Label14(9).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.8
Label14(10).Caption = Label14(10).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.8
Label14(11).Caption = Label14(11).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.8
Label14(12).Caption = Label14(12).Caption + "浪味仙-2.8；"
Label74(13).Caption = Label74(13).Caption - 1
End Select

End Sub

Private Sub Command19_Click() '听装可乐
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 2.1  '扣除金额
Label14(0).Caption = Label14(0).Caption + "听可乐-2.1；"                '添加记录
Label74(14).Caption = Label74(14).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 2.1
Label14(1).Caption = Label14(1).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 2.1
Label14(2).Caption = Label14(2).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 2.1
Label14(3).Caption = Label14(3).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 2.1
Label14(4).Caption = Label14(4).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 2.1
Label14(5).Caption = Label14(5).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 2.1
Label14(6).Caption = Label14(6).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 2.1
Label14(7).Caption = Label14(7).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 2.1
Label14(8).Caption = Label14(8).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 2.1
Label14(9).Caption = Label14(9).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 2.1
Label14(10).Caption = Label14(10).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 2.1
Label14(11).Caption = Label14(11).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 2.1
Label14(12).Caption = Label14(12).Caption + "听可乐-2.1；"
Label74(14).Caption = Label74(14).Caption - 1
End Select
End Sub
Private Sub Command21_Click() '八宝粥
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 3.4  '扣除金额
Label14(0).Caption = Label14(0).Caption + "八宝粥-3.4；"                '添加记录
Label74(15).Caption = Label74(15).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 3.4
Label14(1).Caption = Label14(1).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 3.4
Label14(2).Caption = Label14(2).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 3.4
Label14(3).Caption = Label14(3).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 3.4
Label14(4).Caption = Label14(4).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 3.4
Label14(5).Caption = Label14(5).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 3.4
Label14(6).Caption = Label14(6).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 3.4
Label14(7).Caption = Label14(7).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 3.4
Label14(8).Caption = Label14(8).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 3.4
Label14(9).Caption = Label14(9).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 3.4
Label14(10).Caption = Label14(10).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 3.4
Label14(11).Caption = Label14(11).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 3.4
Label14(12).Caption = Label14(12).Caption + "八宝粥-3.4；"
Label74(15).Caption = Label74(15).Caption - 1
End Select

End Sub

Private Sub Command22_Click()  '百醇
Select Case Combo1(0).Text

Case Label20(0).Caption:  Label22(0).Caption = Label22(0).Caption - 6.2  '扣除金额
Label14(0).Caption = Label14(0).Caption + "百醇-6.2；"                '添加记录
Label74(16).Caption = Label74(16).Caption - 1                              '减少库存
Case Label20(1).Caption:  Label22(1).Caption = Label22(1).Caption - 6.2
Label14(1).Caption = Label14(1).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(2).Caption:  Label22(2).Caption = Label22(2).Caption - 6.2
Label14(2).Caption = Label14(2).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(3).Caption:  Label22(3).Caption = Label22(3).Caption - 6.2
Label14(3).Caption = Label14(3).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(4).Caption:  Label22(4).Caption = Label22(4).Caption - 6.2
Label14(4).Caption = Label14(4).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(5).Caption:  Label22(5).Caption = Label22(5).Caption - 6.2
Label14(5).Caption = Label14(5).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1

Case Label20(6).Caption:  Label22(6).Caption = Label22(6).Caption - 6.2
Label14(6).Caption = Label14(6).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(7).Caption:  Label22(7).Caption = Label22(7).Caption - 6.2
Label14(7).Caption = Label14(7).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(8).Caption:  Label22(8).Caption = Label22(8).Caption - 6.2
Label14(8).Caption = Label14(8).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(9).Caption:  Label22(9).Caption = Label22(9).Caption - 6.2
Label14(9).Caption = Label14(9).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(10).Caption:  Label22(10).Caption = Label22(10).Caption - 6.2
Label14(10).Caption = Label14(10).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(11).Caption:  Label22(11).Caption = Label22(11).Caption - 6.2
Label14(11).Caption = Label14(11).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
Case Label20(12).Caption:  Label22(12).Caption = Label22(12).Caption - 6.2
Label14(12).Caption = Label14(12).Caption + "百醇-6.2；"
Label74(16).Caption = Label74(16).Caption - 1
End Select
End Sub
Private Sub Command6_Click() ' 添加姓名按钮
a = GetSetting(App.Title, "Set", "变量a")  '从注册表取得变量a的值，用于判断下一个添加的姓名是第几个
Load Label20(a + 1)   '加载下一个姓名
Label20(a + 1).Top = Label20(a).Top + 500 '调整位置
Label20(a + 1).Visible = True
Label20(a + 1).Caption = Text1.Text   '设置姓名
SaveSetting App.Title, "Set", "姓名" & (a + 2), Label20(a + 1).Caption  '保存姓名的数据到注册表，下一次加载窗体时读取
a = a + 1
SaveSetting App.Title, "Set", "变量a", a
Combo1(0).AddItem Text1.Text
Combo1(1).AddItem Text1.Text

End Sub

Private Sub Command7_Click()
For i = 0 To 12          '卡号数-1
Label14(i).Caption = ""
Next
End Sub


Private Sub Form_Load()
Dim j, k As Integer
For j = 0 To 12         '卡号数-1
Label22(j).Caption = GetSetting(App.Title, "Set", "余额" & j + 1)   '读取余额数据
Label14(j).Caption = GetSetting(App.Title, "Set", "记录" & j + 1)   '读取记录数据
Next
    
For k = 0 To 16         '商品数-1
Label74(k).Caption = GetSetting(App.Title, "Set", "库存" & k + 1)   '读取库存数据
Next

  a = GetSetting(App.Title, "Set", "变量a")
If a <> 0 Then
Dim i As Integer
i = 0
  While i <> a
  Load Label20(i + 1)
  Label20(i + 1).Top = Label20(i).Top + 500
  Label20(i + 1).Visible = True
  Label20(i + 1).Caption = GetSetting(App.Title, "Set", "姓名" & (i + 2))   '读取姓名数据
  Combo1(0).AddItem Label20(i + 1).Caption
  Combo1(1).AddItem Label20(i + 1).Caption
  i = i + 1
  Wend
End If
 
End Sub

