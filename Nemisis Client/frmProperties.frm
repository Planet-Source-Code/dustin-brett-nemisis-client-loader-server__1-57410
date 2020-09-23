VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox chkAttrib 
      Caption         =   "Hidden"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   16
      Top             =   1980
      Width           =   735
   End
   Begin VB.CheckBox chkAttrib 
      Caption         =   "Read-only"
      Height          =   195
      Index           =   0
      Left            =   1140
      TabIndex        =   15
      Top             =   1980
      Width           =   915
   End
   Begin VB.TextBox txtRename 
      Height          =   255
      Left            =   1140
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label8 
      Caption         =   "Attributes:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1980
      Width           =   915
   End
   Begin VB.Label Label6 
      Height          =   1575
      Left            =   4080
      TabIndex        =   14
      Top             =   420
      Width           =   75
   End
   Begin VB.Label lblAccessed 
      AutoSize        =   -1  'True
      Height          =   165
      Left            =   1140
      TabIndex        =   13
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label Label9 
      Caption         =   "Accessed:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label lblModified 
      AutoSize        =   -1  'True
      Height          =   165
      Left            =   1140
      TabIndex        =   11
      Top             =   1440
      Width           =   45
   End
   Begin VB.Label Label7 
      Caption         =   "Modified:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label lblCreated 
      AutoSize        =   -1  'True
      Height          =   165
      Left            =   1140
      TabIndex        =   9
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label Label5 
      Caption         =   "Created:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   915
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Height          =   165
      Left            =   1140
      TabIndex        =   7
      Top             =   960
      Width           =   45
   End
   Begin VB.Label Label4 
      Caption         =   "Size:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   915
   End
   Begin VB.Label lblLocation 
      AutoSize        =   -1  'True
      Height          =   165
      Left            =   1140
      TabIndex        =   5
      Top             =   720
      Width           =   45
   End
   Begin VB.Label Label3 
      Caption         =   "Location:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   915
   End
   Begin VB.Label lblFileType 
      AutoSize        =   -1  'True
      Height          =   165
      Left            =   1140
      TabIndex        =   3
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label2 
      Caption         =   "Rename:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Type of file:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   915
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Me.Visible = True
End Sub

Private Sub txtRename_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        If txtRename = vbNullString Then
            MsgBox "Field Blank"
            Exit Sub
        End If
        frmClient.wskConnect.SendData "043|" & lblLocation.Caption & Left$(Me.Caption, Len(Me.Caption) - 11) & "|" & lblLocation.Caption & txtRename.Text & "|"
        DoEvents
        frmProperties.Caption = txtRename.Text & " Properties"
    End If
End Sub


