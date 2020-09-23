VERSION 5.00
Begin VB.Form frmChatClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Window"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChatClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtNickname 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3540
      TabIndex        =   2
      Text            =   "Client"
      Top             =   4620
      Width           =   1335
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      Height          =   4395
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4755
   End
   Begin VB.TextBox txtBar 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Press Enter to Send"
      Top             =   4620
      Width           =   3375
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   4980
   End
   Begin VB.Line Line 
      Index           =   1
      X1              =   4980
      X2              =   4980
      Y1              =   0
      Y2              =   4980
   End
   Begin VB.Line Line 
      Index           =   2
      X1              =   0
      X2              =   4980
      Y1              =   4980
      Y2              =   4980
   End
   Begin VB.Line Line 
      Index           =   3
      X1              =   0
      X2              =   4980
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "frmChatClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If frmClient.wskConnect.State = 7 Then
        frmClient.wskConnect.SendData "012"
    End If
End Sub

Private Sub txtBar_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        Dim NickName As String, SData As String
        
        NickName = txtNickname.Text
        SData = "<" & NickName & "> " & txtBar.Text & vbCrLf
        frmClient.wskConnect.SendData "011" & SData
        txtChat.Text = txtChat.Text & "[" & Format$(Now, "HH:mm:ss") & "] " & SData
        txtBar.Text = vbNullString
    End If
End Sub

Private Sub txtChat_Change()
    On Error Resume Next
    txtChat.SelStart = Len(txtChat.Text) - 1
End Sub


