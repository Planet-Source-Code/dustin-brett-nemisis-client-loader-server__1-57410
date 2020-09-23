VERSION 5.00
Begin VB.Form frmScreenShot 
   BorderStyle     =   0  'None
   Caption         =   "ScreenShot"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScreenShot.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Image picSS 
      Appearance      =   0  'Flat
      Height          =   3255
      Left            =   0
      Stretch         =   -1  'True
      ToolTipText     =   "Double click to close preview."
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmScreenShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Dim WW As Integer
    Dim HH As Integer
    
    picSS.Picture = LoadPicture(App.Path & "\ScreenShot.bmp")
    
    WW = Screen.Width
    HH = Screen.Height
    
    picSS.Width = WW
    picSS.Height = HH
    
    Me.Top = 0
    Me.Left = 0
    Me.Width = WW
    Me.Height = HH
    Me.Visible = True
End Sub

Private Sub picSS_DblClick()
    On Error Resume Next
    Unload Me
End Sub


