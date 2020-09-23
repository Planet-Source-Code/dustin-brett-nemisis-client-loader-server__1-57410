VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmProc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process List"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ImgList 
      Left            =   3660
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProc.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProc.frx":02A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProc.frx":03FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProc.frx":0998
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProc.frx":0F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProc.frx":108C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProc.frx":11E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAuto 
      Caption         =   "Auto Refresh On"
      Height          =   255
      Left            =   1260
      TabIndex        =   3
      Top             =   2820
      Width           =   1215
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   3780
      Top             =   120
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   3540
      TabIndex        =   1
      Top             =   2820
      Width           =   735
   End
   Begin VB.CommandButton cmdKillProc 
      Caption         =   "Kill Process"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   2820
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstProc 
      Height          =   2715
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4789
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblProcesses 
      Caption         =   "Processes:"
      Height          =   195
      Left            =   2580
      TabIndex        =   2
      Top             =   2850
      Width           =   855
   End
End
Attribute VB_Name = "frmProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAuto_Click()
    On Error Resume Next
    If cmdAuto.Caption = "Auto Refresh On" Then
        tmrRefresh.Enabled = True
        frmProc.Caption = "Process List - Refresh in 20s - (" & Format$(Now, "HH:mm:ss") & ")"
        cmdAuto.Caption = "Auto Refresh Off"
    Else
        tmrRefresh.Enabled = False
        frmProc.Caption = "Process List"
        cmdAuto.Caption = "Auto Refresh On"
    End If
End Sub

Private Sub cmdKillProc_Click()
    On Error Resume Next
    If Len(lstProc.SelectedItem.Text) > 0 Then
        frmClient.wskConnect.SendData "046" & lstProc.SelectedItem.Index
        lstProc.ListItems.Remove (lstProc.SelectedItem.Index)
        lblProcesses.Caption = "Processes: " & Int(Mid$(lblProcesses.Caption, 12, Len(lblProcesses.Caption)) - 1)
    End If
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "046"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If frmProc.Visible = True Then
        Cancel = 1
        frmProc.Visible = False
    End If
End Sub

Private Sub tmrRefresh_Timer()
    On Error Resume Next
    If frmProc.Visible = True Then
        frmClient.wskConnect.SendData "046"
        frmProc.Caption = "Process List - Refresh in 20s - (" & Format$(Now, "HH:mm:ss") & ")"
    Else
        tmrRefresh.Enabled = False
        frmProc.Caption = "Process List"
    End If
End Sub


