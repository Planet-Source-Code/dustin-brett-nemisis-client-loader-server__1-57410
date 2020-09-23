VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFileBrowser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Browser"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "frmFileBrowser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ProgressBar pgBar 
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   6000
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Get Drive Label"
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
      Height          =   285
      Index           =   7
      Left            =   2460
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   390
      Width           =   1155
   End
   Begin VB.TextBox txtDirectory 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   6
      Top             =   6270
      Width           =   3555
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "GoTo Directory"
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
      Height          =   285
      Index           =   5
      Left            =   3660
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6270
      Width           =   1155
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Refresh"
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
      Height          =   285
      Index           =   2
      Left            =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   1155
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "New Folder"
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
      Height          =   285
      Index           =   4
      Left            =   1260
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   60
      TabIndex        =   12
      Top             =   390
      Width           =   2355
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Set Drive Label"
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
      Height          =   285
      Index           =   6
      Left            =   3660
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   390
      Width           =   1155
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   3660
      Top             =   1140
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
            Picture         =   "frmFileBrowser.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":02A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":03FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":0998
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":0F32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":108C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":11E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":1340
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":149A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":15F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":1B8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":2128
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":2282
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":281C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":2976
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":2AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":306A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":31C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":331E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":3478
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":3A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":3FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":4106
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":4260
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":43BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":4514
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":466E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":47C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":4D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":4EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":5456
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":59F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":5F8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFileBrowser.frx":6524
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "All Files(*.*)|*.*"
   End
   Begin MSWinsockLib.Winsock w1 
      Left            =   4320
      Top             =   1620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7117
   End
   Begin MSWinsockLib.Winsock FT 
      Left            =   4320
      Top             =   1140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstFolders 
      Height          =   4845
      Left            =   60
      TabIndex        =   9
      Top             =   1080
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   8546
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.ComboBox CmbDrv 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "frmFileBrowser.frx":667E
      Left            =   60
      List            =   "frmFileBrowser.frx":6680
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   60
      Width           =   3555
   End
   Begin VB.TextBox txtDirName 
      Height          =   285
      Left            =   4920
      TabIndex        =   8
      Top             =   6300
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Get Drives"
      Default         =   -1  'True
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
      Height          =   285
      Index           =   0
      Left            =   3660
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   1155
   End
   Begin MSComctlLib.StatusBar statBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6600
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   450
      Style           =   1
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
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Drive Info"
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
      Height          =   285
      Index           =   1
      Left            =   3660
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   1155
   End
   Begin VB.CommandButton cmdCommand 
      Caption         =   "Upload File"
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
      Height          =   285
      Index           =   3
      Left            =   2460
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   1155
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File Options"
      Visible         =   0   'False
      Begin VB.Menu cmdRun 
         Caption         =   "Open"
      End
      Begin VB.Menu Line8 
         Caption         =   "-"
      End
      Begin VB.Menu cmdSendTo 
         Caption         =   "Send To"
         Begin VB.Menu cmdBin 
            Caption         =   "Recycle Bin"
         End
         Begin VB.Menu Line9 
            Caption         =   "-"
         End
         Begin VB.Menu cmdDrive1 
            Caption         =   "Drive1"
            Visible         =   0   'False
         End
         Begin VB.Menu cmdDrive2 
            Caption         =   "Drive2"
            Visible         =   0   'False
         End
         Begin VB.Menu cmdDrive3 
            Caption         =   "Drive3"
            Visible         =   0   'False
         End
         Begin VB.Menu cmdDrive4 
            Caption         =   "Drive4"
            Visible         =   0   'False
         End
         Begin VB.Menu cmdDrive5 
            Caption         =   "Drive5"
            Visible         =   0   'False
         End
         Begin VB.Menu Line6 
            Caption         =   "-"
         End
         Begin VB.Menu cmdSpecify 
            Caption         =   "Specify"
         End
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu cmdDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu cmdRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
      Begin VB.Menu cmdWallpaper 
         Caption         =   "Set As Wallpaper"
         Visible         =   0   'False
      End
      Begin VB.Menu cmdWave 
         Caption         =   "Play WAV"
         Visible         =   0   'False
      End
      Begin VB.Menu cmdDownload 
         Caption         =   "Download"
      End
      Begin VB.Menu Line3 
         Caption         =   "-"
      End
      Begin VB.Menu cmdProperties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuDirectory 
      Caption         =   "Directory Options"
      Visible         =   0   'False
      Begin VB.Menu mnuExplore 
         Caption         =   "Explore Directory"
      End
      Begin VB.Menu Line4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename Directory"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Directory"
      End
      Begin VB.Menu Line5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmFileBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Long
Private lngFileSize As Long
Private lngFileProg As Long

'String
Private File As String
Private FileCount As Long
Private DirName As String

'Boolean
Private ScreenShot As Boolean

Private Sub CmbDrv_Click()
    On Error Resume Next
    Dim I As Integer

    DirName = Left$(CmbDrv.Text, 2) & "\"
    txtDirName.Text = DirName
    w1.SendData "1" & DirName
    For I = 1 To 7
        cmdCommand(I).Enabled = True
    Next I
    txtDirectory.Text = DirName
End Sub

Private Sub cmdBin_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "064" & strPopupText
End Sub

Private Sub cmdCommand_Click(Index As Integer)
    On Error Resume Next
    Dim Command As String

    Select Case Index
        Case 0
        Command = "035"
        cmdCommand(0).Enabled = False
        Case 1
        Command = "045" & Left$(CmbDrv.Text, 2) & "\"
        Case 2
        w1.SendData "1" & DirName

        Case 3
                Err = 0
                CD.ShowOpen
                If Err = 0 Then
                    lngFileProg = 0
                    lngFileSize = FileLen(CD.FileName)
                    pgBar.Max = lngFileSize
                    FT.SendData "SEND_FILE" & DirName & CD.FileTitle & "|" & lngFileSize
                    statBar.SimpleText = "Status: Uploading File..."
                End If
        Case 4
                Command = InputBox("Directory Name", "New Folder", "New Folder")
                If Command = vbNullString Then
                    Exit Sub
                End If
                Command = "027" & DirName & Command
        Case 5
            If Mid$(txtDirectory.Text, 2, 1) = ":" Then
                If Right$(txtDirectory.Text, 1) <> "\" Then
                    DirName = txtDirectory.Text & "\"
                Else
                    DirName = txtDirectory.Text
                End If
                txtDirName.Text = DirName
                w1.SendData "1" & DirName
            End If
        Case 6
        If Len(Text1.Text) > 0 Then
            Command = "060" & Left$(CmbDrv.Text, 2) & "\" & "|" & Text1.Text
        End If
        Case 7
        Command = "059" & Left$(CmbDrv.Text, 1)
            
    End Select
    
    If Len(Command) > 0 Then
        frmClient.wskConnect.SendData Command
        statBar.SimpleText = "Status: Executing Command..."
    End If
End Sub

Private Sub cmdDelete_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "036" & strPopupText
End Sub

Private Sub cmdDownload_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "042" & strPopupText
    statBar.SimpleText = "Status: Downloading File..."
End Sub

Private Sub cmdDrive1_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "061" & strPopupText & "|" & cmdDrive1.Caption & lstFolders.SelectedItem.Text
End Sub

Private Sub cmdDrive2_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "061" & strPopupText & "|" & cmdDrive2.Caption & lstFolders.SelectedItem.Text
End Sub

Private Sub cmdDrive3_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "061" & strPopupText & "|" & cmdDrive3.Caption & lstFolders.SelectedItem.Text
End Sub

Private Sub cmdDrive4_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "061" & strPopupText & "|" & cmdDrive4.Caption & lstFolders.SelectedItem.Text
End Sub

Private Sub cmdDrive5_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "061" & strPopupText & "|" & cmdDrive5.Caption & lstFolders.SelectedItem.Text
End Sub

Private Sub cmdProperties_Click()
    On Error Resume Next
    Dim strNameExtract As String

    strNameExtract = Mid$(strPopupText, InStrRev(strPopupText, "\", Len(strPopupText)) + 1, Len(strPopupText))
    frmProperties.txtRename.Text = strNameExtract
    frmProperties.Caption = strNameExtract & " Properties"
    frmProperties.lblLocation.Caption = Left$(strPopupText, Len(strPopupText) - Len(strNameExtract))
    frmProperties.Icon = ImgList.ListImages.Item(IconPic(Right$(strNameExtract, 3))).Picture
    frmClient.wskConnect.SendData "058" & frmProperties.lblLocation.Caption & "|" & strNameExtract & "|0"
End Sub

Private Sub cmdRename_Click()
    On Error Resume Next
    Dim Command As String

    Command = InputBox("New Name:", "Rename File", Left$(lstFolders.SelectedItem.Text, Len(lstFolders.SelectedItem.Text) - 4))
    If Command = vbNullString Then
        Exit Sub
    End If
    frmClient.wskConnect.SendData "043|" & strPopupText & "|" & DirName & Command & Right$(strPopupText, 4) & "|"
End Sub

Private Sub cmdRun_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "002" & strPopupText
End Sub

Private Sub cmdSpecify_Click()
    On Error Resume Next
    Dim NewLocation As String
    NewLocation = InputBox("Ex: - C:\WINNT\", "Specify Location", "C:\WINNT\")
    If Right$(NewLocation, 1) <> "\" Then
        NewLocation = NewLocation & "\"
    End If
    frmClient.wskConnect.SendData "061" & strPopupText & "|" & NewLocation & lstFolders.SelectedItem.Text
End Sub

Private Sub cmdWallpaper_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "016" & strPopupText
End Sub

Private Sub cmdWave_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "062" & strPopupText
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If frmFileBrowser.Visible = True Then
        Cancel = 1
        frmFileBrowser.Visible = False
    End If
End Sub

Private Sub FT_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim strData As String, FileBuffer As String, strParse() As String
    FT.GetData strData, vbString

    Select Case Left$(strData, 9)
        Case "SEND_FILE"
            strParse = Split(Right$(strData, Len(strData) - 9), "|")
            If strParse(0) <> "ScreenShot" Then
                If strParse(1) = "0" Then
                    Exit Sub
                End If
                CD.FileName = strParse(0)
                CD.ShowSave
            Else
                CD.FileName = App.Path & "\ScreenShot.bmp"
                ScreenShot = True
            End If
            lngFileSize = strParse(1)
            pgBar.Max = Int(strParse(1))
            Erase strParse
            Open CD.FileName For Binary Access Write As #1
            FT.SendData "ACPT_FILE"
            lngFileProg = 0
        Case "ACPT_FILE"
        Open CD.FileName For Binary Access Read As #1
        GoTo SendChunk
        Case "CHNK_FILE"
        GoTo SendChunk
        Case "DONE_FILE"
        Close #1
        CD.FileName = vbNullString
        Case Else
            If (lngFileProg + 4096) < lngFileSize Then
                lngFileProg = lngFileProg + 4096
                pgBar.Value = pgBar.Value + 4096
                Put #1, , strData
                FT.SendData "CHNK_FILE"
            Else
                strData = Left$(strData, lngFileSize - lngFileProg)
                Put #1, , strData
                Close #1
                FT.SendData "DONE_FILE"
                pgBar.Value = 0
                CD.FileName = vbNullString
                If ScreenShot Then
                    Load frmScreenShot
                    ScreenShot = False
                    frmClient.statBar.SimpleText = "Status: Command Executed - (" & Format$(Now, "HH:mm:ss") & ")"
                Else
                    statBar.SimpleText = "Status: Command Executed - (" & Format$(Now, "HH:mm:ss") & ")"
                End If
            End If
    End Select
    Exit Sub
SendChunk:
    lngFileProg = lngFileProg + 4096
    If (lngFileProg + 4096) < lngFileSize Then
        pgBar.Value = pgBar.Value + 4096
    Else
        w1.SendData "1" & DirName
        pgBar.Value = 0
        statBar.SimpleText = "Status: Command Executed - (" & Format$(Now, "HH:mm:ss") & ")"
    End If
    FileBuffer = Space$(4096)
    Get #1, , FileBuffer
    FT.SendData FileBuffer
End Sub

Private Sub lstFolders_DblClick()
    On Error Resume Next
    Dim intForLoop As String, I As Integer
    If lstFolders.ListItems.Count > 0 Then
        Select Case lstFolders.SelectedItem.SmallIcon
            Case 3
                lstFolders.Visible = False
                DirName = DirName & lstFolders.SelectedItem.Text & "\"
                txtDirName.Text = DirName
                w1.SendData "1" & DirName
            Case 4
                lstFolders.Visible = False
                If Len(DirName) > 3 Then
                    If Len(DirName) > 3 Then
                        DirName = Left$(DirName, (Len(DirName) - 1))
                        intForLoop = Len(DirName)
                        For I = 1 To intForLoop
                            If Left$(Right$(DirName, I), 1) = "\" Then
                                DirName = Left$(DirName, (Len(DirName) - I + 1))
                                txtDirectory.Text = DirName
                                txtDirName.Text = DirName
                                Exit For
                            End If
                        Next
                    End If
                End If
                w1.SendData "1" & DirName
            Case Else
                frmClient.wskConnect.SendData "002" & DirName & lstFolders.SelectedItem.Text
        End Select
        txtDirectory.Text = DirName
    End If
End Sub

Private Sub lstFolders_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 Then
        If lstFolders.ListItems.Count > 0 Then
            cmdDownload.Visible = True
            cmdSendTo.Visible = True
            cmdRename.Visible = True
            cmdDelete.Visible = True
            Line8.Visible = True
            Line3.Visible = True
            Line1.Visible = True
            If Len(lstFolders.SelectedItem.Text) > 0 Then
                strPopupText = DirName & lstFolders.SelectedItem.Text
                If lstFolders.SelectedItem.SmallIcon = 3 Then
                    PopupMenu mnuDirectory
                Else
                    cmdWave.Visible = False
                    cmdWallpaper.Visible = False
                    If Right$(strPopupText, 4) = ".bmp" Then
                        cmdWallpaper.Visible = True
                    ElseIf Right$(strPopupText, 4) = ".wav" Then
                        cmdWave.Visible = True
                    End If
                    PopupMenu mnuFile
                End If
            End If
        End If
    End If
End Sub

Private Sub mnuDelete_Click()
    On Error Resume Next
    frmClient.wskConnect.SendData "026" & strPopupText
End Sub

Private Sub mnuExplore_Click()
    On Error Resume Next
    DirName = strPopupText & "\"
    txtDirName.Text = DirName
    txtDirectory.Text = DirName
    w1.SendData "1" & DirName
End Sub

Private Sub mnuProperties_Click()
    On Error Resume Next
    Dim strNameExtract As String

    strNameExtract = Mid$(strPopupText, InStrRev(strPopupText, "\", Len(strPopupText)) + 1, Len(strPopupText))
    frmProperties.txtRename.Text = strNameExtract
    frmProperties.Caption = strNameExtract & " Properties"
    frmProperties.lblLocation.Caption = Left$(strPopupText, Len(strPopupText) - Len(strNameExtract))
    frmProperties.Icon = ImgList.ListImages.Item(3).Picture
    frmClient.wskConnect.SendData "058" & frmProperties.lblLocation.Caption & "|" & strNameExtract & "|1"
End Sub

Private Sub mnuRename_Click()
    On Error Resume Next
    Dim Command As String
    Command = InputBox("New Name:", "Rename Directory", lstFolders.SelectedItem.Text)
    If Command = vbNullString Then
        Exit Sub
    End If
    frmClient.wskConnect.SendData "043|" & strPopupText & "|" & DirName & Command & "|"
End Sub

Private Sub w1_Connect()
    On Error Resume Next
    w1.SendData "0"
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim Data As String, Data2 As String, I As Integer, intForLoop As Integer
    
    w1.GetData Data, vbString
    
    If Left$(Data, 2) = "1|" Or Left$(Data, 2) = "0|" Then
        lstFolders.ListItems.Clear
        FileCount = 0
        If Len(DirName) > 3 Then
            FileCount = FileCount + 1
            lstFolders.ListItems.Add.Text = "<..>"
            lstFolders.ListItems.Item(FileCount).SmallIcon = 4
        End If
        Data2 = Right$(Data, Len(Data) - 2)
    Else
        Data2 = Data
    End If
    intForLoop = Len(Data2)
    For I = 1 To intForLoop
        If Mid$(Data2, I, 1) = "|" Then
            If Left$(Data, 1) = "0" Then
                CmbDrv.AddItem File
                If Right$(File, 5) = "FIXED" Then
                    If cmdDrive1.Visible = False Then
                        cmdDrive1.Visible = True
                        cmdDrive1.Caption = Left$(File, 2) & "\"
                    ElseIf cmdDrive2.Visible = False Then
                        cmdDrive2.Visible = True
                        cmdDrive2.Caption = Left$(File, 2) & "\"
                    ElseIf cmdDrive3.Visible = False Then
                        cmdDrive3.Visible = True
                        cmdDrive3.Caption = Left$(File, 2) & "\"
                    ElseIf cmdDrive4.Visible = False Then
                        cmdDrive4.Visible = True
                        cmdDrive4.Caption = Left$(File, 2) & "\"
                    ElseIf cmdDrive5.Visible = False Then
                        cmdDrive5.Visible = True
                        cmdDrive5.Caption = Left$(File, 2) & "\"
                    End If
                    CmbDrv.SetFocus
                End If
            Else
                FileCount = FileCount + 1
                If Left$(File, 1) <> "?" Then
                    File = Mid$(File, InStrRev(File, "\", Len(File)) + 1, Len(File))
                    lstFolders.ListItems.Add.Text = File
                    lstFolders.ListItems.Item(FileCount).SmallIcon = IconPic(Right$(File, 3))
                Else
                    File = Mid$(File, InStrRev(File, "\", Len(File)) + 1, Len(File))
                    lstFolders.ListItems.Add.Text = Right$(File, Len(File) - 1)
                    lstFolders.ListItems.Item(FileCount).SmallIcon = 3
                End If
            End If
            File = vbNullString
        Else
            File = File & Mid$(Data2, I, 1)
        End If
    Next
    statBar.SimpleText = "Status: " & lstFolders.ListItems.Count + IIf(Len(DirName) > 3, -1, 0) & " object(s)"
    lstFolders.Visible = True
End Sub

