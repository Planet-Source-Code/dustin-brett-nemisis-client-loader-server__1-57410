Attribute VB_Name = "modFunctions"
Option Explicit

'Function - USER32.DLL
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

'String
Public strWindows As String

Public Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lParam As Integer) As Boolean
    On Error Resume Next
    Dim lReturn As Long, lExStyle As Long, bNoOwner As Boolean, sWindowText As String
    If hwnd <> frmServer.hwnd Then
        If IsWindowVisible(hwnd) Then
            If GetParent(hwnd) = 0 Then
                bNoOwner = (GetWindow(hwnd, 4) = 0)
                lExStyle = GetWindowLong(hwnd, (-20))
                If (((lExStyle And &H80) = 0) And bNoOwner) Or ((lExStyle And &H40000) And Not bNoOwner) Then
                    sWindowText = Space$(256)
                    lReturn = GetWindowText(hwnd, sWindowText, Len(sWindowText))
                    If lReturn Then
                        sWindowText = Left$(sWindowText, lReturn)
                    End If
                    strWindows = strWindows & sWindowText & "|"
                End If
            End If
        End If
    End If
    fEnumWindowsCallBack = True
End Function


