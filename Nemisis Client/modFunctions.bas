Attribute VB_Name = "modFunctions"
Option Explicit

Public strPopupText As String

Public Function IconPic(ByVal FileName As String) As Integer
    On Error Resume Next
    Select Case LCase$(FileName)
        Case "dll"
        IconPic = 2
        Case "sys"
        IconPic = 2
        Case "vxd"
        IconPic = 2
        Case "cpl"
        IconPic = 2
        Case "lnk"
        IconPic = 5
        Case "exe"
        IconPic = 6
        Case "com"
        IconPic = 6
        Case "bat"
        IconPic = 7
        Case "scr"
        IconPic = 8
        Case "avi"
        IconPic = 9
        Case "asx"
        IconPic = 9
        Case "asf"
        IconPic = 9
        Case "vob"
        IconPic = 9
        Case "mov"
        IconPic = 10
        Case "mpa"
        IconPic = 11
        Case "mpe"
        IconPic = 11
        Case "mpg"
        IconPic = 11
        Case "m1v"
        IconPic = 11
        Case "m2v"
        IconPic = 11
        Case "wma"
        IconPic = 11
        Case "wmv"
        IconPic = 11
        Case "cda"
        IconPic = 12
        Case "mp1"
        IconPic = 13
        Case "mp2"
        IconPic = 13
        Case "mp3"
        IconPic = 13
        Case "m3u"
        IconPic = 13
        Case "wav"
        IconPic = 14
        Case "voc"
        IconPic = 15
        Case "mid"
        IconPic = 15
        Case "bmp"
        IconPic = 16
        Case "gif"
        IconPic = 17
        Case "jpg"
        IconPic = 18
        Case "pcx"
        IconPic = 19
        Case "tif"
        IconPic = 19
        Case "pdf"
        IconPic = 20
        Case "psd"
        IconPic = 21
        Case "fon"
        IconPic = 22
        Case "rtf"
        IconPic = 23
        Case "doc"
        IconPic = 23
        Case "ini"
        IconPic = 24
        Case "inf"
        IconPic = 24
        Case "css"
        IconPic = 24
        Case "ttf"
        IconPic = 25
        Case "txt"
        IconPic = 25
        Case "dat"
        IconPic = 25
        Case "log"
        IconPic = 25
        Case "cfg"
        IconPic = 25
        Case "nfo"
        IconPic = 25
        Case "vbs"
        IconPic = 26
        Case "vbe"
        IconPic = 26
        Case "jse"
        IconPic = 27
        Case "js"
        IconPic = 27
        Case "htm"
        IconPic = 28
        Case "url"
        IconPic = 28
        Case "reg"
        IconPic = 29
        Case "key"
        IconPic = 29
        Case "hlp"
        IconPic = 30
        Case "chm"
        IconPic = 30
        Case "cab"
        IconPic = 31
        Case "jar"
        IconPic = 31
        Case "rar"
        IconPic = 32
        Case "zip"
        IconPic = 33
        Case "par"
        IconPic = 34
        Case Else
        IconPic = 1
    End Select
End Function

