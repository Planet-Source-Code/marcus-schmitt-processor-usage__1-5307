Attribute VB_Name = "modOS"
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
  
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Dim OS As OSVERSIONINFO


Function IsNT()
    OS.dwOSVersionInfoSize = Len(OS)
    GetVersionEx OS
    If OS.dwMajorVersion = 4 And OS.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        IsNT = True
    Else
        IsNT = False
    End If
End Function

