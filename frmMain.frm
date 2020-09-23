VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "processor usage"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4290
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   3360
      Top             =   600
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "00 %"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long


Private Const REG_DWORD = 4
Private Const HKEY_DYN_DATA = &H80000006

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub InitCPU()
    Dim lData As Long, lType As Long, lSize As Long
    Dim hKey As Long
    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StartStat", hKey)
    If Qry <> 0 Then
          MsgBox "Error reading registry", vbCritical, "Error"
          End
    End If
    lType = REG_DWORD
    lSize = 4
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
    Qry = RegCloseKey(hKey)
End Sub

Private Sub OnTop()
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    If SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS) = True Then
        success% = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

Private Sub Form_Load()
    If IsNT Then
        MsgBox "Windows NT ist not supported !", vbCritical, "Error"
        End
    End If
    Call InitCPU
    Call OnTop
End Sub

Private Sub Timer1_Timer()
    Dim lData As Long, lType As Long, lSize As Long
    Dim hKey As Long
    Qry = RegOpenKey(HKEY_DYN_DATA, "PerfStats\StatData", hKey)
    If Qry <> 0 Then
        MsgBox "Error reading registry", vbCritical, "Error"
        End
    End If
    lType = REG_DWORD
    lSize = 4
    Qry = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, lType, lData, lSize)
    v = Trim(Str(Int(lData)))
    ProgressBar1.Value = v
    Label1.Caption = v & " %"
    Qry = RegCloseKey(hKey)
End Sub
