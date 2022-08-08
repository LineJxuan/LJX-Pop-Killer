VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "LJX弹窗杀手-主程序"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   4410
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
 Const NIM_ADD = &H0
 Const NIM_DELETE = &H2
 Const NIF_ICON = &H2
 Const NIF_MESSAGE = &H1
 Const NIF_TIP = &H4
 Const WM_MOUSEMOVE = &H200
 Const WM_LBUTTONDBLCLK = &H203
Private Type NOTIFYICONDATA
cbSize As Long
hWnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Dim tray As NOTIFYICONDATA
Function Icon_Del() As Long
Dim IconVa As NOTIFYICONDATA
Dim L As Long
With IconVa
.hWnd = iHwnd
.uId = lIndex
.cbSize = Len(IconVa)
End With
Icon_Del = Shell_NotifyIcon(NIM_DELETE, IconVa)
End Function

Public Function LoadAll()
On Error GoTo Errt
Dim q
Dim Stri As String
Dim r As Long
Dim X As Long
Dim Y As Long
r = 0

For X = 0 To 1023
    PopsURL(X) = ""
    Pops(X) = ""
Next

For X = 0 To 1023
    Stri = ""
    q = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx")
    If q <> "" Then
        Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx") For Input As #1
        Line Input #1, Stri
        If Stri <> "" Then
            PopsURL(r) = Stri
            MaxPops = MaxPops + 1
        End If
        Close #1
        Dim t As Long
        Dim tmp As String
        t = Len(PopsURL(r))
        For Y = 1 To t
            tmp = Mid(PopsURL(r), (t - Y), 1)
            If tmp = "\" Then
                Exit For
            End If
        Next
        Pops(r) = Mid(PopsURL(r), (t - Y + 1), Y)
        r = r + 1
    End If
Next

If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode1.ltmp") <> "" Then
    Timer1.Interval = 1
    Mode = 1
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode2.ltmp") <> "" Then
    Timer1.Interval = 1
    Mode = 2
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode3.ltmp") <> "" Then
    Timer1.Interval = 50
    Mode = 3
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode4.ltmp") <> "" Then
    Timer1.Interval = 150
    Mode = 4
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode5.ltmp") <> "" Then
    Timer1.Interval = 500
    Mode = 5
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode6.ltmp") <> "" Then
    Timer1.Interval = 1000
    Mode = 6
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode7.ltmp") <> "" Then
    Timer1.Interval = 5000
    Mode = 7
End If
Exit Function
Errt:
MsgBox ("LJX弹窗杀手启动错误：" & Err.Number & vbCrLf & Err.Description)
End
End Function
Private Sub Form_Load()
On Error GoTo Errt

Me.Hide
tray.cbSize = Len(tray)
tray.uId = vbNull
tray.hWnd = Me.hWnd
tray.uFlags = NIF_TIP Or NIF_MESSAGE Or NIF_ICON
tray.uCallBackMessage = WM_MOUSEMOVE
tray.hIcon = Me.Icon
tray.szTip = "LJX弹窗杀手 - 正在全面拦截弹窗" & vbCrLf & "双击以进入弹窗杀手控制面板" & vbNullChar
Shell_NotifyIcon NIM_ADD, tray


If Dir("C:\Users\" & Environ("USERNAME") & "\AppData\Roaming\LJXPopKiller", vbDirectory + vbSystem) = "" Then
    MsgBox ("请先运行“LJX弹窗杀手控制面板.exe”！")
    Shell_NotifyIcon NIM_DELETE, tray
    End
End If
If Dir("C:\Users\" & Environ("USERNAME") & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Start.ltmp") = "" Then
    MsgBox ("要启动请运行“LJX弹窗杀手启动程序.exe”！")
    Shell_NotifyIcon NIM_DELETE, tray
End
End If
If Dir("C:\Users\" & Environ("USERNAME") & "\AppData\Roaming\LJXPopKiller\Inf\Act.prove", vbSystem + vbHidden + vbReadOnly) = "" Then
    Call MsgBox("这是一个没有被激活的LJX弹窗杀手！请打开“LJX弹窗杀手控制面板”以进行激活。", vbCritical, "软件未激活")
    Shell_NotifyIcon NIM_DELETE, tray
    End
End If
MaxPops = 0
Me.Hide
MyName = Environ("USERNAME")
Timer1.Enabled = False
Call LoadAll

If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode1.ltmp") <> "" Then
    Timer1.Interval = 1
    Mode = 1
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode2.ltmp") <> "" Then
    Timer1.Interval = 1
    Mode = 2
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode3.ltmp") <> "" Then
    Timer1.Interval = 50
    Mode = 3
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode4.ltmp") <> "" Then
    Timer1.Interval = 150
    Mode = 4
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode5.ltmp") <> "" Then
    Timer1.Interval = 500
    Mode = 5
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode6.ltmp") <> "" Then
    Timer1.Interval = 1000
    Mode = 6
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode7.ltmp") <> "" Then
    Timer1.Interval = 1500
    Mode = 7
End If
If Dir("C:\Users\" & Environ("USERNAME") & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Start.ltmp") <> "" Then
    Kill ("C:\Users\" & Environ("USERNAME") & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Start.ltmp")
End If
Timer1.Enabled = True
Exit Sub
Errt:
Shell_NotifyIcon NIM_DELETE, tray
Call MsgBox("LJX弹窗杀手错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
LoadAll
If Mode = 7 Then
    For X = 0 To 1023
        If Dir(PopsURL(X)) <> "" Then
            Kill (PopsURL(X))
        End If
    Next
Else
    L = 0
    For X = 0 To 1023
        r = Pops(X)
        If r <> "" Then
            L = L + 1
        End If
    Next
    For X = 0 To L
        Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
        s = Pops(X)
        Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='" & s & "'")
        For Each objProcess In colProcessList
            objProcess.Terminate
        Next
    Next
    Set objProcess = Nothing
    Set colProcessList = Nothing
    Set objWMIService = Nothing
    If Mode = 1 Then
        For X = 0 To 1023
            If Dir(PopsURL(X)) <> "" Then
                Kill (PopsURL(X))
            End If
        Next
    End If
End If
If Dir("C:\Users\" & Environ("USERNAME") & "\AppData\Roaming\LJXPopKiller\Inf\Temp\EndAll.ltmp") <> "" Then
    Shell_NotifyIcon NIM_DELETE, tray
    End
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
msg = X / 15
If msg = WM_LBUTTONDBLCLK Then
If Dir(App.Path & "\LJX弹窗杀手控制面板.exe") = "" Then
    Call MsgBox("LJX弹窗杀手的文件LJX弹窗杀手控制面板.exe已经缺失！", vbOKOnly + vbCritical)
    Exit Sub
End If
Shell (App.Path & "\LJX弹窗杀手控制面板.exe")
End If
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, tray
End
End Sub


