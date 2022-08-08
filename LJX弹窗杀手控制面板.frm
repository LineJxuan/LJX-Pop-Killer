VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LJX弹窗杀手-控制面板"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9015
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   4920
      Top             =   5640
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   6000
      Top             =   5640
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "软件检测报告"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "查看LJX弹窗杀手的第三方检测报告"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "强制重启LJX弹窗杀手"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "强制重启正在运行的LJX弹窗杀手"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "强制停止LJX弹窗杀手"
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "强制停止正在运行的LJX弹窗杀手"
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   6480
      Top             =   5520
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080FFFF&
      Caption         =   "重启LJX弹窗杀手"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "重新启动LJX弹窗杀手"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FF00FF&
      Caption         =   "检测启动情况"
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "检测LJX弹窗杀手的启动情况"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Caption         =   "启动LJX弹窗杀手"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "启动LJX弹窗杀手"
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF00FF&
      Caption         =   "添加到开机启动"
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "将LJX弹窗杀手添加到开机启动项"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "拦截性能设置(选择合适的模式来适应电脑的性能)："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   9015
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "仅删除（只删除弹窗的程序，没有任何影响）"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   6135
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "极低模式  (没有任何影响，但是有拦截延迟)  "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   6855
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "低性能模式  (对电脑性能几乎没有影响，会有拦截延迟)  "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   6735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "普通模式  (对电脑性能会有微小的影响，会有拦截延迟)  "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   6495
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "强力模式  (可能会对电脑性能有一定影响，会有微小的延迟)  "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   7335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "超级模式  (可能会对电脑性能有较多影响，几乎没有延迟)  "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   6975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "毁灭模式  (会对电脑性能有较多影响、会直接粉碎弹窗的程序)  "
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "关于LJX弹窗杀手"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "显示关于LJX弹窗杀手的信息"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "设置要拦截的弹窗"
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "查看当前被设置为“拦截”的弹窗"
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "停止LJX弹窗杀手的运行"
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "使正在运行的LXJ弹窗杀手停止运行"
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "关闭此控制面板"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "保存设定并关闭控制面板"
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   8400
      Picture         =   "LJX弹窗杀手控制面板.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "最小化当前窗口"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   8400
      Picture         =   "LJX弹窗杀手控制面板.frx":0A26
      Stretch         =   -1  'True
      ToolTipText     =   "最小化当前窗口"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   8400
      Picture         =   "LJX弹窗杀手控制面板.frx":144C
      Stretch         =   -1  'True
      ToolTipText     =   "最小化当前窗口"
      Top             =   0
      Width           =   600
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   9000
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "当前有？？？？个弹窗被拦截"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2
Private ExeNumber
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Timer1.Enabled = True
Else
    Timer1.Enabled = False
End If
End Sub

Private Sub Command11_Click()
Form4.Show
End Sub

Private Sub Command2_Click()
On Error GoTo Errt
Dim a
a = MsgBox("确定要强制停止LJX弹窗杀手吗？这可能会导致意外事故。", vbOKCancel, "系统提示")
If a = vbOK Then
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='LJX弹窗杀手主程序.exe'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next

    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='LJX弹窗杀手启动程序.exe'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next
    Set objProcess = Nothing
    Set colProcessList = Nothing
    Set objWMIService = Nothing
    MsgBox ("停止成功！")
End If
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("错误：" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command3_Click()
On Error GoTo Errt
Dim a
a = MsgBox("确定要强制重启LJX弹窗杀手吗？这可能会导致意外事故。", vbOKCancel, "系统提示")
If a = vbOK Then
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='LJX弹窗杀手主程序.exe'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next

    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='LJX弹窗杀手启动程序.exe'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next
    Set objProcess = Nothing
    Set colProcessList = Nothing
    Set objWMIService = Nothing
    Call GetProgramRunState
    Shell (App.Path & "\LJX弹窗杀手启动程序.exe")
    MsgBox ("重启成功！")
End If
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("错误：" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Visible = False
Image3.Visible = False
End Sub

Private Sub Timer1_Timer()
Shell (App.Path & "\LJX弹窗杀手启动程序.exe")
Form1.Enabled = True
Form1.Caption = "LJX弹窗杀手-控制面板"
Timer1.Enabled = False
End Sub

Private Sub Command1_Click()
On Error GoTo Errt
i = MsgBox("确定要关闭LJX弹窗杀手控制面板吗？(所有设定将会保存)", vbOKCancel, "系统提示")
If i = vbOK And Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp") <> "" Then
    Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp")
    End
End If
Exit Sub
Errt:
MsgBox ("关闭错误：" & Err.Number & vbCrLf & Err.Description)
End
End Sub

Private Sub Command10_Click()
On Error GoTo Errt
i = MsgBox("确定要重启LJX弹窗杀手吗？", vbOKCancel, "系统提示")
If i = vbOK Then
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\EndAll.ltmp") For Output As #1
    Close #1
    Form1.Enabled = False
    Form1.Caption = "LJX弹窗杀手-控制面板   [正在重启LJX弹窗杀手]"
    Timer1.Interval = 6000
    Timer1.Enabled = True
End If
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("错误：" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command4_Click()
On Error GoTo Errt
i = MsgBox("确定停止LJX弹窗杀手的运行吗？停止后LJX弹窗杀手将无法再为你拦截弹窗！", vbOKCancel, "系统提示")
If i = vbOK Then
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\EndAll.ltmp") For Output As #1
    Close #1
    MsgBox ("LJX弹窗杀手停止成功！")
End If
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("错误：" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command5_Click()
LoadAll
Form2.Show
End Sub

Private Sub Command6_Click()
frmAbout.Show
End Sub

Private Sub Command7_Click()
On Error GoTo Errt

Dim a
Dim b
Dim c
Dim d
Dim e
Dim f

Dim i
i = MsgBox("确定要将LJX弹窗杀手添加到开机启动吗？", vbOKCancel, "系统提示")
If i = vbCancel Then
    Exit Sub
End If
a = Dir(App.Path & "\LJX弹窗杀手控制面板.exe")
'b = Dir(App.Path & "\LJX弹窗杀手运行验证程序.exe")
'c = Dir(App.Path & "\LJX弹窗杀手托盘程序.exe")
d = Dir(App.Path & "\LJX弹窗杀手主程序.exe")
e = Dir(App.Path & "\LJX弹窗杀手启动程序.exe")
If a = "" Or d = "" Or e = "" Then
    f = MsgBox("LJX弹窗杀手的文件不完整，无法添加注册表启动项！", vbOKOnly, "无法添加启动项！")
    Exit Sub
End If
If i = vbOK Then
    Set w = CreateObject("wscript.shell")
    w.regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & "LJX弹窗杀手", App.Path & "\LJX弹窗杀手启动程序.exe"
End If
MsgBox ("开机启动项添加成功！")
Exit Sub
Errt:
MsgBox ("错误：" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command8_Click()
On Error GoTo Errt
Shell (App.Path & "\LJX弹窗杀手启动程序.exe")
Exit Sub
Errt:
MsgBox ("错误:" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command9_Click()
Dim a As Boolean
Dim b As Boolean
Dim c As Boolean
Dim q
Call GetProgramRunState
'a = CheckExeIsRun("LJX弹窗杀手运行验证程序.exe")
'b = CheckExeIsRun("LJX弹窗杀手托盘程序.exe")
c = CheckExeIsRun("LJX弹窗杀手主程序.exe")

If c = True Then
    Call MsgBox("程序已经打开！", vbInformation, "程序已打开")
Else
    Call MsgBox("程序并没有打开！", vbInformation, "程序未打开")
End If
End Sub

Private Sub Form_Load()
On Error GoTo Errt
Image2.Visible = False
Image3.Visible = False
Timer1.Enabled = False
WindowState = vbNormal
Form1.Label1.ForeColor = &H0&
Call GetProgramRunState
'
Exit Sub
Errt:
Call MsgBox("启动时窗口Form1_Load初始化错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Visible = False
Image3.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
For x = 0 To 10000
    Cancel = -1
Next
Cancel = -1
Cancel = -1
Cancel = -1
Cancel = -1
Cancel = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
For x = 0 To 10000
    Cancel = -1
Next
Cancel = -1
Cancel = -1
Cancel = -1
Cancel = -1
Cancel = -1
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If x >= 0 And x <= Image1.Width And Y >= 0 And Y <= Image1.Height Then
    Image2.Visible = True
Else
    Image2.Visible = False
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Not (x >= 0 And x <= Image1.Width And Y >= 0 And Y <= Image1.Height) Then
    Image2.Visible = False
End If
End Sub

Private Sub Image2_Click()
WindowState = vbMinimized
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image3.Visible = True
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Image3.Visible = True And (x >= 0 And x <= Image2.Width And Y >= 0 And Y <= Image2.Height) Then
    WindowState = vbMinimized
End If
If Not (x >= 0 And x <= Image1.Width And Y >= 0 And Y <= Image1.Height) Then
    Image3.Visible = False
End If
End Sub

Private Sub Image3_Click()
WindowState = vbMinimized
End Sub

Private Sub Option1_Click()
On Error GoTo Errt:
If Option1.Value = True And Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode1.ltmp") = "" Then
    a = MsgBox("你确定要启动“毁灭模式”吗？这会删除弹窗的程序！", vbOKCancel + vbExclamation, "警告")
    If a = vbOK Then
        If Option1.Value = True Then
            Dim x
            For x = 1 To 7
                If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp") <> "" Then
                    Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp")
                End If
            Next
            Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode1.ltmp") For Output As #1
            Close #1
        End If
    Else
         Option1.Value = False
         Call LoadModes
    End If
End If
Exit Sub
Errt:
Call MsgBox("更改失败：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option2_Click()
On Error GoTo Errt
If Option2.Value = True Then
    Dim x
    For x = 1 To 7
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp")
        End If
    Next
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode2.ltmp") For Output As #1
    Close #1
End If
Exit Sub
Errt:
Call MsgBox("更改失败：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option3_Click()
On Error GoTo Errt
If Option3.Value = True Then
    Dim x
    For x = 1 To 7
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp")
        End If
    Next
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode3.ltmp") For Output As #1
    Close #1
End If
Exit Sub
Errt:
Call MsgBox("更改失败：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option4_Click()
On Error GoTo Errt
If Option4.Value = True Then
    Dim x
    For x = 1 To 7
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp")
        End If
    Next
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode4.ltmp") For Output As #1
    Close #1
End If
Exit Sub
Errt:
Call MsgBox("更改失败：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option5_Click()
On Error GoTo Errt
If Option5.Value = True Then
    Dim x
    For x = 1 To 7
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp")
        End If
    Next
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode5.ltmp") For Output As #1
    Close #1
End If
Exit Sub
Errt:
Call MsgBox("更改失败：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option6_Click()
On Error GoTo Errt
If Option6.Value = True Then
    Dim x
    For x = 1 To 7
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp")
        End If
    Next
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode6.ltmp") For Output As #1
    Close #1
End If
Exit Sub
Errt:
Call MsgBox("更改失败：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option7_Click()
On Error GoTo Errt
If Option7.Value = True And Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode7.ltmp") = "" Then
    a = MsgBox("你确定要启动“仅删除”吗？这会删除弹窗的程序！", vbOKCancel + vbExclamation, "警告")
    If a = vbOK Then
        If Option7.Value = True Then
            Dim x
            For x = 1 To 73
                If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp") <> "" Then
                    Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode" & x & ".ltmp")
                End If
            Next
            Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode7.ltmp") For Output As #1
            Close #1
        End If
    Else
        Option7.Value = False
        Call LoadModes
    End If
End If
Exit Sub
Errt:
Call MsgBox("更改失败：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp") = "" Then
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp") For Output As #1
    Close #1
End If
Exit Sub
End Sub

Private Sub Timer3_Timer()
Call GetProgramRunState
End Sub
