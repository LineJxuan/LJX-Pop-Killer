VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "软件检测报告"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8445
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8445
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "病毒检测"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "性能占用检测"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   8415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "检测项目类型："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8400
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8400
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "以下是LJX弹窗杀手的检测报告："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Text1.Text = "以下是检测电脑的配置：" & vbCrLf & "CPU: Intel Core i5-9300H @2.4GHz" & vbCrLf & "显卡： Intel UHD Graphics 630（核显）" & vbCrLf & "内存：海力士DDR4 2666MHz 16GB(8GB×2)" & vbCrLf & "磁盘：Samsung SSD 970 EVO Plus 1TB" & vbCrLf & "系统：Microsoft Windows 10 家庭中文版21H1(19043.1706)" & vbCrLf & vbCrLf & "以下是性能占用检测结果(拦截一个弹窗)：" & vbCrLf & "CPU占用（毁灭模式，发现的最大值）： 15.8%" & vbCrLf & "CPU占用（超级模式,发现的最大值）: 2.3%" & vbCrLf & "CPU占用（极低模式，发现的最大值）： 1.5% " & vbCrLf & "CPU占用（仅删除模式,发现的占用最大值）：0.8%" & vbCrLf & "内存占用：（预估的最大值）3 MB"
End Sub

Private Sub Command3_Click()
Text1.Text = "以下是病毒检测情况（2021年4月12日检测）：" & vbCrLf & "360安全卫士13.1.0.1001：无病毒" & vbCrLf & "360杀毒5.0.0.8170：无病毒" & vbCrLf & "火绒安全软件5.0.62.4：对LJX弹窗杀手控制面板报毒(Trojan/VBCode.aj)" & vbCrLf & "Windows Defender ：无病毒" & vbCrLf & vbCrLf & "VirSCAN.org(2021年8月17日 上午 11:59检测) ：" & vbCrLf & "    对LJX弹窗杀手控制面板：0%（无杀毒软件报毒）" & vbCrLf & "    对LJX弹窗杀手启动程序：2% (Vba32报毒) " & vbCrLf & "    对LJX弹窗杀手托盘程序：3%(IKARUS、安天 报毒)" & vbCrLf & "    对LJX弹窗杀手运行验证程序：0%(无杀毒软件报毒)" & vbCrLf & "    对LJX弹窗杀手主程序：0%(无杀毒软件报毒)"
End Sub

Private Sub Form_Load()
Form1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub
