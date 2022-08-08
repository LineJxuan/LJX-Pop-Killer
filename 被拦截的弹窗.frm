VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "被拦截的弹窗"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10875
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   10875
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "更改项目"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "删除项目"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "添加项目"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "关闭窗口"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "关闭并保存更改"
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "被拦截的弹窗.frx":0000
      Top             =   600
      Width           =   10815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "刷新"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7320
      TabIndex        =   7
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "如果删除按钮不见了就点我"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   7440
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "这里是被设置为“拦截”的弹窗"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Form3.Show
End Sub

Private Sub Command3_Click()
On Error GoTo Errt
Form2.SetFocus
Dim a
Dim b
a = InputBox("请输入要删除的项目的编号。", "删除项目")
If a <> Cancel Then
    b = MsgBox("确定要删除编号为" & a & "号的项目吗？", vbOKCancel, "删除确认")
    Form2.SetFocus
    If b = vbOK Then
        Call DelFile(a)
    End If
    Call LoadAll
    Unload Form2
    Load Form2
    Call Form2.Refresh
    Form2.Show
End If
Exit Sub
Errt:
Call MsgBox("项目删除错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub
Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
On Error GoTo Errt
Dim a
Dim b
Dim c
a = InputBox("请输入要更改的项目编号。")
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & a & ".ltx") = "" Then
    MsgBox ("所选项目不存在！")
    Exit Sub
End If
b = InputBox("你要将该项目的新数据修改为什么？(输入新路径)")
c = MsgBox("你确定要将第" & a & "号的项目修改为" & b & "？", vbOKCancel, "系统提示")
If c = vbOK Then
    Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & a & ".ltx")
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & a & ".ltx") For Output As #1
    Print #1, b
    Close #1
    Unload Form2
    LoadAll
    Form2.Show
End If
Exit Sub
Errt:
Call MsgBox("项目编号修改错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Form_Load()
On Error GoTo Errt
Form1.Hide
Text1.Text = ""
Form1.Enabled = False
For x = 0 To 1023
    If Pops(x) <> "" Then
        Text1.Text = Text1.Text & x & "、" & Pops(x) & vbCrLf
    End If
Next
If Text1.Text = "" Then
    Text1.Text = "这个列表是空的！"
End If
Exit Sub
Errt:
Call MsgBox("加载Form2\Form_Load错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
Form1.Show
End Sub

Private Sub Label2_Click()
Command3.Value = True
End Sub

Private Sub Label3_Click()
Unload Form2
LoadAll
Form2.Show
End Sub

Private Sub Timer1_Timer()

End Sub
