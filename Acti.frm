VERSION 5.00
Begin VB.Form Acti 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "激活 LJX弹窗杀手"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5565
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5565
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command3 
      Caption         =   "显示密钥"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   1005
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5040
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "取消"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "确定"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "请输入LJX提供的序列号(密钥)"
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
      Left            =   725
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "Acti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Open ("C:\Users\admin\AppData\Roaming\LJXPopKiller\Inf\Act.prove") For Output As #1
'    Print #1, (Mid(Key, 1, 7) & Mid(Key(11, 5)))
'    Close #1
'    Call SetAttr("C:\Users\admin\AppData\Roaming\LJXPopKiller\Inf\Act.prove", vbSystem + vbHidden + vbReadOnly)

Private Sub Command1_Click()
On Error GoTo Errt
Dim Key As String
For x = 1 To Len(Key)
    If Val(Mid(Key, x, 1)) = 0 And x <> 1 Then
        GoTo ErrProve
    End If
Next
Key = Key & Mid(Str(Mid(Text1.Text, 2, 5)), 2, 5)
Key = Key & Mid(Str(Mid(Text1.Text, 7, 5)), 2, 5)
Key = Key & Mid(Str(Mid(Text1.Text, 12, 5)), 2, 5)
Key = Key & Mid(Str(Mid(Text1.Text, 17, 5)), 2, 5)
Key = Key & Mid(Str(Mid(Text1.Text, 22, 5)), 2, 5)
Key = Key & Mid(Str(Mid(Text1.Text, 27, 2)), 2, 2)
Dim NowTime
NowTime = Mid(Str(Format(Now, "yyyymmddhhmmss")), 2, 14)
If Len(Key) <> 27 Then
    GoTo ErrProve
Else
    Dim tArray(0 To 13) As String
    Dim rKey As String
    rKey = ""
    For x = 0 To 13
        tArray(x) = Mid(NowTime, x + 1, 1)
    Next
    For x = 0 To 13
        If x = 0 Then
            rKey = tArray(x)
        Else
            rKey = rKey & tArray(x) + tArray(x - 1)
        End If
    Next
    Dim fx
    Dim mx
    Dim lx
    For x = 1 To 27
        If x Mod 2 = 1 Then
            fx = Mid(Key, x, 1)
            If x <> 27 Then
                lx = Mid(Key, x + 2, 1)
                mx = Val(fx) + Val(lx)
                If Len(mx) = 2 Then
                    mx = Mid(mx, 1, 1)
                End If
                If mx <> Val(Mid(Key, x + 1, 1)) Then
                    GoTo ErrProve
                End If
            End If
        End If
    Next
    If rKey <> Key Then
        Dim tKey As String
        tKey = ""
        For x = 1 To 35
            If x Mod 2 = 1 Then
                tKey = tKey & Mid(Key, x, 1)
            End If
        Next
        Dim nYear, kYear
        Dim nMonth, kMonth
        Dim nDay, kDay
        Dim nHour, kHour
        Dim nMinute, kMinute
        Dim nSecond, kSecond
        nYear = Mid(NowTime, 1, 4): kYear = Mid(tKey, 1, 4)
        nMonth = Mid(NowTime, 5, 2): kMonth = Mid(tKey, 5, 2)
        nDay = Mid(NowTime, 7, 2): kDay = Mid(tKey, 7, 2)
        nHour = Mid(NowTime, 9, 2): kHour = Mid(tKey, 9, 2)
        nMinute = Mid(NowTime, 11, 2): kMinute = Mid(tKey, 11, 2)
        nSecond = Mid(NowTime, 13, 2): kSecond = Mid(tKey, 13, 2)
        If nYear > kYear Then
            GoTo OldKey
        ElseIf nYear < kYear Then
            GoTo FinishAction
        ElseIf nYear = kYear Then
            If nMonth > kMonth Then
                GoTo OldKey
            ElseIf nMonth < kMonth Then
                GoTo FinishAction
            ElseIf nMonth = kMonth Then
                If nDay > kDay Then
                    GoTo OldKey
                ElseIf nDay < kDay Then
                    GoTo FinishAction
                ElseIf nDay = kDay Then
                    If nHour > kHour Then
                        GoTo OldKey
                    ElseIf nHour < kHour Then
                        GoTo FinishAction
                    ElseIf nHour = kHour Then
                        If nMinute > kMinute Then
                            GoTo OldKey
                        ElseIf nMinute < kMinute Then
                            GoTo FinishAction
                        ElseIf nMinute = kMinute Then
                            If nSecond >= kSecond Then
                                GoTo OldKey
                            ElseIf nSecond < kSecond Then
                                GoTo FinishAction
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Else
        GoTo ErrProve
    End If
End If

GoTo ErrProve
Exit Sub
ErrProve:
    Call MsgBox("“U" & Key & "”" & "是一个无效的序列号(密钥)！", vbCritical, "无效的序列号")
Exit Sub
ErrProveU:
    Call MsgBox("“" & Text1.Text & "”" & "是一个不合规的序列号(密钥)！", vbCritical, "不合规的序列号")
Exit Sub
OldKey:
    Call MsgBox("“U" & Key & "”" & "是一个已经过期的序列号(密钥)！", vbCritical, "过期的序列号")
    
Exit Sub
FinishAction:
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Act.prove") For Output As #1
    Print #1, (Mid(Key, 1, 7) & Mid(Key, 11, 7))
    Close #1
    Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Act.prove", vbSystem + vbHidden + vbReadOnly)
    yn = MsgBox("激活成功！要现在开始使用“LJX弹窗杀手”吗？", vbYesNo + vbInformation, "激活成功！")
    If yn = vbYes Then
        Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp")
        Shell (App.Path & "/LJX弹窗杀手控制面板.exe")
        Timer1.Enabled = True
    Else
        End
    End If
Exit Sub
Errt:
    If Err.Number = 13 Then
        GoTo ErrProveU
    End If
    Call MsgBox("激活时错误：" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical, "激活时错误")
End Sub

Private Sub Command2_Click()
a = MsgBox("你确定要取消这次激活吗？", vbOKCancel + vbExclamation, "取消激活")
If a = vbOK Then
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp") <> "" Then
        Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp")
        End
    End If
End If
End Sub

Private Sub Command3_Click()
If Command3.Caption = "显示密钥" Then
    Text1.PasswordChar = ""
    Command3.Caption = "隐藏密钥"
ElseIf Command3.Caption = "隐藏密钥" Then
    Text1.PasswordChar = "*"
    Command3.Caption = "显示密钥"
End If
End Sub

Private Sub Form_Load()
Timer1.Enabled = False
End Sub

Private Sub Text1_Change()
On Error Resume Next
Dim x
If Len(Text1.Text) > 27 Then
    Text1.Text = Mid(Text1.Text, 1, 28)
    e = 28
End If
If Len(Text1.Text) = 28 Then
    Label2.Caption = "序列号最大为28位！你已经输入了28位。"
Else
    Label2.Caption = ""
End If
If Mid(Text1.Text, 1, 1) <> "U" Then
    Text1.Text = "U" & Mid(Text1.Text, 1, Len(Text1.Text) - 1)
    Text1.SelStart = Len(Text1.Text)
    Label2.Caption = "序列号必须以U开头！"
ElseIf Len(Text1.Text) <> 28 Then
    Label2.Caption = ""
End If
If Label2.Caption = "" Then
    Acti.Height = 1815
Else
    Acti.Height = 2070
End If
Dim hStr As Boolean
hStr = False
For x = 1 To Len(Text1.Text)
    If x > 1 Then
        If Mid(Text1.Text, x, 1) <> "0" And Mid(Text1.Text, x, 1) <> "1" And Mid(Text1.Text, x, 1) <> "2" And Mid(Text1.Text, x, 1) <> "3" And Mid(Text1.Text, x, 1) <> "4" And Mid(Text1.Text, x, 1) <> "5" And Mid(Text1.Text, x, 1) <> "6" And Mid(Text1.Text, x, 1) <> "7" And Mid(Text1.Text, x, 1) <> "8" And Mid(Text1.Text, x, 1) <> "9" Then
            hStr = True
        End If
    End If
Next
If hStr = True Then
    Acti.Height = 2070
    Label2.Caption = "序列号中除了前缀U之外的字符必须都是数字！"
ElseIf hStr = False And Label2.Caption = "序列号中除了前缀U之外的字符必须都是数字！" Then
    Acti.Height = 1815
    Label2.Caption = ""
End If
End Sub

Private Sub Timer1_Timer()
End
End Sub
