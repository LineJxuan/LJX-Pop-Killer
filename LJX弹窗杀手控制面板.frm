VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LJX����ɱ��-�������"
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
   StartUpPosition =   2  '��Ļ����
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
      Caption         =   "�����ⱨ��"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "�鿴LJX����ɱ�ֵĵ�������ⱨ��"
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "ǿ������LJX����ɱ��"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "ǿ�������������е�LJX����ɱ��"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "ǿ��ֹͣLJX����ɱ��"
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "ǿ��ֹͣ�������е�LJX����ɱ��"
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
      Caption         =   "����LJX����ɱ��"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "��������LJX����ɱ��"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FF00FF&
      Caption         =   "����������"
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "���LJX����ɱ�ֵ��������"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0000FF00&
      Caption         =   "����LJX����ɱ��"
      Height          =   375
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "����LJX����ɱ��"
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FF00FF&
      Caption         =   "��ӵ���������"
      Height          =   375
      Left            =   4680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "��LJX����ɱ����ӵ�����������"
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������������(ѡ����ʵ�ģʽ����Ӧ���Ե�����)��"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "��ɾ����ֻɾ�������ĳ���û���κ�Ӱ�죩"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����ģʽ  (û���κ�Ӱ�죬�����������ӳ�)  "
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "������ģʽ  (�Ե������ܼ���û��Ӱ�죬���������ӳ�)  "
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��ͨģʽ  (�Ե������ܻ���΢С��Ӱ�죬���������ӳ�)  "
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ǿ��ģʽ  (���ܻ�Ե���������һ��Ӱ�죬����΢С���ӳ�)  "
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����ģʽ  (���ܻ�Ե��������н϶�Ӱ�죬����û���ӳ�)  "
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "����ģʽ  (��Ե��������н϶�Ӱ�졢��ֱ�ӷ��鵯���ĳ���)  "
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "����LJX����ɱ��"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "��ʾ����LJX����ɱ�ֵ���Ϣ"
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF00&
      Caption         =   "����Ҫ���صĵ���"
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "�鿴��ǰ������Ϊ�����ء��ĵ���"
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "ֹͣLJX����ɱ�ֵ�����"
      Height          =   375
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "ʹ�������е�LXJ����ɱ��ֹͣ����"
      Top             =   4560
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�رմ˿������"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "�����趨���رտ������"
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   8400
      Picture         =   "LJX����ɱ�ֿ������.frx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "��С����ǰ����"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   8400
      Picture         =   "LJX����ɱ�ֿ������.frx":0A26
      Stretch         =   -1  'True
      ToolTipText     =   "��С����ǰ����"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   8400
      Picture         =   "LJX����ɱ�ֿ������.frx":144C
      Stretch         =   -1  'True
      ToolTipText     =   "��С����ǰ����"
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
      Caption         =   "��ǰ�У�������������������"
      BeginProperty Font 
         Name            =   "����"
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
a = MsgBox("ȷ��Ҫǿ��ֹͣLJX����ɱ��������ܻᵼ�������¹ʡ�", vbOKCancel, "ϵͳ��ʾ")
If a = vbOK Then
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='LJX����ɱ��������.exe'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next

    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='LJX����ɱ����������.exe'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next
    Set objProcess = Nothing
    Set colProcessList = Nothing
    Set objWMIService = Nothing
    MsgBox ("ֹͣ�ɹ���")
End If
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("����" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command3_Click()
On Error GoTo Errt
Dim a
a = MsgBox("ȷ��Ҫǿ������LJX����ɱ��������ܻᵼ�������¹ʡ�", vbOKCancel, "ϵͳ��ʾ")
If a = vbOK Then
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    
    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='LJX����ɱ��������.exe'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next

    Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='LJX����ɱ����������.exe'")
    For Each objProcess In colProcessList
        objProcess.Terminate
    Next
    Set objProcess = Nothing
    Set colProcessList = Nothing
    Set objWMIService = Nothing
    Call GetProgramRunState
    Shell (App.Path & "\LJX����ɱ����������.exe")
    MsgBox ("�����ɹ���")
End If
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("����" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Image2.Visible = False
Image3.Visible = False
End Sub

Private Sub Timer1_Timer()
Shell (App.Path & "\LJX����ɱ����������.exe")
Form1.Enabled = True
Form1.Caption = "LJX����ɱ��-�������"
Timer1.Enabled = False
End Sub

Private Sub Command1_Click()
On Error GoTo Errt
i = MsgBox("ȷ��Ҫ�ر�LJX����ɱ�ֿ��������(�����趨���ᱣ��)", vbOKCancel, "ϵͳ��ʾ")
If i = vbOK And Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp") <> "" Then
    Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp")
    End
End If
Exit Sub
Errt:
MsgBox ("�رմ���" & Err.Number & vbCrLf & Err.Description)
End
End Sub

Private Sub Command10_Click()
On Error GoTo Errt
i = MsgBox("ȷ��Ҫ����LJX����ɱ����", vbOKCancel, "ϵͳ��ʾ")
If i = vbOK Then
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\EndAll.ltmp") For Output As #1
    Close #1
    Form1.Enabled = False
    Form1.Caption = "LJX����ɱ��-�������   [��������LJX����ɱ��]"
    Timer1.Interval = 6000
    Timer1.Enabled = True
End If
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("����" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command4_Click()
On Error GoTo Errt
i = MsgBox("ȷ��ֹͣLJX����ɱ�ֵ�������ֹͣ��LJX����ɱ�ֽ��޷���Ϊ�����ص�����", vbOKCancel, "ϵͳ��ʾ")
If i = vbOK Then
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\EndAll.ltmp") For Output As #1
    Close #1
    MsgBox ("LJX����ɱ��ֹͣ�ɹ���")
End If
Call GetProgramRunState
Exit Sub
Errt:
MsgBox ("����" & Err.Number & vbCrLf & Err.Description)
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
i = MsgBox("ȷ��Ҫ��LJX����ɱ����ӵ�����������", vbOKCancel, "ϵͳ��ʾ")
If i = vbCancel Then
    Exit Sub
End If
a = Dir(App.Path & "\LJX����ɱ�ֿ������.exe")
'b = Dir(App.Path & "\LJX����ɱ��������֤����.exe")
'c = Dir(App.Path & "\LJX����ɱ�����̳���.exe")
d = Dir(App.Path & "\LJX����ɱ��������.exe")
e = Dir(App.Path & "\LJX����ɱ����������.exe")
If a = "" Or d = "" Or e = "" Then
    f = MsgBox("LJX����ɱ�ֵ��ļ����������޷����ע��������", vbOKOnly, "�޷���������")
    Exit Sub
End If
If i = vbOK Then
    Set w = CreateObject("wscript.shell")
    w.regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\" & "LJX����ɱ��", App.Path & "\LJX����ɱ����������.exe"
End If
MsgBox ("������������ӳɹ���")
Exit Sub
Errt:
MsgBox ("����" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command8_Click()
On Error GoTo Errt
Shell (App.Path & "\LJX����ɱ����������.exe")
Exit Sub
Errt:
MsgBox ("����:" & Err.Number & vbCrLf & Err.Description)
End Sub

Private Sub Command9_Click()
Dim a As Boolean
Dim b As Boolean
Dim c As Boolean
Dim q
Call GetProgramRunState
'a = CheckExeIsRun("LJX����ɱ��������֤����.exe")
'b = CheckExeIsRun("LJX����ɱ�����̳���.exe")
c = CheckExeIsRun("LJX����ɱ��������.exe")

If c = True Then
    Call MsgBox("�����Ѿ��򿪣�", vbInformation, "�����Ѵ�")
Else
    Call MsgBox("����û�д򿪣�", vbInformation, "����δ��")
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
Call MsgBox("����ʱ����Form1_Load��ʼ������" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
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
    a = MsgBox("��ȷ��Ҫ����������ģʽ�������ɾ�������ĳ���", vbOKCancel + vbExclamation, "����")
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
Call MsgBox("����ʧ�ܣ�" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
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
Call MsgBox("����ʧ�ܣ�" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
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
Call MsgBox("����ʧ�ܣ�" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
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
Call MsgBox("����ʧ�ܣ�" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
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
Call MsgBox("����ʧ�ܣ�" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
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
Call MsgBox("����ʧ�ܣ�" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub

Private Sub Option7_Click()
On Error GoTo Errt
If Option7.Value = True And Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode7.ltmp") = "" Then
    a = MsgBox("��ȷ��Ҫ��������ɾ���������ɾ�������ĳ���", vbOKCancel + vbExclamation, "����")
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
Call MsgBox("����ʧ�ܣ�" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
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
