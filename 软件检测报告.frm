VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����ⱨ��"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�������"
      Height          =   375
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "����ռ�ü��"
      Height          =   375
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�����Ŀ���ͣ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "������LJX����ɱ�ֵļ�ⱨ�棺"
      BeginProperty Font 
         Name            =   "����"
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
Text1.Text = "�����Ǽ����Ե����ã�" & vbCrLf & "CPU: Intel Core i5-9300H @2.4GHz" & vbCrLf & "�Կ��� Intel UHD Graphics 630�����ԣ�" & vbCrLf & "�ڴ棺����ʿDDR4 2666MHz 16GB(8GB��2)" & vbCrLf & "���̣�Samsung SSD 970 EVO Plus 1TB" & vbCrLf & "ϵͳ��Microsoft Windows 10 ��ͥ���İ�21H1(19043.1706)" & vbCrLf & vbCrLf & "����������ռ�ü����(����һ������)��" & vbCrLf & "CPUռ�ã�����ģʽ�����ֵ����ֵ���� 15.8%" & vbCrLf & "CPUռ�ã�����ģʽ,���ֵ����ֵ��: 2.3%" & vbCrLf & "CPUռ�ã�����ģʽ�����ֵ����ֵ���� 1.5% " & vbCrLf & "CPUռ�ã���ɾ��ģʽ,���ֵ�ռ�����ֵ����0.8%" & vbCrLf & "�ڴ�ռ�ã���Ԥ�������ֵ��3 MB"
End Sub

Private Sub Command3_Click()
Text1.Text = "�����ǲ�����������2021��4��12�ռ�⣩��" & vbCrLf & "360��ȫ��ʿ13.1.0.1001���޲���" & vbCrLf & "360ɱ��5.0.0.8170���޲���" & vbCrLf & "���ް�ȫ���5.0.62.4����LJX����ɱ�ֿ�����屨��(Trojan/VBCode.aj)" & vbCrLf & "Windows Defender ���޲���" & vbCrLf & vbCrLf & "VirSCAN.org(2021��8��17�� ���� 11:59���) ��" & vbCrLf & "    ��LJX����ɱ�ֿ�����壺0%����ɱ�����������" & vbCrLf & "    ��LJX����ɱ����������2% (Vba32����) " & vbCrLf & "    ��LJX����ɱ�����̳���3%(IKARUS������ ����)" & vbCrLf & "    ��LJX����ɱ��������֤����0%(��ɱ���������)" & vbCrLf & "    ��LJX����ɱ��������0%(��ɱ���������)"
End Sub

Private Sub Form_Load()
Form1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub
