Attribute VB_Name = "Module1"
Public MyName As String         '�û���
Public Pops(0 To 1023) As String        '�����صĵ���·��
Public PopsP(0 To 1023) As String       '�����صĵ�������
Public MaxPops As Long          'Ŀǰ���Ŀ�д�뵯��
Public StartMe As Boolean
Public IsOne As Boolean
Public Act As Boolean           '����Ѿ����
Public Function CheckExeIsRun(ExeName As String) As Boolean
On Error GoTo Errt
Dim WMI
Dim Obj
Dim Objs
CheckExeIsRun = False
Set WMI = GetObject("WinMgmts:")
Set Objs = WMI.InstancesOf("Win32_Process")
For Each Obj In Objs
    If (InStr(UCase(ExeName), UCase(Obj.Description)) <> 0) Then
        CheckExeIsRun = True
        ExeNumber = InStr(UCase(ExeName), UCase(Obj.Description))
        If Not Objs Is Nothing Then Set Objs = Nothing
        If Not WMI Is Nothing Then Set WMI = Nothing
        Exit Function
    End If
Next
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
Exit Function
Errt:
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
End Function
Function LoadModes()
Dim u As Boolean
u = False
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode1.ltmp") <> "" Then
    Form1.Option1.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode2.ltmp") <> "" Then
    Form1.Option2.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode3.ltmp") <> "" Then
    Form1.Option3.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode4.ltmp") <> "" Then
    Form1.Option4.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode5.ltmp") <> "" Then
    Form1.Option5.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode6.ltmp") <> "" Then
    Form1.Option6.Value = True
    u = True
End If
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode7.ltmp") <> "" Then
    Form1.Option7.Value = True
    u = True
End If
If u = False Then
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\Mode4.ltmp") For Output As #1
    Close #1
    u = True
    Form1.Option4.Value = True
End If
End Function
Sub Main()
Load Form1
On Error GoTo Errt
If IsOne = False Then
    IsOne = True
    WindowState = vbNormal
End If
Call LoadAll
Call LoadModes
Call GetProgramRunState
Exit Sub
Errt:
Call MsgBox("����ʱ����ļ�����" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Sub
Public Function LoadAll()
On Error GoTo Errt
MaxPops = 0
MyName = Environ("USERNAME")
'�������
For X = 0 To 1023
    Pops(X) = ""
    PopsP(X) = ""
Next
'----------
Dim a
Dim b
Dim c
Dim d
Dim e
Dim f
Dim g

a = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller", vbDirectory + vbSystem)
If a = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller", vbSystem)

b = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf", vbDirectory + vbSystem)
If b = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf", vbSystem)

c = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Number", vbDirectory + vbSystem)
If c = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Number")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Number", vbSystem)

d = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Log", vbDirectory + vbSystem)
If d = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Log")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Log", vbSystem)

e = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops", vbDirectory + vbSystem)
If e = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops", vbSystem)

f = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp", vbDirectory + vbSystem)
If f = "" Then
    MkDir ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp")
End If
Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp", vbSystem)

g = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\���ڸ��ļ��е�˵��.txt", vbSystem + vbReadOnly)
If g = "" Then
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\���ڸ��ļ��е�˵��.txt") For Output As #1
    Print #1, "����LJX����ɱ�ֵ���Ҫ�ļ��У��벻Ҫ����������κ��ļ���"
    Close #1
    Call SetAttr("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\���ڸ��ļ��е�˵��.txt", vbSystem + vbReadOnly)
End If

Dim q
Dim Stri As String
Dim r As Long
r = 0
For X = 0 To 1023
    Stri = ""
    q = Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx", vbSystem)
    If q <> "" Then
        If FileLen("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx") = 0 Then
            Close #1
            DelFile (X)
        Else
            Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx") For Input As #1
            Line Input #1, Stri
            Pops(r) = Stri
            MaxPops = MaxPops + 1
            Close #1
            r = r + 1
        End If
    End If
Next


'ɾ�����ļ�
Dim Y As String
For X = 0 To 1023
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx") <> "" Then
        Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx") For Input As #1
        Line Input #1, Y
        If Y = "" Then
            Close #1
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx")
        End If
        Close #1
    End If
Next



'����Ƿ����һ���������

If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp") <> "" And StartMe = False And CheckExeIsRun("LJX����ɱ�ֿ������.exe") = True Then
    
    Call MsgBox("��һ��LJX����ɱ�ֿ��������������,�벻Ҫ�ظ���LJX����ɱ�ֿ�����壡" & vbCrLf & "����������������������ʹ�����׵ġ�LJX����ɱ���޸����򡱣���ѡ���޸����͵ĵ�һ�", vbOKOnly + vbExclamation, "ϵͳ��ʾ")
    End
    Exit Function
End If
Refform1
Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp") For Output As #1
Close #1
StartMe = True

If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Act.prove", vbSystem + vbHidden + vbReadOnly) = "" Then
    Act = False
    bo = MsgBox("����һ��û�б�����ġ�LJX����ɱ�֡�������ϵLJX�Խ��м��������Ѿ�ӵ�������к�(��Կ)��������ȷ������ť�Կ�ʼ���", vbOKCancel + vbExclamation, "���ȼ������")
    If bo = vbOK Then
        Acti.Show
        Exit Function
    Else
        If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp") <> "" Then
            Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Temp\ControlRun.ltmp")
            End
        End If
    End If
End If


Form1.Show

Exit Function

Errt:
Call MsgBox("��������" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End

End Function

Public Function DelFile(FileNumber)
On Error GoTo Errt
Form2.Hide
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & FileNumber & ".ltx") = "" Then
    Call MsgBox("û���ҵ�ָ���ĺ�����", vbOKOnly + vbExclamation)
End If
'�ȱ���Ҫ�����ŵ��ļ���
Dim X As Long
Dim r(0 To 1023) As String
Dim Maxr As Long
Maxr = 0
For X = (FileNumber + 1) To 1023
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx") <> "" Then
        Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx") For Input As #1
        Line Input #1, r(X)
        Maxr = Maxr + 1
        Close #1
    Else
        Exit For
    End If
Next
Close #1
'ɾ���ļ�
If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & FileNumber & ".ltx") <> "" Then
    Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & FileNumber & ".ltx")
    MaxPops = MaxPops - 1
End If
'ɾ�����ļ�
For X = (FileNumber + 1) To 1023
    If Dir("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx") <> "" Then
        Kill ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx")
    End If
Next
'�������ļ�
For X = FileNumber To Maxr
    Open ("C:\Users\" & MyName & "\AppData\Roaming\LJXPopKiller\Inf\Pops\" & X & ".ltx") For Output As #1
    Print #1, r(X + 1)
    Close #1
Next
Exit Function
LoadAll
Unload Form1
Unload frmAbout
Unload Form2
Unload Form3
Unload Form4
Unload Acti
Errt:
Call MsgBox("����" & Err.Number & vbCrLf & Err.Description, vbOKOnly + vbCritical)
End Function
Function GetProgramRunState()
Form1.Label1.ForeColor = &H0&
If CheckExeIsRun("LJX����ɱ��������.exe") = False Then
    Form1.Label1.ForeColor = &HFF&
    Form1.Label1.Caption = "LJX����ɱ��δ�����У�"
Else
    Form1.Label1.ForeColor = &H0&
    Form1.Label1.Caption = "��ǰ��" & MaxPops & "������������"
End If
End Function

Public Function Refform1()
If Pops(0) = "" Then
    MaxPops = 0
End If
Form1.Label1.Caption = "��ǰ��" & MaxPops & "������������"
End Function

