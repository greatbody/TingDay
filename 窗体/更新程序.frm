VERSION 5.00
Begin VB.Form frm���³��� 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������"
   ClientHeight    =   3390
   ClientLeft      =   8565
   ClientTop       =   4365
   ClientWidth     =   4845
   Icon            =   "���³���.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "���³���.frx":57E2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CommandButton cmdȷ������ 
      Appearance      =   0  'Flat
      Caption         =   "ȷ������"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ�汾��v2.1"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���°汾��v2.1"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1260
   End
End
Attribute VB_Name = "frm���³���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Dim url���µĵ�ַ() As String '����һ����̬���飬��ʼ��ֻ��һ��0
Private Sub cmdȷ������_Click()
Dim i As Integer
Dim a
Dim Free As Integer
Dim p������ As String
Dim FileName As String
If Dir(App.path & "\temp\", vbDirectory) = "" Then MkDir App.path & "\temp"

i = MsgBox("�ò�������رճ����Ƿ���У�", vbYesNo)
If i = vbNo Then
    '����Ӧ��Ҫɾ���ļ��� temp��
    
    Exit Sub
    Unload Me
End If
cmdȷ������.Enabled = False
    For i = 0 To UBound(url���µĵ�ַ) - 1
        a = Split(url���µĵ�ַ(i), "��")
        FileName = Right(a(1), Len(a(1)) - InStrRev(a(1), "/"))
        Text1.Text = "���������ļ���" & vbCrLf & a(0)
        Call DeleteUrlCacheEntry(a(1))
        Call URLDownloadToFile(0, a(1), App.path & "\temp\" & FileName, 0, 0)
    Next i
    '���ص������ļ����������������
    Free = FreeFile
    Open App.path & "\copyfile.bat" For Output As #Free
        Print #Free, "xcopy /e /h /r /q %cd%\temp\*.* %cd% /y &del copyfile.bat & rd /s /Q %cd%\temp" '�����������˼�ǰ�temp�ļ��������ļ����Ƶ���ǰ�ļ��в���ɾ������
    Close #Free
    Shell App.path & "\copyfile.bat"
End
End Sub

Private Sub Form_Load()
Dim txt���µ����� As String
ReDim url���µĵ�ַ(0) As String '���¶���һ��ά��
Dim url�汾�ļ���ַ As String
Dim t��ǰ�汾 As Long
Dim t���°汾 As Long
Dim s״̬ As String
Dim Free As Integer
Dim vFile As String
Dim temp As String
Me.BackColor = GetSetting(App.EXEName, "����", "������ɫ", RGBtoLong("086154209"))
vFile = "d:\version.ini"
Free = FreeFile
t��ǰ�汾 = CLng(GetSetting(App.EXEName, "�汾����", "��ǰ�汾", 0))  '��õ�ǰ�汾��
url�汾�ļ���ַ = "http://www.putaot.cn/player/version.ini"  '�����ļ���ַ
Call DeleteUrlCacheEntry(url�汾�ļ���ַ)
DoEvents
Call URLDownloadToFile(0, url�汾�ļ���ַ, vFile, 0, 0)
If exitFile(vFile) = False Then Exit Sub
Open vFile For Input As Free
Do While EOF(Free) = False
    Line Input #Free, temp
   If InStr(temp, "���°汾�ţ�") <> 0 Then
        t���°汾 = CLng(Replace(temp, "���°汾�ţ�", ""))  '������°汾��
    ElseIf InStr(temp, "���µ����ݣ�") <> 0 Then
        s״̬ = "�ռ���������"
    ElseIf InStr(temp, "���µ��ļ���") <> 0 Then
        s״̬ = "�ռ������ļ�"
    ElseIf InStr(temp, "�������������˵��") <> 0 Then
        s״̬ = "�����ļ���ȷ"
    Else
        If s״̬ = "�ռ���������" Then
            If txt���µ����� = "" Then
                txt���µ����� = temp
            Else
                txt���µ����� = txt���µ����� & vbCrLf & temp
            End If
        ElseIf s״̬ = "�ռ������ļ�" Then
            If UBound(url���µĵ�ַ) = 0 Then
                    url���µĵ�ַ(0) = temp
                    ReDim Preserve url���µĵ�ַ(UBound(url���µĵ�ַ) + 1) '���ض�һ��
            Else
                    url���µĵ�ַ(UBound(url���µĵ�ַ)) = temp
                    ReDim Preserve url���µĵ�ַ(UBound(url���µĵ�ַ) + 1) '���ض�һ��
            End If
        Else
            MsgBox "�������:" & Err.description, vbCritical
            End
        End If
    End If
Loop
Close #Free
Kill vFile 'ɾ���汾�ļ�
Label1.Caption = "���°汾��" & t���°汾
Label2.Caption = "��ǰ�汾��" & t��ǰ�汾
Text1.Text = txt���µ�����

If t���°汾 > t��ǰ�汾 Then
    cmdȷ������.Enabled = True
Else
    cmdȷ������.Enabled = False
End If
End Sub
