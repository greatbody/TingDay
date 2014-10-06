VERSION 5.00
Begin VB.Form frm�������� 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�����������"
   ClientHeight    =   4305
   ClientLeft      =   5400
   ClientTop       =   3300
   ClientWidth     =   5985
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2910
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1080
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "����"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Text            =   "��������������߸�������"
      Top             =   480
      Width           =   4455
   End
   Begin VB.Menu m_R 
      Caption         =   "�Ҽ�"
      Visible         =   0   'False
      Begin VB.Menu m_addsong 
         Caption         =   "��ӵ��輯"
      End
      Begin VB.Menu m_New 
         Caption         =   "�����¸輯"
      End
      Begin VB.Menu �Ҽ�_�ָ��� 
         Caption         =   "-"
      End
      Begin VB.Menu m_ȫѡ 
         Caption         =   "ȫѡ"
      End
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim soGeJi As GeJi '����ֻ���Ƕ���輯�ˣ�
Dim KeyWord As String


Private Sub Command1_Click()
Dim i As Integer
KeyWord = Text1.Text
If KeyWord = "��������������߸�������" Then Exit Sub
'���ѹ��л�ȡ����
soGeJi = sogouMusics(KeyWord, False)
Call refresh_list2
End Sub

'ˢ�¸輯
Private Sub refresh_list2()
    List1.Clear
    If soGeJi.N = 0 Then Exit Sub
    '�Ѹ輯��ӵ������б�
    With soGeJi
    For i = 1 To .N
         If .Gsong(i).mSinger = "" Then
                List1.AddItem .Gsong(i).mTitle
        Else
                List1.AddItem .Gsong(i).mSinger & "-" & .Gsong(i).mTitle
        End If
    Next i
    End With
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
App.TaskVisible = False
Call refresh_list2
Me.Left = frmTingDay1.Left
Me.Top = frmTingDay1.Top
Me.BackColor = GetSetting(App.EXEName, "����", "������ɫ", RGBtoLong("086154209"))
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub List1_DblClick()
Dim temp_singer As String
Dim temp_title As String
Dim temp
Dim temp_url As String
If List1.ListCount = 0 Or List1.SelCount = 0 Then Exit Sub
temp = Split(List1.Text, "-")
If UBound(temp) > 0 Then
    temp_singer = temp(0)
    temp_title = temp(1)
Else
    temp_title = temp(0)
End If
Call addsong(yGeJi, temp_title, temp_singer, "", False)
Call playGeJi(yGeJi, yGeJi.N)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If List1.SelCount = 0 Then
        m_addsong.Enabled = False
        m_New.Enabled = False
    Else
        m_addsong.Enabled = True
        m_New.Enabled = True
    End If
If Button = 2 And List1.ListCount > 0 Then
    Me.PopupMenu m_R
End If
End Sub

'��ӵ��輯��
Private Sub m_addsong_Click()
Dim temp_singer As String
Dim temp_title As String
Dim temp
Dim temp_url As String
If List1.ListCount = 0 Or List1.SelCount = 0 Then Exit Sub
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
    temp = Split(List1.List(i), "-")
    If UBound(temp) > 0 Then
        temp_singer = temp(0)
        temp_title = temp(1)
    Else
        temp_title = temp(0)
    End If
     Call addsong(yGeJi, temp_title, temp_singer, "", False)
End If
Next i
End Sub
 '����Ϊһ���µĸ輯
Private Sub m_New_Click()
On Error Resume Next
If List1.ListCount = 0 Or List1.SelCount = 0 Then Exit Sub
soGeJi.NameG = InputBox("��Ϊ������ר����һ�����������ְɣ�", "ר������", KeyWord)
If soGeJi.NameG = "" Or soGeJi.NameG = "��������������߸�������" Then Exit Sub
With CommonDialog1
    .Filter = "�輯(*.yyu)|*.yyu"
    .FileName = soGeJi.NameG
    .ShowSave
    soGeJi.pathG = .FileName
    If .FileName = "" Or .FileName = soGeJi.NameG Then Exit Sub
End With
Call saveGeJi(soGeJi, soGeJi.pathG)
End Sub

Private Sub m_ȫѡ_Click()
    Dim i As Integer
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = True
    Next i
End Sub

Private Sub Text1_GotFocus()
Text1.Text = ""
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then Call Command1_Click
End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then Text1.Text = "��������������߸�������"
End Sub

