VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "TingDay"
   ClientHeight    =   2655
   ClientLeft      =   8160
   ClientTop       =   4110
   ClientWidth     =   4455
   Icon            =   "frmTingDay3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   720
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   2040
      Top             =   1200
   End
   Begin VB.Menu m_BackColor 
      Caption         =   "������ɫ"
      Begin VB.Menu ColorList 
         Caption         =   "��ɫ�б�"
         Checked         =   -1  'True
      End
      Begin VB.Menu skyBlue 
         Caption         =   "�����"
      End
      Begin VB.Menu perpol 
         Caption         =   "��ɫ"
      End
      Begin VB.Menu �ָ���_��ɫ3 
         Caption         =   "-"
      End
      Begin VB.Menu ��� 
         Caption         =   "���"
      End
      Begin VB.Menu MoLv 
         Caption         =   "ī��"
      End
      Begin VB.Menu coffee 
         Caption         =   "����ɫ"
      End
      Begin VB.Menu �ָ���_��ɫ1 
         Caption         =   "-"
      End
      Begin VB.Menu orange 
         Caption         =   "��ɫ"
      End
      Begin VB.Menu Caolv 
         Caption         =   "����"
      End
      Begin VB.Menu Red 
         Caption         =   "õ���"
      End
   End
   Begin VB.Menu �˵�1 
      Caption         =   "�˵�1"
      Begin VB.Menu ����ļ��� 
         Caption         =   "����ļ���"
      End
      Begin VB.Menu �򿪸赥 
         Caption         =   "�򿪸赥"
      End
      Begin VB.Menu �ָ���_�˵�1 
         Caption         =   "-"
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
      End
      Begin VB.Menu �ָ���2_�˵�1 
         Caption         =   "-"
      End
      Begin VB.Menu �½��赥 
         Caption         =   "�½��赥"
      End
      Begin VB.Menu ɾ���赥 
         Caption         =   "ɾ���赥"
      End
      Begin VB.Menu ����赥 
         Caption         =   "����赥"
      End
   End
   Begin VB.Menu �˵� 
      Caption         =   "�˵�"
      Begin VB.Menu ��ʾ 
         Caption         =   "TingDay"
      End
      Begin VB.Menu �ָ���3 
         Caption         =   "-"
      End
      Begin VB.Menu �����ȼ� 
         Caption         =   "�����ȼ�"
      End
      Begin VB.Menu TingDay�ָ���1 
         Caption         =   "-"
      End
      Begin VB.Menu ��һ�� 
         Caption         =   "��һ��"
      End
      Begin VB.Menu ͣ�� 
         Caption         =   "����"
      End
      Begin VB.Menu ��һ�� 
         Caption         =   "��һ��"
      End
      Begin VB.Menu �ָ���1 
         Caption         =   "-"
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
      End
      Begin VB.Menu �ָ���2 
         Caption         =   "-"
      End
      Begin VB.Menu �˳� 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu �˵�2 
      Caption         =   "�˵�2"
      Begin VB.Menu ����ļ� 
         Caption         =   "����ļ�"
      End
      Begin VB.Menu ���ļ�λ�� 
         Caption         =   "���ļ�λ��"
      End
      Begin VB.Menu ɾ������ 
         Caption         =   "ɾ������"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub coffee_Click()
setBackColor "096056017"
allFalse
coffee.Checked = True
End Sub

Private Sub ColorList_Click()
    If ColorList.Checked = True Then
        listbackcolor False
    Else
        listbackcolor True
    End If
End Sub

Private Sub Form_Load()
Me.Visible = False
preProc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf GetProc
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
msg = X / 15
Select Case msg

Case WM_LBUTTONDBLCLK
    frmTingDay.Show
    If frmTingDay.Top < 0 Then frmTingDay.Top = 100
Case WM_RBUTTONDOWN
    Call SetForegroundWindow(Me.hWnd)
    Me.PopupMenu �˵�
    
End Select
End Sub

Private Sub MoLv_Click()
setBackColor "000070048"
allFalse
MoLv.Checked = True
End Sub

Private Sub orange_Click()
setBackColor "255097000"
allFalse
orange.Checked = True
End Sub

Private Sub perpol_Click()
setBackColor "153051225"
allFalse
perpol.Checked = True
End Sub

Private Sub Red_Click()
setBackColor "252033094"
allFalse
Red.Checked = True
End Sub
Private Sub allFalse()
skyBlue.Checked = False
perpol.Checked = False
MoLv.Checked = False
coffee.Checked = False
orange.Checked = False
Caolv.Checked = False
Red.Checked = False
���.Checked = False
End Sub
Private Sub skyBlue_Click()
setBackColor "030144225"
allFalse
skyBlue.Checked = True

End Sub

Private Sub Timer1_Timer()
Dim FileName As String


FileName = GetSetting(App.EXEName, "����࿪", "�����ļ�", "")
If FileName = "" Then Exit Sub
FileName = Mid(FileName, 2, Len(FileName) - 2)
If exitFile(FileName) = True Then
    FileOpen FileName
    SaveSetting App.EXEName, "����࿪", "�����ļ�", ""
End If

End Sub

Private Sub caolv_Click()
setBackColor "000199140"
allFalse
Caolv.Checked = True
End Sub

Private Sub Timer2_Timer()
'ע���ȼ�
RegHotKeyAll
End Sub

Private Sub ����赥_Click()
Dim FileName As String
Dim msg As Boolean
Dim mList As list
FileName = ShowSave("����赥", "�赥�ļ���*.tdl��" & Chr(0) & "*.tdl", Me.hWnd)
mList = playList
msg = saveList(mList, FileName)
saveGeJi GeJiPath
If msg = vbYes Then MessageBox frmMenu.hWnd, "����ɹ�", "��ʾ", vbOKOnly
End Sub

Private Sub �򿪸赥_Click()
    Dim FileName As String
    Dim mList As list
    FileName = ShowOpen("�򿪸赥", "�赥�ļ���*.tdl��" & Chr(0) & "*.tdl", Me.hWnd)
    If FileName = "" Then Exit Sub
    openList mList, FileName
    addList GeJiA, mList
    saveGeJi GeJiPath
    GeJiA.Index = GeJiA.Count
    refreshList listA(GeJiA.Index)
    refreshGeJi
End Sub

Private Sub ���ļ�λ��_Click()
On Error Resume Next
Dim File As String
Dim FileName As String
FileName = SongA(playList.Index).url
If FileName = "" Then Exit Sub
File = Dir(FileName)
Shell "explorer /select," & FileName, vbNormalFocus
End Sub

Private Sub ������Ϣ_Click()

End Sub

Private Sub �����ȼ�_Click()
If �����ȼ�.Checked = False Then
    Timer2.Enabled = False
    �����ȼ�.Checked = True
    UnRegHot
Else
    Timer2.Enabled = True
    �����ȼ�.Checked = False
End If
End Sub

Private Sub ���_Click()
setBackColor "000000000"
allFalse
���.Checked = True
End Sub

Private Sub ɾ���赥_Click()
Dim msg As Integer
If GeJiA.Index > 0 Then
    msg = MessageBox(Me.hWnd, " ɾ���ĸ赥�����ָܻ�,�Ƿ����", "ѯ��", vbYesNo)
    If msg = vbNo Then Exit Sub
    deleteList GeJiA, GeJiA.Index
    If GeJiA.Count > 0 Then
        If GeJiA.Index > GeJiA.Count Then
            GeJiA.Index = GeJiA.Count
            zplay_Stop mTingDay
            playList.isPlay = False
            playList.Count = 0
            updataUI
        End If
        If GeJiA.playWhere > GeJiA.Index Then GeJiA.playWhere = GeJiA.playWhere - 1
        listShow GeJiA.Index
    Else
        
    End If
    refreshList listA(GeJiA.Index)
    refreshGeJi
    saveGeJi GeJiPath
End If
End Sub

Private Sub ɾ������_Click()
Dim msg As Integer
Dim i As Integer
i = frmTingDay.List1.listIndex + 1
If i > 0 Then
    msg = MessageBox(Me.hWnd, "���׸���ֻ�Ǵ��б����Ƴ��Ƿ����", "ѯ��", vbYesNo)
    If msg = vbNo Then Exit Sub
    deleteSong listA(GeJiA.Index), i
    refreshList listA(GeJiA.Index)
    saveList listA(GeJiA.Index), listA(GeJiA.Index).path
    playList = listA(GeJiA.Index)
End If
End Sub

Private Sub ��һ��_Click()
playPre
updataUI
End Sub

Private Sub �������_Click()
Rndplay
updataUI
End Sub

Private Sub ����ļ�_Click()
    Dim mSong As Song
    With mSong
        .url = ShowOpen("��Ӹ���", "MP3�ļ���*.mp3��" & Chr(0) & "*.mp3", Me.hWnd)
        If .url = "" Or exitFile(.url) = False Then Exit Sub
        .title = delTuo(.url, True)
        .Islocal = True
    End With
    addsong listA(GeJiA.Index), mSong
    refreshList listA(GeJiA.Index)
    saveList listA(GeJiA.Index), listA(GeJiA.Index).path
End Sub

Private Sub ����ļ���_Click()
    Dim FileName As String
    Dim Caption As String
    Dim mList As list
    Dim Free As Long
    Free = FreeFile
    FileName = showFolder("����ļ���", Me.hWnd)
    If FileName = "" Then Exit Sub
    Caption = FolderName(FileName)
    With mList
        .Name = Caption
        .path = FileName
        seachFiles mList, FileName, ".mp3", True
        .Index = 1
        .playWhere = 0
    End With
    addList GeJiA, mList
    saveGeJi GeJiPath
    refreshList mList
    refreshGeJi
    GeJiA.Index = GeJiA.Count
    frmTingDay.List1.listIndex = mList.Index - 1
    frmTingDay.List2.listIndex = GeJiA.Index - 1
End Sub

Private Sub ͣ��_Click()
    If playList.Count = 0 Then
        FileOpen ShowOpen("�򿪸輯", "�����ļ�", Me.hWnd)
        Exit Sub
    End If
    If mStatus.fPlay = 1 Then
        zplay_Pause mTingDay
        frmTingDay.Label4.Caption = "4"
        ͣ��.Caption = "����"
    ElseIf mStatus.fPause = 1 Then
        zplay_Play mTingDay
        frmTingDay.Label4.Caption = ";"
        ͣ��.Caption = "��ͣ"
    Else
        playList = playList
        listPlay playList, playList.Index + 1
        frmTingDay.Label4.Caption = ";"
        ͣ��.Caption = "��ͣ"
    End If
    updataUI
End Sub

Private Sub �˳�_Click()
EndmTingDay
End Sub

Private Sub ��һ��_Click()
If mPlayStyle = 1 Then
    Rndplay
Else
    playNext '������һ��
End If
updataUI
End Sub

Private Sub ��ʾ_Click()
showMe
End Sub

Private Sub �½��赥_Click()
Dim Caption As String
Dim FileName As String
Dim mTuo As String
Dim mList As list
FileName = ShowSave("����赥", "�輯�ļ���*.tdl��" & Chr(0) & "*.tdl", Me.hWnd)
If FileName = "" Then Exit Sub
Caption = delTuo(FileName, True)
If Caption = "" Then Exit Sub

With mList
    .path = FileName
    .Name = Caption
End With
saveList mList, FileName
addList GeJiA, mList
saveGeJi GeJiPath
refreshList mList
refreshGeJi
End Sub

Private Sub ������_Click()
Dim Name As String
Name = InputBox("�������µ����ƣ�", "������")
If Name = "" Then Exit Sub
listA(GeJiA.Index).Name = Name
saveGeJi GeJiPath
refreshGeJi
End Sub
