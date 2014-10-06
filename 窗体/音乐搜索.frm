VERSION 5.00
Begin VB.Form frm音乐搜索 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "搜索网络歌曲"
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
      Caption         =   "搜索"
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
      Text            =   "请输入歌手名或者歌曲名称"
      Top             =   480
      Width           =   4455
   End
   Begin VB.Menu m_R 
      Caption         =   "右键"
      Visible         =   0   'False
      Begin VB.Menu m_addsong 
         Caption         =   "添加到歌集"
      End
      Begin VB.Menu m_New 
         Caption         =   "创建新歌集"
      End
      Begin VB.Menu 右键_分割线 
         Caption         =   "-"
      End
      Begin VB.Menu m_全选 
         Caption         =   "全选"
      End
   End
End
Attribute VB_Name = "frm音乐搜索"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim soGeJi As GeJi '看来只能是多个歌集了！
Dim KeyWord As String


Private Sub Command1_Click()
Dim i As Integer
KeyWord = Text1.Text
If KeyWord = "请输入歌手名或者歌曲名称" Then Exit Sub
'从搜狗中获取歌曲
soGeJi = sogouMusics(KeyWord, False)
Call refresh_list2
End Sub

'刷新歌集
Private Sub refresh_list2()
    List1.Clear
    If soGeJi.N = 0 Then Exit Sub
    '把歌集添加到搜索列表
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
Me.BackColor = GetSetting(App.EXEName, "设置", "背景颜色", RGBtoLong("086154209"))
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

'添加到歌集中
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
 '保存为一个新的歌集
Private Sub m_New_Click()
On Error Resume Next
If List1.ListCount = 0 Or List1.SelCount = 0 Then Exit Sub
soGeJi.NameG = InputBox("请为你这张专辑起一个好听的名字吧！", "专辑名称", KeyWord)
If soGeJi.NameG = "" Or soGeJi.NameG = "请输入歌手名或者歌曲名称" Then Exit Sub
With CommonDialog1
    .Filter = "歌集(*.yyu)|*.yyu"
    .FileName = soGeJi.NameG
    .ShowSave
    soGeJi.pathG = .FileName
    If .FileName = "" Or .FileName = soGeJi.NameG Then Exit Sub
End With
Call saveGeJi(soGeJi, soGeJi.pathG)
End Sub

Private Sub m_全选_Click()
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
If Text1.Text = "" Then Text1.Text = "请输入歌手名或者歌曲名称"
End Sub

