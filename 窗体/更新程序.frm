VERSION 5.00
Begin VB.Form frm更新程序 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "检查更新"
   ClientHeight    =   3390
   ClientLeft      =   8565
   ClientTop       =   4365
   ClientWidth     =   4845
   Icon            =   "更新程序.frx":0000
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
      Text            =   "更新程序.frx":57E2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.CommandButton cmd确定更新 
      Appearance      =   0  'Flat
      Caption         =   "确定更新"
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
      Caption         =   "当前版本：v2.1"
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
      Caption         =   "最新版本：v2.1"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   1260
   End
End
Attribute VB_Name = "frm更新程序"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Declare Function DeleteUrlCacheEntry Lib "wininet" Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
Dim url更新的地址() As String '这是一个动态数组，初始化只有一个0
Private Sub cmd确定更新_Click()
Dim i As Integer
Dim a
Dim Free As Integer
Dim p批处理 As String
Dim FileName As String
If Dir(App.path & "\temp\", vbDirectory) = "" Then MkDir App.path & "\temp"

i = MsgBox("该操作将会关闭程序，是否进行？", vbYesNo)
If i = vbNo Then
    '这里应该要删除文件夹 temp的
    
    Exit Sub
    Unload Me
End If
cmd确定更新.Enabled = False
    For i = 0 To UBound(url更新的地址) - 1
        a = Split(url更新的地址(i), "：")
        FileName = Right(a(1), Len(a(1)) - InStrRev(a(1), "/"))
        Text1.Text = "正在升级文件：" & vbCrLf & a(0)
        Call DeleteUrlCacheEntry(a(1))
        Call URLDownloadToFile(0, a(1), App.path & "\temp\" & FileName, 0, 0)
    Next i
    '下载到缓存文件夹了用批处理更新
    Free = FreeFile
    Open App.path & "\copyfile.bat" For Output As #Free
        Print #Free, "xcopy /e /h /r /q %cd%\temp\*.* %cd% /y &del copyfile.bat & rd /s /Q %cd%\temp" '该批处理的意思是把temp文件夹所有文件复制到当前文件夹并且删除自身
    Close #Free
    Shell App.path & "\copyfile.bat"
End
End Sub

Private Sub Form_Load()
Dim txt更新的内容 As String
ReDim url更新的地址(0) As String '重新定义一下维数
Dim url版本文件地址 As String
Dim t当前版本 As Long
Dim t最新版本 As Long
Dim s状态 As String
Dim Free As Integer
Dim vFile As String
Dim temp As String
Me.BackColor = GetSetting(App.EXEName, "设置", "背景颜色", RGBtoLong("086154209"))
vFile = "d:\version.ini"
Free = FreeFile
t当前版本 = CLng(GetSetting(App.EXEName, "版本更新", "当前版本", 0))  '获得当前版本号
url版本文件地址 = "http://www.putaot.cn/player/version.ini"  '更新文件地址
Call DeleteUrlCacheEntry(url版本文件地址)
DoEvents
Call URLDownloadToFile(0, url版本文件地址, vFile, 0, 0)
If exitFile(vFile) = False Then Exit Sub
Open vFile For Input As Free
Do While EOF(Free) = False
    Line Input #Free, temp
   If InStr(temp, "最新版本号：") <> 0 Then
        t最新版本 = CLng(Replace(temp, "最新版本号：", ""))  '获得最新版本号
    ElseIf InStr(temp, "更新的内容：") <> 0 Then
        s状态 = "收集更新内容"
    ElseIf InStr(temp, "更新的文件：") <> 0 Then
        s状态 = "收集更新文件"
    ElseIf InStr(temp, "随心听程序更新说明") <> 0 Then
        s状态 = "更新文件正确"
    Else
        If s状态 = "收集更新内容" Then
            If txt更新的内容 = "" Then
                txt更新的内容 = temp
            Else
                txt更新的内容 = txt更新的内容 & vbCrLf & temp
            End If
        ElseIf s状态 = "收集更新文件" Then
            If UBound(url更新的地址) = 0 Then
                    url更新的地址(0) = temp
                    ReDim Preserve url更新的地址(UBound(url更新的地址) + 1) '加载多一个
            Else
                    url更新的地址(UBound(url更新的地址)) = temp
                    ReDim Preserve url更新的地址(UBound(url更新的地址) + 1) '加载多一个
            End If
        Else
            MsgBox "程序出错:" & Err.description, vbCritical
            End
        End If
    End If
Loop
Close #Free
Kill vFile '删除版本文件
Label1.Caption = "最新版本：" & t最新版本
Label2.Caption = "当前版本：" & t当前版本
Text1.Text = txt更新的内容

If t最新版本 > t当前版本 Then
    cmd确定更新.Enabled = True
Else
    cmd确定更新.Enabled = False
End If
End Sub
