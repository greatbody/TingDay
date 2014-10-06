VERSION 5.00
Begin VB.Form frm主题选择器 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "随心听主题选择"
   ClientHeight    =   3360
   ClientLeft      =   9765
   ClientTop       =   1785
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3090
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frm主题选择器"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim themePath As String
Dim themes(0 To 19) As String
Private Sub Form_Load()
Dim themePath2 As String
Dim i As Integer
Me.BackColor = GetSetting(App.EXEName, "设置", "背景颜色", RGBtoLong("086154209"))
i = 0
'检查主题文件夹是否存在
themePath = App.path & "\theme"
themePath2 = App.path & "\theme\*.txt"
If Dir(themePath2) = "" Then
    MsgBox "主题不存在，请更新程序！", vbCritical, "错误"
    Exit Sub
End If
'主题存在，开始装载程序
themes(i) = Dir(themePath2)
List1.AddItem Left(themes(i), Len(themes(i)) - 4)
Do
i = i + 1
themes(i) = Dir()
If themes(i) = "" Then Exit Do
List1.AddItem Left(themes(i), Len(themes(i)) - 4)
Loop
End Sub

Private Sub List1_Click()
    Call applyTheme(themePath & "\" & themes(List1.ListIndex))
    
End Sub
