VERSION 5.00
Begin VB.Form frm����ѡ���� 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "����������ѡ��"
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
Attribute VB_Name = "frm����ѡ����"
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
Me.BackColor = GetSetting(App.EXEName, "����", "������ɫ", RGBtoLong("086154209"))
i = 0
'��������ļ����Ƿ����
themePath = App.path & "\theme"
themePath2 = App.path & "\theme\*.txt"
If Dir(themePath2) = "" Then
    MsgBox "���ⲻ���ڣ�����³���", vbCritical, "����"
    Exit Sub
End If
'������ڣ���ʼװ�س���
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
