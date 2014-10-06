VERSION 5.00
Begin VB.Form frmLRC 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "LRC"
   ClientHeight    =   10620
   ClientLeft      =   7110
   ClientTop       =   0
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10424.18
   ScaleMode       =   0  'User
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frmLRC.frx":0000
      Top             =   7560
      Width           =   4575
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4560
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   4560
      Top             =   1320
   End
   Begin VB.Label lrcinfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "∏Ë ÷:"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   465
   End
   Begin VB.Label lrcinfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "◊®º≠:"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lrcinfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "∏Ë«˙:"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   465
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   10
      Left            =   1440
      TabIndex        =   10
      Top             =   6840
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   9
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   1440
      TabIndex        =   8
      Top             =   3840
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   7
      Top             =   3360
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   2880
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   1440
      TabIndex        =   5
      Top             =   6360
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   1440
      TabIndex        =   4
      Top             =   5880
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   1440
      TabIndex        =   3
      Top             =   5400
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   1920
      TabIndex        =   2
      Top             =   4920
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   5
      Left            =   960
      TabIndex        =   1
      Top             =   4320
      Width           =   90
   End
   Begin VB.Label lb_End 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "frmLRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iiii As Integer
Private Sub Form_Load()
Dim i As Integer
Me.Height = frmTingDay.Height
'Dim backpic As String
Dim Color As Long
'backpic = App.path & "\backpic.jpg"
Color = GetSetting(App.EXEName, "…Ë÷√", "±≥æ∞…´", RGBtoLong("000000000"))
Me.BackColor = Color
Text1.BackColor = Color
'If exitFile(backpic) = True Then Me.Picture = LoadPicture(backpic)
setTrsp 210, Me.hWnd
'≈‰÷√∏Ë¥ 
    For i = 1 To 10
        frmLRC.Label2(i).Caption = ""
    Next i
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '“∆∂Ø¥∞ÃÂ
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub



Private Sub lb_End_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
Dim mForm As RECT
mForm.Left = ScaleX(Me.Left, vbTwips, vbPixels)
mForm.Right = ScaleX(Me.Width, vbTwips, vbPixels) + ScaleX(Me.Left, vbTwips, vbPixels)
mForm.Top = ScaleY(Me.Top, vbTwips, vbPixels)
mForm.Bottom = ScaleY(Me.Height, vbTwips, vbPixels) + ScaleY(Me.Top, vbTwips, vbPixels)
'¥∞ÃÂ∂•≤ø“˛≤ÿ
hit mForm, 2
End Sub


Private Sub Timer2_Timer()
On Error Resume Next
Dim i As Integer
If UBound(Lyric) <= 5 Then
    For i = 0 To 2
        frmLRC.lrcinfo(i).Caption = ""
    Next i
    frmLRC.Label2(5).Caption = "µ±«∞∏Ë«˙√ª”–∏Ë¥ "
        
    frmLRC.Label2(6).Caption = "TingDay“Ù¿÷£¨◊£ƒ„ÃÏÃÏ∫√–ƒ«È"
    For i = 5 To 6
        frmLRC.Label2(i).Left = (frmLRC.Width - frmLRC.Label2(i).Width) / 2
    Next i
    Exit Sub
End If
lrcinfo(0).Caption = "∏Ë«˙:" & sGM '∏Ë«˙√˚≥∆
lrcinfo(1).Caption = "◊®º≠:" & sZJ '◊®º≠
lrcinfo(2).Caption = "∏Ë ÷:" & sGS '∏Ë ÷–≈œ¢
drawLRC curLength
End Sub
