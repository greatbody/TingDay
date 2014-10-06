Attribute VB_Name = "歌词模块"
Option Explicit
Public Type lrc
Caption As String
Time As Long
End Type
Public mlrc(1 To 500) As lrc
Public lrc_i As Long
Public lrc_Count As Long

Public Sub loadLRC()
Dim i As Integer
For i = 1 To 500
mlrc(i).Caption = i & "偷偷告诉你，歌词模块正在制作当中"
mlrc(i).Time = i
Next i
End Sub

Public Sub drawLRC()
    Dim i As Integer
    With frmLRC
        For i = 1 To 10
            If lrc_i - 5 + i > 0 Then
                .Label2(i).Caption = mlrc(lrc_i - 5 + i).Caption
                .Label2(i).Left = (frmLRC.Width - .Label2(i).Width) / 2
            Else
                .Label2(i).Caption = ""
            End If
        Next i
    End With
    
End Sub

Public Sub getcurLRC(ByVal curTime As Long)
Dim i As Integer
For i = 1 To 499
    If mlrc(i).Time = curTime Then
        lrc_i = i
        Exit Sub
    End If
Next i
End Sub
