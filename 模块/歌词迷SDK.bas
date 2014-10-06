Attribute VB_Name = "�����SDK"
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Sub dowloadlrc(ByRef mSong As Song)
Dim allLRC As String '
Dim one As Integer
Dim two As Integer
Dim lrcURL As String

If mSong.Singer <> "" Then
    allLRC = htmlStr("http://geci.me/api/lyric/" & mSong.title & "/" & mSong.Singer)
Else
    allLRC = htmlStr("http://geci.me/api/lyric/" & mSong.title)
End If
'ֱ��ѡȡ��һ�����
one = InStr(allLRC, "http://")
If one = 0 Then Exit Sub

two = InStr(allLRC, ".lrc") + 4
lrcURL = Mid(allLRC, one, two)
'��ʼ���ظ��
mSong.lrcPath = App.path & "\���\" & mSong.title & ".lrc"
URLDownloadToFile 0, lrcURL, mSong.lrcPath, 0, 0
End Sub
