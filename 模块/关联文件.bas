Attribute VB_Name = "�����ļ�"
Option Explicit
Private Const REG_SZ = 1
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Sub setLink()
'ע���ļ�����.mp3Ϊmp3��Ƶ�ļ�
RegSetValue HKEY_CLASSES_ROOT, ".mp3", REG_SZ, "mp3��Ƶ�ļ�", 7
'�����ļ�����mp3��Ƶ�ļ���ͼ���뱾���ֲ�������ͼ����ͬ
RegSetValue HKEY_CLASSES_ROOT, "mp3��Ƶ�ļ�\DefaultIcon", REG_SZ, App.path & "\" & App.EXEName & ".exe,0", 24
'����mp3��Ƶ�ļ���ȱʡ�򿪷�ʽΪread
RegSetValue HKEY_CLASSES_ROOT, "mp3��Ƶ�ļ�\Shell", REG_SZ, "open", 4
'����mp3��Ƶ�ļ����Ҽ��˵�read��ʾ�Ĳ˵�������Ϊ"������������"
RegSetValue HKEY_CLASSES_ROOT, "mp3��Ƶ�ļ�\Shell\open", REG_SZ, "TingDay!", 12
'�����ļ�����mp3��Ƶ�ļ���ȱʡ�򿪷�ʽΪ����������
RegSetValue HKEY_CLASSES_ROOT, "mp3��Ƶ�ļ�\Shell\open\Command", REG_SZ, App.path & "\" & App.EXEName & ".exe ""%1""", 22
End Sub

