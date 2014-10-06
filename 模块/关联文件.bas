Attribute VB_Name = "关联文件"
Option Explicit
Private Const REG_SZ = 1
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Sub setLink()
'注册文件类型.mp3为mp3音频文件
RegSetValue HKEY_CLASSES_ROOT, ".mp3", REG_SZ, "mp3音频文件", 7
'设置文件类型mp3音频文件的图标与本音乐播放器的图标相同
RegSetValue HKEY_CLASSES_ROOT, "mp3音频文件\DefaultIcon", REG_SZ, App.path & "\" & App.EXEName & ".exe,0", 24
'设置mp3音频文件的缺省打开方式为read
RegSetValue HKEY_CLASSES_ROOT, "mp3音频文件\Shell", REG_SZ, "open", 4
'设置mp3音频文件的右键菜单read显示的菜单项名称为"用随心听播放"
RegSetValue HKEY_CLASSES_ROOT, "mp3音频文件\Shell\open", REG_SZ, "TingDay!", 12
'设置文件类型mp3音频文件的缺省打开方式为用随心听打开
RegSetValue HKEY_CLASSES_ROOT, "mp3音频文件\Shell\open\Command", REG_SZ, App.path & "\" & App.EXEName & ".exe ""%1""", 22
End Sub

