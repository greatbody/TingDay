Attribute VB_Name = "��������Ƿ�����"
'*************************windows xp+sp3,vb6.0 ����ͨ��
Option Explicit
'����/����
Private Declare Function InternetDial Lib "wininet.dll" (ByVal hwndParent As Long, ByVal lpszConnectoid As String, ByVal dwFlags As Long, lpdwConnection As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetHangUp Lib "wininet.dll" (ByVal dwConnection As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Const INTERNET_DIALSTATE_DISCONNECTED = 1
Private Const INTERNET_AUTODIAL_FORCE_ONLINE = 1
Private Const INTERNET_AUTODIAL_FORCE_UNATTENDED = 2
Private Const INTERNET_DIAL_UNATTENDED = &H8000
Private Handle As Long

'����״̬
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Const INTERNET_CONNECTION_MODEM As Long = &H1 '��ϵͳʹ�õ��ƽ����������������
Private Const INTERNET_CONNECTION_LAN As Long = &H2 '��ϵͳͨ��LAN������������
Private Const INTERNET_CONNECTION_PROXY As Long = &H4 '��ϵͳʹ��proxy���������������������
Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8 'δʹ��
Private Const INTERNET_RAS_INSTALLED As Long = &H10
Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40

'ö����������
Private Const RAS_MaxDeviceType = 16
Private Const RAS95_MaxDeviceName = 128
Private Const RAS95_MaxEntryName = 256
Private Type RASCONN95
dwSize As Long
hRasConn As Long
szEntryName(RAS95_MaxEntryName) As Byte
szDeviceType(RAS_MaxDeviceType) As Byte
szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Private Type RASENTRYNAME95
dwSize As Long
szEntryName(RAS95_MaxEntryName) As Byte
End Type
Private Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lprasconn As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long
Private Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long

'����
Public Function DialUp(LinkName As String) As Boolean
InternetDial 0, LinkName, INTERNET_AUTODIAL_FORCE_UNATTENDED, Handle, 0
DialUp = (Handle <> 0)
End Function
'����
Public Sub HangUp()
If Handle <> 0 Then
InternetHangUp Handle, 0
Handle = 0
End If
End Sub
'ö����������
Public Sub EnumConnectName(Value() As String)
Dim s As Long, l As Long, ln As Long, a As String
ReDim r(255) As RASENTRYNAME95

r(0).dwSize = 264
s = 256 * r(0).dwSize
l = RasEnumEntries(vbNullString, vbNullString, r(0), s, ln)
ReDim Value(ln - 1)
For l = 0 To ln - 1
a = StrConv(r(l).szEntryName(), vbUnicode)
Value(l) = Left$(a$, InStr(a$, Chr$(0)) - 1)
Next
End Sub

'�ж��Ƿ�����
Public Function Online() As Boolean
Online = InternetGetConnectedState(0&, 0&)
End Function
'�ж��Ƿ����߲��������ӷ�ʽ
Public Property Get OnlineOfLinkName(LinkName As String) As Boolean
LinkName = Space$(128)
OnlineOfLinkName = InternetGetConnectedStateEx(0, LinkName, 128, 0&)
End Property

'�����ͨ��LAN�����ӣ��򷵻�True
Public Function IsNetConnectViaLAN() As Boolean
Dim dwFlags As Long
Call InternetGetConnectedState(dwFlags, 0&)
IsNetConnectViaLAN = dwFlags And INTERNET_CONNECTION_LAN
End Function

'�����ͨ�����ƽ���������ӣ��򷵻�True
Public Function IsNetConnectViaModem() As Boolean
Dim dwFlags As Long
Call InternetGetConnectedState(dwFlags, 0&)
IsNetConnectViaModem = dwFlags And INTERNET_CONNECTION_MODEM
End Function

'�����ͨ��Proxy��������������ӣ��򷵻�True
Public Function IsNetConnectViaProxy() As Boolean
Dim dwFlags As Long
Call InternetGetConnectedState(dwFlags, 0&)
IsNetConnectViaProxy = dwFlags And INTERNET_CONNECTION_PROXY
End Function

'����Ѱ�װ��RAS���򷵻�True
Public Function IsNetRASInstalled() As Boolean
Dim dwFlags As Long
Call InternetGetConnectedState(dwFlags, 0&)
IsNetRASInstalled = dwFlags And INTERNET_RAS_INSTALLED
End Function
'���ص�ǰ����״̬��Ϣ�ַ���
Public Function GetNetConnectString() As Boolean
Dim dwFlags As Long
Dim Msg As String
If InternetGetConnectedState(dwFlags, 0&) Then
If dwFlags And INTERNET_CONNECTION_CONFIGURED Then
'Msg = MsgBox("ϵͳ��������������") & vbCrLf
End If
If dwFlags And INTERNET_CONNECTION_LAN Then
'Msg = MsgBox("ϵͳͨ����������������������")
End If
If dwFlags And INTERNET_CONNECTION_PROXY Then
'Msg = MsgBox("��ʹ����Proxy�������")

End If
If dwFlags And INTERNET_CONNECTION_MODEM Then
'Msg = MsgBox("ϵͳʹ�õ��ƽ������������������")
End If
If dwFlags And INTERNET_CONNECTION_OFFLINE Then
'Msg = MsgBox("ϵͳ��ǰ��������״̬")
End If
If dwFlags And INTERNET_CONNECTION_MODEM_BUSY Then
'Msg = MsgBox("ϵͳ�ĵ��ƽ����δ���ӵ�������")
End If
If dwFlags And INTERNET_RAS_INSTALLED Then
'Msg = MsgBox("��ϵͳ��װ��Զ�̷��ʷ���")
End If
GetNetConnectString = True
Else
'Msg = MsgBox("��ǰδ������������")
GetNetConnectString = False
End If
End Function

