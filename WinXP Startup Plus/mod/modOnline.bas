Attribute VB_Name = "modOnline"
Option Explicit
Public Declare Function InternetGetConnectedState _
    Lib "wininet.dll" (ByRef lpdwFlags As Long, _
    ByVal dwReserved As Long) As Long
    Public Const INTERNET_CONNECTION_MODEM As Long = &H1
    Public Const INTERNET_CONNECTION_LAN As Long = &H2
    Public Const INTERNET_CONNECTION_PROXY As Long = &H4
    Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
    Public Const INTERNET_RAS_INSTALLED As Long = &H10
    Public Const INTERNET_CONNECTION_OFFLINE As Long = &H20
    Public Const INTERNET_CONNECTION_CONFIGURED As Long = &H40

Public Function Online() As Boolean
    Online = InternetGetConnectedState(0&, 0&)
End Function


