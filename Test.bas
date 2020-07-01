Attribute VB_Name = "Test"
Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Const FTP_TRANSFER_TYPE_ASCII = &H1
Const FTP_TRANSFER_TYPE_BINARY = &H2
Const INTERNET_DEFAULT_FTP_PORT = 21 ' default for FTP servers
Const INTERNET_SERVICE_FTP = 1
Const INTERNET_FLAG_PASSIVE = &H8000000 ' used for FTP connections
Const INTERNET_OPEN_TYPE_PRECONFIG = 0 ' use registry configuration
Const INTERNET_OPEN_TYPE_DIRECT = 1 ' direct to net
Const INTERNET_OPEN_TYPE_PROXY = 3 ' via named proxy
Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4 ' prevent using java/script/INS
Const MAX_PATH = 260

Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Const PassiveConnection As Boolean = True

Sub testea()

End Sub

Sub testea2()
Dim hConnection As Long, hOpen As Long, sOrgPath As String
'open an internet connection
hOpen = InternetOpen("API-Guide sample program", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
'connect to the FTP server
hConnection = InternetConnect(hOpen, "217.125.105.227", INTERNET_DEFAULT_FTP_PORT, "epanaturalcom", "jordi", INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
'create a buffer to store the original directory
sOrgPath = String(MAX_PATH, 0)
'get the directory
FtpGetCurrentDirectory hConnection, sOrgPath, Len(sOrgPath)
'create a new directory 'testing'
FtpCreateDirectory hConnection, "testing"
'set the current directory to 'root/testing'
FtpSetCurrentDirectory hConnection, "testing"
'upload the file 'test.htm'
FtpPutFile hConnection, "C:\\aa.Txt", "aa.Txt", FTP_TRANSFER_TYPE_UNKNOWN, 0
'rename 'test.htm' to 'apiguide.htm'
FtpRenameFile hConnection, "aa.Txt", "apiguide.htm"
'enumerate the file list from the current directory ('root/testing')
EnumFiles hConnection
'retrieve the file from the FTP server
FtpGetFile hConnection, "apiguide.htm", "c:\\apiguide.htm", False, 0, FTP_TRANSFER_TYPE_UNKNOWN, 0
'delete the file from the FTP server
FtpDeleteFile hConnection, "apiguide.htm"
'set the current directory back to the root
FtpSetCurrentDirectory hConnection, sOrgPath
'remove the direcrtory 'testing'
FtpRemoveDirectory hConnection, "testing"
'close the FTP connection
InternetCloseHandle hConnection
'close the internet connection
InternetCloseHandle hOpen
End Sub


Public Sub EnumFiles(hConnection As Long)
Dim pData As WIN32_FIND_DATA, hFind As Long, lRet As Long
'set the graphics mode to persistent
'create a buffer
pData.cFileName = String(MAX_PATH, 0)
'find the first file
hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
'if there's no file, then exit sub
If hFind = 0 Then Exit Sub
'show the filename
Debug.Print Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
Do
'create a buffer
pData.cFileName = String(MAX_PATH, 0)
'find the next file
lRet = InternetFindNextFile(hFind, pData)
'if there's no next file, exit do
If lRet = 0 Then Exit Do
'show the filename
Debug.Print Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
Loop
'close the search handle
InternetCloseHandle hFind
End Sub

Sub ShowError()
Dim lErr As Long, sErr As String, lenBuf As Long
'get the required buffer size
InternetGetLastResponseInfo lErr, sErr, lenBuf
'create a buffer
sErr = String(lenBuf, 0)
'retrieve the last respons info
InternetGetLastResponseInfo lErr, sErr, lenBuf
'show the last response info
MsgBox "Error " + CStr(lErr) + ": " + sErr, vbOKOnly + vbCritical
End Sub


