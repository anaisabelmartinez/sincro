Attribute VB_Name = "Importacions"

Option Explicit

'Icmp constants converted from
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_pingstatus.asp

Private Const ICMP_SUCCESS As Long = 0
Private Const ICMP_STATUS_BUFFER_TO_SMALL = 11001                   'Buffer Too Small
Private Const ICMP_STATUS_DESTINATION_NET_UNREACH = 11002           'Destination Net Unreachable
Private Const ICMP_STATUS_DESTINATION_HOST_UNREACH = 11003          'Destination Host Unreachable
Private Const ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH = 11004      'Destination Protocol Unreachable
Private Const ICMP_STATUS_DESTINATION_PORT_UNREACH = 11005          'Destination Port Unreachable
Private Const ICMP_STATUS_NO_RESOURCE = 11006                       'No Resources
Private Const ICMP_STATUS_BAD_OPTION = 11007                        'Bad Option
Private Const ICMP_STATUS_HARDWARE_ERROR = 11008                    'Hardware Error
Private Const ICMP_STATUS_LARGE_PACKET = 11009                      'Packet Too Big
Private Const ICMP_STATUS_REQUEST_TIMED_OUT = 11010                 'Request Timed Out
Private Const ICMP_STATUS_BAD_REQUEST = 11011                       'Bad Request
Private Const ICMP_STATUS_BAD_ROUTE = 11012                         'Bad Route
Private Const ICMP_STATUS_TTL_EXPIRED_TRANSIT = 11013               'TimeToLive Expired Transit
Private Const ICMP_STATUS_TTL_EXPIRED_REASSEMBLY = 11014            'TimeToLive Expired Reassembly
Private Const ICMP_STATUS_PARAMETER = 11015                         'Parameter Problem
Private Const ICMP_STATUS_SOURCE_QUENCH = 11016                     'Source Quench
Private Const ICMP_STATUS_OPTION_TOO_BIG = 11017                    'Option Too Big
Private Const ICMP_STATUS_BAD_DESTINATION = 11018                   'Bad Destination
Private Const ICMP_STATUS_NEGOTIATING_IPSEC = 11032                 'Negotiating IPSEC
Private Const ICMP_STATUS_GENERAL_FAILURE = 11050                   'General Failure

Public Const WINSOCK_ERROR = "Windows Sockets not responding correctly."
Public Const INADDR_NONE As Long = &HFFFFFFFF
Public Const WSA_SUCCESS = 0
Public Const WS_VERSION_REQD As Long = &H101

'Clean up sockets.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512

Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long

'Open the socket connection.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long

'Create a handle on which Internet Control Message Protocol (ICMP) requests can be issued.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmpcreatefile.asp
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long

'Convert a string that contains an (Ipv4) Internet Protocol dotted address into a correct address.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winsock/wsapiref_4esy.asp
Private Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal cp As String) As Long

'Close an Internet Control Message Protocol (ICMP) handle that IcmpCreateFile opens.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmpclosehandle.asp

Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long

'Information about the Windows Sockets implementation
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Private Type WSADATA
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To 256) As Byte
   szSystemStatus(0 To 128) As Byte
   iMaxSockets As Long
   iMaxUDPDG As Long
   lpVendorInfo As Long
End Type

'Send an Internet Control Message Protocol (ICMP) echo request, and then return one or more replies.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcetcpip/htm/cerefIcmpSendEcho.asp
Private Declare Function IcmpSendEcho Lib "icmp.dll" _
   (ByVal IcmpHandle As Long, _
    ByVal DestinationAddress As Long, _
    ByVal RequestData As String, _
    ByVal RequestSize As Long, _
    ByVal RequestOptions As Long, _
    ReplyBuffer As ICMP_ECHO_REPLY, _
    ByVal ReplySize As Long, _
    ByVal TimeOut As Long) As Long
 
'This structure describes the options that will be included in the header of an IP packet.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcetcpip/htm/cerefIP_OPTION_INFORMATION.asp
Private Type IP_OPTION_INFORMATION
   Ttl             As Byte
   Tos             As Byte
   Flags           As Byte
   OptionsSize     As Byte
   OptionsData     As Long
End Type

'This structure describes the data that is returned in response to an echo request.
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wcesdkr/htm/_wcesdk_icmp_echo_reply.asp
Public Type ICMP_ECHO_REPLY
   address         As Long
   Status          As Long
   RoundTripTime   As Long
   DataSize        As Long
   Reserved        As Integer
   ptrData                 As Long
   Options        As IP_OPTION_INFORMATION
   data            As String * 250
End Type
'Ftp
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
    Const PassiveConnection As Boolean = True '
  


Sub SincronitzaTangram(Cad As String, FTPPATH As String, FTPUSER As String, FTPPASS As String)
    Dim hConnection As Long, hOpen As Long, codiBotiga As String, cadArray, Dia, Sql As String, connMDB, FilePath
    Dim IdCabeTPVLin, IdCabeTPV, rs As rdoResultset, Rs2 As rdoResultset, Rs3 As ADODB.Recordset
    Dim Total As Double, TotalDesc As Double, TotalIva1 As Double, TotalIva2 As Double, TotalIva3 As Double, CosteIva As Double
    Dim CosteIva1 As Double, CosteIva2 As Double, CosteIva3 As Double, tipoIva As Integer, client, ClientAnt
    Dim BaseImpon As Double, CuotaIva, clientsCad As String
    Dim obj_FSO As Object
    Dim Txt As Object, err
    Dim a As Integer
   On Error Resume Next
    Dim rstest As New ADODB.Recordset

    a = 0
    If a <> 1 Then
    'Insert TPVCabe y TPVLine
    ClientAnt = ""
    Dia = ""
    clientsCad = ""
    cadArray = Split(Cad, "|")
    If UBound(cadArray) > 0 Then
      Dia = cadArray(0)
      Dia = Replace(Dia, "[", "")
      Dia = Replace(Dia, "]", "")
      Dia = FormatDateTime(Dia, vbShortDate)
      clientsCad = cadArray(1)
      clientsCad = Replace(clientsCad, "[", "")
      clientsCad = Replace(clientsCad, "]", "")
    End If
    'Codi botiga
    If codiBotiga = "" Then
        Set rs = Db.OpenResultset("select valor from ConstantsClient where Codi='" & Left(Right(clientsCad, Len(clientsCad) - 2), Len(Right(clientsCad, Len(clientsCad) - 2)) - 2) & "' and Variable='CodiFtp'")
        If Not rs.EOF Then
            codiBotiga = rs("valor")
        Else
            codiBotiga = Left(Right(clientsCad, Len(clientsCad) - 2), Len(Right(clientsCad, Len(clientsCad) - 2)) - 2)
        End If
    End If
    'Copia mdb
    FilePath = AppPath & "\Tmp\GESTIONT" & codiBotiga & ".MDB"
    FileCopy AppPath & "\Tmp\GESTIONT_VACIO.MDB", FilePath
    If Dia <> "" Then
        Sql = "Select case isnull(cast(ap.valor as varchar),a.codi) when '' then a.codi else isnull(cast(ap.valor as varchar),a.codi) end codigointerno, [TimeStamp] DiaDesti2,YEAR([TimeStamp]) Anyo,"
        Sql = Sql & "s.*,a.NOM,a.PREU,a.PreuMajor,i.iva,isnull(ap.CodiArticle,a.codi) CodiArticle2 "
        Sql = Sql & "from [" & DonamNomTaulaServit(CDate(Dia)) & "] s "
        Sql = Sql & "left join Articles a on (a.Codi=s.CodiArticle)"
        Sql = Sql & "left join ArticlesPropietats ap on (a.Codi=ap.CodiArticle and ap.Variable='CODI_PROD') "
        Sql = Sql & "left join TipusIva i on (a.TipoIva=i.Tipus) "
        Sql = Sql & "where QuantitatServida<>0 and client in " & clientsCad & " "
        Sql = Sql & " and ISNULL(a.codi,0) <> 0"
        
        If CInt(Hour(Now)) >= 16 Then
            Sql = Sql & "and viatge='Tarda' "
        Else
            Sql = Sql & "and viatge='Mati' "
        End If
        Set Rs2 = Db.OpenResultset(Sql)
        If Not Rs2.EOF Then
             While Not Rs2.EOF
                 client = Rs2("client")
                 If client <> ClientAnt Then
                     'Codi botiga
                     Set rs = Db.OpenResultset("select valor from ConstantsClient where Codi='" & client & "' and Variable='CodiFtp'")
                     If Not rs.EOF Then codiBotiga = rs("valor")
                     'Cadena conexiom ; recMDB(sqlMDB)
                     Set connMDB = CreateObject("ADODB.Connection")
                     connMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & AppPath & "\Tmp\GESTIONT" & codiBotiga & ".mdb;User Id=admin;Password=;Persist Security Info=False"
                     'Insert TPVCabe
                     If ClientAnt <> "" Then 'Update cabecera en curso sino es la primera
                         Sql = "UPDATE [TempTPVCabe] set [Bruto]=" & Total & ",[Bruto_Dto]=" & TotalDesc & ","
                         Sql = Sql & "[BaseImpon1]=" & TotalIva1 & ",[CuotaIva1]=" & CosteIva1 & ","
                         Sql = Sql & "[BaseImpon2]=" & TotalIva2 & ",[CuotaIva2]=" & CosteIva2 & ","
                         Sql = Sql & "[BaseImpon3]=" & TotalIva3 & ",[CuotaIva3]=" & CosteIva3 & " "
                         Sql = Sql & "where [IdCabeTPV]= " & IdCabeTPV & " "
                         'commmdb.ActiveConnection = connMDB
                         'commmdb.CommandTimeout = 1500
                         'commmdb.CommandText = sql
                         connMDB.Execute Sql
                         'Set Rs3 = recMDB(sql)
                     End If
                     'Nueva cabecera
                     Sql = " INSERT into [TempTPVCabe] ([Ejercicio],[CodCli],[IdCaja],[Serie],[NumTiquet],[NumCliTienda],[Fecha],[Arqueado],[Bruto],[DescuentoLin],[Bruto_Dto],"
                     Sql = Sql & "[PorProntoPago],[DescuentoPP],[Parcial1],[TipoIva1],[BaseImpon1],[PorIva1],[CuotaIva1],[PorRec1],[CuotaRecEqu1],[TipoIva2],[BaseImpon2],[PorIva2],[CuotaIva2],"
                     Sql = Sql & "[PorRec2],[CuotaRecEqu2],[TipoIva3],[BaseImpon3],[PorIva3],[CuotaIva3],[PorRec3],[CuotaRecEqu3],[Liquido],[BaseComision],[IdOperario],[Hora],[FormaPago],[ImportePago],"
                     Sql = Sql & "[Vencimiento],[Listado],[Efectivo],[LiquidoEuros],[IdTiquetReducido],[RecibidoEuros],[RecibidoPtas],[EfectivoEuros],[NumTarifa],[IdAgencia],[Iva],[RecEqu],[CodCliTienda],"
                     Sql = Sql & "[NumTiquetReg],[SolicitadoCliente],[PasadoEstadis],[PagoDeuda]) values ("
                     'sql = sql & Rs2("Anyo") & ",'" & CodiBotiga & "',1,19,1,1,'" & Rs2("DiaDesti2") & "',1," & Total & ",0," & TotalDesc & ","
                     'sql = sql & "0,0,0,4," & TotalIva1 & ",4," & CosteIva1 & ",0,0,8," & TotalIva2 & ",8," & CosteIva2 & ",0,0,18," & TotalIva3 & ",18," & CosteIva3 & ",0,0,"
                     'sql = sql & Total & ",0,0,'12:00',0,0,'" & Rs2("DiaDesti2") & "',1,0,0,0,0,0,0,0,'0',0,0,'0',0,1,0,1) "
                     Sql = Sql & Year(Dia) & ",'" & codiBotiga & "',1,19,1,1,'" & Dia & "',0,0,0,0,"
                     Sql = Sql & "0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,"
                     Sql = Sql & "0,0,501,'12:00',0,0,'" & Dia & "',0,0,0,0,0,0,0,3,'0',0,0,'0',0,0,0,1) "
                     'Set Rs3 = recMDB(sql)
                     connMDB.Execute Sql
                     
                     Total = 0
                     TotalDesc = 0
                     
                     TotalIva1 = 0
                     TotalIva2 = 0
                     TotalIva3 = 0
                     CosteIva1 = 0
                     CosteIva2 = 0
                     CosteIva3 = 0
                     ClientAnt = client
                     Sql = "Select top 1 [IdCabeTPV] as num from [TempTPVCabe] order by [IdCabeTPV] desc"
                     rstest.ActiveConnection = connMDB
                     'Set Rs3 = recMDB(sql)
                     rstest.Source = Sql
                     rstest.Open
                     
                     If Not rstest.EOF Then IdCabeTPV = rstest!Num
                     rstest.Close
                 End If
                 tipoIva = Rs2("iva")
                 BaseImpon = CDbl(Rs2("Preu") * Rs2("QuantitatServida"))
                 CosteIva = ((CDbl(Rs2("Preu") * Rs2("QuantitatServida"))) * tipoIva) * 100
                 'Insert TPVLine
                 Sql = " INSERT into [TempTPVLine] ([IdCabeTPV],[IdArticulo],[Descripcion1],[Precio],[UnidadesEntregadas],[Bruto],[PorDescu],[ImpDescuento],[Parcial1],"
                 Sql = Sql & "[TipoIva],[BaseImpon],[CuotaIva],[CuotaRecEqu],[ImpRetencion],[Liquido],[Familia],[CodArt],[Comentario],[ImpCocina],[CobradoEnPartes],"
                 Sql = Sql & "[PrecioSinOferta],[EsRegalo],[CodArtGenRegalo],[NuevoRegistro]) values ("
                 'sql = sql & IdCabeTPV & "," & Rs2("CodiArticle2") & ",'" & Rs2("nom") & "'," & Rs2("preu") & "," & Rs2("QuantitatServida") & "," & BaseImpon & ","
                 'sql = sql & "0,0," & Round(BaseImpon, 2) & "," & TipoIva & "," & BaseImpon & "," & CosteIva & ",0,0," & Round(BaseImpon, 2) & ",0," & Rs2("CodiArticle2") & ",'0',0,0,0,0,'0',0)"
                 Sql = Sql & IdCabeTPV & "," & Rs2("CodiArticle2") & ",'" & Rs2("nom") & "'," & Rs2("preu") & "," & Rs2("QuantitatServida") & ",0,"
                 Sql = Sql & "0,0,0," & tipoIva & ",0,0,0,0," & BaseImpon & ",0," & Rs2("codigointerno") & ",'0',0,0,0,0,'0',0)"
                 'Set Rs3 = recMDB(sql)
                 connMDB.Execute Sql
                 Total = Total + BaseImpon
                 Select Case tipoIva
                   Case "4"
                       TotalIva1 = TotalIva1 + BaseImpon
                       CosteIva1 = CosteIva1 + CosteIva
                   Case "8"
                       TotalIva2 = TotalIva2 + BaseImpon
                       CosteIva2 = CosteIva2 + CosteIva
                   Case "18"
                       TotalIva3 = TotalIva3 + BaseImpon
                       CosteIva3 = CosteIva3 + CosteIva
                 End Select
                 Rs2.MoveNext
             Wend
             'Actualizamos ultima cabecera
             Sql = "UPDATE [TempTPVCabe] set [Bruto]=" & Total & ",[Bruto_Dto]=" & TotalDesc & ","
             Sql = Sql & "[BaseImpon1]=" & TotalIva1 & ",[CuotaIva1]=" & CosteIva1 & ","
             Sql = Sql & "[BaseImpon2]=" & TotalIva2 & ",[CuotaIva2]=" & CosteIva2 & ","
             Sql = Sql & "[BaseImpon3]=" & TotalIva3 & ",[CuotaIva3]=" & CosteIva3
             Sql = Sql & " where [IdCabeTPV]=" & IdCabeTPV & " "
            'Set Rs3 = recMDB(sql)
            connMDB.Execute Sql
            '************************* TANQUEM CONEXIO AL FITXER MDB.
             connMDB.Close
             Set connMDB = Nothing
             '************************* TANQUEM CONEXIO AL FITXER MDB.
        End If
    End If
    err = 0
    '**************************************************COPIA A DIRECTORIOS NUEVOS
    'Copiamos mdb y txt en la nueva estructura de directorios
    'FilePath = AppPath & "\Tmp\GESTIONT" & codiBotiga & ".MDB"
    'FileCopy FilePath, AppPath & "\Tmp\4300001261\4300001261\" & codiBotiga & "1\GESTIONT.MDB"
    'Crear txt
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    'Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\4300001261\4300001261\" & codiBotiga & "1\TESTIGO.txt", True)
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\TESTIGO.txt", True)
    Txt.WriteLine "Importar"
    Txt.Close
    '**************************************************SUBIDA FTP
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\TESTIGO.txt", True)
    Txt.WriteLine "Importar"
    Txt.Close
    Kill "X:\FtpGdt\magatzem\armengol\ftp\4300001261\4300001261\" & codiBotiga & "1\GESTIONT.MDB"
    FileCopy FilePath, "X:\FtpGdt\magatzem\armengol\ftp\4300001261\4300001261\" & codiBotiga & "1\GESTIONT.MDB"
    FileCopy AppPath & "\Tmp\TESTIGO.txt", "X:\FtpGdt\magatzem\armengol\ftp\4300001261\4300001261\" & codiBotiga & "1\TESTIGO.TXT"
    'FTPPATH = Mid(FTPPATH, 2, Len(FTPPATH) - 2)
    'FTPPATH = "217.113.245.74"
    'FTPPATH = "localhost"
    'FTPUSER = Mid(FTPUSER, 2, Len(FTPUSER) - 2)
    'FTPPASS = Mid(FTPPASS, 2, Len(FTPPASS) - 2)
    'Pujar fitxer mdb
    'hOpen = InternetOpen("Puja MDB", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    'hConnection = InternetConnect(hOpen, FTPPATH, INTERNET_DEFAULT_FTP_PORT, FTPUSER, FTPPASS, INTERNET_SERVICE_FTP, 0, 0)
    'If hOpen = 0 Or hConnection = 0 Then 'No s'ha fet be la conexió
        'ShowErrorFTP
    'Else
        'Pujar MDB
     '   FtpPutFile hConnection, FilePath, "/4300001261/4300001261/" & codiBotiga & "1/GESTIONT.MDB", FTP_TRANSFER_TYPE_UNKNOWN, 0
        'ShowErrorFTP
        'Pujar txt
        'FtpPutFile hConnection, AppPath & "\Tmp\TESTIGO.TXT", "/4300001261/4300001261/" & codiBotiga & "1/TESTIGO.TXT", FTP_TRANSFER_TYPE_UNKNOWN, 0
        'ShowErrorFTP
    'End If
    'InternetCloseHandle hConnection
    'InternetCloseHandle hOpen
    Set obj_FSO = Nothing
    Set Txt = Nothing
    If err = 0 Then
        'Borrar fitxers temp
        Kill AppPath & "\Tmp\TESTIGO.txt"
        'Kill FilePath
    End If
End If
End Sub
    Sub ShowErrorFTP()
    Dim lErr As Long, sErr As String, lenBuf As Long
    InternetGetLastResponseInfo lErr, sErr, lenBuf
    sErr = String(lenBuf, 0)
    InternetGetLastResponseInfo lErr, sErr, lenBuf
    If lErr <> 0 Then MsgBox "Error " + CStr(lErr) + ": " + sErr, vbOKOnly + vbCritical
End Sub


Function TePing(strIPAddress As String) As Boolean

   Dim Reply As ICMP_ECHO_REPLY
   Dim lngSuccess As Long

   
   TePing = False
   'Get the sockets ready.
   If SocketsInitialize() Then
      
    'Address to ping
   ' strIPAddress = "192.168.1.1"
    
    'Ping the IP that is passing the address and get a reply.
    lngSuccess = ping(strIPAddress, Reply)
    If EvaluatePingResponse(lngSuccess) = "Success!" Then TePing = True
    
    'Display the results.
'    Debug.Print "Address to Ping: " & strIPAddress
'    Debug.Print "Raw ICMP code: " & lngSuccess
'    Debug.Print "Ping Response Message : " & EvaluatePingResponse(lngSuccess)
'    Debug.Print "Time : " & Reply.RoundTripTime & " ms"
      
    'Clean up the sockets.
    SocketsCleanup
      
   Else
   
   'Winsock error failure, initializing the sockets.
 '  Debug.Print WINSOCK_ERROR
   
   End If
   
End Function

'Clean up the sockets.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Public Sub SocketsCleanup()
   
   WSACleanup
    
End Sub


'Convert the ping response to a message that you can read easily from constants.
'For more information about these constants, visit the following Microsoft Web site:
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/win32_pingstatus.asp

Public Function EvaluatePingResponse(PingResponse As Long) As String

  Select Case PingResponse
    
  'Success
  Case ICMP_SUCCESS: EvaluatePingResponse = "Success!"
            
  'Some error occurred
  Case ICMP_STATUS_BUFFER_TO_SMALL:    EvaluatePingResponse = "Buffer Too Small"
  Case ICMP_STATUS_DESTINATION_NET_UNREACH: EvaluatePingResponse = "Destination Net Unreachable"
  Case ICMP_STATUS_DESTINATION_HOST_UNREACH: EvaluatePingResponse = "Destination Host Unreachable"
  Case ICMP_STATUS_DESTINATION_PROTOCOL_UNREACH: EvaluatePingResponse = "Destination Protocol Unreachable"
  Case ICMP_STATUS_DESTINATION_PORT_UNREACH: EvaluatePingResponse = "Destination Port Unreachable"
  Case ICMP_STATUS_NO_RESOURCE: EvaluatePingResponse = "No Resources"
  Case ICMP_STATUS_BAD_OPTION: EvaluatePingResponse = "Bad Option"
  Case ICMP_STATUS_HARDWARE_ERROR: EvaluatePingResponse = "Hardware Error"
  Case ICMP_STATUS_LARGE_PACKET: EvaluatePingResponse = "Packet Too Big"
  Case ICMP_STATUS_REQUEST_TIMED_OUT: EvaluatePingResponse = "Request Timed Out"
  Case ICMP_STATUS_BAD_REQUEST: EvaluatePingResponse = "Bad Request"
  Case ICMP_STATUS_BAD_ROUTE: EvaluatePingResponse = "Bad Route"
  Case ICMP_STATUS_TTL_EXPIRED_TRANSIT: EvaluatePingResponse = "TimeToLive Expired Transit"
  Case ICMP_STATUS_TTL_EXPIRED_REASSEMBLY: EvaluatePingResponse = "TimeToLive Expired Reassembly"
  Case ICMP_STATUS_PARAMETER: EvaluatePingResponse = "Parameter Problem"
  Case ICMP_STATUS_SOURCE_QUENCH: EvaluatePingResponse = "Source Quench"
  Case ICMP_STATUS_OPTION_TOO_BIG: EvaluatePingResponse = "Option Too Big"
  Case ICMP_STATUS_BAD_DESTINATION: EvaluatePingResponse = "Bad Destination"
  Case ICMP_STATUS_NEGOTIATING_IPSEC: EvaluatePingResponse = "Negotiating IPSEC"
  Case ICMP_STATUS_GENERAL_FAILURE: EvaluatePingResponse = "General Failure"
            
  'Unknown error occurred
  Case Else: EvaluatePingResponse = "Unknown Response"
        
  End Select

End Function

'-- Ping a string representation of an IP address.
' -- Return a reply.
' -- Return long code.
Public Function ping(sAddress As String, Reply As ICMP_ECHO_REPLY) As Long

Dim hIcmp As Long
Dim lAddress As Long
Dim lTimeOut As Long
Dim StringToSend As String

'Short string of data to send
StringToSend = "hello"

'ICMP (ping) timeout
lTimeOut = 50 'ms

'Convert string address to a long representation.
lAddress = inet_addr(sAddress)

If (lAddress <> -1) And (lAddress <> 0) Then
        
    'Create the handle for ICMP requests.
    hIcmp = IcmpCreateFile()
    
    If hIcmp Then
        'Ping the destination IP address.
        Call IcmpSendEcho(hIcmp, lAddress, StringToSend, Len(StringToSend), 0, Reply, Len(Reply), lTimeOut)

        'Reply status
        ping = Reply.Status
        
        'Close the Icmp handle.
        IcmpCloseHandle hIcmp
    Else
        Debug.Print "failure opening icmp handle."
        ping = -1
    End If
Else
    ping = -1
End If

End Function


'Get the sockets ready.
'http://support.microsoft.com/default.aspx?scid=kb;EN-US;q154512
Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA

   SocketsInitialize = WSAStartup(WS_VERSION_REQD, WSAD) = ICMP_SUCCESS

End Function



Function NetejaNum(a As String) As String
    
    NetejaNum = Trim(a)
    
    While InStr(NetejaNum, "€") > 0
        NetejaNum = Left(NetejaNum, InStr(NetejaNum, "€") - 1) & Right(NetejaNum, Len(NetejaNum) - InStr(NetejaNum, "€"))
    Wend
    
    While InStr(NetejaNum, ",") > 0
        NetejaNum = Left(NetejaNum, InStr(NetejaNum, ",") - 1) & "." & Right(NetejaNum, Len(NetejaNum) - InStr(NetejaNum, ","))
    Wend
    NetejaNum = Trim(NetejaNum)
    
End Function

Function PillaNominasHtmlCar(f) As String
    Dim aa
    
    PillaNominasHtmlCar = ""
    While Not EOF(f) And PillaNominasHtmlCar = ""
        Line Input #f, aa
        If Left(aa, 3) = "<P " And Right(aa, 4) = "</P>" And Not Right(aa, 5) = "></P>" Then
           PillaNominasHtmlCar = Join(Split(Split(Split(aa, ">")(1), "<")(0), "&nbsp;"), " ")
           If IsNumeric(PillaNominasHtmlCar) Then
            PillaNominasHtmlCar = Join(Split(PillaNominasHtmlCar, "."), "")
            PillaNominasHtmlCar = Join(Split(PillaNominasHtmlCar, ","), ".")
           End If
           Debug.Print PillaNominasHtmlCar
        End If
    Wend
    
End Function

Sub CarregaClientsXls(nomfile)
   Dim MsExcel As Excel.Application, Libro As Excel.Workbook, Hoja, Qr1 As rdoQuery, Qr2  As rdoQuery, i, Fi As Boolean, Codi, Variable, Valor, K, Codis As String, rs, TarifasN(), TarifasC(), TarifaCodi, p1, P2, Fami1, Fami2, Fami3, CasiNum As String, TecN(), TecC()
   Dim botiga As Double
   
    If Not frmSplash.Debugant Then On Error GoTo nok

    InformaMiss "Excel CarregantClients"
    Set MsExcel = CreateObject("Excel.Application")
    MsExcel.Visible = frmSplash.Debugant
    MsExcel.Workbooks.Open fileName:=nomfile
    Set Libro = MsExcel.Workbooks(1)
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Codis = ""
    
    If InStr(nomfile, "INVENTARI") Then
            Dim H As Integer
            'Copia seguretat
            ExecutaComandaSql "Select * into [" & NomTaulaInventari(Now) & "" & Format(Now, "yyyymmddhhnnss") & "]  from  [" & NomTaulaInventari(Now) & "]"
            
            For H = 1 To Libro.Sheets.Count
                Set Hoja = Libro.Sheets(H)
                
                Set rs = Db.OpenResultset("Select * From Clients where nom = '" & Trim(Hoja.Name) & "' or [Nom Llarg] = '" & Trim(Hoja.Name) & "' ")
                If Not rs.EOF Then
                    botiga = rs("Codi")
                    i = 2
                    While Not Hoja.Cells(i, 5).Value = ""
                        Valor = Hoja.Cells(i, 10).Value
                        If Len(Valor) > 0 Then
                            Set rs = Db.OpenResultset("Select * From Articles Where codi = '" & Trim(Hoja.Cells(i, 5).Value) & "' ")
                            If Not rs.EOF Then
                                Codi = rs("codi")
                                ExecutaComandaSql "Insert Into [" & NomTaulaInventari(Now) & "] (Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta) Values(" & botiga & ", getdate(), 1, 1, '', " & Codi & "," & Valor & ",0,1) "
                            End If
                        End If
                        i = i + 1
                    Wend
                End If
            Next
    Else
        Select Case Trim(Hoja.Name)
            Case "Cálculo Inventario", "Calculo Inventario"
                Fi = False
                i = 3
                Dim data As Date
                data = Hoja.Cells(4, 3).Value
                
                Set rs = Db.OpenResultset("Select * From Clients where nom = '" & Hoja.Cells(2, 3).Value & "' or [Nom Llarg] = '" & Hoja.Cells(2, 3).Value & "' ")
                If rs.EOF Then Exit Sub
                botiga = rs("Codi")
                
                ExecutaComandaSql "Select * into [" & NomTaulaInventari(data) & "" & Format(Now, "yyyymmddhhnnss") & "]  from  [" & NomTaulaInventari(data) & "]"
                i = 16
                While Not Hoja.Cells(i, 3).Value = ""
                    Valor = Hoja.Cells(i, 6).Value
                    If Len(Valor) > 0 Then
                        Set rs = Db.OpenResultset("Select * From Articles Where nom = '" & Trim(Hoja.Cells(i, 3).Value) & "' ")
                        If Not rs.EOF Then
                            InformaMiss "Excel Inventari " & Hoja.Cells(i, 3).Value
                            Codi = rs("Codi")
                            ExecutaComandaSql "Insert Into [" & NomTaulaInventari(data) & "] (Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta) Values(" & botiga & ",'" & data & "',1,1,''," & Codi & "," & Valor & ",0,2) "
                        End If
                    End If
                    i = i + 1
                Wend
            Case "Clients"
                Fi = False
                i = 3
                ExecutaComandaSql "Select * into [Clients" & Format(Now, "yyyymmddhhnnss") & "]  from Clients "
                ExecutaComandaSql "Select * into [ConstantsClient" & Format(Now, "yyyymmddhhnnss") & "]  from ConstantsClient "
                
                While Not Hoja.Cells(i, 2).Value = ""
                    InformaMiss "Excel CarregantClients " & i
                    Codi = Hoja.Cells(i, 1).Value
                    If Codi = "" Then
                        Codi = DonamSql("Select Max(Codi) from Clients ") + 1
                        If Codi = "" Then Codi = 1
                        ExecutaComandaSql "Insert into Clients (Codi,nom) Values (" & Codi & ",'') "
                    End If
                    
                    Set rs = Db.OpenResultset("Select Codi From Clients Where codi = " & Codi & " ")
                    If rs.EOF Then
                        ExecutaComandaSql "Insert into clients (codi,nom) Values (" & Codi & ",'') "
                    Else
                        rs.MoveNext
                        If Not rs.EOF Then  ' !!!!!!!!!!!!!!!!! TENIM REPESSSS
                            ExecutaComandaSql "Delete Clients Where Codi = " & Codi
                            ExecutaComandaSql "Insert into clients (codi,nom) Values (" & Codi & ",'') "
                        End If
                    End If
                    
                    If Not Codis = "" Then Codis = Codis & ","
                    Codis = Codis & Codi
                    K = 2
                    While Not Hoja.Cells(1, K).Value = ""
                        DoEvents
                        Variable = Hoja.Cells(1, K).Value
                        Valor = Hoja.Cells(i, K).Value
                        Valor = Join(Split(Valor, Chr(9)), "")  ' Treu tabuladors
                        Valor = Join(Split(Valor, "'"), " ")    ' Treu apostrof
                        Valor = Join(Split(Valor, """"), " ")   ' Treu Cometes
                        
                        Select Case Variable
                            Case "NomCurt"
                                ExecutaComandaSql "Update Clients Set Nom = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "Nif"
                                ExecutaComandaSql "Update Clients Set nif = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "Adresa"
                                ExecutaComandaSql "Update Clients Set Adresa = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "ciutat"
                                ExecutaComandaSql "Update Clients Set Ciutat = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "cp"
                                ExecutaComandaSql "Update Clients Set Cp = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "lliure"
                                ExecutaComandaSql "Update Clients Set lliure = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "Nom Llarg"
                                ExecutaComandaSql "Update Clients Set [Nom Llarg] = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "Tipus Iva"
                                ExecutaComandaSql "Update Clients Set [Tipus Iva] = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "Preu Base"
                                ExecutaComandaSql "Update Clients Set [Preu Base] = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "Desconte"
    '                            ExecutaComandaSql "Update Clients Set [Desconte ProntoPago] = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "ProntoPago"
                                ExecutaComandaSql "Update Clients Set [Desconte ProntoPago] = '" & Valor & "' Where Codi = " & Codi & " "
                            Case "Tarifa"
                                TarifaCodi = DonamSql("select top 1 TarifaCodi from tarifesespecials where tarifanom = '" & Valor & "' ")
                                If Not TarifaCodi = "" Then
                                    ExecutaComandaSql "Update Clients Set [Desconte 5] = '" & TarifaCodi & "' Where Codi = " & Codi & " "
                                Else
                                Valor = Valor
                                End If
                            Case "AlbaraValorat", "Grup_client", "EsCalaix", "Fax", "Drebuts", "COPIES_ALB", "Nrebuts", "Adr_Entrega", "Acreedor", "DescTe", "EsClient", "CFINAL", "Email", "DiaPagament", "codiContable", "P_Contacte", "Venciment", "Per_Facturacio", "Idioma", "NomClientFactura", "Tel", "CompteCorrent", "FormaPago", "FormaPagoLlista", "NoDevolucions", "NoPagaEnTienda", "AlbaransValorats", "CodiContable"
                                ExecutaComandaSql "Delete constantsclient Where Codi = " & Codi & " and Variable = '" & Variable & "'"
                                If Not Valor = "" Then ExecutaComandaSql "Insert into constantsclient (Codi,Variable,Valor) Values ('" & Codi & "','" & Variable & "','" & Valor & "')"
                        End Select
                        K = K + 1
                    Wend
                    i = i + 1
                Wend
                ExecutaComandaSql "Delete Clients Where not codi in (" & Codis & ") "
                ExecutaComandaSql "Update ConstantsClient  set   variable = 'CodiContable' where variable = 'codiContable'    "
                ExecutaComandaSql "Delete ConstantsClient  Where variable = 'CodiContable' And  not codi in (" & Codis & ")  "
                
                ExecutaComandaSql "update clients Set [Desconte prontopago] = 0 where [Desconte prontopago] is null"
                ExecutaComandaSql "update clients Set [Desconte 1] = 0 where [Desconte 1] is null"
                ExecutaComandaSql "update clients Set [Desconte 2] = 0 where [Desconte 2] is null"
                ExecutaComandaSql "update clients Set [Desconte 3] = 0 where [Desconte 3] is null"
                ExecutaComandaSql "update clients Set [Desconte 4] = 0 where [Desconte 4] is null"
                ExecutaComandaSql "update clients Set [Desconte 5] = 0 where [Desconte 5] is null"
                ExecutaComandaSql "update clients Set [AlbaraValorat] = 0 where [AlbaraValorat] is null"
                
            Case "Preus"
                
                Set Qr1 = Db.CreateQuery("", "Update Articles Set Preu = ? Where codi =  ? ")
                Set Qr2 = Db.CreateQuery("", "Update Articles Set [PreuMajor] = ? Where codi =  ? ")
    
                ExecutaComandaSql "Select * into [articleshistorial" & Format(Now, "yyyymmddhhnnss") & "]  from articleshistorial "
                ExecutaComandaSql "Select * into [Articles" & Format(Now, "yyyymmddhhnnss") & "]  from articles "
                ExecutaComandaSql "Select * into [ArticlesPropietats" & Format(Now, "yyyymmddhhnnss") & "]  from ArticlesPropietats "
                ExecutaComandaSql "Select * into [TarifesEspecials" & Format(Now, "yyyymmddhhnnss") & "]  from TarifesEspecials "
                ExecutaComandaSql "Select * into [tarifeshistorial" & Format(Now, "yyyymmddhhnnss") & "]  from tarifeshistorial "
                ExecutaComandaSql "Select * into [CodisBarres" & Format(Now, "yyyymmddhhnnss") & "]  from CodisBarres "
                ExecutaComandaSql "Select * into [tarifesespecialsclients" & Format(Now, "yyyymmddhhnnss") & "]  from tarifesespecialsclients "
                
    
                
                'Guardem Articles Per Historial
                ExecutaComandaSql "insert into articleshistorial  (codi,nom,preu,preumajor,Desconte,EsSumable,Familia,CodiGenetic,TipoIva,fechaModif,UsuarioModif) select codi,nom,preu,preumajor,Desconte,EsSumable,Familia,CodiGenetic,TipoIva,GetDate(),'Excel' from articles "
                'Guardem Tarifes Per Historial
                ExecutaComandaSql "Insert into tarifeshistorial (TarifaCodi,TarifaNom,Codi,Preu,PreuMajor,fechaModif,UsuarioModif) select TarifaCodi,TarifaNom,Codi,Preu,PreuMajor,GetDate(),'Excel' from TarifesEspecials "
    
    
                ReDim TarifasN(0), TarifasC(0)
                i = 1
                While Hoja.Cells(1, 8 + (2 * i)).Value = "Tarifa"
                    ReDim Preserve TarifasN(i), TarifasC(i)
                    TarifasC(i) = -1
                    TarifasN(i) = Hoja.Cells(1, 9 + (2 * i)).Value
                    Set rs = Db.OpenResultset("Select Distinct TarifaCodi From TarifesEspecials Where  Tarifanom = '" & TarifasN(i) & "' ")
                    If rs.EOF Then
                        Set rs = Db.OpenResultset("Select isnull(max(tarifacodi),0) + 1 as TarifaCodi  From TarifesEspecials ")
                        TarifasC(i) = rs("TarifaCodi")
                        ExecutaComandaSql "insert into TarifesEspecials (Tarifacodi,Tarifanom,codi,preu,preumajor) values (" & TarifasC(i) & ",'" & TarifasN(i) & "',-1,1,1) "
                    Else
                        TarifasC(i) = rs("TarifaCodi")
                    End If
                    i = i + 1
                Wend
                
                ReDim TecN(0), TecC(0)
                While Hoja.Cells(1, 8 + (2 * i)).Value = "Client"
                    ReDim Preserve TecN(i), TecC(i)
                    TecC(i) = -1
                    TecN(i) = Hoja.Cells(2, 8 + (2 * i)).Value
                    Set rs = Db.OpenResultset("Select codi from clients where nom = '" & TecN(i) & "' ")
                    If rs.EOF Then
                        TecC(i) = -1
                    Else
                        TecC(i) = rs("Codi")
                    End If
                    i = i + 1
                Wend
                
                
                i = 3
                While Not (Hoja.Cells(i, 1).Value = "" And Hoja.Cells(i, 2).Value = "" And Hoja.Cells(i, 3).Value = "")
    ' codi,nom,preu,preumajor,desconte,essumable,familia,codigenetic,tipoiva,nodescontesespecials
    'If Hoja.Cells(i, 3).Value = "Pa galleg 280 grms ( 20 u )" Then
    'If i = 1089 Then
    '    i = i
    'End If
                    Codi = Hoja.Cells(i, 2).Value
                    If Codi = "" Or Not IsNumeric(Codi) Then
                        Codi = DonamSql("Select Max(Codi) from Articles ") + 1
                        If Codi = "" Then Codi = 1
                        ExecutaComandaSql "Insert into Articles (codi,nom,preu,preumajor,desconte,essumable,familia,codigenetic,tipoiva,nodescontesespecials) Values (" & Codi & ",'',0,0,0,0,''," & Codi & ",1,0) "
                    End If
                    
                    Set rs = Db.OpenResultset("Select Codi From Articles Where codi = " & Codi & " ")
                    If rs.EOF Then
                        ExecutaComandaSql "Insert into Articles (codi,nom,preu,preumajor,desconte,essumable,familia,codigenetic,tipoiva,nodescontesespecials) Values (" & Codi & ",'',0,0,0,0,''," & Codi & ",1,0) "
                    Else
                        rs.MoveNext
                        If Not rs.EOF Then  ' !!!!!!!!!!!!!!!!! TENIM REPESSSS
                            ExecutaComandaSql "Delete Articles Where Codi = " & Codi
                            ExecutaComandaSql "Insert into Articles (codi,nom,preu,preumajor,desconte,essumable,familia,codigenetic,tipoiva,nodescontesespecials) Values (" & Codi & ",'',0,0,0,0,''," & Codi & ",1,0) "
                        End If
                    End If
                    
                    If Not Codis = "" Then Codis = Codis & ","
                    Codis = Codis & Codi
                    K = 2
                    
                    ' El nom
                    Dim nom As String
                    nom = Normalitza(Trim(Hoja.Cells(i, 3).Value))
                    InformaMiss "Excel Carregant Preus " & i
                    ExecutaComandaSql "Update Articles set nom = '" & nom & "' Where codi = " & Codi & " "
                    'El Codi Mp
                    ExecutaComandaSql "Delete articlespropietats where Variable = 'CODI_PROD' And codiArticle = " & Codi & " "
                    If Not Hoja.Cells(i, 1).Value = "" Then ExecutaComandaSql "insert into articlespropietats  (CodiArticle,Variable,Valor) values (" & Codi & ",'CODI_PROD','" & Hoja.Cells(i, 1).Value & "') "
                    If Len(Hoja.Cells(i, 1).Value) = 13 And IsNumeric(Hoja.Cells(i, 1).Value) Then
                        ExecutaComandaSql "Delete CodisBarres Where Producte = " & Codi & " "
                        ExecutaComandaSql "Delete CodisBarres Where Codi = " & Hoja.Cells(i, 1).Value & " "
                        ExecutaComandaSql "insert into CodisBarres (Producte,Codi) values (" & Codi & "," & Hoja.Cells(i, 1).Value & ") "
                    End If
                    
                    ' El Preu
                    
                    If Not CStr(Hoja.Cells(i, 8).Value) = "" Then
                        CasiNum = NetejaNum(CStr(Hoja.Cells(i, 8).Value))
                        If IsNumeric(CasiNum) Then
                            Qr1.rdoParameters(0) = CasiNum
                            Qr1.rdoParameters(1) = Codi
                            Qr1.Execute
                        End If
                    End If
                    
                    ' El Preu Major
                    If Not CStr(Hoja.Cells(i, 9).Value) = "" Then
                        CasiNum = NetejaNum(CStr(Hoja.Cells(i, 9).Value))
                        If IsNumeric(CasiNum) Then
                            Qr2.rdoParameters(0) = CasiNum
                            Qr2.rdoParameters(1) = Codi
                            Qr2.Execute
                        End If
                    End If
                    ' El Iva
                    If Hoja.Cells(i, 7).Value = 4 Or Hoja.Cells(i, 7).Value = 4 Then ExecutaComandaSql "Update Articles set TipoIva = 1 Where codi = " & Codi & " "
                    If Hoja.Cells(i, 7).Value = 7 Or Hoja.Cells(i, 7).Value = 8 Then ExecutaComandaSql "Update Articles set TipoIva = 2 Where codi = " & Codi & " "
                    If Hoja.Cells(i, 7).Value = 16 Or Hoja.Cells(i, 7).Value = 18 Then ExecutaComandaSql "Update Articles set TipoIva = 3 Where codi = " & Codi & " "
                    
                    ' La Familia
                    If Not Hoja.Cells(i, 6).Value = "" Then
                        ExecutaComandaSql "Update Articles set Familia = '" & Normalitza(Hoja.Cells(i, 6).Value) & "' Where codi = " & Codi & " "
                        If Not Hoja.Cells(i, 5).Value = "" Then
                            Fami1 = Normalitza(Hoja.Cells(i, 4).Value)
                            Fami2 = Normalitza(Hoja.Cells(i, 5).Value)
                            Fami3 = Normalitza(Hoja.Cells(i, 6).Value)
                            
                            If Fami3 = Fami2 Then Fami2 = Fami2 & "."
                            If Fami3 = Fami1 Then Fami1 = Fami1 & "."
                            If Fami2 = Fami1 Then Fami1 = Fami1 & "."
                            
                            ExecutaComandaSql "Delete Families where nom = '" & Fami3 & "' "
                            ExecutaComandaSql "Insert Into Families (nom,pare,estatus,nivell,utilitza) Values ('" & Fami3 & "','" & Fami2 & "',0,3,0) "
                            If Not Hoja.Cells(i, 4).Value = "" Then
                                ExecutaComandaSql "Delete Families where nom = '" & Fami2 & "' "
                                ExecutaComandaSql "Insert Into Families (nom,pare,estatus,nivell,utilitza) Values ('" & Fami2 & "','" & Fami1 & "',0,2,0) "
                                
                                ExecutaComandaSql "Delete Families where nom = '" & Fami1 & "' "
                                ExecutaComandaSql "Insert Into Families (nom,pare,estatus,nivell,utilitza) Values ('" & Fami1 & "','Article',0,1,0) "
                             End If
                        End If
                    End If
                    
                    ExecutaComandaSql "Delete TarifesEspecials Where  Codi= " & Codi
                    For K = 1 To UBound(TarifasC)
                        p1 = 0
                        P2 = 0
                        If Not Hoja.Cells(i, 8 + (2 * K)).Value = "" Then p1 = Hoja.Cells(i, 8 + (2 * K)).Value
                        If Not Hoja.Cells(i, 9 + (2 * K)).Value = "" Then P2 = Hoja.Cells(i, 9 + (2 * K)).Value
                        If p1 > 0 Or P2 > 0 Then ExecutaComandaSql "Insert Into TarifesEspecials (TarifaCodi,TarifaNom,Codi,Preu,PreuMajor  ) Values (" & TarifasC(K) & ",'" & TarifasN(K) & "'," & Codi & ",'" & p1 & "','" & P2 & "') "
                    Next
                    
                    ExecutaComandaSql "Delete tarifesespecialsclients Where  Codi= " & Codi
                    For K = K To UBound(TecC)
                        p1 = 0
                        P2 = 0
                        If Not Hoja.Cells(i, 8 + (2 * K)).Value = "" Then p1 = Hoja.Cells(i, 8 + (2 * K)).Value
                        If Not Hoja.Cells(i, 9 + (2 * K)).Value = "" Then P2 = Hoja.Cells(i, 9 + (2 * K)).Value
                        If p1 > 0 Or P2 > 0 Then
                            ExecutaComandaSql "Insert Into tarifesespecialsclients (Client,Codi,Preu,PreuMajor ,qmin ) Values (" & TecC(K) & "," & Codi & ",'" & p1 & "','" & P2 & "',0) "
                        End If
                    Next
                    
                    i = i + 1
                Wend
                
                ExecutaComandaSql "Delete Families where nom = 'Article' "
                ExecutaComandaSql "Insert Into Families (nom,pare,estatus,nivell,utilitza) Values ('Article','',0,0,0) "
                ExecutaComandaSql "Delete Articles Where not codi in (" & Codis & ") "
                ExecutaComandaSql "Delete Articles Where nom = '' "
                ExecutaComandaSql "Delete TarifesEspecials Where codi = -1 "
                
        End Select
    
    End If
    
nok:
    TancaExcel MsExcel, Libro
End Sub


Sub CarregaDependentesXls(nomfile)
   Dim MsExcel As Excel.Application, Libro As Excel.Workbook, Hoja, Qr1 As rdoQuery, Qr2  As rdoQuery, i, Fi As Boolean, Codi, Variable, Valor, K, Codis As String, rs, TarifasN(), TarifasC(), TarifaCodi, p1, P2, Fami1, Fami2, Fami3, CasiNum As String, TecN(), TecC()
   Dim botiga As Double
   
    If Not frmSplash.Debugant Then On Error GoTo nok

    InformaMiss "Excel CarregantDependentes"
    Set MsExcel = CreateObject("Excel.Application")
    MsExcel.Visible = frmSplash.Debugant
    MsExcel.Workbooks.Open fileName:=nomfile
    Set Libro = MsExcel.Workbooks(1)
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Codis = ""
    

    Fi = False
    i = 3
    'Copia de seguretat
    ExecutaComandaSql "Select * into [Dependentes" & Format(Now, "yyyymmddhhnnss") & "]  from Dependentes "
    ExecutaComandaSql "Select * into [DependentesExtes" & Format(Now, "yyyymmddhhnnss") & "]  from DependentesExtes "
                
    While Not Hoja.Cells(i, 1).Value = ""
        InformaMiss "Excel Carregant Dependentes " & i
        Codi = Hoja.Cells(i, 1).Value
        If Codi = "" Then 'Nova dependenta
            'Codi = DonamSql("Select Max(Codi) from Clients ") + 1
            'If Codi = "" Then Codi = 1
            'ExecutaComandaSql "Insert into Clients (Codi,nom) Values (" & Codi & ",'') "
        End If
                    
        Set rs = Db.OpenResultset("Select Codi From Dependentes Where codi = " & Codi & " ")
        If rs.EOF Then
            ExecutaComandaSql "Insert into Dependentes (codi, nom) Values (" & Codi & ", '') "
        Else
            rs.MoveNext
            If Not rs.EOF Then  ' !!!!!!!!!!!!!!!!! TENIM REPESSSS
                ExecutaComandaSql "Delete Dependentes Where Codi = " & Codi
                ExecutaComandaSql "Insert into Dependentes (codi, nom) Values (" & Codi & ", '') "
            End If
        End If
                    
        If Not Codis = "" Then Codis = Codis & ","
        Codis = Codis & Codi
        K = 2
        While Not Hoja.Cells(2, K).Value = ""
            DoEvents
            Variable = Hoja.Cells(2, K).Value
            Valor = Hoja.Cells(i, K).Value
            Valor = Join(Split(Valor, Chr(9)), "")  ' Treu tabuladors
            Valor = Join(Split(Valor, "'"), " ")    ' Treu apostrof
            Valor = Join(Split(Valor, """"), " ")   ' Treu Cometes
                        
            'CODI | NOM | TIPUS | ADREÇA | TELÈFON | TLF MÒBIL | LLOC TREBALL | DNI | Carnet Manipulador | E-MAIL | DATA NAIXEMENT
            Select Case Variable
                Case "NOM"
                    ExecutaComandaSql "Update Dependentes Set NOM = '" & Valor & "' Where Codi = " & Codi & " "
                'Case "TIPUS"  'NO SE PUEDE MODIFICAR EL TIPO DE TRABAJADOR DESDE EL EXCEL!!
                '    ExecutaComandaSql "Update DependentesExtes set valor= '" & Valor & "' where Id='" & Codi & "' and Nom='TIPUSTREBALLADOR'"
                Case "ADREÇA"
                    ExecutaComandaSql "Update Dependentes Set ADREÇA = '" & Valor & "' Where Codi = " & Codi & " "
                Case "TELÈFON"
                    ExecutaComandaSql "Update Dependentes Set TELEFON = '" & Valor & "' Where Codi = " & Codi & " "
                Case "TLF MÒBIL"
                    ExecutaComandaSql "Update DependentesExtes set valor= '" & Valor & "' where Id='" & Codi & "' and Nom='TLF_MOBIL'"
                Case "LLOC TREBALL"
                    ExecutaComandaSql "Update DependentesExtes set valor= '" & Valor & "' where Id='" & Codi & "' and Nom='LLOC_TREBALL'"
                Case "DNI"
                    ExecutaComandaSql "Update DependentesExtes set valor= '" & Valor & "' where Id='" & Codi & "' and Nom='DNI'"
                Case "Carnet Manipulador"
                    ExecutaComandaSql "Update DependentesExtes set valor= '" & Valor & "' where Id='" & Codi & "' and Nom='cManipulador'"
                Case "E-MAIL"
                    ExecutaComandaSql "Update DependentesExtes set valor= '" & Valor & "' where Id='" & Codi & "' and Nom='EMAIL'"
                Case "DATA NAIXEMENT"
                    ExecutaComandaSql "Update DependentesExtes set valor= '" & Valor & "' where Id='" & Codi & "' and Nom='DATA_NAIXEMENT'"
            End Select
            K = K + 1
        Wend
        i = i + 1
    Wend
    
    'Borrar los que no estan en el fichero excel???
    ExecutaComandaSql "insert into Dependentes_Zombis (select getdate(), * from Dependentes where codi not in (" & Codis & ")) "
    ExecutaComandaSql "Delete Dependentes Where not codi in (" & Codis & ") "
                
nok:
    TancaExcel MsExcel, Libro
End Sub



Sub ImportaNominasExcel(nomfile, IdFile)
    Dim MsExcel As Excel.Application, Libro As Excel.Workbook, Hoja
    Dim empleado As String, fecha As String, data As String, cCargo As String, cAbono As String, Sql As String
    Dim i_liquid As Double, i_Tc1 As Double, i_Brut As Double, i_Irpf As Double, i_ssTre As Double
    Dim i As Integer

    If Not frmSplash.Debugant Then On Error GoTo nok
    
    InformaMiss "Excel CarregantNomines"
    Set MsExcel = CreateObject("Excel.Application")
    MsExcel.Visible = frmSplash.Debugant
    MsExcel.Workbooks.Open fileName:=nomfile
    Set Libro = MsExcel.Workbooks(1)
    Set Hoja = Libro.Sheets(1)
 
    i = 2
    empleado = Right("0000" & Hoja.Cells(i, 6).Value, 4)
    
    fecha = Hoja.Cells(i, 2)
    data = Split(fecha, "/")(2) & Split(fecha, "/")(1) & Split(fecha, "/")(0)
    While Not Hoja.Cells(i, 1).Value = ""
        If empleado <> Right("0000" & Hoja.Cells(i, 6), 4) Then
            Sql = "insert into sousNominaImportats (idFichero, Origen, CodiEmpresa, data, treb, i_liquid, i_irpf, i_tc1, i_brut, i_ssEmp, i_ssTre, RetibEspecie, IrpfEspecies) values "
            Sql = Sql & "('" & IdFile & "', 'Ficher', 0, '" & data & "', '" & empleado & "', " & i_liquid & ", " & i_Irpf & ", " & i_Tc1 & ", " & i_Brut & ", " & i_Tc1 - i_ssTre & ", " & i_ssTre & ", 0, 0)"
            ExecutaComandaSql Sql
            
            i_liquid = 0
            i_Tc1 = 0
            i_Irpf = 0
            i_Brut = 0
            i_ssTre = 0
        End If
        cCargo = Hoja.Cells(i, 3)
        cAbono = Hoja.Cells(i, 4)
        If Left(cAbono, 4) = "4650" Then i_liquid = Hoja.Cells(i, 5)
        If Left(cAbono, 4) = "4760" Then i_Tc1 = i_Tc1 + Hoja.Cells(i, 5)
        If Left(cAbono, 4) = "4751" Then i_Irpf = Hoja.Cells(i, 5)
        If Left(cCargo, 4) = "6400" Then i_Brut = Hoja.Cells(i, 5)
        If Left(cCargo, 4) = "6420" Then i_ssTre = Hoja.Cells(i, 5)
    
        empleado = Right("0000" & Hoja.Cells(i, 6), 4)
        i = i + 1
    Wend
    
    If empleado <> "" Then
        Sql = "insert into sousNominaImportats (idFichero, Origen, CodiEmpresa, data, treb, i_liquid, i_irpf, i_tc1, i_brut, i_ssEmp, i_ssTre, RetibEspecie, IrpfEspecies) values "
        Sql = Sql & "('" & IdFile & "', 'Ficher', 0, '" & data & "', '" & empleado & "', " & i_liquid & ", " & i_Irpf & ", " & i_Tc1 & ", " & i_Brut & ", " & i_Tc1 - i_ssTre & ", " & i_ssTre & ", 0, 0)"
        ExecutaComandaSql Sql
    End If

nok:
    TancaExcel MsExcel, Libro
End Sub



Sub ImportaNominasPANET(nomfile, IdFile)
    Dim MsExcel As Excel.Application, Libro As Excel.Workbook, Hoja
    Dim empleado As String, fecha As String, data As String, Empresa As String, Sql As String
    Dim i_liquid As Double, i_Tc1 As Double, i_Brut As Double, i_Irpf As Double, i_SsEmp As Double, i_especies As Double
    Dim i As Integer, H As Integer

    If Not frmSplash.Debugant Then On Error GoTo nok
    
    InformaMiss "Excel CarregantNomines PANET"
    Set MsExcel = CreateObject("Excel.Application")
    MsExcel.Visible = frmSplash.Debugant
    MsExcel.Workbooks.Open fileName:=nomfile
    Set Libro = MsExcel.Workbooks(1)
    
    For H = 1 To Libro.Sheets.Count
    
        Set Hoja = Libro.Sheets(H)
     
        Empresa = Hoja.Cells(5, 1)
        Empresa = Replace(Empresa, "Empresa:", "")
        
        i = 10
        While Not Hoja.Cells(i, 2).Value = ""
            empleado = Hoja.Cells(i, 2).Value & " " & Hoja.Cells(i, 3).Value
            fecha = Hoja.Cells(i, 6)
            data = Split(fecha, "/")(2) & Split(fecha, "/")(1) & Split(fecha, "/")(0)
    
            i_liquid = Hoja.Cells(i, 20)
            i_Tc1 = Hoja.Cells(i, 13)
            i_Irpf = Hoja.Cells(i, 19)
            i_Brut = Hoja.Cells(i, 15)
            i_SsEmp = Hoja.Cells(i, 8)
            i_especies = Hoja.Cells(i, 16)
            
            Sql = "insert into sousNominaImportats ([IdFichero],[Origen],[CodiEmpresa],[data],[treb],[i_liquid],[I_Irpf],[I_Tc1],[I_Brut],[i_SsEmp],[RetibEspecie],[IrpfEspecies]) values "
            Sql = Sql & "('" & IdFile & "', 'Ficher', '" & Empresa & "', '" & data & "', '" & empleado & "', " & i_liquid & ", " & i_Irpf & ", " & i_Tc1 & ", " & i_Brut & ", " & i_SsEmp & ", " & i_especies & ", 0)"
            ExecutaComandaSql Sql
            
            i = i + 1
        Wend
    Next
    
nok:
    TancaExcel MsExcel, Libro
End Sub




Sub CarregaNominasHtml(nomfile)
   Dim MsExcel As Excel.Application, Libro As Excel.Workbook, Hoja, Qr1 As rdoQuery, Qr2  As rdoQuery, i, Fi As Boolean, Codi, Variable, Valor, K, Codis As String, rs

    If Not frmSplash.Debugant Then On Error GoTo nok

    InformaMiss "Excel CarregantClients"
    Set MsExcel = CreateObject("Excel.Application")
    MsExcel.Visible = frmSplash.Debugant
    MsExcel.Workbooks.Open fileName:=nomfile
    Set Libro = MsExcel.Workbooks(1)
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Codis = ""
    
'    Empresa = Hoja.Cells(7, 1).Value
'    Periodo = Hoja.Cells(15, 1).Value
    
    
    
'    While Not Hoja.Cells(i, 2).Value = ""
'        Codi = Hoja.Cells(i, 1).Value
'        ExecutaComandaSql "Insert into Clients (Codi,nom) Values (" & Codi & ",'') "
'        k = k + 1
'    Wend
'    ExecutaComandaSql "Delete dffhfghhf f hfgh fgh ffgArticles Where not codi in (" & Codis & ") "

nok:
    TancaExcel MsExcel, Libro
End Sub


Sub ExternCarregaFtp()
   Dim CalBorrarArticles As Boolean, Files() As String, rs As rdoResultset, Botis() As String, K As Integer, p1 As String, P2 As String
   Dim i As Integer, Tipus() As String, Param() As String, Contingut As String, j As Integer, LaAgafem As Boolean
   Dim nom As String, Interesa As Integer, Esborrem As Integer, EsCfg As Boolean, ClientsGenerats As Boolean, EnviaArticles As Boolean
   
   InformaMiss "Configurant"
   frmSplash.IpConexio.InteresPerContingutReset
   frmSplash.IpConexio.InteresPerContingut "Tot", True, True
   frmSplash.IpConexio.LabelEstat = frmSplash.Estat
   frmSplash.IpConexio.LabelEstatDbg = frmSplash.lblVersion
   
   InformaMiss "Connectant Extern ... "
   frmSplash.IpConexio.CarterUltimaData = BuscaLastData
   frmSplash.IpConexio.CarterUltimaData = DateAdd("m", -3, Now)
   frmSplash.IpConexio.Cfg_ConnexioString = NomServerInternet
   frmSplash.IpConexio.Cfg_AppPath = AppPath
   If frmSplash.IpConexio.EnviaIReb("Reb") Then SetLastData DateAdd("s", -3, frmSplash.IpConexio.CarterUltimaData)

End Sub

Sub FileDeBinaTxt(Fil As String)
    Dim f, F2, aa
    Dim a As String
    Dim K
    Dim Ultim13, Ultim10
    
    MyKill Fil & ".Txt"
    f = FreeFile
    Open Fil & ".Bin" For Binary Access Read As #f
    F2 = FreeFile
    Open Fil & ".Txt" For Output As #F2
    a = Space(1)
    K = 1
    Ultim13 = False
    Ultim10 = False
    While Not EOF(f)
         Get #f, K, a
         K = K + 1
         If Not (Asc(a) = 13 Or Asc(a) = 10) Then
            If Asc(a) <= 26 Then a = " "
            If Asc(a) >= 130 Then a = " "
         End If
         
         If Asc(a) = 10 And Not Ultim13 Then Print #F2, Chr(13);
         Print #F2, a;
         Ultim13 = False
         Ultim10 = False
         
         If Asc(a) = 13 Then Ultim13 = True
         If Asc(a) = 10 Then Ultim10 = True
         
         DoEvents
    Wend
    Close f
    Close F2
End Sub

Sub ImportaClients(Database As String, Clau As String)

   ExecutaComandaSql "create table " & Database & ".Dbo.Clients_Imp_Codis (Codi [int],Unio [nvarchar] (255))"
   ExecutaComandaSql "Drop table " & Database & ".Dbo.Clients_Imp_Codis_Tmp "
   ExecutaComandaSql "Select param_1 as Unio,IDENTITY(int,1000,1) as codi into " & Database & ".Dbo.Clients_Imp_Codis_Tmp from importat_client N left join " & Database & ".Dbo.Clients_Imp_Codis V on N.param_1 = v.Unio where v.codi is null"
   ExecutaComandaSql "Insert Into " & Database & ".Dbo.Clients_Imp_Codis (unio,codi) select t.unio,t.codi + r.c from " & Database & ".Dbo.Clients_Imp_Codis_Tmp t, (Select isnull(max(codi),0) c from " & Database & ".Dbo.Clients_Imp_Codis) r "
   ExecutaComandaSql "Delete " & Database & ".Dbo.clients where codi in(Select Codi From " & Database & ".Dbo.Clients_Imp_Codis) "
   ExecutaComandaSql "Insert Into " & Database & ".Dbo.clients (nif,codi,Nom,Adresa,Cp,Ciutat,[Nom Llarg],lliure,[Tipus Iva],[Preu Base],[Desconte ProntoPago],[Desconte 1],[Desconte 2],[Desconte 3],[Desconte 4] , [Desconte 5], [AlbaraValorat])  select '',codi,Param_1 as Nom,Param_8 as Adresa,'' as Cp,'' as Ciutat,Param_3  as [Nom Llarg] ,Param_9 as lliure,1,2,0,0,0,0,0,param_6,1 From importat_client join  " & Database & ".Dbo.Clients_Imp_Codis on param_1 = unio where param_0='" & Clau & "'  And param_5='S' "
   ExecutaComandaSql "update " & Database & ".Dbo.clients set [Desconte 5] = 0 where [Desconte 5] = 1"
   
End Sub

Sub ImportaComanda(Database As String, sDia As String, Taula As String)
   Dim d As Date, Ser As String
   Dim t As String
On Error GoTo nor
   
   d = CVDate(sDia)
   t = Taula
   
   ExecutaComandaSql "Use " & Database
   If Not ExisteixTaula("ComandesModificades") Then ExecutaComandaSql "CREATE TABLE [ComandesModificades] ([Id] [uniqueidentifier] NULL ,[TimeStamp] [datetime] NULL ,  [TaulaOrigen] [nvarchar] (255) NULL ) ON [PRIMARY] "
   Ser = DonamNomTaulaServit(d)
   ExecutaComandaSql "Use Integraciones"
   
   ExecutaComandaSql "Delete from " & Database & ".dbo.[" & Ser & "] where Client in(Select distinct c.Codi from [" & t & "] s join " & Database & ".dbo.Clients_Imp_Codis c on s.Param_0 = c.unio) "
   ExecutaComandaSql "insert into " & Database & ".dbo.[" & Ser & "]  ([Id],[TimeStamp],[QuiStamp],[Client],[CodiArticle] ,[PluUtilitzat],[Viatge],[Equip],[QuantitatDemanada],[QuantitatTornada] ,[QuantitatServida],[MotiuModificacio],[Hora],[TipusComanda],[Comentari],[ComentariPer],[Atribut] ,[CitaDemanada],[CitaServida],[CitaTornada]) select  newid(),getdate(),'',c.codi,a.codi,a.codi,'Inicial','Inicial',s.param_7,0,0,0,69,0,0,'',0,'','','' from [" & t & "] s join " & Database & ".dbo.Clients_Imp_Codis c on s.Param_0 = c.unio join " & Database & ".dbo.Articles_Imp_Codis a on s.Param_3 = a.unio where s.param_2 = '" & sDia & "' "
   
   

nor:

End Sub



Sub ImportaArticles(Database As String, Clau As String, TarifaCodi As String)
   
   ExecutaComandaSql "create table " & Database & ".Dbo.Articles_Imp_Codis (Codi [int],Unio [nvarchar] (255))"
   ExecutaComandaSql "Drop table " & Database & ".Dbo.Articles_Imp_Codis_Tmp "
   ExecutaComandaSql "Select n.param_2 as Unio,IDENTITY(int,1,1) as codi into " & Database & ".Dbo.Articles_Imp_Codis_Tmp from importat_Articu N left join " & Database & ".Dbo.Articles_Imp_Codis V on N.param_2 = v.Unio join importat_ArtEmp E on N.param_2 = E.Param_1 And E.Param_0 = '" & Clau & "' where v.codi is null "
   ExecutaComandaSql "Insert Into " & Database & ".Dbo.Articles_Imp_Codis (unio,codi) select t.unio,t.codi + r.c from " & Database & ".Dbo.Articles_Imp_Codis_Tmp t, (Select isnull(max(codi),0) c from " & Database & ".Dbo.Articles_Imp_Codis) r "
   ExecutaComandaSql "Delete " & Database & ".Dbo.Articles"
   ExecutaComandaSql "Insert Into " & Database & ".Dbo.Articles (Codi,Codigenetic,nom,essumable,Familia,Tipoiva,Desconte,NoDescontesEspecials) select Distinct codi,codi,param_3,CASE Param_4 WHEN 'PIEZAS' THEN 1  ELSE 0  END ,param_1,CASE param_5 WHEN 'RR4'  THEN 1  WHEN 'RR7'  THEN 2  WHEN 'RR16' THEN 3  WHEN 'RRE4'  THEN 1  ELSE 2  END,1,0  From importat_articu join  " & Database & ".Dbo.Articles_Imp_Codis on param_2 = unio "
'ExecutaComandaSql "Insert Into " & Database & ".Dbo.Articles (Codi,Codigenetic,nom,essumable,Familia,Tipoiva,Desconte,NoDescontesEspecials) select Distinct codi,codi,param_3,1 ,param_1,CASE param_5 WHEN 'RR4'  THEN 1  WHEN 'RR7'  THEN 2  WHEN 'RR16' THEN 3  WHEN 'RRE4'  THEN 1  ELSE 2  END,1,0  From importat_articu join  " & Database & ".Dbo.Articles_Imp_Codis on param_2 = unio "
   ExecutaComandaSql "Update " & Database & ".Dbo.Articles Set Preu = 0, Desconte = 1 , PreuMajor = 0 "
   
   ImportaActualitzaNoms Database, Clau
   
   ExecutaComandaSql "update A set nom = nom + ' Cnj'    From " & Database & ".Dbo.Articles a join " & Database & ".Dbo.Articles_Imp_Codis i on a.codi = i.codi where unio like '%+%' "
   ExecutaComandaSql "update A set nom = nom + ' Preco'  From " & Database & ".Dbo.Articles a join " & Database & ".Dbo.Articles_Imp_Codis i on a.codi = i.codi where unio like '%&%' "
   ExecutaComandaSql "update A set nom = nom + ' Tallat' From " & Database & ".Dbo.Articles a join " & Database & ".Dbo.Articles_Imp_Codis i on a.codi = i.codi where unio like '%/%' "
   
   ImportaTarifa Database, TarifaCodi, Clau
   ExecutaComandaSql "update " & Database & ".Dbo.Articles set preu = 0.01 where nom like 'Reventa Centralizada'"

End Sub

Sub ImportaFitchersArticles(Fitcher)

'   ExecutaComandaSql "create table " & Database & ".Dbo.Articles_Imp_Codis (Codi [int],Unio [nvarchar] (255))"
'   ExecutaComandaSql "Drop table " & Database & ".Dbo.Articles_Imp_Codis_Tmp "
'   ExecutaComandaSql "Select n.param_2 as Unio,IDENTITY(int,1,1) as codi into " & Database & ".Dbo.Articles_Imp_Codis_Tmp from importat_Articu N left join " & Database & ".Dbo.Articles_Imp_Codis V on N.param_2 = v.Unio join importat_ArtEmp E on N.param_2 = E.Param_1 And E.Param_0 = '" & CLau & "' where v.codi is null "
'   ExecutaComandaSql "Insert Into " & Database & ".Dbo.Articles_Imp_Codis (unio,codi) select t.unio,t.codi + r.c from " & Database & ".Dbo.Articles_Imp_Codis_Tmp t, (Select isnull(max(codi),0) c from " & Database & ".Dbo.Articles_Imp_Codis) r "
'   ExecutaComandaSql "Delete " & Database & ".Dbo.Articles"
'   ExecutaComandaSql "Insert Into " & Database & ".Dbo.Articles (Codi,Codigenetic,nom,essumable,Familia,Tipoiva,Desconte,NoDescontesEspecials) select Distinct codi,codi,param_3,CASE Param_4 WHEN 'PIEZAS' THEN 1  ELSE 0  END ,param_1,CASE param_5 WHEN 'RR4'  THEN 1  WHEN 'RR7'  THEN 2  WHEN 'RR16' THEN 3  WHEN 'RRE4'  THEN 1  ELSE 2  END,1,0  From importat_articu join  " & Database & ".Dbo.Articles_Imp_Codis on param_2 = unio "
''ExecutaComandaSql "Insert Into " & Database & ".Dbo.Articles (Codi,Codigenetic,nom,essumable,Familia,Tipoiva,Desconte,NoDescontesEspecials) select Distinct codi,codi,param_3,1 ,param_1,CASE param_5 WHEN 'RR4'  THEN 1  WHEN 'RR7'  THEN 2  WHEN 'RR16' THEN 3  WHEN 'RRE4'  THEN 1  ELSE 2  END,1,0  From importat_articu join  " & Database & ".Dbo.Articles_Imp_Codis on param_2 = unio "
'   ExecutaComandaSql "Update " & Database & ".Dbo.Articles Set Preu = 0, Desconte = 1 , PreuMajor = 0 "
'
'   ImportaActualitzaNoms Database, CLau
'
'   ExecutaComandaSql "update A set nom = nom + ' Cnj'    From " & Database & ".Dbo.Articles a join " & Database & ".Dbo.Articles_Imp_Codis i on a.codi = i.codi where unio like '%+%' "
'   ExecutaComandaSql "update A set nom = nom + ' Preco'  From " & Database & ".Dbo.Articles a join " & Database & ".Dbo.Articles_Imp_Codis i on a.codi = i.codi where unio like '%&%' "
'   ExecutaComandaSql "update A set nom = nom + ' Tallat' From " & Database & ".Dbo.Articles a join " & Database & ".Dbo.Articles_Imp_Codis i on a.codi = i.codi where unio like '%/%' "
'
'   ImportaTarifa Database, TarifaCodi, CLau
'   ExecutaComandaSql "update " & Database & ".Dbo.Articles set preu = 0.01 where nom like 'Reventa Centralizada'"

End Sub


Sub ImportaFitchersClientsMasterpan(Idrs)
    Dim Sql As String, rs As New ADODB.Recordset, L As String, Codi, nom, nif, Direccio, Telefon, P2, P, Tarifa
    
    ExecutaComandaSql "Drop Table Clients_Importat_Masterpan "
    ExecutaComandaSql "create table Clients_Importat_Masterpan  (Tarifa [nvarchar](255),Codi [nvarchar](255),Nom [nvarchar](255),Direccio [nvarchar](255),Telefon [nvarchar](255),nif [nvarchar](255)) "
    
    Sql = "Select Linea "
    Sql = Sql & "From Archivolines "
    Sql = Sql & "Where id = '" & Idrs & "' And IsNumeric(SUBSTRING(linea, 1, 5)) = 1 "
    Set rs = rec(Sql)
    
    While Not rs.EOF
        L = rs("Linea")
        L = Join(Split(L, Chr(9)), "")  ' Treu tabuladors
        L = Join(Split(L, "'"), " ")    ' Treu apostrof
        L = Join(Split(L, """"), " ")   ' Treu Cometes
        Codi = ""
        nom = ""
        nif = ""
        Direccio = ""
        Telefon = ""
        Codi = RTrim(Mid(L, 1, 7))
        nom = RTrim(Mid(L, 8, 40))
        nif = RTrim(Mid(L, 48, 15))
        Direccio = RTrim(Mid(L, 63, 30))
        Telefon = RTrim(Mid(L, 93, 12))
        Tarifa = "Tarifa-" & RTrim(Mid(L, 106, 1))
        If Codi <> "" Then ExecutaComandaSql "Insert Into Clients_Importat_Masterpan (tarifa,telefon,codi,nom,nif,direccio) Values ('" & Tarifa & "','" & Telefon & "','" & Codi & "','" & nom & "','" & nif & "','" & Direccio & "')"
        rs.MoveNext
    Wend
    rs.Close
    
    
    ExecutaComandaSql "create table Clients_Imp_Codis (Codi [int],Unio [nvarchar] (255))"
    ExecutaComandaSql "Drop table Clients_Imp_Codis_Tmp "
    Sql = "Select n.codi as Unio,IDENTITY(int,1,1) as codi "
    Sql = Sql & "into Clients_Imp_Codis_Tmp  "
    Sql = Sql & "from Clients_Importat_Masterpan N left join Clients_Imp_Codis V on N.codi = v.Unio Where v.codi is null  "
    
    ExecutaComandaSql Sql ' insertem els nous a tmp
    
    Sql = "Insert Into Clients_Imp_Codis (unio,codi) "
    Sql = Sql & "select t.unio,t.codi + r.c from Clients_Imp_Codis_Tmp t, "
    Sql = Sql & "(Select isnull(max(codi),0) c from Clients_Imp_Codis) r  "
    
    ExecutaComandaSql Sql  ' els pasem a Clients_Imp_Codis
    
    ExecutaComandaSql "Delete Clients"
    
    Sql = "insert into clients (codi,nom,nif,adresa ,[Desconte 5]) "
    Sql = Sql & "select "
    Sql = Sql & "c.codi as codi,m.nom as nom,m.nif as nif, m.direccio as adresa ,isnull(t.tarifacodi,0) as [Desconte 5] "
    Sql = Sql & "From Clients_Importat_Masterpan m join  Clients_Imp_Codis c  on m.Codi = c.unio "
    Sql = Sql & "left Join Tarifesespecials T On t.tarifanom = m.tarifa "
    
    ExecutaComandaSql Sql
    
    ExecutaComandaSql "delete constantsclient where variable='Tel' and codi in (Select c.codi From Clients_Importat_Masterpan m join  Clients_Imp_Codis c  on m.Codi = c.unio) "
    ExecutaComandaSql "insert into constantsclient (codi,variable,valor) select c.codi,'Tel' ,m.telefon  From Clients_Importat_Masterpan m join  Clients_Imp_Codis c  on m.Codi = c.unio "
    
    ExecutaComandaSql "delete constantsclient where variable='CodiContable' and codi in (Select c.codi From Clients_Importat_Masterpan m join  Clients_Imp_Codis c  on m.Codi = c.unio) "
    ExecutaComandaSql "insert into constantsclient (codi,variable,valor) select c.codi,'CodiContable' ,m.Codi  From Clients_Importat_Masterpan m join  Clients_Imp_Codis c  on m.Codi = c.unio "
    
    
End Sub



Sub ImportaFitchersPreusEspecialsMasterpan(Idrs)
    Dim Sql As String, rs As New ADODB.Recordset, L As String, Codi, nom, nif, Direccio, Telefon, P2, P, Tarifa
    
    ExecutaComandaSql "Drop Table PreusEspecials_Importat_Masterpan "
    ExecutaComandaSql "Create Table PreusEspecials_Importat_Masterpan  (Client [nvarchar](255),Article [nvarchar](255),Preu [float]) "
    
    Sql = "Select Linea "
    Sql = Sql & "From Archivolines "
    Sql = Sql & "Where id = '" & Idrs & "' And IsNumeric(SUBSTRING(linea, 1, 5)) = 1 "
    Set rs = rec(Sql)
    
    While Not rs.EOF
        L = rs("Linea")
        L = Join(Split(L, Chr(9)), "")  ' Treu tabuladors
        L = Join(Split(L, "'"), " ")    ' Treu apostrof
        L = Join(Split(L, """"), " ")   ' Treu Cometes
        L = Join(Split(L, ","), ".")   ' Treu Cometes
        Codi = ""
        nom = ""
        nif = ""
        Direccio = ""
        Telefon = ""
        Codi = RTrim(Mid(L, 1, 7))
        nom = RTrim(Mid(L, 28, 5))
        nif = RTrim(Mid(L, 59, 8))
        If Codi <> "" Then ExecutaComandaSql "Insert Into PreusEspecials_Importat_Masterpan (Client,Article,Preu) Values ('" & Codi & "','" & nom & "','" & nif & "')"
        rs.MoveNext
    Wend
    rs.Close
    
    ExecutaComandaSql "Delete tarifesespecialsclients"
    ExecutaComandaSql "insert into tarifesespecialsclients (id,client,codi,preu,preumajor,Qmin) select newid(),c.codi,a.codi,preu,preu,0 from PreusEspecials_Importat_Masterpan m join Articles_Imp_Codis a on a.unio = m.article join Clients_Imp_Codis  c on c.unio = m.client     "

End Sub




Sub ImportaFitchersTarifesMasterpan(Idrs)
    Dim Sql As String, rs As New ADODB.Recordset, L As String, Codi, nom, nif, Direccio, Telefon, P2, P, Tarifa, i
    Dim nomTarifa() As String
    
    ExecutaComandaSql "Drop Table TarifesEspecials_Importat_Masterpan "
    ExecutaComandaSql "Create Table TarifesEspecials_Importat_Masterpan  (TarifaNom [nvarchar](255),Article [nvarchar](255),Preu [float]) "
    
    Sql = "Select Linea "
    Sql = Sql & "From Archivolines "
    Sql = Sql & "Where id = '" & Idrs & "' order by numlinea "
    Set rs = rec(Sql)
    
'Art¡Nombre Art¡culo     BCN-2     S.FELIU-4 HOSP.-5   S.BOI-7   PRAT-8    CORNELLA9
'1000BARRA DE 1/2 SUPERLA  1,050     1,100     0,950     1,100     1,000     0,950
    
    ReDim nomTarifa(9)
    nomTarifa(0) = "BCN-2"
    nomTarifa(1) = "S.FELIU-4"
    nomTarifa(2) = "HOSP.-5"
    nomTarifa(3) = "S.BOI-7"
    nomTarifa(4) = "PRAT-8"
    nomTarifa(5) = "CORNELLA-9"
    
    While Not rs.EOF
        L = rs("Linea")
        L = Join(Split(L, Chr(9)), "")  ' Treu tabuladors
        L = Join(Split(L, "'"), " ")    ' Treu apostrof
        L = Join(Split(L, """"), " ")   ' Treu Cometes
        L = Join(Split(L, ","), ".")    ' Punt decimal
        If IsNumeric(Left(L, 4)) Then
            Codi = RTrim(Mid(L, 1, 4))
            
            nom = RTrim(Mid(L, 25, 7))
            ExecutaComandaSql "Insert Into TarifesEspecials_Importat_Masterpan (Tarifanom,Article,Preu) Values ('" & nomTarifa(0) & "','" & Codi & "','" & nom & "')"
            
            nom = RTrim(Mid(L, 35, 7))
            ExecutaComandaSql "Insert Into TarifesEspecials_Importat_Masterpan (Tarifanom,Article,Preu) Values ('" & nomTarifa(1) & "','" & Codi & "','" & nom & "')"
            
            nom = RTrim(Mid(L, 43, 9))
            ExecutaComandaSql "Insert Into TarifesEspecials_Importat_Masterpan (Tarifanom,Article,Preu) Values ('" & nomTarifa(2) & "','" & Codi & "','" & nom & "')"
            
            nom = RTrim(Mid(L, 54, 8))
            ExecutaComandaSql "Insert Into TarifesEspecials_Importat_Masterpan (Tarifanom,Article,Preu) Values ('" & nomTarifa(3) & "','" & Codi & "','" & nom & "')"
            
            nom = RTrim(Mid(L, 65, 7))
            ExecutaComandaSql "Insert Into TarifesEspecials_Importat_Masterpan (Tarifanom,Article,Preu) Values ('" & nomTarifa(4) & "','" & Codi & "','" & nom & "')"
            
            nom = RTrim(Mid(L, 74, 8))
            ExecutaComandaSql "Insert Into TarifesEspecials_Importat_Masterpan (Tarifanom,Article,Preu) Values ('" & nomTarifa(5) & "','" & Codi & "','" & nom & "')"
            
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    ExecutaComandaSql "Delete tarifesespecials"
    For i = 0 To 5
        ExecutaComandaSql "insert into Tarifesespecials (TarifaCodi,TarifaNom,Codi,preu,preumajor) select " & i + 1 & ",'" & nomTarifa(i) & "',a.codi,preu,preu from TarifesEspecials_Importat_Masterpan m join Articles_Imp_Codis a on a.unio = m.article where tarifanom = '" & nomTarifa(i) & "' "
    Next


End Sub





Sub ImportaFitchersArticlesMasterpan(Idrs)
    Dim Sql As String, rs As New ADODB.Recordset, L As String
    Dim Codi, nom, nif, Direccio, Telefon
    Dim P2, P, Tarifa
    Dim C1 As String, C2 As String, C3 As String, C4 As String, C5 As String, C6 As String, C7 As String, C8 As String, C9 As String, C10 As String
    
    ExecutaComandaSql "Drop Table Articles_Importat_Masterpan "
    ExecutaComandaSql "create table Articles_Importat_Masterpan  (Codi [nvarchar](255),nom [nvarchar](255),Familia [nvarchar](255),Unidades [nvarchar](255),Iva [nvarchar](255)) "
    
    Sql = "Select Linea "
    Sql = Sql & "From Archivolines "
    Sql = Sql & "Where id = '" & Idrs & "' And IsNumeric(SUBSTRING(linea, 1, 8)) = 1 "
    Set rs = rec(Sql)
    
    While Not rs.EOF
        L = rs("Linea")
        L = Join(Split(L, Chr(9)), "")  ' Treu tabuladors
        L = Join(Split(L, Chr(25)), "")  ' Treu tabuladors
        L = Join(Split(L, "'"), " ")    ' Treu apostrof
        L = Join(Split(L, """"), " ")   ' Treu Cometes
        L = Join(Split(L, "¥"), "Ñ")   ' Treu Cometes
        
        C1 = ""
        C2 = ""
        C3 = ""
        C4 = ""
        C5 = ""
        C1 = RTrim(Mid(L, 1, 7))
        C2 = RTrim(Mid(L, 9, 30))
        C3 = RTrim(Mid(L, 46, 20))
        C4 = RTrim(Mid(L, 87, 2))
        C5 = RTrim(Mid(L, 90, 5))
        If C1 <> "" Then ExecutaComandaSql "Insert Into Articles_Importat_Masterpan (Codi ,nom ,Familia ,Unidades ,Iva) Values ('" & C1 & "','" & C2 & "','" & C3 & "','" & C4 & "','" & C5 & "')"
        
        
 
        rs.MoveNext
    Wend
    rs.Close
    
    
    ExecutaComandaSql "create table Articles_Imp_Codis (Codi [int],Unio [nvarchar] (255))"
    ExecutaComandaSql "Drop table Articles_Imp_Codis_Tmp "
    Sql = "Select n.codi as Unio,IDENTITY(int,1,1) as codi "
    Sql = Sql & "into Articles_Imp_Codis_Tmp  "
    Sql = Sql & "from Articles_Importat_Masterpan N left join Articles_Imp_Codis V on N.codi = v.Unio Where v.codi is null  "
    
    ExecutaComandaSql Sql ' insertem els nous a tmp
    
    Sql = "Insert Into Articles_Imp_Codis (unio,codi) "
    Sql = Sql & "select t.unio,t.codi + r.c from Articles_Imp_Codis_Tmp t, "
    Sql = Sql & "(Select isnull(max(codi),0) c from Articles_Imp_Codis) r  "
    
    ExecutaComandaSql Sql  ' els pasem a Articles_Imp_Codis
    
    ExecutaComandaSql "Delete Articles"
    
    ExecutaComandaSql "insert into Articles (codi,nom,familia,TipoIva,essumable) select c.codi as codi,m.nom as nom,m.familia as familia, CASE m.iva WHEN ' 4,00' THEN 1 WHEN ' 7,00' THEN 2 WHEN '16,00' THEN 3 ELSE 0  END as tipoiva,CASE m.Unidades WHEN 'un' THEN 1  ELSE 0  END  as essumable From Articles_Importat_Masterpan m join  Articles_Imp_Codis c  on m.Codi = c.unio "
    
    ExecutaComandaSql "Delete ArticlesPropietats where variable='CODI_PROD' and codiArticle in (Select c.codi From Articles_Importat_Masterpan m join  Articles_Imp_Codis c  on m.Codi = c.unio) "
    ExecutaComandaSql "Insert Into articlespropietats (codiArticle,variable,valor)  select c.codi,'CODI_PROD' ,m.Codi   From Articles_Importat_Masterpan m join  Articles_Imp_Codis c   on m.Codi = c.unio  "

    ExecutaComandaSql "update articles set preu = 0 where preu is null "
    ExecutaComandaSql "update articles set [preumajor] = 0 where [preumajor] is null "
    
    ExecutaComandaSql "Delete families"
    ExecutaComandaSql "insert into families select distinct familia ,familia+'.',0 estatus,0 nivell,0 utilitza from articles "
    ExecutaComandaSql "insert into families select distinct familia+'.',familia+'..',0 estatus,0 nivell,0 utilitza from articles"
    ExecutaComandaSql "insert into families select distinct familia+'..','Article',0 estatus,0 nivell,0 utilitza from articles"
    ExecutaComandaSql "insert into families (nom,pare,estatus,nivell,utilitza) values ('Article','',0,0,0)"
    
    
End Sub




Sub ImportaFitchersAlbaransMasterpan(Idrs)
    Dim Sql As String, rs As New ADODB.Recordset, L As String, Viatge, equip, Codi, nom, nif, Direccio, Telefon, P2, P, Tarifa, C1 As String, C2 As String, C3 As String, C4 As String, C5 As String, C6 As String, C7 As String, C8 As String, C9 As String, C10 As String, client As String, data As String, linea, Qr As rdoQuery
    Dim V1 As String, tT As String, mes As String, An As String, V2 As String, V3 As String, V4 As String, e As String, Dia As String
    
    ExecutaComandaSql "Drop Table Albarans_Importat_Masterpan "
    ExecutaComandaSql "create table Albarans_Importat_Masterpan  (Data [datetime],Client [nvarchar](255),Article [nvarchar](255),Nom [nvarchar](255),E [nvarchar](255),V1 [nvarchar](255),V2 [nvarchar](255),V3 [nvarchar](255),V4 [nvarchar](255),Q [float]) "
    
    Sql = "Select Linea "
    Sql = Sql & "From Archivolines "
    Sql = Sql & "Where id = '" & Idrs & "' order by numlinea " 'And IsNumeric(SUBSTRING(linea, 1, 8)) = 1 "
    Set rs = rec(Sql)
    
    linea = 0
    While Not rs.EOF
        L = rs("Linea")
        DoEvents
        If linea Mod 100 = 0 Then InformaMiss "Important Albarans Mp. " & linea, True
        linea = linea + 1
'        If InStr(L, "ANTIGA 150 GRS BARROT") Then
'                   L = L
'        End If
        'If Left(L, 11) = " Fecha    :" Then data = CVDate(Mid(L, 12, 12))
        If InStr(L, "FECHA:") Then
            tT = Trim(Mid(L, InStr(L, "FECHA:") + 8, 18))
            P = InStr(tT, "/")
            If P > 0 Then
                Dia = Left(tT, P - 1)
                tT = Right(tT, Len(tT) - P)
                P = InStr(tT, "/")
                If P > 0 Then
                    mes = Left(tT, P - 1)
                    tT = Right(tT, Len(tT) - P)
                    P = 4
                    If P > 0 Then
                        An = Left(tT, P)
                        data = DateSerial(An, mes, Dia)
                    End If
                End If
            End If
        End If
        If InStr(L, "CLIENTE:") Then client = LTrim(RTrim(Mid(L, InStr(L, "CLIENTE:") + 8, 10)))
        
        If IsNumeric(Left(L, 8)) Then
            C1 = ""
            C2 = ""
            C3 = ""
            V1 = ""
            V2 = ""
            V3 = ""
            V4 = ""
            C1 = RTrim(LTrim(Mid(L, 1, 8)))
            C2 = RTrim(LTrim(Mid(L, 10, 30)))
            C3 = Join(Split(RTrim(LTrim(Mid(L, 120, 8))), ","), ".")   ' Treu ,
            V1 = Join(Split(RTrim(LTrim(Mid(L, 45, 11))), ","), ".")   ' Treu ,
            V2 = Join(Split(RTrim(LTrim(Mid(L, 58, 8))), ","), ".")   ' Treu ,
            V3 = Join(Split(RTrim(LTrim(Mid(L, 67, 8))), ","), ".")   ' Treu ,
            V4 = Join(Split(RTrim(LTrim(Mid(L, 76, 8))), ","), ".")   ' Treu ,
            e = RTrim(LTrim(Mid(L, 105, 5)))
            If C1 <> "" Then ExecutaComandaSql "Insert Into Albarans_Importat_Masterpan (Data ,Client ,Article ,Nom ,e,v1,v2,v3,v4,Q ) Values ('" & data & "','" & client & "','" & C1 & "','" & C2 & "','" & e & "','" & V1 & "','" & V2 & "','" & V3 & "','" & V4 & "','" & C3 & "')"
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    ExecutaComandaSql "CREATE TABLE [ComandesModificades] ([Id] [uniqueidentifier] NULL , [TimeStamp] [datetime] NULL , [TaulaOrigen] [nvarchar] (255) NULL) ON [PRIMARY] "
    ' Asignem els equips
    ExecutaComandaSql "update b set b.e = c.equip from Albarans_Importat_Masterpan b join articlespropietats A  on A.valor = B.article and a.variable = 'CODI_PROD' join (Select distinct c.CodiArticle,Equip  from (Select max([timestamp]) t ,CodiArticle From comandesmemotecnicperclient where not equip is null group by  CodiArticle ) c join comandesmemotecnicperclient cc on c.CodiArticle = cc.CodiArticle and c.t = cc.[timestamp]) c on c.codiarticle = a.CodiArticle "
    
    Set rs = rec("Select distinct data from Albarans_Importat_Masterpan order by data")
    While Not rs.EOF
        Debug.Print DonamNomTaulaServit(rs(0))
        DoEvents
        data = rs(0)
        ExecutaComandaSql "Delete  [" & DonamNomTaulaServit(rs(0)) & "] Where  not tipuscomanda = 2 "
        Sql = "Insert Into [" & DonamNomTaulaServit(rs(0)) & "] (CodiArticle,PluUtilitzat, Client,QuantitatDemanada,QuantitatServida,QuantitatTornada,Viatge,Equip)  "
        Sql = Sql & "Select ar.codi ,ar.codi ,c.codi as client ,b.v1,b.v1,0 ,'Viaje 1',E "
        Sql = Sql & "From Albarans_Importat_Masterpan B "
        Sql = Sql & "join articlespropietats A  on A.valor = B.article and a.variable = 'CODI_PROD'  join Articles Ar  on cast(Ar.codi as nvarchar) = A.Codiarticle "
        Sql = Sql & "join ConstantsClient  C  on c.valor = B.Client  collate Modern_Spanish_CI_AS and c.variable = 'CodiContable' "
        Sql = Sql & " Where Data = '" & data & "' and not V1='' and not c.codi is null and not ar.codi is null and isnumeric (b.v1) = 1 "
        ExecutaComandaSql Sql
        
        Sql = "Insert Into [" & DonamNomTaulaServit(rs(0)) & "] (CodiArticle,PluUtilitzat, Client,QuantitatDemanada,QuantitatServida,QuantitatTornada,Viatge,Equip)  "
        Sql = Sql & "Select ar.codi ,ar.codi ,c.codi as client ,b.v2,b.v2,0 ,'Viaje 2',E "
        Sql = Sql & "From Albarans_Importat_Masterpan B "
        Sql = Sql & "join articlespropietats A  on A.valor = B.article and a.variable = 'CODI_PROD'  join Articles Ar  on cast(Ar.codi as nvarchar) = A.Codiarticle "
        Sql = Sql & "join ConstantsClient  C  on c.valor = B.Client  collate Modern_Spanish_CI_AS and c.variable = 'CodiContable' "
        Sql = Sql & " Where Data = '" & data & "' and not V2='' and not c.codi is null and not ar.codi is null and isnumeric (b.v2) = 1 "
        ExecutaComandaSql Sql
        
        Sql = "Insert Into [" & DonamNomTaulaServit(rs(0)) & "] (CodiArticle,PluUtilitzat, Client,QuantitatDemanada,QuantitatServida,QuantitatTornada,Viatge,Equip)  "
        Sql = Sql & "Select ar.codi ,ar.codi ,c.codi as client ,b.v3,b.v3,0 ,'Viaje 3',E "
        Sql = Sql & "From Albarans_Importat_Masterpan B "
        Sql = Sql & "join articlespropietats A  on A.valor = B.article and a.variable = 'CODI_PROD'  join Articles Ar  on cast(Ar.codi as nvarchar) = A.Codiarticle "
        Sql = Sql & "join ConstantsClient  C  on c.valor = B.Client  collate Modern_Spanish_CI_AS and c.variable = 'CodiContable' "
        Sql = Sql & " Where Data = '" & data & "' and not V3='' and not c.codi is null and not ar.codi is null and isnumeric (b.v3) = 1 "
        ExecutaComandaSql Sql
        
        Sql = "Insert Into [" & DonamNomTaulaServit(rs(0)) & "] (CodiArticle,PluUtilitzat, Client,QuantitatDemanada,QuantitatServida,QuantitatTornada,Viatge,Equip)  "
        Sql = Sql & "Select ar.codi ,ar.codi ,c.codi as client ,b.v4,b.v4,0 ,'Viaje 4',E "
        Sql = Sql & "From Albarans_Importat_Masterpan B "
        Sql = Sql & "join articlespropietats A  on A.valor = B.article and a.variable = 'CODI_PROD'  join Articles Ar  on cast(Ar.codi as nvarchar) = A.Codiarticle "
        Sql = Sql & "join ConstantsClient  C  on c.valor = B.Client  collate Modern_Spanish_CI_AS and c.variable = 'CodiContable' "
        Sql = Sql & " Where Data = '" & data & "' and not V4='' and not c.codi is null and not ar.codi is null and isnumeric (b.v4) = 1 "
        ExecutaComandaSql Sql
        
        Sql = "insert into AlbaransNum (Data,Client,Viatge) "
        Sql = Sql & "Select distinct '" & Day(rs(0)) & "/" & Month(rs(0)) & "/" & Year(rs(0)) & "',client,'' "
        Sql = Sql & "from [" & DonamNomTaulaServit(rs(0)) & "] Where Not Comentari like '%Albara:%' or comentari is null "
        ExecutaComandaSql Sql

        Sql = "Update S set comentari = '[IdAlbara:' + cast(a.Codi as nvarchar) +  ']' + isnull(comentari,'')  "
        Sql = Sql & "From [" & DonamNomTaulaServit(rs(0)) & "] s Join AlbaransNum a "
        Sql = Sql & "on s.client = a.client and a.data = '" & Day(rs(0)) & "/" & Month(rs(0)) & "/" & Year(rs(0)) & "' and (Not (Comentari like '%Albara:%') or comentari is null) "
        ExecutaComandaSql Sql
        
        rs.MoveNext
    Wend
    rs.Close
End Sub




Sub ImportaFitchersAlbaransTangram(Idrs)
    Dim Rs2 As rdoResultset, Sql As String, rs As rdoResultset, L As String, Viatge, equip, Codi, nom, nif, Direccio, Telefon, P2, P, Tarifa, C1 As String, C2 As String, C3 As String, C4 As String, C5 As String, C6 As String, C7 As String, C8 As String, C9 As String, C10 As String, client As String, data As String, linea, Qr As rdoQuery
    Dim V1 As String, tT As String, mes As String, An As String, V2 As String, V3 As String, V4 As String, e As String, Dia As String
    Dim TenimClient As Boolean, CliNom As String, CliAdresa As String, CliNif As String, Clicodi As String, Clialbara As String, CliData As String
    Dim EstemDincsAlbara As Boolean
    ExecutaComandaSql "Drop Table Albarans_Importat_Tangram "
    ExecutaComandaSql "create table Albarans_Importat_Tangram  (Data [datetime],Client [nvarchar](255),Article [nvarchar](255),Nom [nvarchar](255),E [nvarchar](255),V1 [nvarchar](255),V2 [nvarchar](255),V3 [nvarchar](255),V4 [nvarchar](255),Q [float]) "
    
    rec "Delete Archivolines Where id = '" & Idrs & "' and ltrim(rtrim(Linea)) = ''"
    rec "Delete Archivolines Where id = '" & Idrs & "' and ltrim(rtrim(Linea)) = 'OK'"
    Sql = "Select Linea "
    Sql = Sql & "From Archivolines "
    Sql = Sql & "Where id = '" & Idrs & "' order by numlinea " 'And IsNumeric(SUBSTRING(linea, 1, 8)) = 1 "
    Set rs = Db.OpenResultset(Sql)
    
'                                                 CENTRO ARAGONS
'
'                                                 CL Valent¡ Almirall 29
'8206                                                   Sabadell
'                                                 BARCELONA
'                                                 C.I.F.:
' CLIENTE    ALBARAN      FECHA      PAGINA
'-------------------------------------------
'430000011      3266   15/10/12         1
'
'ARTICULO   D E S C R I P C I O N   E  UNIDADES DEVUELTAS    PRECIO   IMPORTE
'----------------------------------------------------------------------------
'1032     BARRETES 1 TALL           1     30,00                0,45    13,50
    
    linea = 0
    TenimClient = False
    EstemDincsAlbara = False
    While Not rs.EOF
        L = rs("Linea")
        DoEvents
        If linea Mod 100 = 0 Then InformaMiss "Important Albarans Tg. " & linea, True
        linea = linea + 1
'        If InStr(L, "ANTIGA 150 GRS BARROT") Then
'                   L = L
'        End If
        'If Left(L, 11) = " Fecha    :" Then data = CVDate(Mid(L, 12, 12))
        If Not TenimClient Then
           CliNom = Trim(L)
           rs.MoveNext
           L = rs("Linea")
           CliAdresa = Trim(L)
           rs.MoveNext
           L = rs("Linea") ' Poblacio
           rs.MoveNext
           L = rs("Linea") ' Ciutat
           rs.MoveNext
           L = rs("Linea") ' Cif
           TenimClient = True
           rs.MoveNext
           L = rs("Linea")
        End If
        
        
        If InStr(L, " CLIENTE    ALBARAN      FECHA      PAGINA") Then
           rs.MoveNext
           L = rs("Linea") ' Lineas
           If InStr(L, "-------------------------------------------") Then
               rs.MoveNext
               L = rs("Linea") '430000011      3266   15/10/12         1
               client = LTrim(RTrim(Mid(L, 1, InStr(L, " "))))
               Set Rs2 = Db.OpenResultset("Select * from ConstantsClient where variable = 'CodiContable' and valor = '" & client & "'")
               If Rs2.EOF Then
                  Set Rs2 = Db.OpenResultset("Select * from clients where nom = '" & CliNom & "'")
                  If Rs2.EOF Then
                    Set Rs2 = Db.OpenResultset("Select max(codi)+1 as codi From Clients ")
                    rec "Insert Into Clients (Codi,Nom) Values (" & Rs2("Codi") & ",'" & CliNom & "')  "
                    rec "Insert Into constantsclient (codi,Variable,Valor) Values ('" & Rs2("Codi") & "','CodiContable','" & client & "') "
                  Else
                    rec "Delete constantsclient where variable = 'CodiContable' and codi ='" & Rs2("Codi") & "' "
                    rec "Insert Into constantsclient (codi,Variable,Valor) Values ('" & Rs2("Codi") & "','CodiContable','" & client & "') "
                  End If
               End If
               
               Clialbara = LTrim(RTrim(Mid(L, InStr(L, " "), InStr(InStr(L, " "), L, " "))))
               tT = LTrim(RTrim(Mid(L, InStr(L, "/") - 2, 8)))
               P = InStr(tT, "/")
               If P > 0 Then
                Dia = Left(tT, P - 1)
                tT = Right(tT, Len(tT) - P)
                P = InStr(tT, "/")
                If P > 0 Then
                    mes = Left(tT, P - 1)
                    tT = Right(tT, Len(tT) - P)
                    P = 4
                    If P > 0 Then
                        An = Left(tT, P)
                        data = DateSerial(An, mes, Dia)
                    End If
                End If
               End If
           End If
           rs.MoveNext
           L = rs("Linea") ' Lineas
        End If
        
        If InStr(L, "ARTICULO   D E S C R I P C I O N   E  UNIDADES DEVUELTAS") Then
           rs.MoveNext
           L = rs("Linea") ' Lineas
           If InStr(L, "-----------------------------") Then
               rs.MoveNext
               L = rs("Linea") '1032     BARRETES 1 TALL           1     30,00                0,45    13,50
               EstemDincsAlbara = True
           End If
       End If
       
       If EstemDincsAlbara And IsNumeric(Left(L, 8)) Then
          C1 = ""
          C2 = ""
          C3 = ""
          V1 = ""
          V2 = ""
          V3 = ""
          V4 = ""
          e = ""
          C1 = RTrim(LTrim(Mid(L, 1, 8))) ' codi article
          C2 = RTrim(LTrim(Mid(L, 10, 26)))  ' Nom Producte
          V1 = Join(Split(Mid(L, 60, 8), ","), ".") ' preu
          C3 = Join(Split(Mid(L, 40, 10), ","), ".") ' Unitats
          If C1 <> "" Then ExecutaComandaSql "Insert Into Albarans_Importat_Tangram (Data ,Client ,Article ,Nom ,e,v1,v2,v3,v4,Q ) Values ('" & data & "','" & client & "','" & C1 & "','" & C2 & "','" & e & "','" & V1 & "','" & V2 & "','" & V3 & "','" & V4 & "','" & C3 & "')"
          C3 = C3
       End If
        
        If InStr(L, "BRUTO          DESCUENTO") Then
           TenimClient = False
           EstemDincsAlbara = False
           rs.MoveNext
           L = rs("Linea") ' Lineas
           rs.MoveNext
           L = rs("Linea") ' Lineas
        End If
        
        rs.MoveNext
    Wend
    rs.Close
    
    ExecutaComandaSql "CREATE TABLE [ComandesModificades] ([Id] [uniqueidentifier] NULL , [TimeStamp] [datetime] NULL , [TaulaOrigen] [nvarchar] (255) NULL) ON [PRIMARY] "
    ' Asignem els equips
    ExecutaComandaSql "Update Albarans_Importat_Tangram set e='No Definit' "
    Sql = "update b set b.e = cc.Equip from Albarans_Importat_Tangram b join "
    Sql = Sql & "articlespropietats A on A.valor = B.article collate Modern_Spanish_CI_AS and a.variable = 'CODI_PROD' join "
    Sql = Sql & "ConstantsClient    C on c.valor = B.Client  collate Modern_Spanish_CI_AS and c.variable = 'CodiContable' join "
    Sql = Sql & "comandesmemotecnicperclient cc on cc.Client =c.Codi  and cc.CodiArticle = a.CodiArticle "
    ExecutaComandaSql Sql
    
    Set rs = Db.OpenResultset("Select distinct data from Albarans_Importat_Tangram order by data")
    While Not rs.EOF
        InformaMiss "Important Factura Tg. " & rs(0), True
        Debug.Print DonamNomTaulaServit(rs(0))
        DoEvents
        data = rs(0)
        ExecutaComandaSql "Delete  [" & DonamNomTaulaServit(rs(0)) & "] Where  not tipuscomanda = 2 "
        Sql = "Insert Into [" & DonamNomTaulaServit(rs(0)) & "] (CodiArticle,PluUtilitzat, Client,QuantitatDemanada,QuantitatServida,QuantitatTornada,Viatge,Equip)  "
        Sql = Sql & "Select ar.codi ,ar.codi ,c.codi as client ,b.q,b.q,0 ,'Viaje 1',E "
        Sql = Sql & "From Albarans_Importat_Tangram B "
        Sql = Sql & "join articlespropietats A  on A.valor = B.article and a.variable = 'CODI_PROD'  join Articles Ar  on cast(Ar.codi as nvarchar) = A.Codiarticle "
        Sql = Sql & "join ConstantsClient  C  on c.valor = B.Client  collate Modern_Spanish_CI_AS and c.variable = 'CodiContable' "
        Sql = Sql & " Where day(data) = " & Day(data) & " and month(data) = " & Month(data) & " and year(data) = " & Year(data) & " and  not c.codi is null and not ar.codi is null "
        ExecutaComandaSql Sql
        
'        Sql = "Insert into AlbaransNum (Data,Client,Viatge) "
'        Sql = Sql & "Select distinct '" & Day(Rs(0)) & "/" & Month(Rs(0)) & "/" & Year(Rs(0)) & "',client,'' "
'        Sql = Sql & "from [" & DonamNomTaulaServit(Rs(0)) & "] Where Not Comentari like '%Albara:%' or comentari is null "
'        ExecutaComandaSql Sql

        Sql = "Update S set comentari = '[IdAlbara:' + cast(a.Codi as nvarchar) +  ']' + isnull(comentari,'')  "
        Sql = Sql & "From [" & DonamNomTaulaServit(rs(0)) & "] s Join AlbaransNum a "
        Sql = Sql & "on s.client = a.client and a.data = '" & Day(rs(0)) & "/" & Month(rs(0)) & "/" & Year(rs(0)) & "' and (Not (Comentari like '%Albara:%') or comentari is null) "
        ExecutaComandaSql Sql
        
        rs.MoveNext
    Wend
    rs.Close
End Sub





Sub ImportaFitchersFacturesMasterpan(Idrs, nomfile)
    Dim BaseFactura2 As Double, BaseFactura As Double, BaseIvaTipus_1 As Double, BaseIvaTipus_2 As Double, BaseIvaTipus_3 As Double, BaseIvaTipus_4 As Double, BaseIvaTipus_1_Iva As Double, BaseIvaTipus_2_Iva As Double, BaseIvaTipus_3_Iva As Double, BaseIvaTipus_4_Iva As Double, iD As String, i As Integer, rs As New ADODB.Recordset, Q As rdoQuery, FacturaTotal As Double, BaseRecTipus_1 As Double, BaseRecTipus_2 As Double, BaseRecTipus_3 As Double, BaseRecTipus_4  As Double, BaseRecTipus_1_Rec As Double, BaseRecTipus_2_Rec As Double, BaseRecTipus_3_Rec As Double, BaseRecTipus_4_Rec As Double, NumFacNoAria As Double, DescontePp As Double, TipusFacturacio As Double, EmpCodi As Double, EmpSerie As String, Tot_BaseIvaTipus_1 As Double, Tot_BaseIvaTipus_2 As Double, Tot_BaseIvaTipus_3 As Double, Tot_BaseIvaTipus_4 As Double, Tot_BaseRecTipus_1 As Double, Tot_BaseRecTipus_2 As Double, Tot_BaseRecTipus_3  As Double, Tot_BaseRecTipus_4  As Double, VencimentActual As Date, AceptaDevolucions
    Dim Cclient As Double, clientNom As String, ClientNif As String, clientAdresa As String, clientCp As String, ClientLliure As String, EmpNom As String, empNif As String, empAdresa As String, empCp As String, EmpLliure As String, empTel, empFax, empEMail, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCampMercantil, empCiutat, ClientNomComercial As String, Tarifa As Integer, Impostos As Double, ClientCodiFact, TipIva As String
    Dim Sql As String, valorIva1 As Double, valorIva2 As Double, valorIva3 As Double, valorIva4 As Double, valorRec1 As Double, valorRec2 As Double, valorRec3 As Double, valorRec4 As Double, IvaRec1 As Double, IvaRec2 As Double, IvaRec3 As Double, IvaRec4 As Double, PreusActuals As Boolean, DiesVenciment, DiaPagament, FormaPagoLlista, Clis() As String, baseIva1, baseIva2, vaseiva3, tipiva1, tipiva2, tipiva3, totalfactura, quotaiva1, quotaiva2, quotaiva3
    Dim L As String, Viatge, equip, Codi, nom, nif, Direccio, Telefon, P2, P, C1 As String, C2 As String, C3 As String, C4 As String, C5 As String, C6 As String, C7 As String, C8 As String, C9 As String, C10 As String, client As String, data As String, linea, Qr As rdoQuery, V1 As String, tT As String, mes As String, An As String, V2 As String, V3 As String, V4 As String, e As String, Dia As String, numFactura As String, fecha As Date
    Dim Estat As String, idFactura As String, Preu As String, import As String, tipusIva As String, iva As String, Referencia, RsAdo As New ADODB.Recordset
    Dim cap, Peu, baseIva3
    Dim CliSerie As String
    
    cap = "GS REGISTRO DE CONTROL" & Chr(13) & Chr(10)
    Peu = ""
    Set rs = rec("select newid()")
    If Not rs.EOF Then iD = rs(0)
    
    baseIva1 = 0: quotaiva1 = 0: baseIva2 = 0: quotaiva2 = 0: baseIva3 = 0: quotaiva3 = 0
    
    Sql = "Select Linea "
    Sql = Sql & "From Archivolines "
    Sql = Sql & "Where id = '" & Idrs & "' order by numlinea " 'And IsNumeric(SUBSTRING(linea, 1, 8)) = 1 "
    Set rs = rec(Sql)
    
    linea = 0
    Estat = "Capselera"
    While Not rs.EOF
        L = rs("Linea")
        If linea Mod 100 = 0 Then InformaMiss "Important Factura Mp. " & linea, True
        
        If Estat = "Capselera" And InStr(L, "F A C T U R A :") > 0 Then
            L = L
            numFactura = Right(L, Len(L) - InStr(L, ":"))
            If InStr(numFactura, "/") > 0 Then numFactura = Right(numFactura, Len(numFactura) - InStr(numFactura, "/"))
        End If
        
'        If Estat = "Capselera" And Len(Trim(Left(L, 40))) = 0 And Len(Trim(Mid(L, 44, 10))) > 0 Then
'            NumFac = Trim(Mid(L, 44, 10))
'        End If
        If Estat = "Capselera" And InStr(L, "Cliente  :") > 0 Then
            L = L
            client = Right(L, Len(L) - InStr(L, ":"))
            Cclient = ClientCodiExternCodiIntern(client)
        End If
        
        If Estat = "Capselera" And InStr(L, "Fecha    :") > 0 Then
            L = L
            fecha = CVDate(Left(Right(L, Len(L) - InStr(L, ":")), 12))
        End If
        
        If Estat = "Lineas" Then
            If IsDate(Left(L, 10)) Then
                Dim d As Date, Al As String, brut As Double, Total As Double
                L = Join(Split(L, ","), ".")
                d = CVDate(Mid(L, 1, 11))
                Al = Mid(L, 12, 8)
                brut = Mid(L, 20, 14)
                iva = Mid(L, 35, 25)
                Total = Mid(L, 60, 25)
                tipusIva = 1
                
                Sql = "Insert Into [" & NomTaulaFacturaData(fecha) & "] ([IdFactura],[Data],[Client],[Producte],[ProducteNom],[Acabat],[Preu],[Import],[Desconte],[TipusIva],[Iva],[Rec],[Referencia],[Servit],[Tornat]) "
                Sql = Sql & " Values "
                Sql = Sql & "('" & iD & "','" & d & "','" & Cclient & "','1','Albara Num " & Al & "' ,'1','" & brut & "','" & brut & "','','" & tipusIva & "','" & tipusIva & "','0','[Data:" & Format(d, "yyyy-mm-dd") & "][IdAlbara:" & Al & "]','1','0')"
                Peu = Peu & "D" & Format(d, "yyyymmdd") & Format(Al, "00000000") & Format(Int(brut * 100), "0000000000000") & Format(0, "0000000000000") & Format(Int(brut * 100), "0000000000000") & Chr(13) & Chr(10)
                ExecutaComandaSql Sql
            End If
        End If
        
        If Estat = "peu" And Len(Trim(L)) > 0 Then
            ' Join(Split(Join(Split(L, "."), ""), ","), ".")
            
            TipIva = Val(Join(Split(Join(Split(Mid(L, 39, 6), "."), ""), ","), "."))
            If TipIva = 4 Then
                baseIva1 = Val(Join(Split(Join(Split(Mid(L, 29, 10), "."), ""), ","), "."))
                quotaiva1 = Val(Join(Split(Join(Split(Mid(L, 45, 10), "."), ""), ","), "."))
            End If
            If TipIva = 7 Then
                baseIva2 = Val(Join(Split(Join(Split(Mid(L, 29, 10), "."), ""), ","), "."))
                quotaiva2 = Val(Join(Split(Join(Split(Mid(L, 45, 10), "."), ""), ","), "."))
            End If
            If TipIva = 16 Then
                baseIva3 = Val(Join(Split(Join(Split(Mid(L, 29, 10), "."), ""), ","), "."))
                quotaiva3 = Val(Join(Split(Join(Split(Mid(L, 45, 10), "."), ""), ","), "."))
            End If
            If Len(Trim(Mid(L, 65, 13))) > 0 Then
               totalfactura = Val(Join(Split(Join(Split(Mid(L, 65, 13), "."), ""), ","), "."))
            End If
        End If
        
        If Trim(L) = "TOT.BRUTO                     BASE   %IVA  IMP.IVA  %REC  IMP.EQ TOTAL FRA." Then
            Estat = "peu"
        End If
        
        If L = "   Fecha    Albar n  Total Bruto                  I.V.A.    Total ALBARAN  " Then
            Estat = "Lineas"
        End If
        
        If Estat = "peu" And L = "Forma de Pago:                                  " Then
            FacturacioCreaTaulesBuides fecha
            Estat = "Fin "
            Cclient = ClientCodiExternCodiIntern(client)
            CarregaDadesEmpresa Cclient, EmpCodi, EmpSerie, EmpNom, empNif, empAdresa, empCp, EmpLliure, empTel, empFax, empEMail, ClientCampMercantil, empCiutat, 0
            CarregaDadesClient Cclient, clientNom, ClientNomComercial, ClientNif, clientAdresa, clientCp, ClientLliure, clientTel, clientFax, clienteMail, ClienteFormaPago, clientCiutat, ClientCodiFact, AceptaDevolucions, CliSerie
            Set Q = Db.CreateQuery("", "Insert Into [" & NomTaulaFacturaIva(fecha) & "] ([IdFactura],[EmpresaCodi],[Serie],[NumFactura], [DataInici], [DataFi], [DataFactura], [DataEmissio], [DataVenciment], [FormaPagament], [Total],[ClientCodi] , [ClientNom], [ClientNif], [ClientAdresa], [ClientCp], [Tel], [Fax], [eMail], [ClientLliure], [EmpNom], [EmpNif], [EmpAdresa], [EmpCp], [EmpTel], [EmpFax], [EmpeMail], [EmpLliure], [BaseIva1], [Iva1], [BaseIva2], [Iva2], [BaseIva3], [Iva3], [BaseIva4], [Iva4], [BaseRec1], [Rec1], [BaseRec2], [Rec2], [BaseRec3], [Rec3], [BaseRec4], [Rec4],[ClientCiutat], [CampMercantil],[EmpCiutat],ClientCodiFac,[valorIva1] , [valorIva2], [valorIva3], [valorIva4], [valorRec1], [valorRec2], [valorRec3], [valorRec4], [IvaRec1], [IvaRec2], [IvaRec3], [IvaRec4],Reservat)  Values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ")
            Q.rdoParameters(0) = iD '[IdFactura]
            Q.rdoParameters(1) = EmpCodi '[EmpresaCodi]
            Q.rdoParameters(2) = EmpSerie '[Serie]
            Q.rdoParameters(3) = numFactura '[NumFactura]
            Q.rdoParameters(4) = fecha '[DataInici]
            Q.rdoParameters(5) = fecha '[DataFi]
            Q.rdoParameters(6) = fecha '[DataFactura]
            Q.rdoParameters(7) = Now '[DataEmissio]
            Q.rdoParameters(8) = VencimentActual '[DataVenciment]
            Q.rdoParameters(9) = ClienteFormaPago '[FormaPagament]
            Q.rdoParameters(10) = totalfactura  '[Total]
            Q.rdoParameters(11) = Cclient '[ClientCodi]
            Q.rdoParameters(12) = clientNom '[ClientNom]
            Q.rdoParameters(13) = ClientNif '[ClientNif]
            Q.rdoParameters(14) = clientAdresa '[ClientAdresa]
            Q.rdoParameters(15) = clientCp '[ClientCp]
            Q.rdoParameters(16) = clientTel '[Tel]
            Q.rdoParameters(17) = clientFax '[Fax]
            Q.rdoParameters(18) = clienteMail '[email]
            Q.rdoParameters(19) = ClientLliure '[ClientLliure]
            Q.rdoParameters(20) = EmpNom '[EmpNom]
            Q.rdoParameters(21) = empNif '[EmpNif]
            Q.rdoParameters(22) = empAdresa '[EmpAdresa]
            Q.rdoParameters(23) = empCp '[EmpCp]
            Q.rdoParameters(24) = empTel '[EmpTel]
            Q.rdoParameters(25) = empFax '[EmpFax]
            Q.rdoParameters(26) = empEMail '[Empemail]
            Q.rdoParameters(27) = EmpLliure '[EmpLliure]
            Q.rdoParameters(28) = baseIva1 '[Base1]
            Q.rdoParameters(29) = quotaiva1 '[Iva1]
            Q.rdoParameters(30) = baseIva2 '[Base2]
            Q.rdoParameters(31) = quotaiva2 '[Iva2]
            Q.rdoParameters(32) = baseIva3 '[Base3]
            Q.rdoParameters(33) = quotaiva3 '[Iva3]
            Q.rdoParameters(34) = 0 '[Base4]
            Q.rdoParameters(35) = 0 '[Iva4]
            Q.rdoParameters(36) = 0 '[Rec1]
            Q.rdoParameters(37) = 0 '[BaseRec1]
            Q.rdoParameters(38) = 0 '[Rec2]
            Q.rdoParameters(39) = 0 '[BaseRec1]
            Q.rdoParameters(40) = 0 '[Rec3]
            Q.rdoParameters(41) = 0 '[BaseRec1]
            Q.rdoParameters(42) = 0 '[Rec4]
            Q.rdoParameters(43) = 0 '[Rec4]
            Q.rdoParameters(44) = clientCiutat '[Rec4]
            Q.rdoParameters(45) = ClientCampMercantil '[Rec4]
            Q.rdoParameters(46) = empCiutat '[Rec4]
            Q.rdoParameters(47) = ClientCodiFact '[ClientCodiFac]
            Q.rdoParameters(48) = 4 '[valorIva1]
            Q.rdoParameters(49) = 7 '[valorIva2]
            Q.rdoParameters(50) = 16 '[valorIva3]
            Q.rdoParameters(51) = 0 '[valorIva4]
            Q.rdoParameters(52) = 0 '[valorRec1]
            Q.rdoParameters(53) = 0 '[valorRec2]
            Q.rdoParameters(54) = 0 '[valorRec3]
            Q.rdoParameters(55) = 0 '[valorRec4]
            Q.rdoParameters(56) = 0 '[IvaRec1]
            Q.rdoParameters(57) = 0 '[IvaRec2]
            Q.rdoParameters(58) = 0 '[IvaRec3]
            Q.rdoParameters(59) = 0 '[IvaRec4]
            Q.rdoParameters(60) = "V1.ImportadaMp" '[Reservat]
            Q.Execute
            creaRebuts fecha, iD, EmpCodi, Cclient
            cap = cap & "C" & Right("00000000" & numFactura, 8) & Format(fecha, "yyyymmdd") & "056" & Format(Int((baseIva1 + baseIva2 + baseIva3) * 100), "0000000000000") & Format(0, "0000000000000") & Format(Int((baseIva1) * 100), "0000000000000") & Format(400, "00000") & Format(Int((quotaiva1) * 100), "0000000000000") & Format(Int((baseIva2) * 100), "0000000000000") & Format(700, "00000") & Format(Int((quotaiva2) * 100), "0000000000000") & Format(Int((baseIva3) * 100), "0000000000000") & Format(1600, "00000") & Format(Int((quotaiva3) * 100), "0000000000000") & Format(Int(totalfactura * 100), "0000000000000") & Format(fecha, "yyyymmdd") & Format(Int(totalfactura * 100), "0000000000000") & "00000000" & Format(Int(0 * 100), "0000000000000") & "00000000" & Format(Int(0 * 100), "0000000000000") & Space(60) & Chr(13) & Chr(10)
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    Dim FicheroCutre
    Dim s() As Byte
    FicheroCutre = cap & Peu & "I" & Left(numFactura & ".Pdf" & Space(100), 100) & Chr(13) & Chr(10)
    s = FicheroCutre
        Set RsAdo = rec("select * from " & tablaArchivo() & " where fecha=(select max(fecha) from " & tablaArchivo() & ")", True)
        RsAdo.AddNew
        RsAdo("id").Value = iD
        RsAdo("nombre").Value = Left(nomfile, InStr(nomfile, ".") - 1) & ".Txt"
        RsAdo("descripcion").Value = "Ficher Traspas a System2010"
        RsAdo("extension").Value = "Txt"
        RsAdo("mime").Value = "text/plain"
        RsAdo("propietario").Value = ""
        RsAdo("archivo").Value = s
        RsAdo("fecha").Value = Now
        RsAdo("tmp").Value = 0
        RsAdo("down").Value = 1
        RsAdo.Update
        RsAdo.Close
    
'
'            ExecutaComandaSql "Insert into "
    
    

End Sub



Sub ImportaFitchersAlbaransMasterpanFormat1(Idrs)
    Dim Sql As String, rs As New ADODB.Recordset, L As String, Viatge, equip, Codi, nom, nif, Direccio, Telefon, P2, P, Tarifa, C1 As String, C2 As String, C3 As String, C4 As String, C5 As String, C6 As String, C7 As String, C8 As String, C9 As String, C10 As String, client As String, data As String, linea, Qr As rdoQuery
    Dim V1 As String
    Dim V2 As String, V3 As String, V4 As String
    Dim e As String
    
    ExecutaComandaSql "Drop Table Albarans_Importat_Masterpan "
    ExecutaComandaSql "create table Albarans_Importat_Masterpan  (Data [datetime],Client [nvarchar](255),Article [nvarchar](255),Nom [nvarchar](255),E [nvarchar](255),V1 [nvarchar](255),V2 [nvarchar](255),V3 [nvarchar](255),V4 [nvarchar](255),Q [float]) "
    
    Sql = "Select Linea "
    Sql = Sql & "From Archivolines "
    Sql = Sql & "Where id = '" & Idrs & "' order by numlinea " 'And IsNumeric(SUBSTRING(linea, 1, 8)) = 1 "
    Set rs = rec(Sql)
' Fecha    : 04/01/2001
' N§Cliente: 28118                                                 P g.:   1
'1002     BARRA DE 1/2                                   10,00 un.
    linea = 0
    While Not rs.EOF
        L = rs("Linea")
        DoEvents
        If linea Mod 100 = 0 Then InformaMiss "Important Albarans Mp. " & linea, True
        linea = linea + 1
        If Left(L, 11) = " Fecha    :" Then data = CVDate(Mid(L, 12, 12))
        If Left(L, 11) = " N§Cliente:" Then client = LTrim(RTrim(Mid(L, 12, 10)))
        If IsNumeric(Left(L, 8)) Then
            C1 = ""
            C2 = ""
            C3 = ""
            C1 = RTrim(LTrim(Mid(L, 1, 8)))
            C2 = RTrim(LTrim(Mid(L, 10, 30)))
            C3 = Join(Split(RTrim(LTrim(Mid(L, 120, 8))), ","), ".")   ' Treu ,
            V1 = Join(Split(RTrim(LTrim(Mid(L, 54, 11))), ","), ".")   ' Treu ,
            V2 = Join(Split(RTrim(LTrim(Mid(L, 73, 8))), ","), ".")   ' Treu ,
            V3 = Join(Split(RTrim(LTrim(Mid(L, 85, 8))), ","), ".")   ' Treu ,
            V4 = Join(Split(RTrim(LTrim(Mid(L, 97, 8))), ","), ".")   ' Treu ,
            e = RTrim(LTrim(Mid(L, 105, 5)))
'If V2 <> "" Then
'V2 = V2
'End If
            If C1 <> "" Then ExecutaComandaSql "Insert Into Albarans_Importat_Masterpan (Data ,Client ,Article ,Nom ,e,v1,v2,v3,v4,Q ) Values ('" & data & "','" & client & "','" & C1 & "','" & C2 & "','" & e & "','" & V1 & "','" & V2 & "','" & V3 & "','" & V4 & "','" & C3 & "')"
        End If
        rs.MoveNext
    Wend
    rs.Close
    
    ExecutaComandaSql "CREATE TABLE [ComandesModificades] ([Id] [uniqueidentifier] NULL , [TimeStamp] [datetime] NULL , [TaulaOrigen] [nvarchar] (255) NULL) ON [PRIMARY] "
    
    
    Set rs = rec("Select distinct data from Albarans_Importat_Masterpan order by data")
    While Not rs.EOF
        Debug.Print DonamNomTaulaServit(rs(0))
        DoEvents
        data = rs(0)
        ExecutaComandaSql "Delete  [" & DonamNomTaulaServit(rs(0)) & "]"
        
        Sql = "Insert Into [" & DonamNomTaulaServit(rs(0)) & "] (CodiArticle,PluUtilitzat, Client,QuantitatDemanada,QuantitatServida,QuantitatTornada,Viatge,Equip)  "
        Sql = Sql & "Select a.codi ,a.codi ,c.codi as client ,b.v1 ,0,0 ,'Viaje 1',E "
        Sql = Sql & "From Albarans_Importat_Masterpan B "
        Sql = Sql & "join Articles_Imp_Codis A  on A.unio = B.article "
        Sql = Sql & "join Articles Ar  on Ar.codi = A.Codi "
        Sql = Sql & "join Clients_Imp_Codis  C  on c.unio = B.Client Where Data = '" & data & "' and not V1='' "
        ExecutaComandaSql Sql
        
        Sql = "Insert Into [" & DonamNomTaulaServit(rs(0)) & "] (CodiArticle,PluUtilitzat, Client,QuantitatDemanada,QuantitatServida,QuantitatTornada,Viatge,Equip)  "
        Sql = Sql & "Select a.codi ,a.codi ,c.codi as client ,b.v2 ,0,0 ,'Viaje 2',E "
        Sql = Sql & "From Albarans_Importat_Masterpan B "
        Sql = Sql & "join Articles_Imp_Codis A  on A.unio = B.article "
        Sql = Sql & "join Articles Ar  on Ar.codi = A.Codi "
        Sql = Sql & "join Clients_Imp_Codis  C  on c.unio = B.Client Where Data = '" & data & "' and not V2='' "
        ExecutaComandaSql Sql
        
        Sql = "Insert Into [" & DonamNomTaulaServit(rs(0)) & "] (CodiArticle,PluUtilitzat, Client,QuantitatDemanada,QuantitatServida,QuantitatTornada,Viatge,Equip)  "
        Sql = Sql & "Select a.codi ,a.codi ,c.codi as client ,b.v3 ,0,0 ,'Viaje 3',E "
        Sql = Sql & "From Albarans_Importat_Masterpan B "
        Sql = Sql & "join Articles_Imp_Codis A  on A.unio = B.article "
        Sql = Sql & "join Articles Ar  on Ar.codi = A.Codi "
        Sql = Sql & "join Clients_Imp_Codis  C  on c.unio = B.Client Where Data = '" & data & "' and not V3='' "
        ExecutaComandaSql Sql
        
        Sql = "Insert Into [" & DonamNomTaulaServit(rs(0)) & "] (CodiArticle,PluUtilitzat, Client,QuantitatDemanada,QuantitatServida,QuantitatTornada,Viatge,Equip)  "
        Sql = Sql & "Select a.codi ,a.codi ,c.codi as client ,b.v4 ,0,0 ,'Viaje 4',E "
        Sql = Sql & "From Albarans_Importat_Masterpan B "
        Sql = Sql & "join Articles_Imp_Codis A  on A.unio = B.article "
        Sql = Sql & "join Articles Ar  on Ar.codi = A.Codi "
        Sql = Sql & "join Clients_Imp_Codis  C  on c.unio = B.Client Where Data = '" & data & "' and not V4='' "
        ExecutaComandaSql Sql
        
        rs.MoveNext
    Wend
    rs.Close
    
'    data = ""
'    linea = 0
'    Set Rs = rec("Select Data,a.codi As Codi,C.codi as client ,b.q As Q ,ar.familia As Familia from Albarans_Importat_Masterpan B join Articles_Imp_Codis A  on A.unio = B.article join Articles Ar  on Ar.codi = A.Codi join Clients_Imp_Codis  C  on c.unio = B.Client order by data,client,article,q ")
'    While Not Rs.EOF
'        DoEvents
'        If data <> Rs("Data") Then
'           Set Qr = Db.CreateQuery("", "Insert into [" & DonamNomTaulaServit(Rs("Data")) & "] (CodiArticle,PluUtilitzat, Client,QuantitatDemanada,QuantitatServida,QuantitatTornada,Viatge,Equip)  Values (?,?,?,?,?,?,?,?)  ")
'           ExecutaComandaSql "Delete * form [" & DonamNomTaulaServit(Rs("Data")) & "] "
'           data = Rs("Data")
'        End If
'        Qr.rdoParameters(0) = Rs("Codi")  'CodiArticle
'        Qr.rdoParameters(1) = Rs("Codi")  'PluUtilitzat
'        Qr.rdoParameters(2) = Rs("client")  'Client
'        Qr.rdoParameters(3) = Rs("Q")  'QuantitatDemanada
'        Qr.rdoParameters(4) = 0  'QuantitatServida
'        Qr.rdoParameters(5) = 0  'QuantitatTornada
'        Qr.rdoParameters(6) = Rs("Familia")  'Viatge
'        Qr.rdoParameters(7) = Rs("Familia")  'Equip
'        Qr.Execute
'        linea = linea + 1
'        If linea Mod 100 = 0 Then InformaMiss "Important Lin Albarans Mp. " & linea, True
'        Rs.MoveNext
'    Wend
'    Rs.Close

    
    
    
    
    
End Sub




Sub ImportarFamilia(Database As String, Clau As String, TarifaCodi As String)

   ExecutaComandaSql "update a set a.familia = i.param_1 from " & Database & ".Dbo.Articles A join " & Database & ".Dbo.Articles_Imp_Codis U on a.codi = U.codi Join importat_articu i on i.param_2 = u.unio "
   ExecutaComandaSql "update a set a.familia = f.param_2 + f.param_1 from " & Database & ".Dbo.Articles A join importat_famili f on a.familia = f.Param_1"
   ExecutaComandaSql "Delete " & Database & ".Dbo.Families "
   ExecutaComandaSql "Insert Into " & Database & ".dbo.families (nom,pare,estatus,nivell,utilitza) select distinct familia,'',4,3,'' from " & Database & ".dbo.articles"
   ExecutaComandaSql "update fa set pare = ff.Param_1 from " & Database & ".dbo.families fa join (select distinct c.Param_0 + c.Param_1  as Param_1,f.Param_2  + f.Param_1 as Param_2 from importat_articu a join importat_famili f on f.param_1 = a.param_1 join importat_casta c on a.Param_0 = c.Param_0) ff on fa.nom = ff.Param_2"
   ExecutaComandaSql "insert into " & Database & ".dbo.families (nom,pare,estatus,nivell,utilitza) select distinct f.pare,'" & Clau & "',4,2,'' from " & Database & ".dbo.families f "
   ExecutaComandaSql "insert into " & Database & ".dbo.families (nom,pare,estatus,nivell,utilitza)  Values ('" & Clau & "','Article',4,1,'' )"
   ExecutaComandaSql "insert into " & Database & ".dbo.families (nom,pare,estatus,nivell,utilitza)  Values ('Article','',4,1,'' )"
   
   Missatges_CalEnviar "Tpv_Families_", ""
   
End Sub


Sub ImportarPromocio(Database As String, Clau As String, TarifaCodi As String)

' select Param_4,param_6,Param_7   from importat_promoc where not param_6 ='' and not param_7 ='' and not param_4 ='' and param_0 ='S' and convert(datetime,param_1) < getdate() and convert(datetime,param_2) > getdate()
'   ExecutaComandaSql "update a set a.familia = i.param_1 from " & Database & ".Dbo.Articles A join " & Database & ".Dbo.Articles_Imp_Codis U on a.codi = U.codi Join importat_articu i on i.param_2 = u.unio "
'   ExecutaComandaSql "update a set a.familia = f.param_2 from " & Database & ".Dbo.Articles A join importat_famili f on a.familia = f.Param_1"
'   ExecutaComandaSql "Delete " & Database & ".Dbo.Families "
'   ExecutaComandaSql "Insert Into " & Database & ".dbo.families (nom,pare,estatus,nivell,utilitza) select distinct familia,'',4,3,'' from " & Database & ".dbo.articles"
'   ExecutaComandaSql "update fa set pare = ff.Param_1 from " & Database & ".dbo.families fa join (select distinct c.Param_1,f.Param_2 from importat_articu a join importat_famili f on f.param_1 = a.param_1 join importat_casta c on a.Param_0 = c.Param_0) ff on fa.nom = ff.Param_2"
'   ExecutaComandaSql "insert into iblatpa.dbo.families (nom,pare,estatus,nivell,utilitza) select distinct c.Param_1,'" & CLau & "',4,2,'' from importat_articu a join importat_famili f on f.param_1 = a.param_1 join importat_casta c on a.Param_0 = c.Param_0"
'   ExecutaComandaSql "insert into iblatpa.dbo.families (nom,pare,estatus,nivell,utilitza)  Values ('" & CLau & "','Article',4,1,'' )"
'   ExecutaComandaSql "insert into iblatpa.dbo.families (nom,pare,estatus,nivell,utilitza)  Values ('Article','',4,1,'' )"
'   ExecutaComandaSql "delete " & Database & ".Dbo.tarifesespecials   where tarifacodi in (select distinct c.codi from importat_promoc p join " & Database & ".Dbo.articles_imp_codis I on p.Param_6 = unio  join " & Database & ".Dbo.clients_imp_codis c On  p.Param_4 = C.Unio where not p.param_6 ='' and not p.param_7 ='' and not p.param_4 ='' and p.param_0 ='" & CLau & "' and convert(datetime,p.param_1) < getdate() and convert(datetime,p.param_2) > getdate())"
'   ExecutaComandaSql "delete " & Database & ".Dbo.tarifesespecials "
'   ExecutaComandaSql "insert into " & Database & ".Dbo.tarifesespecials  (tarifacodi,tarifanom,codi,preu,preumajor) select c.codi as tarifacodi,p.Param_4  as tarifanom,i.codi as codi ,p.Param_7   as preu,p.Param_7   as preumajor from importat_promoc p join " & Database & ".Dbo.articles_imp_codis I on p.Param_6 = unio  join " & Database & ".Dbo.clients_imp_codis c On  p.Param_4 = C.Unio where not p.param_6 ='' and not p.param_7 ='' and not p.param_4 ='' and p.param_0 ='" & CLau & "' and convert(datetime,p.param_1) < getdate() and convert(datetime,p.param_2) > getdate()"
'   ExecutaComandaSql "update " & Database & ".Dbo.clients set [desconte 5] = codi where codi in (select distinct c.codi from importat_promoc p join " & Database & ".Dbo.articles_imp_codis I on p.Param_6 = unio  join " & Database & ".Dbo.clients_imp_codis c On  p.Param_4 = C.Unio where not p.param_6 ='' and not p.param_7 ='' and not p.param_4 ='' and p.param_0 ='" & CLau & "' and convert(datetime,p.param_1) < getdate() and convert(datetime,p.param_2) > getdate())"
   
   ExecutaComandaSql " drop   TABLE " & Database & ".Dbo.[ProductesPromocionats] "
   ExecutaComandaSql " CREATE TABLE " & Database & ".Dbo.[ProductesPromocionats] ([Id] [nvarchar] (50) NULL ,  [Di] [datetime] NULL ,  [Df] [datetime] NULL ,  [D_Producte] [float] NULL ,    [D_Quantitat] [float] NULL ,   [S_Producte] [float] NULL ,    [S_Quantitat] [float] NULL ,   [S_Preu] [float] NULL ,    [Client] [nvarchar] (50) NULL ) ON [PRIMARY]"
   ExecutaComandaSql " insert into  " & Database & ".Dbo.[ProductesPromocionats] ([Id],[Di],[Df],[D_Producte],[D_Quantitat],[S_Producte],[S_Quantitat],[S_Preu],[Client]) select newid() as id,convert(datetime,p.param_1) as Di,convert(datetime,p.param_2) As Df,i.codi ,1,i.codi,1,p.Param_7 ,c.codi from importat_promoc p join " & Database & ".Dbo.articles_imp_codis I on p.Param_6 = unio join " & Database & ".Dbo.clients_imp_codis c On  p.Param_4 = C.Unio where not p.param_6 ='' and not p.param_7 ='' and not p.param_4 ='' and p.param_0 ='" & Clau & "' and convert(datetime,p.param_2)  > getdate() "
   ExecutaComandaSql " insert into  " & Database & ".Dbo.[ProductesPromocionats] ([Id],[Di],[Df],[D_Producte],[D_Quantitat],[S_Producte],[S_Quantitat],[S_Preu],[Client]) select newid() as id,convert(datetime,p.param_1) as Di,convert(datetime,p.param_2) As Df,i.codi ,1,i.codi,1,p.Param_7 ,0 from importat_promoc p join " & Database & ".Dbo.articles_imp_codis I on p.Param_6 = unio where not p.param_6 ='' and not p.param_7 ='' and p.param_4 ='' and p.param_0 ='" & Clau & "' " ' and convert(datetime,p.param_2)  >= getdate() "
   
End Sub



Sub ImportaActualitzaNoms(Database As String, Clau As String)

   ExecutaComandaSql "update a set a.Nom = t.Param_2 from  " & Database & ".Dbo.articles a join " & Database & ".Dbo.Articles_Imp_Codis c on a.codi = c.codi join Importat_artemp t on t.param_1 = c.unio  and t.Param_0 ='" & Clau & "'"
         
End Sub

Sub ImportaTarifa(Database As String, TarifaCodi As String, Empresa As String)

   ExecutaComandaSql "update a set a.preu = t.param_3 from " & Database & ".Dbo.articles a join " & Database & ".Dbo.Articles_Imp_Codis c on a.codi = c.codi join importat_tarifa t on t.param_2 = c.unio and t.param_1 = " & TarifaCodi & " and t.param_0 = '" & Empresa & " '"
   ExecutaComandaSql "delete " & Database & ".Dbo.articles where preu = 0"
   ExecutaComandaSql "update " & Database & ".Dbo.Articles set preu = 0.01 where nom like 'Reventa Centralizada'"
   ExecutaComandaSql "Insert Into " & Database & ".dbo.tarifesespecials (Tarifacodi,tarifanom,codi,preu,preumajor) select t.param_1,t.Param_1,u.codi,t.param_3,t.param_3  from importat_tarifa t join  " & Database & ".Dbo.Articles_Imp_Codis U on t.Param_2 = U.Unio and t.param_0 = '" & Empresa & "' "
      
   'Missatges_CalEnviar "Tarifa", ""
   
End Sub


Sub ImportarClients(ByVal p1 As String, ByVal P2 As String, ByVal P3 As String)
   Dim rs As rdoResultset
   
   ExecutaComandaSql "CREATE TABLE [Importat_Params] ([Clau]    [nvarchar] (255) NULL ,[Valor_1] [nvarchar] (255) NULL ,[Valor_2] [nvarchar] (255) NULL ,[Valor_3] [nvarchar] (255) NULL ,[Valor_4] [nvarchar] (255) NULL ,[Valor_5] [nvarchar] (255) NULL ) ON [PRIMARY]"
   
   Set rs = Db.OpenResultset("select valor_1,Valor_2 from Importat_Params where Clau = 'Empresa'")
   While Not rs.EOF
      ImportaClients rs(0), rs(1)
      rs.MoveNext
   Wend
   rs.Close
   Missatges_CalEnviar "Clients", ""
   
End Sub


Sub AjuntaTpvS(ByVal An As String, ByVal mes As String, Dia As String)
    Dim rs As rdoResultset, NomTaula As String
    
    NomTaula = NomTaulaVentas(DateSerial(An, mes, 1))
    
    Set rs = Db.OpenResultset("select * from " & DonamNomTaulaTpvEquivalents & " ")
    If Not rs.EOF Then ExecutaComandaSql "Update [" & NomTaula & "] Set estat = botiga where estat = ''"
    While Not rs.EOF
        ExecutaComandaSql "Update [" & NomTaula & "] Set Botiga = " & rs("valor2") & " where Botiga = " & rs("Valor1") & " "
        ExecutaComandaSql "Update [" & DonamNomTaulaServit(DateSerial(An, mes, Dia)) & "] Set Client = " & rs("valor2") & " where Client = " & rs("Valor1") & "  And Viatge = 'Auto' And Equip = 'Auto' "
        rs.MoveNext
    Wend
    rs.Close

End Sub



Sub ImportarComanda(ByVal p1 As String, ByVal P2 As String, ByVal P3 As String, Taula As String)
    Dim rs As rdoResultset, Rs2 As rdoResultset, Staula As String
On Error GoTo 0
    Set rs = Db.OpenResultset("select valor_1,Valor_2 from Importat_Params where Clau = 'Empresa' And Valor_2 = '" & Car(P3) & "' ")
    Staula = Car(Taula)
    Staula = Left(Staula, InStr(Staula, ".") - 1)
     
    While Not rs.EOF
       Set Rs2 = Db.OpenResultset("select distinct Param_2 from [" & Staula & "] ")
       While Not Rs2.EOF
          If Not IsNull(Rs2(0)) Then ImportaComanda rs(0), Rs2(0), Staula
          Rs2.MoveNext
       Wend
       Rs2.Close
       rs.MoveNext
    Wend
    rs.Close

'   Set rs2 = Db.OpenResultset("select distinct Param_2 from [" & Staula & "] ")
'   While Not rs2.EOF
'      If Not IsNull(rs2(0)) Then ImportaComanda rs(0), rs2(0), Staula
'      rs2.MoveNext
'   Wend
'   rs2.Close
'
'   Missatges_CalEnviar "Comandes", "[" & Bo & "][" & Format(D, "dd-mm-yyyy") & "]"
   
End Sub

Sub ImportarArticles(ByVal p1 As String, ByVal P2 As String, ByVal P3 As String)
   Dim rs As rdoResultset
   
   ExecutaComandaSql "CREATE TABLE [Importat_Params] ([Clau]    [nvarchar] (255) NULL ,[Valor_1] [nvarchar] (255) NULL ,[Valor_2] [nvarchar] (255) NULL ,[Valor_3] [nvarchar] (255) NULL ,[Valor_4] [nvarchar] (255) NULL ,[Valor_5] [nvarchar] (255) NULL ) ON [PRIMARY]"
   
   Set rs = Db.OpenResultset("select valor_1,Valor_2,Valor_3 from Importat_Params where Clau = 'Empresa'")
   While Not rs.EOF
      ImportaArticles rs(0), rs(1), rs(2)
      Missatges_CalEnviar "Articles", "", False, rs(0)
      rs.MoveNext
   Wend
   rs.Close
   
End Sub



Sub ImportaFitchersLogiis()
   Dim rs As rdoResultset, Fitcher, f, P, P2
   Dim rst As New ADODB.Recordset
   Dim adoConn As New ADODB.Connection
   Dim mStream As New ADODB.Stream
   Dim strDescr, vChunk, K As Double
   Dim aa As String, Sql As String
   Dim OrId As String, OrData As Date
   Dim Qi As rdoQuery
   
   Dim Mode As String, client As String, producte As String, pes As String
   
'Exit Sub
   If Not ExisteixTaula("monas") Then
      Sql = "CREATE TABLE monas ( "
      Sql = Sql & " [modo]     [nvarchar] (255) NULL ,"
      Sql = Sql & " [client]   [float]  NULL ,"
      Sql = Sql & " [Producte] [float]  NULL ,"
      Sql = Sql & " [Pes]      [float]  NULL ,"
      Sql = Sql & " [NumLinea] [float] NULL)"
      ExecutaComandaSql Sql
   End If
   
    
On Error GoTo 0
   
    Set Qi = Db.CreateQuery("", "Insert into [monas] (modo,client,producte,pes) Values (?,?,?,?) ")
        
        f = FreeFile
        Open "c:\aa.txt" For Input As #f
        K = 0
        While Not EOF(f)
           Line Input #f, aa
           K = K + 1
           If InStr(aa, "217.125.105.234") > 0 Then
            If InStr(aa, "/Facturacion/Recorda/facturar/Save.asp") > 0 Then
                If InStr(aa, "modo=SAVE") Then Mode = "Save"
                If InStr(aa, "modo=DELETE") Then Mode = "Delete"
                
                client = ""
                producte = ""
                pes = ""
    If Mode = "Save" Then
                P = InStr(aa, "cliente=")
                P2 = InStr(P, aa, "&")
                client = Mid(aa, P + 8, P2 - P - 8)
                
                P = InStr(aa, "newArt=")
                P2 = InStr(P, aa, "&")
                producte = Mid(aa, P + 7, P2 - P - 7)
                
                P = InStr(aa, "qd=")
                P2 = InStr(P, aa, "&")
                pes = Mid(aa, P + 3, P2 - P - 3)
                
                P = InStr(pes, ",")
                If P > 0 Then pes = Left(pes, P - 1) & "." & Right(pes, Len(pes) - P)
                
                
                Qi.rdoParameters(0) = Mode
                Qi.rdoParameters(1) = client
                Qi.rdoParameters(2) = producte
                Qi.rdoParameters(3) = pes
                Qi.Execute
End If
            End If
           End If
           DoEvents
        Wend
   
End Sub




Sub ImportaFitchers()
   Dim rs As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset, Rs4 As rdoResultset, Fitcher, f, LlistaShlinder As String, nomfile As String, tipusFile As String, Extensions As String, CodiEmpresa, data, Treb, i_liquid, i_Irpf, SSTrabaj, TotalBr, SSempre, Nnomfile As String, ppp As Integer, rst As New ADODB.Recordset, K As Double, P As Integer, adoConn As New ADODB.Connection, mStream As New ADODB.Stream, aa As String, Sql As String, Qi As rdoQuery, Empresa As String, Banco As String, Oficina As String, numcuenta As String, Fi As String, Ff As String, SaldoInicial As String, NombreCliente As String, CodigoCliente As String, Idrs As String, Insertala As Boolean, IDArchivo, NumLinea, FechaNomina As Date, FechaNominaStr As String, PEsp, IrpEsp, empresaborrada As String, Line, SSTrabaj2, Patro
   Set rs = Db.OpenResultset("select Id from " & tablaArchivo() & "  Where descripcion not like 'Interpretat %' and descripcion Like  '%Interpreta%' and not id in (Select distinct id from " & DonamNomTaulaArchivoLines() & " ) ")
On Error GoTo error

    If Not rs.EOF Then
        On Error Resume Next
        db2.Close
        On Error GoTo 0
        Set db2 = New ADODB.Connection
        db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"
   
        Set Qi = Db.CreateQuery("", "Insert into " & DonamNomTaulaArchivoLines & " (Id,nombre,fecha,linea,numlinea) Values (?,?,?,?,?) ")
        While Not rs.EOF
            Set rst = rec("select * from Archivo Where id  = '" & rs("Id") & "' ")
            GuardaHistoric Cnf.llicencia, Now, "ImportaFicher", rst("Id"), rst("Nombre"), rst("fecha")
            nomfile = UCase(rst("Nombre"))
            Nnomfile = nomfile
            Debug.Print rst("fecha") & " -- " & rst("Nombre")
            Qi.rdoParameters(0) = rst("Id")
            Qi.rdoParameters(1) = rst("Nombre")
            Qi.rdoParameters(2) = rst("fecha")
            
            tipusFile = "N43"
            If InStr(nomfile, ".PRN") > 0 Then tipusFile = "Masterpan"
            If InStr(nomfile, ".HTA") > 0 Then tipusFile = "N43"
            If InStr(nomfile, ".HTM") > 0 Then tipusFile = "NominasHtml"
            If InStr(nomfile, ".NMSG") > 0 Then tipusFile = "NominasHtml"
            If InStr(nomfile, ".NLC") > 0 Then tipusFile = "SousLc"
            If InStr(nomfile, ".NLC2") > 0 Then tipusFile = "SousLc2"
            If InStr(nomfile, "SUENLACE") > 0 Then tipusFile = "SousA3"
            If InStr(nomfile, ".NMSP") > 0 Then tipusFile = "SousSp"
            If InStr(nomfile, ".GEN") > 0 Then tipusFile = "SousSILEMA"
            If InStr(nomfile, "SOUS") > 0 Then tipusFile = "SousA3"
            If InStr(nomfile, "NOMINESPANET") > 0 Then tipusFile = "SousPANET"
            If InStr(nomfile, "Interpreta.Xls") > 0 Then tipusFile = "Interpreta"
            If InStr(nomfile, ".ALB") > 0 Then tipusFile = "AlbaransMasterpan"
            
'            If InStr(NomFile , ".DAT") > 0 Then tipusFile = "SousA3"
            InformaMiss rst("Nombre"), True
            ExecutaComandaSql "Delete " & DonamNomTaulaArchivoLines & " Where Id = '" & rs("Id") & "' "
            Dim campo As ADODB.Field
            
            ExecutaComandaSql "Update Archivo Set Descripcion = 'Interpretant...' + cast(getdate() as nvarchar)  Where id  = '" & rs("Id") & "' "
            IDArchivo = rs("Id")
            If InStr(UCase(nomfile), ".XLS") > 0 Then
                ExecutaComandaSql "Update Archivo Set Descripcion = 'Interpretando... ' + cast(getdate() as nvarchar)  Where id  = '" & IDArchivo & "' "
                Set Rs2 = Db.OpenResultset("select Archivo from Archivo Where id  = '" & rs("Id") & "' ", rdOpenKeyset)
                MyKill "c:\" & nomfile
                ColumnToFile Rs2.rdoColumns("Archivo"), "c:\" & nomfile, 102400, Rs2("Archivo").ColumnSize
                Rs2.Close
                If tipusFile = "NominasHtml" Then CarregaNominasHtml "c:\" & nomfile
                If InStr(UCase(nomfile), ".XLS") > 0 Then
                    If InStr(UCase(nomfile), "NOMINES") <= 0 Then
                        If UCase(nomfile) = "DEPENDENTES.XLS" Then
                            CarregaDependentesXls "c:\" & nomfile
                        Else
                            CarregaClientsXls "c:\" & nomfile
                        End If
                        MyKill "c:\" & nomfile
                    End If
                End If

                Qi.rdoParameters(3) = "Ok"
                Qi.rdoParameters(4) = 0
                Qi.Execute
                ExecutaComandaSql "Update Archivo Set Descripcion = 'Interpretat ' + cast(getdate() as nvarchar)  Where id  = '" & IDArchivo & "' "
            ElseIf InStr(UCase(nomfile), ".NMSP") Or InStr(UCase(nomfile), ".NLC") Or InStr(UCase(nomfile), ".GEN") Then
                Set Rs2 = Db.OpenResultset("select Archivo from Archivo Where id  = '" & rs("Id") & "' ", rdOpenKeyset)
                MyKill "c:\aa.Bin"
                ColumnToFile Rs2.rdoColumns("Archivo"), "c:\aa.Bin", 102400, Rs2("Archivo").ColumnSize
                Rs2.Close
                Qi.rdoParameters(3) = "Ok"
                Qi.rdoParameters(4) = 0
                Qi.Execute
                FileDeBinaTxt "c:\aa"
                
                f = FreeFile
                Open "c:\aa.txt" For Input As #f
                K = 0
                
                While Not EOF(f)
                    Line Input #f, aa
                    Debug.Print aa
                    Qi.rdoParameters(3) = CStr(aa)
                    Qi.rdoParameters(4) = K
                    Qi.Execute
                    DoEvents
                    If K Mod 100 = 0 Then InformaMiss rst("Nombre") & " " & K, True
                    K = K + 1
                Wend
                Close #f
            
            Else
                Set Rs2 = Db.OpenResultset("select Archivo from Archivo Where id  = '" & rs("Id") & "' ", rdOpenKeyset)
                MyKill "c:\aa.Bin"
                ColumnToFile Rs2.rdoColumns("Archivo"), "c:\aa.Bin", 102400, Rs2("Archivo").ColumnSize
                Rs2.Close
                Qi.rdoParameters(3) = "Ok"
                Qi.rdoParameters(4) = 0
                Qi.Execute
                FileDeBinaTxt "c:\aa"
                
                f = FreeFile
                Open "c:\aa.txt" For Input As #f
                K = 0
                
                While Not EOF(f)
                    Line Input #f, aa
                    Debug.Print aa
                    Insertala = True
                    If K = 0 Or Empresa = "" Then
                        Insertala = False
                        Empresa = ""
                        Banco = ""
                        Oficina = ""
                        numcuenta = ""
                        Fi = ""
                        Ff = ""
                        SaldoInicial = ""
                        NombreCliente = ""
                        CodigoCliente = ""
                        If tipusFile = "N43" And K = 0 Then
                            aa = LTrim(aa)
                            aa = Right(aa, Len(aa) - P + 1)
                            Empresa = EmpresaActual
                            Banco = Mid(aa, 3, 4)
                            Oficina = Mid(aa, 7, 4)
                            numcuenta = Mid(aa, 11, 10)
                            Fi = Mid(aa, 21, 6)
                            Ff = Mid(aa, 27, 6)
                            SaldoInicial = Mid(aa, 34, 14)
                            NombreCliente = Mid(aa, 52, 26)
                            CodigoCliente = Mid(aa, 78, 3)
                        End If
                        If (tipusFile = "AlbaransMasterpan" Or tipusFile = "Masterpan" Or tipusFile = "SousA3" Or tipusFile = "SousLc" Or tipusFile = "SousLc2" Or tipusFile = "SousSp") And K = 0 Then
                            Empresa = "Ok"
                            Insertala = True
                            aa = Right(aa, Len(aa) - P + 1)
                        End If
                    End If
                    K = K + 1
                    
                    If Insertala Then
                        Qi.rdoParameters(3) = CStr(aa)
                        Qi.rdoParameters(4) = K
                        Qi.Execute
                    End If
                    DoEvents
                    If K Mod 100 = 0 Then InformaMiss rst("Nombre") & " " & K, True
                Wend
                Close #f
            End If
            
            InformaMiss rst("Nombre") & " " & K, True
            Idrs = rs("Id")
            
            If UCase(Right(nomfile, 4)) = ".HTA" Then nomfile = UCase("norma43.n43")
            If UCase(Right(nomfile, 4)) = UCase(".n43") Then nomfile = UCase("norma43.n43")
            'If UCase(nomfile) = "NOMINES_HIT.XLS" Then nomfile = "NominasXls"
            If tipusFile = "SousA3" Then nomfile = "SousA3"
            If tipusFile = "SousLc" Then nomfile = "SousLc"
            If tipusFile = "SousLc2" Then nomfile = "SousLc2"
            If tipusFile = "SousSp" Then nomfile = "SousSp"
            If tipusFile = "SousSILEMA" Then nomfile = "SousSILEMA"
            If tipusFile = "AlbaransMasterpan" Then nomfile = UCase("Albarans.Prn")
            If tipusFile = "Masterpan" Then nomfile = "Masterpan"
            If tipusFile = "SousPANET" Then nomfile = "SousPANET"
            If tipusFile = "Interpreta" Then nomfile = "Interpreta.Xls"
            If tipusFile = "NominasHtml" Then nomfile = "NominasHtml"
            
            'If UCase(rst("Descripcion")) = UCase("Interpreta") Then NomFile = "Interpreta.Xls"
            
            
            Select Case nomfile
                Case "Masterpan"
                    ImportaFitchersFacturesMasterpan IDArchivo, Nnomfile
                Case "TRASART.TXT"
                    'ImportaFitchersArticles Fitcher
                    'Missatges_CalEnviar "Articles", "", False, Rs(0)
                Case UCase("Interpreta.Xls")
'                    CarregaClientsXls "c:\aa.Bin"
                Case UCase("Tarifes.PRN")
                    ImportaFitchersTarifesMasterpan Idrs
                Case UCase("PREESP.PRN")
                    ImportaFitchersPreusEspecialsMasterpan Idrs
                Case UCase("Clients.Prn")
                    ImportaFitchersClientsMasterpan Idrs
                    Missatges_CalEnviar "Clients", "", False
                Case UCase("Articles.Prn")
                    ImportaFitchersArticlesMasterpan Idrs
                    Missatges_CalEnviar "Articles", "", False
                Case UCase("Tangram.Txt")
                    ImportaFitchersAlbaransTangram Idrs
                Case UCase("Albarans.Prn")
                    ImportaFitchersAlbaransMasterpan Idrs
                    'Missatges_CalEnviar "Articles", "", False
                Case "NominasHtml"
                    PillaNominasHtml "c:\aa.txt", Idrs
                Case "SousSILEMA"
                    FechaNomina = "01-01-01"
                    If IsNumeric(Left(Nnomfile, 5)) Then FechaNomina = "28" & "-" & Mid(Nnomfile, 5, 2) & "-" & Mid(Nnomfile, 3, 2)
                    ImportaNominasSILEMA "c:\aa.txt", FechaNomina, Idrs
                Case "SousPANET"
                    ImportaNominasPANET "c:\" & Nnomfile, Idrs
                    MyKill "c:\" & Nnomfile
                Case "NOMINES_HIT.XLS"
                    ImportaNominasExcel "c:\" & nomfile, Idrs
                    MyKill "c:\" & nomfile
                Case UCase("norma43.n43")
'                    Idrs = "{0A19DC36-8357-41F4-ACB0-010C4995DA4C}"
                    ExecutaComandaSql "Delete " & DonamNomTaulaNorma43 & " where IdFichero = '" & Idrs & "' "
                    ExecutaComandaSql "Drop Table " & DonamNomTaulaNorma43 & "_Tmp "
                    
                    'sql = "Select "
                    'sql = sql & "a1.Id as IdFichero,"
                    'sql = sql & "Newid() as IdNorma43,"
                    'sql = sql & "'" & Empresa & "' as [Comu_Empresa],"
                    'sql = sql & "'" & Banco & "' as [Comu_Banco],"
                    'sql = sql & "'" & Oficina & "' as [Comu_Oficina],"
                    'sql = sql & "'" & numcuenta & "' as [Comu_numcuenta],"
                    'sql = sql & "'" & Fi & "' as [Comu_Fi],"
                    'sql = sql & "'" & Ff & "' as [Comu_Ff],"
                    'sql = sql & "'" & SaldoInicial & "' as [Comu_SaldoInicial],"
                    'sql = sql & "'" & NombreCliente & "' as [Comu_NombreCliente],"
                    'sql = sql & "'" & CodigoCliente & "' as [Comu_CodigoCliente],"
                    'sql = sql & "a1.numlinea As numlineaA ,a2.numlinea As numlineaB, "
                    'sql = sql & "SUBSTRING (a1.linea ,1,2 ) tipusRegistreA,"
                    'sql = sql & "SUBSTRING (a1.linea ,3,4 ) LliureA,"
                    'sql = sql & "SUBSTRING (a1.linea ,7,4 ) Oficina,"
                    'sql = sql & "SUBSTRING (a1.linea ,11,6 ) DataOperacio,"
                    'sql = sql & "SUBSTRING (a1.linea ,17,6 ) DataValor,"
                    'sql = sql & "SUBSTRING (a1.linea ,23,2 ) ConceptoComun,"
                    'sql = sql & "SUBSTRING (a1.linea ,25,3 ) ConceptoPropio,"
                    'sql = sql & "SUBSTRING (a1.linea ,28,1 ) DeveHaver,"
                    'sql = sql & "SUBSTRING (a1.linea ,29,14 ) Importe,"
                    'sql = sql & "SUBSTRING (a1.linea ,43,10 ) Documento,"
                    'sql = sql & "SUBSTRING (a1.linea ,53,11 ) Referencia,"
                    'sql = sql & "SUBSTRING (a1.linea ,64,16 ) Referencia2,"
                    'sql = sql & "SUBSTRING (a2.linea ,1,2 ) tipusRegistreB,"
                    'sql = sql & "SUBSTRING (a2.linea ,3,2 ) CodigoDato,"
                    'sql = sql & "SUBSTRING (a2.linea ,5,38 ) Concepto1,"
                    'sql = sql & "SUBSTRING (a2.linea ,43,38 ) Concepto2,"
                    'sql = sql & "SUBSTRING (a2.linea ,81,38 ) Concepto3,"
                    'sql = sql & "SUBSTRING (a2.linea ,119,38 ) Concepto4,"
                    'sql = sql & "SUBSTRING (a2.linea ,157,38 ) Concepto5 "
                    'sql = sql & " Into " & DonamNomTaulaNorma43 & "_Tmp  "
                    'sql = sql & "from Archivolines A1 join Archivolines A2 on a2.numlinea = a1.numlinea + 1 "
                    'sql = sql & "where a1.id = '" & Idrs & "' and a2.id = '" & Idrs & "' and SUBSTRING (a1.linea ,1,2 )='22' and SUBSTRING (a2.linea ,1,2 )='23' "
'                    Sql = Sql & " And (SUBSTRING (a1.linea ,17,6 ) > (Select isnull(Max(dataOperacio),'000000') from Norma43 Where [Comu_numcuenta] = '" & numcuenta & "' ) or   SUBSTRING (a1.linea ,17,6 ) < (Select isnull(Min(dataOperacio),'999999') from Norma43 Where [Comu_numcuenta] = '" & numcuenta & "' ))"
                    
                    
                    Sql = "Select "
                    Sql = Sql & "a1.Id as IdFichero,"
                    Sql = Sql & "Newid() as IdNorma43,"
                    Sql = Sql & "'" & Empresa & "' as [Comu_Empresa],"
                    Sql = Sql & "'" & Banco & "' as [Comu_Banco],"
                    Sql = Sql & "'" & Oficina & "' as [Comu_Oficina],"
                    Sql = Sql & "'" & numcuenta & "' as [Comu_numcuenta],"
                    Sql = Sql & "'" & Fi & "' as [Comu_Fi],"
                    Sql = Sql & "'" & Ff & "' as [Comu_Ff],"
                    Sql = Sql & "'" & SaldoInicial & "' as [Comu_SaldoInicial],"
                    Sql = Sql & "'" & NombreCliente & "' as [Comu_NombreCliente],"
                    Sql = Sql & "'" & CodigoCliente & "' as [Comu_CodigoCliente],"
                    Sql = Sql & "a1.numlinea As numlineaA ,isnull(a2.numlinea, -1) As numlineaB,"
                    Sql = Sql & "SUBSTRING (a1.linea ,1,2 ) tipusRegistreA,"
                    Sql = Sql & "SUBSTRING (a1.linea ,3,4 ) LliureA,"
                    Sql = Sql & "SUBSTRING (a1.linea ,7,4 ) Oficina,"
                    Sql = Sql & "SUBSTRING (a1.linea ,11,6 ) DataOperacio,"
                    Sql = Sql & "SUBSTRING (a1.linea ,17,6 ) DataValor,"
                    Sql = Sql & "SUBSTRING (a1.linea ,23,2 ) ConceptoComun,"
                    Sql = Sql & "SUBSTRING (a1.linea ,25,3 ) ConceptoPropio,"
                    Sql = Sql & "SUBSTRING (a1.linea ,28,1 ) DeveHaver,"
                    Sql = Sql & "SUBSTRING (a1.linea ,29,14 ) Importe,"
                    Sql = Sql & "SUBSTRING (a1.linea ,43,10 ) Documento,"
                    Sql = Sql & "SUBSTRING (a1.linea ,53,11 ) Referencia,"
                    Sql = Sql & "SUBSTRING (a1.linea ,64,16 ) Referencia2,"
                    Sql = Sql & "SUBSTRING (isnull(a2.linea, '23') ,1,2 ) tipusRegistreB,"
                    Sql = Sql & "SUBSTRING (isnull(a2.linea, '2301') ,3,2 ) CodigoDato,"
                    Sql = Sql & "case when a2.linea is null then SUBSTRING (a1.linea ,53,11)+SUBSTRING (a1.linea ,64,16 ) else SUBSTRING (a2.linea ,5,38) end Concepto1,"
                    Sql = Sql & "SUBSTRING (isnull(a2.linea, '') ,43,38 ) Concepto2,"
                    Sql = Sql & "SUBSTRING (isnull(a2.linea, '') ,81,38 ) Concepto3,"
                    Sql = Sql & "SUBSTRING (isnull(a2.linea, '') ,119,38 ) Concepto4,"
                    Sql = Sql & "SUBSTRING (isnull(a2.linea, '') ,157,38 ) Concepto5 "
                    Sql = Sql & " Into " & DonamNomTaulaNorma43 & "_Tmp  "
                    Sql = Sql & "From "
                    Sql = Sql & "(select * from ArchivoLines where id='" & Idrs & "' and Linea like '22%') a1 "
                    Sql = Sql & "Left Join "
                    Sql = Sql & "(select * from ArchivoLines where id='" & Idrs & "' and Linea like '23%') a2 on a1.NumLinea+1=a2.NumLinea "
                     
                    ExecutaComandaSql Sql  ' Insertem a temporals
                    
                    Sql = "Insert Into " & DonamNomTaulaNorma43 & " (IdNorma43,IdFichero,[Comu_Empresa],[Comu_Banco],[Comu_Oficina],[Comu_numcuenta],[Comu_Fi],[Comu_Ff],[Comu_SaldoInicial],[Comu_NombreCliente],[Comu_CodigoCliente],numlineaA ,numlineaB,tipusRegistreA,LliureA,Oficina,DataOperacio,DataValor,ConceptoComun,ConceptoPropio,DeveHaver,Importe,Documento,Referencia,Referencia2,tipusRegistreB,CodigoDato,Concepto1,Concepto2,Concepto3,Concepto4,Concepto5) "
                    Sql = Sql & " Select newid(),IdFichero,[Comu_Empresa],[Comu_Banco],[Comu_Oficina],[Comu_numcuenta],[Comu_Fi],[Comu_Ff],[Comu_SaldoInicial],[Comu_NombreCliente],[Comu_CodigoCliente],numlineaA ,numlineaB,tipusRegistreA,LliureA,Oficina,DataOperacio,DataValor,ConceptoComun,ConceptoPropio,DeveHaver,Importe,Documento,Referencia,Referencia2,tipusRegistreB,CodigoDato,Concepto1,Concepto2,Concepto3,Concepto4,Concepto5  From " & DonamNomTaulaNorma43 & "_Tmp O "
                    Sql = Sql & " Where "
                    Sql = Sql & "      O.dataOperacio > (Select isnull(Max(dataOperacio),'000000') from Norma43 Where [Comu_numcuenta] = '" & numcuenta & "' ) "
                    Sql = Sql & "   or O.dataOperacio < (Select isnull(Min(dataOperacio),'999999') from Norma43 Where [Comu_numcuenta] = '" & numcuenta & "' ) "
                    Sql = Sql & "   or O.IdNorma43 in ("
                    Sql = Sql & "              Select T.IdNorma43 From  " & DonamNomTaulaNorma43 & "_Tmp T left Join " & DonamNomTaulaNorma43 & " N on t.[Comu_Empresa] = N.[Comu_Empresa] And  t.[Comu_Banco] = N.[Comu_Banco] And t.[Comu_Oficina] = N.[Comu_Oficina] And t.[Comu_numcuenta] = N.[Comu_numcuenta] And t.[Comu_NombreCliente] = N.[Comu_NombreCliente] And t.[Comu_CodigoCliente] = N.[Comu_CodigoCliente] And t.tipusRegistreA = N.tipusRegistreA And t.LliureA = N.LliureA And t.Oficina = N.Oficina And t.DataOperacio = N.DataOperacio And t.DataValor = N.DataValor And t.ConceptoComun = N.ConceptoComun And t.ConceptoPropio = N.ConceptoPropio And t.DeveHaver = N.DeveHaver And t.Importe = N.Importe And t.Documento = N.Documento And t.Referencia = N.Referencia And t.Referencia2 = N.Referencia2 And t.tipusRegistreB = N.tipusRegistreB And t.CodigoDato = N.CodigoDato And t.Concepto1 = N.Concepto1 And t.Concepto2 = N.Concepto2 And t.Concepto3 = N.Concepto3 And t.Concepto4 = N.Concepto4 And t.Concepto5 = N.Concepto5 "
                    Sql = Sql & "              where N.IdNorma43  is null"
                    Sql = Sql & "                      ) "
                    ExecutaComandaSql Sql
                    
                    'CUADRAR AUTOMÁTICAMENTE LOS MOVIMIENTOS DE TARJETA DE CRÉDITO
                    Dim rsConceptos As rdoResultset, rsCodisComercio As rdoResultset, rsTienda As rdoResultset, rsSubc1 As rdoResultset, rsSubc2 As rdoResultset
                    Dim ArrCodisComercio() As String, nCc As Integer, concepto As String, c As Integer
                    Dim tienda As String, tiendaNom As String, SubC1 As String, SubC2 As String
                    Dim OK As Boolean, dataOperacio As Date
                    Dim nDigitos As Integer
                    nDigitos = 8
                    
                    nCc = 1
                    Set rsCodisComercio = Db.OpenResultset("select * from constantsclient where variable='CodigoComercio'")
                    While Not rsCodisComercio.EOF
                        ReDim Preserve ArrCodisComercio(nCc)
                        ArrCodisComercio(nCc) = rsCodisComercio("codi") & "|" & rsCodisComercio("valor")
                        nCc = nCc + 1
                        rsCodisComercio.MoveNext
                    Wend
                                        
                    Set rsConceptos = Db.OpenResultset("select * from norma43 where (concepto1 like '%ABONO TPV%' or concepto1 like 'BANKIA S.A.U.%' or concepto1 like '%COMISIONES%') and idFichero='" & Idrs & "'")
                    While Not rsConceptos.EOF
                        OK = False
                        concepto = rsConceptos("concepto1")
                        For c = 1 To UBound(ArrCodisComercio)
                            If InStr(concepto, Split(ArrCodisComercio(c), "|")(1)) Then
                                OK = True
                                Exit For
                            End If
                        Next
                        If OK Then
                            tienda = Split(ArrCodisComercio(c), "|")(0)
                            Set rsTienda = Db.OpenResultset("select * from clients where codi = " & tienda)
                            If Not rsTienda.EOF Then
                                tiendaNom = rsTienda("Nom")
                                SubC1 = ""
                                Set rsSubc1 = Db.OpenResultset("select * from constantsclient where variable='EmpresaVendesCC' and codi=" & tienda)
                                If Not rsSubc1.EOF Then
                                    SubC1 = "572" & Right("000000000000" & rsSubc1("valor"), nDigitos - 3)
                                    dataOperacio = CDate(Right(rsConceptos("DataOperacio"), 2) & "/" & Mid(rsConceptos("DataOperacio"), 3, 2) & "/" & "20" & Left(rsConceptos("DataOperacio"), 2))
                                    
                                    If InStr(concepto, "ABONO") Then
                                        Set rsSubc2 = Db.OpenResultset("select * from constantsclient where variable='CodiContable' and codi=" & tienda)
                                        If Not rsSubc2.EOF Then
                                            SubC2 = "43" & Right("000000000000" & rsSubc2("valor"), nDigitos - 2)
                                            
                                            CreaAsientoCtb rsConceptos("idNorma43"), "1", "C", "", dataOperacio, concepto, "", "", SubC1, concepto & " " & tiendaNom & " - NF 0 (" & Now() & ")", rsConceptos("importe") / 100, 0, ""
                                            CreaAsientoCtbBis rsConceptos("idNorma43"), "2", "B", "", dataOperacio, "N/F 0 BANCO", "", "", SubC2, tiendaNom & " - NF 0 (" & Now() & ")", 0, rsConceptos("importe") / 100, "", tienda
                                            ExportaMURANO_Bancs rsConceptos("idNorma43"), dataOperacio
                                        End If
                                    End If
                                    
                                    If InStr(concepto, "COMISIONES") Or InStr(concepto, "BANKIA S.A.U.") Then
                                        SubC2 = "62600000"
                                        
                                        CreaAsientoCtb rsConceptos("idNorma43"), "1", "C", "", dataOperacio, concepto, "", "", SubC1, concepto & " - COMISIONES TARJETAS", 0, rsConceptos("importe") / 100, ""
                                        CreaAsientoCtb rsConceptos("idNorma43"), "2", "B", "", dataOperacio, "COMISIONES TARJETAS", "", "", SubC2, "COMISIONES TARJETAS", rsConceptos("importe") / 100, 0, ""
                                        ExportaMURANO_Bancs rsConceptos("idNorma43"), dataOperacio
                                    End If
                                End If
                            End If
                        End If
                        
                        rsConceptos.MoveNext
                    Wend
                    
                Case "NominasHtmlVell"
'                    ExecutaComandaSql "Delete " & SousNominaImportats & " where IdFichero = '" & Idrs & "' "
'                    Sql = ""
'                    Sql = Sql & " delete s  from Archivolines  a join SousNominaImportats  s on "
'                    Sql = Sql & "s.Data = '20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  and "
'                    Sql = Sql & "s.Treb = SUBSTRING (linea ,61,5) and "
'                    Sql = Sql & "s.i_liquid = Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float) "
'                    Sql = Sql & "where id = '" & Idrs & "' and "
'                    Sql = Sql & "SUBSTRING (linea ,46,15) = 'TRANSFERENCIAS '"
'                    ExecutaComandaSql Sql
'
'                    Sql = "Insert Into " & SousNominaImportats & " "
'                    Sql = Sql & "Select "
'                    Sql = Sql & " '" & Idrs & "','Ficher',CodiEmpresa,data,treb,sum(i_liquid)i_liquid,sum(I_Brut) - sum(I_Liquid) -sum(I_Irpf) I_Irpf,sum(I_Irpf) + sum(I_Tc1) I_Tc1,sum(I_Brut) I_Brut,sum(i_SsEmp) i_SsEmp , sum(i_SsTre) i_SsTre  "
'                    Sql = Sql & "From ( "
'                    Sql = Sql & "      select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,65,5) Treb,0.0 I_Liquid,0.0 I_Irpf,0.0 I_Tc1,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float)  I_Brut,0.0 I_SsEmp,0.0 I_SsTre from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,20) = 'SUELDOS Y SALARIOS  ' "
'                    Sql = Sql & "union select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,62,5) Treb,0.0 I_Liquid,0.0 I_Irpf,0.0 I_Tc1,0.0 I_Brut,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float)  I_SsEmp,0.0 I_SsTre from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,17) = 'SEG SOCIAL EMPL. '  "
'                    Sql = Sql & "union select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,61,5) Treb,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float) I_Liquid,0.0 I_Irpf,0.0 I_Tc1,0.0 I_Brut,0.0 I_SsEmp,0.0 I_SsTre  from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,15) = 'TRANSFERENCIAS ' "
'                    Sql = Sql & "union select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,62,5) Treb,0.0 I_Liquid,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float) I_Irpf,0.0 I_Tc1,0.0 I_Brut,0.0 I_SsEmp,0.0 I_SsTre  from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,17) = 'SEG SOCIAL EMPL. ' "
'                    Sql = Sql & "union select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,66,5) Treb,0.0 I_Liquid,0.0 I_Irpf,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float)  I_Tc1,0.0 I_Brut,0.0 I_SsEmp,0.0 I_SsTre from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,17) = 'SEG. SOCIAL ACREE' "
'                    Sql = Sql & "union select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,62,5) Treb,0.0 I_Liquid,0.0 I_Irpf,0 I_Tc1,0.0 I_Brut,0.0 I_SsEmp,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float) I_SsTre    from Archivolines where id = '{B2CA376E-FAB8-44AD-B906-1E0AA05EDB62}' and SUBSTRING (linea ,46,17) = 'COSTE DE EMPRESA'  "
'                    Sql = Sql & ") a group by data,treb,CodiEmpresa"
'                    ExecutaComandaSql Sql  ' Insertem a temporals
                Case "SousLc"
                    ExecutaComandaSql "Delete " & SousNominaImportats & " where IdFichero = '" & Idrs & "' "
                    Sql = ""
                    Sql = Sql & " delete s  from Archivolines  a join SousNominaImportats  s on "
                    Sql = Sql & "s.Data = '20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  and "
                    Sql = Sql & "s.Treb = SUBSTRING (linea ,61,5) and "
                    Sql = Sql & "s.i_liquid = Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float) "
                    Sql = Sql & "where id = '" & Idrs & "' and "
                    Sql = Sql & "SUBSTRING (linea ,46,15) = 'TRANSFERENCIAS '"
                    ExecutaComandaSql Sql
                    
                    Sql = "Insert Into " & SousNominaImportats & " ([IdFichero],[Origen],[CodiEmpresa],[data],[treb],[i_liquid],[I_Irpf],[I_Tc1],[I_Brut],[i_SsEmp],[RetibEspecie],[IrpfEspecies]) "
                    Sql = Sql & "Select "
                    Sql = Sql & " '" & Idrs & "','Ficher',CodiEmpresa,data,treb,sum(i_liquid)i_liquid,sum(I_Brut) - sum(I_Liquid) -sum(I_Irpf) I_Irpf,sum(I_Irpf) + sum(I_Tc1) I_Tc1,sum(I_Brut) I_Brut,sum(i_SsEmp) i_SsEmp ,0, 0  "
                    Sql = Sql & "From ( "
                    Sql = Sql & "      select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,65,4) Treb,0.0 I_Liquid,0.0 I_Irpf,0.0 I_Tc1,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float)  I_Brut,0.0 I_SsEmp,0.0 I_SsTre from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,18) = 'SUELDOS Y SALARIOS' "
                    Sql = Sql & "union select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,63,4) Treb,0.0 I_Liquid,0.0 I_Irpf,0.0 I_Tc1,0.0 I_Brut,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float)  I_SsEmp,0.0 I_SsTre from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,14) = 'SEG.SOCIAL. EM' "
                    Sql = Sql & "union select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,61,4) Treb,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float) I_Liquid,0.0 I_Irpf,0.0 I_Tc1,0.0 I_Brut,0.0 I_SsEmp,0.0 I_SsTre  from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,12) = 'TRANSFERENCI' "
                    Sql = Sql & "union select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,63,4) Treb,0.0 I_Liquid,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float) I_Irpf,0.0 I_Tc1,0.0 I_Brut,0.0 I_SsEmp,0.0 I_SsTre  from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,14) =  'SEG.SOCIAL. EM'  "
                    Sql = Sql & "union select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,65,4) Treb,0.0 I_Liquid,0.0 I_Irpf,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float)  I_Tc1,0.0 I_Brut,0.0 I_SsEmp,0.0 I_SsTre from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,15) =  'SEG.SOCIAL ACRE' "
                    Sql = Sql & "union select 0 as CodiEmpresa,'20' + SUBSTRING (linea ,75,2) + SUBSTRING (linea ,73,2) + SUBSTRING (linea ,71,2)  Data,SUBSTRING (linea ,63,4) Treb,0.0 I_Liquid,0.0 I_Irpf,0 I_Tc1,0.0 I_Brut,0.0 I_SsEmp,Cast(STUFF(SUBSTRING (linea ,81,17), 15, 1, '.')  as float) I_SsTre    from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,46,15) = 'COSTE DE EMPRES'   "
                    Sql = Sql & ") a group by data,treb,CodiEmpresa"

                    ExecutaComandaSql Sql  ' Insertem a temporals
                Case "SousLc2"
                    Dim fNomina  As String
                    Dim lin As String
                    Dim trebLc2 As String
                    
                    ExecutaComandaSql "Delete " & SousNominaImportatsTmp()
                    
                    Set Rs4 = Db.OpenResultset("select Id,nombre,   fecha, isnull(linea,'') as linea,Numlinea from ArchivoLines where  id = '" & Idrs & "' order by numLinea")
                    While Not Rs4.EOF
                            lin = Rs4("linea")
                        If Mid(lin, 1, 4) = "GNOM" Then
                            fNomina = Mid(lin, 28, 8)
                        ElseIf Mid(lin, 1, 1) = "D" Then
                            trebLc2 = Mid(lin, 14, 4)
                            If Mid(lin, 9, 5) = "64000" Then 'SUELDOS Y SALARIOS [I_Brut]
                                Sql = "Insert Into " & SousNominaImportatsTmp & " ([IdFichero],[Origen],[CodiEmpresa],[data],[treb],[i_liquid],[I_Irpf],[I_Tc1],[I_Brut],[i_SsEmp],[i_SsTre],[RetibEspecie],[IrpfEspecies]) "
                                Sql = Sql & "values ('" & Idrs & "','Ficher',0, '" & fNomina & "', '" & trebLc2 & "', 0.0 ,0.0 ,0.0," & Mid(lin, 61, 20) & ", 0.0, 0.0, 0.0, 0.0 )"
                                ExecutaComandaSql Sql
                            ElseIf Mid(lin, 9, 5) = "47600" Then 'SEG.SOCIAL. EMPL + Seg.SOCIAL ACRE [I_Tc1]
                                Sql = "Insert Into " & SousNominaImportatsTmp & " ([IdFichero],[Origen],[CodiEmpresa],[data],[treb],[i_liquid],[I_Irpf],[I_Tc1],[I_Brut],[i_SsEmp],[i_SsTre],[RetibEspecie],[IrpfEspecies]) "
                                Sql = Sql & "values ('" & Idrs & "','Ficher',0, '" & fNomina & "', '" & trebLc2 & "', 0.0 ,0.0, " & Mid(lin, 61, 20) & ", 0.0, 0.0 ,0.0 , 0.0, 0.0 )"
                                ExecutaComandaSql Sql
                            ElseIf Mid(lin, 9, 5) = "46500" Then 'TRANSFERENCIES [i_liquid]
                                Sql = "Insert Into " & SousNominaImportatsTmp & " ([IdFichero],[Origen],[CodiEmpresa],[data],[treb],[i_liquid],[I_Irpf],[I_Tc1],[I_Brut],[i_SsEmp],[i_SsTre],[RetibEspecie],[IrpfEspecies]) "
                                Sql = Sql & "values ('" & Idrs & "','Ficher',0, '" & fNomina & "', '" & trebLc2 & "', " & Mid(lin, 61, 20) & " ,0.0 ,0.0, 0.0, 0.0, 0.0, 0.0, 0.0 )"
                                ExecutaComandaSql Sql
                            ElseIf Mid(lin, 9, 5) = "64200" Then 'COSTE DE EMPRES
                                Sql = "Insert Into " & SousNominaImportatsTmp & " ([IdFichero],[Origen],[CodiEmpresa],[data],[treb],[i_liquid],[I_Irpf],[I_Tc1],[I_Brut],[i_SsEmp],[i_SsTre], [RetibEspecie],[IrpfEspecies]) "
                                Sql = Sql & "values ('" & Idrs & "','Ficher',0, '" & fNomina & "', '" & trebLc2 & "', 0.0 , 0.0, 0.0 , 0.0, 0.0, " & Mid(lin, 61, 20) & ", 0.0, 0.0 )"
                                ExecutaComandaSql Sql
                            End If
                        End If
                        
                        Rs4.MoveNext
                    Wend
                    
                    Sql = ""
                    Sql = Sql & "delete s  from " & SousNominaImportatsTmp & "  a join " & SousNominaImportats & "  s on "
                    Sql = Sql & "s.Data = a.Data  and s.Treb = a.Treb and s.i_liquid = a.i_liquid "
                    Sql = Sql & "where a.idFichero = '" & Idrs & "' and a.i_liquid<>0 "
                    ExecutaComandaSql Sql

                    Sql = ""
                    Sql = Sql & "insert into " & SousNominaImportats & " (IdFichero, Origen, codiEmpresa, data, treb, i_liquid, i_irpf, i_tc1, i_brut, i_ssEmp, i_ssTre, retibEspecie, IrpfEspecies) "
                    Sql = Sql & "select idFichero, 'Ficher' origen, codiempresa, data, treb, "
                    Sql = Sql & "sum(i_liquid) i_liquid, sum(I_brut)-sum(I_liquid)-(sum(i_tc1)-sum(i_ssTre)) I_Irpf, "
                    Sql = Sql & "sum(I_Tc1) i_Tc1, sum(I_brut) i_brut, "
                    Sql = Sql & "sum(i_tc1)-sum(i_ssTre) i_ssemp, sum(i_ssTre) i_ssTre, sum(retibEspecie) retibEspecie, "
                    Sql = Sql & "sum(irpfespecies) irpfEspecies "
                    Sql = Sql & "From " & SousNominaImportatsTmp & " "
                    Sql = Sql & "where idFichero='" & Idrs & "' "
                    Sql = Sql & "group by idFichero, data, treb, codiempresa"
                    ExecutaComandaSql Sql

                Case "SousA3"
                    ExecutaComandaSql "Delete " & SousNominaImportats & " where IdFichero = '" & Idrs & "' "
                    Sql = ""
                    Sql = Sql & "delete s  from Archivolines  a join SousNominaImportats  s on s.codiempresa = SUBSTRING (linea ,2,5)  and s.Data = SUBSTRING (linea ,7,8) and s.Treb = SUBSTRING (linea ,59,10)  and s.i_liquid = Cast(SUBSTRING (linea ,100,14) as float) "
                    Sql = Sql & " where id = '" & Idrs & "' "
                    Sql = Sql & "and SUBSTRING (linea ,1,1 ) = '3' and SUBSTRING (linea ,254,1) = 'N' and left(SUBSTRING (linea ,16,12),4) = '4650' "
                    ExecutaComandaSql Sql
                    Sql = "Insert Into " & SousNominaImportats & " "
                    Sql = Sql & "Select "
                    Sql = Sql & " '" & Idrs & "','Ficher',CodiEmpresa,data,treb,sum(i_liquid)i_liquid,sum(I_Irpf) I_Irpf,sum(I_Tc1) I_Tc1,sum(I_Brut) I_Brut,sum(i_SsEmp) i_SsEmp "
                    Sql = Sql & "From ( "
                    Sql = Sql & "      Select SUBSTRING (linea ,2,5) CodiEmpresa ,SUBSTRING (linea ,7,8) Data,SUBSTRING (linea ,59,10) Treb,Cast(SUBSTRING (linea ,100,14) as float) I_Liquid,0.0 I_Irpf,0.0 I_Tc1,0.0 I_Brut,0.0 I_SsEmp  from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,1,1 ) = '3' and SUBSTRING (linea ,254,1) = 'N' and left(SUBSTRING (linea ,16,12),4) = '4650' "
                    Sql = Sql & "union Select SUBSTRING (linea ,2,5) CodiEmpresa ,SUBSTRING (linea ,7,8) Data,SUBSTRING (linea ,59,10) Treb,0.0 I_Liquid, Cast(SUBSTRING (linea ,100,14) as float) I_Irpf,0.0 I_Tc1,0.0 I_Brut,0.0 I_SsEmp from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,1,1 ) = '3' and SUBSTRING (linea ,254,1) = 'N' and left(SUBSTRING (linea ,16,12),4) = '4751' "
                    Sql = Sql & "union Select SUBSTRING (linea ,2,5) CodiEmpresa ,SUBSTRING (linea ,7,8) Data,SUBSTRING (linea ,59,10) Treb,0.0 I_Liquid,0.0 I_Irpf,Cast(SUBSTRING (linea ,100,14) as float) I_Tc1,0.0 I_Brut,0.0 I_SsEmp  from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,1,1 ) = '3' and SUBSTRING (linea ,254,1) = 'N' and left(SUBSTRING (linea ,16,12),4) = '4760' "
                    Sql = Sql & "union Select SUBSTRING (linea ,2,5) CodiEmpresa ,SUBSTRING (linea ,7,8) Data,SUBSTRING (linea ,59,10) Treb,0.0 I_Liquid,0.0 I_Irpf,0.0 I_Tc1,Cast(SUBSTRING (linea ,100,14) as float) I_Brut,0.0 I_SsEmp  from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,1,1 ) = '3' and SUBSTRING (linea ,254,1) = 'N' and left(SUBSTRING (linea ,16,12),4) = '6400' "
                    Sql = Sql & "union Select SUBSTRING (linea ,2,5) CodiEmpresa ,SUBSTRING (linea ,7,8) Data,SUBSTRING (linea ,59,10) Treb,0.0 I_Liquid,0.0 I_Irpf,0.0 I_Tc1,0.0 I_Brut,Cast(SUBSTRING (linea ,100,14) as float) I_SsEmp  from Archivolines where id = '" & Idrs & "' and SUBSTRING (linea ,1,1 ) = '3' and SUBSTRING (linea ,254,1) = 'N' and left(SUBSTRING (linea ,16,12),4) = '6420' "
                    Sql = Sql & ") a group by data,treb,CodiEmpresa"
                    ExecutaComandaSql Sql  ' Insertem a temporals
                Case "SousSpVELL"
                    ExecutaComandaSql "Delete " & SousNominaImportats & " where IdFichero = '" & Idrs & "' "
                    FechaNomina = "01-01-01"
                    Set Rs2 = Db.OpenResultset("Select NumLinea,cast(linea as nvarchar (100)) linea from Archivolines  where id = '" & Idrs & "' and linea like '%ConPeriodo:%' ")
                    If Not Rs2.EOF Then If Not IsNull(Rs2("Linea")) Then FechaNomina = DateSerial(Right(Left(Rs2("Linea"), InStr(Rs2("Linea"), "ConPeriodo:") + 15), 4), Right(Left(Rs2("Linea"), InStr(Rs2("Linea"), "ConPeriodo:") + 18), 2), 28)
                    If FechaNomina = "01-01-01" Then
                        Set Rs2 = Db.OpenResultset("Select NumLinea, left(linea,400) Linea from Archivolines  where id = '" & Idrs & "' and linea like '%Fecha%' ")
                        If Not Rs2.EOF Then
                            Line = Rs2("Linea")
                            If Not IsNull(Rs2("Linea")) Then
                            
                                FechaNomina = Trim(Right(Rs2("Linea"), Len(Rs2("Linea")) - InStr(Rs2("Linea"), "Fecha") - 5))
                            End If
                        End If
                    End If
                    If FechaNomina = "01-01-01" Then FechaNomina = Now
                        
                        
                    FechaNominaStr = Format(FechaNomina, "yyyymmdd")

                    Sql = ""
                    Sql = Sql & "Delete s  From Archivolines a join SousNominaImportats  s on s.codiempresa = SUBSTRING (linea ,2,5)  and s.Data = SUBSTRING (linea ,7,8) and s.Treb = SUBSTRING (linea ,59,10)  and s.i_liquid = Cast(SUBSTRING (linea ,100,14) as float) "
                    Sql = Sql & "Where id = '" & Idrs & "' "
                    Sql = Sql & "and SUBSTRING (linea ,1,1 ) = '3' and SUBSTRING (linea ,254,1) = 'N' and left(SUBSTRING (linea ,16,12),4) = '4650' "
                    ExecutaComandaSql Sql
                    
                    Sql = "Select numlinea,"
                    Sql = Sql & "SUBSTRING (linea ,97,10) IrpfEspecies, "
                    Sql = Sql & "SUBSTRING (linea ,0,43) treb, "
                    Sql = Sql & "SUBSTRING (linea ,141,14) i_liquid, "
                    Sql = Sql & "SUBSTRING (linea ,86,14) I_Irpf, "
                    Sql = Sql & "SUBSTRING (linea ,59,14) SSTrabaj, "
                    Sql = Sql & "SUBSTRING (linea ,127,14) TotalBr, "
                    Sql = Sql & "SUBSTRING (linea ,50,14) SSempre "
                    Sql = Sql & "from Archivolines where id = '" & Idrs & "'  "
                    Sql = Sql & "  And SUBSTRING (linea ,1,2 ) = ' 0' and SUBSTRING (linea ,7,1) = '  '  "
                    Sql = Sql & "Order by cast(numlinea as numeric)"
                    empresaborrada = ""
                    Set Rs2 = Db.OpenResultset(Sql)
                    data = Now
                    While Not Rs2.EOF
                        NumLinea = Rs2("NumLinea")
                        Treb = Rs2("Treb")
                        Treb = Right(Treb, Len(Treb) - 1)
                        Treb = Left(Treb, 5) & " " & Right(Treb, Len(Treb) - 5)
                        
                        If InStr(Treb, "guaman") > 0 Then
                            Debug.Print Treb
                        End If
                        i_liquid = EsNumero(Rs2("TotalBr")) ' EsNumero(Rs2("i_liquid"))
                        i_Irpf = 0 'EsNumero(Rs2("I_Irpf"))
                        SSTrabaj = 0 'EsNumero(Rs2("SSTrabaj"))
                        TotalBr = 0 'EsNumero(Rs2("TotalBr"))
                        SSempre = 0 'EsNumero(Rs2("SSempre"))
                        IrpEsp = 0 'EsNumero(Rs2("IrpfEspecies"))
                        If IsNumeric(TotalBr) And IsNumeric(i_liquid) And IsNumeric(SSTrabaj) And IsNumeric(i_Irpf) And IsNumeric(IrpEsp) Then
                            PEsp = 0 'Round(TotalBr - i_liquid - SSTrabaj - I_Irpf - IrpEsp, 2)
                            SSTrabaj = 0 'CDbl(SSTrabaj) + CDbl(SSempre)
                            SSTrabaj2 = 0 ' EsNumero(Rs2("SSTrabaj"))
                            Set Rs3 = Db.OpenResultset("Select NumLinea,cast(linea as nvarchar (100)) linea from Archivolines  where id = '" & Idrs & "' and (SUBSTRING (linea ,1,1) = '0' or SUBSTRING (linea ,1,1) = '1') and SUBSTRING (linea ,4,1) = ' ' and NumLinea < " & NumLinea & " Order by  cast(numlinea as numeric) desc  ")
                            If Not Rs3.EOF Then If Not IsNull(Rs3("Linea")) Then CodiEmpresa = Rs3("Linea")
                                If Not empresaborrada = CodiEmpresa Then
                                    empresaborrada = CodiEmpresa
                                    ExecutaComandaSql "Delete " & SousNominaImportats & "  Where CodiEmpresa= '" & CodiEmpresa & "' And data = '" & FechaNominaStr & "'"
                                End If
                                ExecutaComandaSql "Insert Into " & SousNominaImportats & " ([IdFichero],[Origen],[CodiEmpresa],[data],[treb],[i_liquid],[I_Irpf],[I_Tc1],[I_Brut],[i_SsEmp],[RetibEspecie],[IrpfEspecies])  Values ('" & Idrs & "','Ficher','" & CodiEmpresa & "','" & FechaNominaStr & "','" & Treb & "'," & i_liquid & "," & i_Irpf & "," & SSTrabaj & "," & TotalBr & "," & SSTrabaj2 & "," & PEsp & "," & IrpEsp & ")"
                         End If
                         Rs2.MoveNext
                    Wend
                    Rs2.Close
                Case "SousSp"
                    ExecutaComandaSql "Delete " & SousNominaImportats & " where IdFichero = '" & Idrs & "' "
                    FechaNomina = "01-01-01"
                    If IsNumeric(Left(Nnomfile, 5)) Then FechaNomina = "28" & "-" & Mid(Nnomfile, 5, 2) & "-" & Mid(Nnomfile, 3, 2)
                    
'                    Set Rs2 = Db.OpenResultset("Select NumLinea,cast(linea as nvarchar (100)) linea from Archivolines  where id = '" & Idrs & "' and linea like '%ConPeriodo:%' ")
'                    If Not Rs2.EOF Then If Not IsNull(Rs2("Linea")) Then FechaNomina = DateSerial(Right(Left(Rs2("Linea"), InStr(Rs2("Linea"), "ConPeriodo:") + 15), 4), Right(Left(Rs2("Linea"), InStr(Rs2("Linea"), "ConPeriodo:") + 18), 2), 28)
                    If FechaNomina = "01-01-01" Then
                        Set Rs2 = Db.OpenResultset("Select NumLinea, left(linea,400) Linea from Archivolines  where id = '" & Idrs & "' and linea like '%Fecha%' ")
                        If Not Rs2.EOF Then
                            Line = Rs2("Linea")
                            If Not IsNull(Rs2("Linea")) Then
                                FechaNomina = Trim(Right(Rs2("Linea"), Len(Rs2("Linea")) - InStr(Rs2("Linea"), "Fecha") - 5))
                            End If
                        End If
                    End If
                    If FechaNomina = "01-01-01" Then FechaNomina = Now
                        
                        
                    FechaNominaStr = Format(FechaNomina, "yyyymmdd")

                    Sql = ""
                    Sql = Sql & "Delete s  From Archivolines a join SousNominaImportats  s on s.codiempresa = SUBSTRING (linea ,2,5)  and s.Data = SUBSTRING (linea ,7,8) and s.Treb = SUBSTRING (linea ,59,10)  and s.i_liquid = Cast(SUBSTRING (linea ,100,14) as float) "
                    Sql = Sql & "Where id = '" & Idrs & "' "
                    Sql = Sql & "and SUBSTRING (linea ,1,1 ) = '3' and SUBSTRING (linea ,254,1) = 'N' and left(SUBSTRING (linea ,16,12),4) = '4650' "
                    ExecutaComandaSql Sql
                    
                    Patro = "C  digo Trabajador                      SS.SS.Trabaj  SS.EMPRESA    TOTAL S.S.      Base irpf       IRPF         Total bruto  Tot. l  quido    COSTE TOTAL               N   de Seguridad IRPF Especi Base irpf e C  d Nombre de empresa"
'                    Set Rs2 = Db.OpenResultset("Select Linea From Archivolines  where id = '" & Idrs & "' and linea like '%Ss.Ss.Trabaj%' ")
'                    If Not Rs2.EOF Then Patro = Rs2("Linea")
                    Sql = "Select numlinea,"
                    Sql = Sql & "SUBSTRING (linea ," & InStr(Patro, "IRPF Especi") & ",12) IrpfEspecies, "
                    Sql = Sql & "SUBSTRING (linea ," & InStr(Patro, "C  digo") & ",37) treb, "
                    Sql = Sql & "SUBSTRING (linea ," & InStr(Patro, "Tot. l  quido") & ",12) i_liquid, "
                    Sql = Sql & "SUBSTRING (linea ," & InStr(Patro, "    IRPF") & ",12) I_Irpf, "
                    Sql = Sql & "SUBSTRING (linea ," & InStr(Patro, "SS.SS.Trabaj") & ",12) SSTrabaj, "
                    Sql = Sql & "SUBSTRING (linea ," & InStr(Patro, "Total bruto") & ",12) TotalBr, "
                    Sql = Sql & "SUBSTRING (linea ," & InStr(Patro, "e C  d Nombre de empresa") - 1 & ",34) CodiEmpresa, "
                    Sql = Sql & "SUBSTRING (linea ," & InStr(Patro, "SS.EMPRESA") & ",12) SSempre "
                    Sql = Sql & "from Archivolines where id = '" & Idrs & "'  "
                    Sql = Sql & "  And SUBSTRING (linea ,1,2 ) = ' 0' and SUBSTRING (linea ,7,1) = '  '  "
                    Sql = Sql & "Order by CodiEmpresa, treb,cast(numlinea as numeric)"
                    empresaborrada = ""
                    Set Rs2 = Db.OpenResultset(Sql)
                    data = Now
                    
                    Dim esborraFitxer As Boolean
                    esborraFitxer = True
                    
                    While Not Rs2.EOF
                    
                        NumLinea = Rs2("NumLinea")
                        Treb = Rs2("Treb")
                        Treb = Right(Treb, Len(Treb) - 1)
                        Treb = Left(Treb, 5) & " " & Right(Treb, Len(Treb) - 5)
                        
                        If InStr(UCase(Treb), "TERAN") > 0 Then
                            Debug.Print Treb
                        End If
                        i_liquid = EsNumero(Rs2("i_liquid"))
                        i_Irpf = EsNumero(Rs2("I_Irpf"))
                        SSTrabaj = EsNumero(Rs2("SSTrabaj"))
                        TotalBr = EsNumero(Rs2("TotalBr"))
                        SSempre = EsNumero(Rs2("SSempre"))
                        IrpEsp = EsNumero(Rs2("IrpfEspecies"))
                        CodiEmpresa = Rs2("CodiEmpresa")
                        If esborraFitxer Then
                            ExecutaComandaSql "Delete " & SousNominaImportats & "  Where CodiEmpresa= '" & CodiEmpresa & "' And data = '" & FechaNominaStr & "'"
                            esborraFitxer = False
                        End If
                        If Trim(IrpEsp) = "" Then IrpEsp = "0"
                        If IsNumeric(TotalBr) And IsNumeric(i_liquid) And IsNumeric(SSTrabaj) And IsNumeric(i_Irpf) And IsNumeric(IrpEsp) Then
                            PEsp = Round(TotalBr - i_liquid - SSTrabaj - i_Irpf - IrpEsp, 2)
                            SSTrabaj = CDbl(SSTrabaj) + CDbl(SSempre)
                            SSTrabaj2 = EsNumero(Rs2("SSTrabaj"))
                            Dim TrebEmpresa As String
                            If Not TrebEmpresa = CodiEmpresa Then
                                ExecutaComandaSql "Delete " & SousNominaImportats & "  Where treb = '" & Treb & "' And CodiEmpresa= '" & CodiEmpresa & "' And data = '" & FechaNominaStr & "'"
                            End If
                            TrebEmpresa = CodiEmpresa
                            ExecutaComandaSql "Insert Into " & SousNominaImportats & " ([IdFichero],[Origen],[CodiEmpresa],[data],[treb],[i_liquid],[I_Irpf],[I_Tc1],[I_Brut],[i_SsEmp],[RetibEspecie],[IrpfEspecies])  Values ('" & Idrs & "','Ficher','" & CodiEmpresa & "','" & FechaNominaStr & "','" & Treb & "'," & i_liquid & "," & i_Irpf & "," & SSTrabaj & "," & TotalBr & "," & SSTrabaj2 & "," & PEsp & "," & IrpEsp & ")"
                         Else
                         Debug.Print "ep"
                         End If
                         
                         Rs2.MoveNext
                    Wend
                    Rs2.Close
               End Select
error:
            rst.Close
            ExecutaComandaSql "Update Archivo Set Descripcion = 'Interpretat ' + cast(getdate() as nvarchar)  Where id  = '" & rs("Id") & "' "
            rs.MoveNext
        Wend
    End If
   
End Sub


Sub PillaNominasHtml(nomfile As String, Idrs As String)
    Dim f, K, aa, NumLinea, Treb, i_liquid, i_Irpf, SSTrabaj, TotalBr, SSempre, IrpEsp, PEsp, Rs3, CodiEmpresa, empresaborrada, FechaNominaStr, Rs2, Preparat, i_SsEmp, NomEmpresa
    
    f = FreeFile
    Open "c:\aa.txt" For Input As #f
    K = 0
    FechaNominaStr = ""
    CodiEmpresa = ""
    Preparat = False
    While Not EOF(f)
        aa = PillaNominasHtmlCar(f)
        K = K + 1
            Select Case aa
                Case "Cost total"
                    CodiEmpresa = PillaNominasHtmlCar(f)
                    ExecutaComandaSql "Delete " & SousNominaImportats & " Where [CodiEmpresa] = '" & CodiEmpresa & "'  And [data] = '" & FechaNominaStr & "' "
                    Preparat = True
                Case "Total Empresa"
                    Preparat = False
                Case "Costes de empresa"
                    aa = PillaNominasHtmlCar(f)
                    aa = PillaNominasHtmlCar(f)
                    aa = PillaNominasHtmlCar(f)
                    aa = PillaNominasHtmlCar(f)
                    FechaNominaStr = PillaNominasHtmlCar(f)
                    FechaNominaStr = "20" & Split(FechaNominaStr, "-")(2) & Split(FechaNominaStr, "-")(1) & Split(FechaNominaStr, "-")(0)
                Case Else
                    If Preparat And IsNumeric(aa) Then
                        Treb = aa & "  " & PillaNominasHtmlCar(f)
                        
                        aa = PillaNominasHtmlCar(f)
                        SSempre = Join(Split(PillaNominasHtmlCar(f), ","), ".")
                        SSTrabaj = Join(Split(PillaNominasHtmlCar(f), ","), ".")
                        aa = PillaNominasHtmlCar(f)
                        i_Irpf = Join(Split(PillaNominasHtmlCar(f), ","), ".")
                        IrpEsp = Join(Split(PillaNominasHtmlCar(f), ","), ".")
                        aa = PillaNominasHtmlCar(f)
                        TotalBr = Join(Split(PillaNominasHtmlCar(f), ","), ".")
                        i_liquid = Join(Split(PillaNominasHtmlCar(f), ","), ".")
                        aa = PillaNominasHtmlCar(f)
                        PEsp = Val(TotalBr) - Val(SSTrabaj) - Val(i_Irpf) - Val(IrpEsp) - Val(i_liquid)
'if PEsp > 100 Then
'PEsp = PEsp
'End If
                        ExecutaComandaSql "Insert Into " & SousNominaImportats & " ([IdFichero],[Origen],[CodiEmpresa],[data],[treb],[i_liquid],[I_Irpf],[I_Tc1],[I_Brut],[i_SsEmp],[RetibEspecie],[IrpfEspecies])  Values ('" & Idrs & "','Ficher','" & CodiEmpresa & "','" & FechaNominaStr & "','" & Treb & "','" & i_liquid & "','" & i_Irpf & "','" & Val(SSempre) + Val(SSTrabaj) & "','" & TotalBr & "','" & SSTrabaj & "','" & PEsp & "','" & IrpEsp & "')"
                    End If
            End Select
       DoEvents
    Wend
    Close #f
                

End Sub

Sub ImportaNominasSILEMA(nomfile As String, FechaNomina As Date, Idrs As String)
    Dim f As Integer, n As Integer
    Dim linesFile As String, FechaNominaStr As String, CodiEmpresa As String, Treb As String, i_liquid As String
    Dim i_Irpf As String, i_Tc1 As String, i_Brut As String, i_SsEmp As String, retibEspecie As String, IrpfEspecies As String
    Dim nominasArr() As String, nominasDetall() As String
    Dim esborraFitxer As Boolean
    
    esborraFitxer = True
    
    f = FreeFile
    Open "c:\aa.txt" For Input As #f
    
    ExecutaComandaSql "Delete " & SousNominaImportats & " where IdFichero = '" & Idrs & "' "
                    
    If FechaNomina = "01-01-01" Then FechaNomina = Now
    FechaNominaStr = Format(FechaNomina, "yyyymmdd")

    Line Input #f, linesFile
    nominasArr = Split(linesFile, "|")
    
    For n = 0 To UBound(nominasArr) - 1
        nominasDetall = Split(nominasArr(n), "#")
        'Código Empresa (0) # Nombre Empresa (1) # Código Trabajador (2) # Nombre Trabajador (3) # 2 Primeros SS (4) # 8 SS (5) # 2 Ultims ss (6) # SS Trab 1 (7) # ss Trab 2 (8) #
        'SS.Empresa (9) # Total SS (10) # Base IRPF (11) # Retencion IRPF (12) # Total bruto (13) # Total liquido (14) # Coste empresa (15) # Importe especie (16) # Retencion especie (17) #
        'Embargo (18) # Anticipo (19) # Total bruto (20) (Sumar al 13) # Total Líquido (21) (Sumar al 14) # retencion IRPF (22) (Sumar al 12) #
        'Seguridad social trabajador contingencias comunes (23) (Sumar al 7) # Seguridad social trabajador accidentes (24) (sumar al 8) # Coste empresa (25) (Sumar al 15) #
        's.s. empresa (26) (sumar al 9) # Total seguridad social (27) (sumar al 10)
        If UBound(nominasDetall) >= 28 Then
            CodiEmpresa = nominasDetall(0) & " " & nominasDetall(1)
            Treb = nominasDetall(2) & " " & nominasDetall(3)
            Treb = Replace(Treb, "'", "''")
            i_liquid = CDbl(EsNumero(nominasDetall(14))) + CDbl(EsNumero(nominasDetall(21)))
            i_Irpf = CDbl(EsNumero(nominasDetall(12))) + CDbl(EsNumero(nominasDetall(22)))
            i_Tc1 = CDbl(EsNumero(nominasDetall(10))) + CDbl(EsNumero(nominasDetall(27)))
            i_Brut = CDbl(EsNumero(nominasDetall(13))) + CDbl(EsNumero(nominasDetall(20)))
            i_SsEmp = CDbl(EsNumero(nominasDetall(7))) + CDbl(EsNumero(nominasDetall(8))) + CDbl(EsNumero(nominasDetall(23))) + CDbl(EsNumero(nominasDetall(24)))
            retibEspecie = CDbl(EsNumero(nominasDetall(16)))
            IrpfEspecies = CDbl(EsNumero(nominasDetall(17)))
        
            If esborraFitxer Then
                ExecutaComandaSql "Delete " & SousNominaImportats & "  Where CodiEmpresa= '" & CodiEmpresa & "' And data = '" & FechaNominaStr & "'"
                esborraFitxer = False
            End If
        
            ExecutaComandaSql "Insert Into " & SousNominaImportats & " ([IdFichero],[Origen],[CodiEmpresa],[data],[treb],[i_liquid],[I_Irpf],[I_Tc1],[I_Brut],[i_SsEmp],[RetibEspecie],[IrpfEspecies]) Values ('" & Idrs & "','Ficher','" & CodiEmpresa & "','" & FechaNominaStr & "','" & Treb & "'," & i_liquid & "," & i_Irpf & "," & i_Tc1 & "," & i_Brut & "," & i_SsEmp & "," & retibEspecie & "," & IrpfEspecies & ")"
        Else
            Debug.Print "ep"
        End If
    Next
    
    Close #f
End Sub


 Private Sub Command1_Click()
'     MousePointer = vbHourglass
'     Dim cn As rdoConnection
'     Dim rs As rdoResultset, TempRs As rdoResultset
'     Dim cnstr As String, sqlstr As String
'     cnstr = "Driver={SQLServer};Server=myserver;Database=pubs;Uid=sa;Pwd="
'     sqlstr = "Select int1, char1, text1, image1 from chunktable"
'
'     rdoEnvironments(0).CursorDriver = rdUseServer
'     Set cn = rdoEnvironments(0).OpenConnection( _
'       "", rdDriverNoPrompt, False, cnstr)
'     On Error Resume Next
'     If cn.rdoTables("chunktable").Updatable Then
'       'table exists
'     End If
'     If Err > 0 Then
'       On Error GoTo 0
'       Debug.Print "Creating new table..."
'       cn.Execute "Create table chunktable(int1 int identity, " & _
'                  "char1 char(30), text1 text, image1 image)"
'       cn.Execute "create unique index int1index on chunktable(int1)"
'     End If
'     On Error GoTo 0
'     Set rs = cn.OpenResultset(Name:=sqlstr, _
'       Type:=rdOpenDynamic, _
'       LockType:=rdConcurRowVer)
'     If rs.EOF Then
'       rs.AddNew
'       rs("char1") = Now
'       rs.Update
'       rs.Requery
'     End If
'     Dim currec As Integer
'     currec = rs("int1")
'     rs.Edit
'     FileToColumn rs.rdoColumns("text1"), App.Path & "\README.TXT", 102400
'     FileToColumn rs.rdoColumns("image1"), App.Path & "\SETUP.BMP", 102400
'     rs("char1") = Now  'need to update at least one non-BLOB column
'     rs.Update
'
'     'this code gets the columnsize of each column
'     Dim text1_len As Long, image1_len As Long
'     If rs("text1").ColumnSize = -1 Then
'       'the function Datalength is SQL Server specific
'       'so you may have to change this for your database
'       sqlstr = "Select Datalength(text1) As text1_len, " & _
'                "Datalength(image1) As image1_len from chunktable " & _
'                "Where int1=" & currec
'       Set TempRs = cn.OpenResultset(Name:=sqlstr, _
'         Type:=rdOpenStatic, _
'         LockType:=rdConcurReadOnly)
'       text1_len = TempRs("text1_len")
'       image1_len = TempRs("image1_len")
'       TempRs.Close
'     Else
'       text1_len = rs("text1").ColumnSize
'       image1_len = rs("image1").ColumnSize
'     End If
'
'     ColumnToFile rs.rdoColumns("text1"), App.Path & "\text1.txt", _
'       102400, text1_len
'     ColumnToFile rs.rdoColumns("image1"), App.Path & "\image1.bmp", _
'       102400, image1_len
'     MousePointer = vbNormal
  End Sub

   Sub ColumnToFile(Col As rdoColumn, ByVal DiskFile As String, _
     BlockSize As Long, ColSize As Long)
     Dim NumBlocks As Integer
     Dim LeftOver As Long
     Dim byteData() As Byte   'Byte array for LongVarBinary
     Dim strData As String    'String for LongVarChar
     Dim DestFileNum As Integer, i As Integer

     ' Remove any existing destination file
     If Len(Dir$(DiskFile)) > 0 Then
       Kill DiskFile
     End If

     DestFileNum = FreeFile
     Open DiskFile For Binary As DestFileNum

     NumBlocks = ColSize \ BlockSize
     LeftOver = ColSize Mod BlockSize
     Select Case Col.Type
       Case rdTypeLONGVARBINARY
         If LeftOver > 0 Then
             byteData() = Col.GetChunk(LeftOver)
             Put DestFileNum, , byteData()
         End If
         For i = 1 To NumBlocks
           byteData() = Col.GetChunk(BlockSize)
           Put DestFileNum, , byteData()
         Next i
       Case rdTypeLONGVARCHAR
         For i = 1 To NumBlocks
           strData = String(BlockSize, 32)
           strData = Col.GetChunk(BlockSize)
           Put DestFileNum, , strData
         Next i
         strData = String(LeftOver, 32)
         strData = Col.GetChunk(LeftOver)
         Put DestFileNum, , strData
       Case Else
         MsgBox "Not a ChunkRequired column."
     End Select
     Close DestFileNum

   End Sub

   Sub FileToColumn(Col As rdoColumn, DiskFile As String, _
   BlockSize As Long)
     'moves a disk file to a ChunkRequired column in the table
     'A Byte array is used to avoid a UNICODE string
     Dim byteData() As Byte   'Byte array for LongVarBinary
     Dim strData As String    'String for LongVarChar
     Dim NumBlocks As Integer
     Dim filelength As Long
     Dim LeftOver As Long
     Dim SourceFile As Integer
     Dim i As Integer
     SourceFile = FreeFile
     Open DiskFile For Binary Access Read As SourceFile
     filelength = LOF(SourceFile) ' Get the length of the file
     If filelength = 0 Then
       Close SourceFile
       MsgBox DiskFile & " empty or not found."
     Else
       ' Calculate number of blocks to read and left over bytes
       NumBlocks = filelength \ BlockSize
       LeftOver = filelength Mod BlockSize
       Col.AppendChunk Null

       Select Case Col.Type
         Case rdTypeLONGVARCHAR
           ' Read the 'left over' amount of LONGVARCHAR data
           strData = String(LeftOver, " ")
           Get SourceFile, , strData
           Col.AppendChunk strData
           strData = String(BlockSize, " ")
           For i = 1 To NumBlocks
             Get SourceFile, , strData
             Col.AppendChunk strData
           Next i
           Close SourceFile
         Case rdTypeLONGVARBINARY
           ' Read the left over amount of LONGVARBINARY data
           ReDim byteData(0, LeftOver)
           Get SourceFile, , byteData()
           Col.AppendChunk byteData()
           ReDim byteData(0, BlockSize)
           For i = 1 To NumBlocks
             Get SourceFile, , byteData()
             Col.AppendChunk byteData()
           Next i
           Close SourceFile
         Case Else
           MsgBox "not a chunkrequired column."
       End Select
     End If

   End Sub

Sub ImportarFamilies(ByVal p1 As String, ByVal P2 As String, ByVal P3 As String)
   Dim rs As rdoResultset
   
   ExecutaComandaSql "CREATE TABLE [Importat_Params] ([Clau]    [nvarchar] (255) NULL ,[Valor_1] [nvarchar] (255) NULL ,[Valor_2] [nvarchar] (255) NULL ,[Valor_3] [nvarchar] (255) NULL ,[Valor_4] [nvarchar] (255) NULL ,[Valor_5] [nvarchar] (255) NULL ) ON [PRIMARY]"
   
   Set rs = Db.OpenResultset("select valor_1,Valor_2,Valor_3 from Importat_Params where Clau = 'Empresa'")
   While Not rs.EOF
      ImportarFamilia rs(0), rs(1), rs(2)
      rs.MoveNext
   Wend
   rs.Close
   
End Sub

Sub ImportarPromocions(ByVal p1 As String, ByVal P2 As String, ByVal P3 As String)
   Dim rs As rdoResultset

   ExecutaComandaSql "CREATE TABLE [Importat_Params] ([Clau]    [nvarchar] (255) NULL ,[Valor_1] [nvarchar] (255) NULL ,[Valor_2] [nvarchar] (255) NULL ,[Valor_3] [nvarchar] (255) NULL ,[Valor_4] [nvarchar] (255) NULL ,[Valor_5] [nvarchar] (255) NULL ) ON [PRIMARY]"
   
   Set rs = Db.OpenResultset("select valor_1,Valor_2,Valor_3 from Importat_Params where Clau = 'Empresa'")
   While Not rs.EOF
      ImportarPromocio rs(0), rs(1), rs(2)
      Missatges_CalEnviar "ProductesPromocionats", "", False, rs(0)
      rs.MoveNext
   Wend
   rs.Close
   
End Sub


Sub ImportarTarifa(ByVal p1 As String, ByVal P2 As String, ByVal P3 As String)
   Dim rs As rdoResultset
   
   ExecutaComandaSql "CREATE TABLE [Importat_Params] ([Clau]    [nvarchar] (255) NULL ,[Valor_1] [nvarchar] (255) NULL ,[Valor_2] [nvarchar] (255) NULL ,[Valor_3] [nvarchar] (255) NULL ,[Valor_4] [nvarchar] (255) NULL ,[Valor_5] [nvarchar] (255) NULL ) ON [PRIMARY]"
   
   Set rs = Db.OpenResultset("select valor_1,Valor_3,Valor_2 from Importat_Params where Clau = 'Empresa'")
   While Not rs.EOF
      ImportaTarifa rs(0), rs(1), rs(2)
      Missatges_CalEnviar "Articles", "", False, rs(0)
      rs.MoveNext
   Wend
   rs.Close
   
End Sub
Sub SincronitzaIntegracionesRebPas2(Path As String)
    Dim Files() As String, f As String, i As Integer, FS, Fss
    Dim Tot() As String, nomfile() As String, EsDirectori() As Boolean, data() As Date, Kb() As Double, OK() As Boolean, EsDir() As Boolean
    Dim d As Date, e As String
    
    Set FS = CreateObject("Scripting.FileSystemObject")
    
    f = Dir(Path & "\*.*")
    BuscaLastData
    ReDim Tot(0)
    ReDim Files(0)
    ReDim nomfile(0)
    ReDim EsDirectori(0)
    ReDim data(0)
    ReDim Kb(0)
    ReDim OK(0)
    ReDim EsDir(0)
    
    While Len(f)
       ReDim Preserve nomfile(UBound(nomfile) + 1)
       ReDim Preserve Tot(UBound(Tot) + 1)
       ReDim Preserve EsDirectori(UBound(EsDirectori) + 1)
       ReDim Preserve data(UBound(data) + 1)
       ReDim Preserve Kb(UBound(Kb) + 1)
       ReDim Preserve OK(UBound(OK) + 1)
       ReDim Preserve EsDir(UBound(EsDir) + 1)
       ReDim Preserve Files(UBound(Files) + 1)
       
       Set Fss = FS.GetFile(Path & "\" & f)
       nomfile(UBound(nomfile)) = f
       Tot(UBound(Tot)) = Path & "\" & f & " " & Fss.DateLastModified
       EsDirectori(UBound(EsDirectori)) = False
       data(UBound(data)) = Fss.DateLastModified
       Kb(UBound(Kb)) = Fss.Size
       OK(UBound(OK)) = True
       EsDir(UBound(EsDir)) = False
       Files(UBound(Files)) = f
       
       f = Dir
    Wend
    
    frmSplash.FiltreFiles Path, Tot, nomfile, EsDirectori, data, Kb, OK, EsDir
    
    For i = 1 To UBound(Files)
       If OK(i) Then
          Select Case tipusFile(Files(i))
             Case "unl":
             FtpExternImportaUnl Path & "\" & Files(i), "Importat_" & Files(i)
             Case "csv": FtpExternImportaCsv Path & "\" & Files(i), "Importat_" & Files(i)
          End Select
          
          frmSplash.FileCarregada Tot(i), Kb(i)
          
          If UCase(Left(Files(i), 10)) = UCase("plantilla_") Then
             d = DateSerial(Mid(Files(i), 13, 2), Mid(Files(i), 15, 2), Mid(Files(i), 17, 2))
             e = Mid(Files(i), 11, 1)
             InsertFeineaAFer "ImportarComanda", "Integraciones", "[" & Format(d, "dd-mm-yyyy") & "]", "[" & e & "]", "[" & "Importat_" & Files(i) & "]"
          End If
          
          If Files(i) = "client.unl" Then InsertFeineaAFer "ImportarClients", "Integraciones", "[" & Format(Now, "dd-mm-yyyy hh:mm:ss") & "]"
          If Files(i) = "artemp.unl" Then InsertFeineaAFer "ImportarArticlesEmpresa", "Integraciones", "[" & Format(Now, "dd-mm-yyyy hh:mm:ss") & "]"
          If Files(i) = "articu.unl" Then InsertFeineaAFer "ImportarArticles", "Integraciones", "[" & Format(Now, "dd-mm-yyyy hh:mm:ss") & "]"
          If Files(i) = "casta.unl" Then InsertFeineaAFer "ImportarFamilies", "Integraciones", "[" & Format(Now, "dd-mm-yyyy hh:mm:ss") & "]"
          If Files(i) = "famili.unl" Then InsertFeineaAFer "ImportarFamilies", "Integraciones", "[" & Format(Now, "dd-mm-yyyy hh:mm:ss") & "]"
          If Files(i) = "promoc.unl" Then InsertFeineaAFer "ImportarPromocions", "Integraciones", "[" & Format(Now, "dd-mm-yyyy hh:mm:ss") & "]"
          If Files(i) = "tarifa.unl" Then InsertFeineaAFer "ImportarTarifa", "Integraciones", "[" & Format(Now, "dd-mm-yyyy hh:mm:ss") & "]"
       End If
    Next
    
End Sub
Sub IntegracionesEnviaValidacionsPas2(Files() As String)
   Dim CampsClaus As String, CampsCreate As String, lin As String, rs As rdoResultset, f, i As Integer, NomFileFeinaFeta As String, Condicio As String, Fd() As String
   Dim d As Date, Df As Date, CondicioEnviamentClient As String, CondicioEnviamentViatge As String, CondicioEnviamentEquip As String, nomfile As String, res As rdoResultset
   Dim LastNomTaula As String, NomTaula As String, Files1() As String, Lineas As Integer
   
   Set res = Db.OpenResultset("Select * From Records Where Concepte = 'Validacions'")
   If res.EOF Then ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('Validacions',DATEADD(Day, -1, GetDate()))"
   res.Close
   
   f = FreeFile
   If ExisteixTaula("Validacions") Then
      Lineas = 0
      Set rs = Db.OpenResultset("Select v.data,v.datavalidat,unio,nom,estat from Validacions v join Clients_Imp_Codis C on c.codi=v.botiga join hit.dbo.web_users s on v.responsable=s.id where DataValidat >= (Select Max(TimeStamp) From  Records Where Concepte = 'Validacions') ")
      If Not rs.EOF Then
         LastNomTaula = ""
         While Not rs.EOF
            If Lineas <= 0 Then
               EnviaFacturacioCreaFile Files, f, CampsClaus, "Validacions_v0.unl"
               LastNomTaula = ""
               Lineas = 500
            End If
            lin = rs(0) & "|" & rs(1) & "|" & rs(2) & "|" & rs(3) & "|" & rs(4)
            Print #f, lin
            rs.MoveNext
            Lineas = Lineas - 1
         Wend
         Print #f, "EOF " & Now()
         rs.Close
         Close f
      End If
   End If
   
   ExecutaComandaSql "Update Records Set [TimeStamp] = GetDate() Where Concepte = 'Validacions' "

End Sub
Sub IntegracionesEnviaFacturacioPas2(Files() As String)
   Dim CampsClaus As String, CampsCreate As String, lin As String, rs As rdoResultset, f, i As Integer, NomFileFeinaFeta As String, Condicio As String, Fd() As String
   Dim d As Date, Df As Date, CondicioEnviamentClient As String, CondicioEnviamentViatge As String, CondicioEnviamentEquip As String, nomfile As String, res As rdoResultset
   Dim LastNomTaula As String, NomTaula As String, Files1() As String, Lineas As Integer
   
   EsborraTaula "TemporalEnviaments"
   CreaTaulaServit "TemporalEnviaments", False
   ExecutaComandaSql "ALTER TABLE TemporalEnviaments ADD DiaDesti VarChar(255)  NULL"
   
   ExecutaComandaSql "Delete Records Where Concepte = 'FacturacioNovaDataBotiguesIntegraciones' "
   ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values('FacturacioNovaDataBotiguesIntegraciones',GetDate())"
   Set res = Db.OpenResultset("Select * From Records Where Concepte = 'FacturacioBotiguesIntegraciones'")
   If res.EOF Then ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('FacturacioBotiguesIntegraciones',DATEADD(Day, -1, GetDate()))"
   res.Close
   
   Set rs = Db.OpenResultset("Select * From TemporalEnviaments ")
   CampsClaus = ""
   CampsCreate = ""
   For i = 0 To rs.rdoColumns.Count - 1
      If Not rs.rdoColumns(i).Name = "DiaDesti" Then
         If Len(CampsCreate) > 0 Then CampsCreate = CampsCreate & ","
         CampsClaus = CampsClaus & "[" & rs.rdoColumns(i).Name & "]" & "[" & DeTypeASt(rs.rdoColumns(i).Type) & "]"
         CampsCreate = CampsCreate & "[" & rs.rdoColumns(i).Name & "]"
      End If
   Next
   
   If ExisteixTaula("ComandesModificades") Then
      Set rs = Db.OpenResultset("Select Distinct TaulaOrigen From ComandesModificades ")
      While Not rs.EOF
         My_DoEvents
         Condicio = " hora <> 69 And Id In (Select Id From ComandesModificades Where  "
         Condicio = Condicio & "   TaulaOrigen = '" & rs("TaulaOrigen") & "'  "
         Condicio = Condicio & "   And [TimeStamp] >  (Select Max(TimeStamp) From  Records Where Concepte = 'FacturacioBotiguesIntegraciones')  "
         Condicio = Condicio & "   And [TimeStamp] <= (Select Max(TimeStamp) From  Records Where Concepte = 'FacturacioNovaDataBotiguesIntegraciones') "
         Condicio = Condicio & "            ) "
         
         EnviaFacturacioDia "TemporalEnviaments", rs("TaulaOrigen"), CampsCreate, CondicioEnviamentClient, CondicioEnviamentViatge, CondicioEnviamentEquip, Condicio
         rs.MoveNext
      Wend
      rs.Close
   End If
   ExecutaComandaSql "Delete ComandesModificades Where  [TimeStamp] <= (Select Max(TimeStamp) From  Records Where Concepte = 'FacturacioNovaDataBotiguesIntegraciones') "
   
   f = FreeFile
   Lineas = 0
   Set rs = Db.OpenResultset("Select c.unio,diadesti,a.unio,quantitatdemanada From TemporalEnviaments t join articles_imp_codis a on t.codiarticle = a.codi join clients_imp_codis c on c.codi = t.client", rdConcurRowVer)
   If Not rs.EOF Then
      LastNomTaula = ""
      CampsClaus = ""
'S024|\ |16/06/05|  A010|12,0|\ |1|12,0|0|0|
'S024|\ |16/06/05|  G035|0,0|\ |1|0,0|0|0|
'S024|\ |16/06/05|  TI16|0,0|\ |1|0,0|0|0|
      
      While Not rs.EOF
         If Lineas <= 0 Then
            EnviaFacturacioCreaFile Files, f, CampsClaus, "Plantilla_v0.unl"
            LastNomTaula = ""
            Lineas = 500
         End If
         Fd = Split(Left(Right(rs(1), 9), 8), "-")
         lin = rs(0) & "||" & Fd(2) & "/" & Fd(1) & "/" & "20" & Fd(0) & "|" & rs(2) & "|" & rs(3) & "||00:00|" & rs(3) & "|0|0|0|"
         Print #f, lin
         rs.MoveNext
         Lineas = Lineas - 1
      Wend
      Print #f, "EOF " & Now()
      rs.Close
      Close f
   End If
   
   EsborraTaula "TemporalEnviaments"
   ExecutaComandaSql "Delete Records Where Concepte = 'Facturacio'"
   ExecutaComandaSql "Update Records Set Concepte = 'Facturacio' Where Concepte = 'FacturacioNovaDataBotiguesIntegraciones' "

End Sub


Sub IntegracionesEnviaValidacions(Files() As String)
      
   ReDim Files(0)
   
   ExecutaComandaSql "Use iblatpa"
      IntegracionesEnviaValidacionsPas2 Files
   ExecutaComandaSql "Use iartpa"
      IntegracionesEnviaValidacionsPas2 Files
   ExecutaComandaSql "Use integraciones"
   
End Sub


Sub IntegracionesEnviaFacturacio(Files() As String)
   
   ReDim Files(0)
   
   ExecutaComandaSql "Use iblatpa"
      IntegracionesEnviaFacturacioPas2 Files
   ExecutaComandaSql "Use iartpa"
      IntegracionesEnviaFacturacioPas2 Files
   ExecutaComandaSql "Use integraciones"


End Sub





Sub SincronitzaIntegracionesReb(Path As String)
   Static LastIntent As Date
   
   DoEvents
   If DateDiff("n", LastIntent, Now) < 5 Then Exit Sub
   
   ExternCarregaFtp
   SincronitzaIntegracionesRebPas2 AppPath & "\tmp"
   
   LastIntent = Now
   
End Sub

Sub FtpExternImportaUnl(Fil As String, NomTaula As String)
   Dim f, lin As String, Creataula As String, Primercop As Boolean, K As Integer, Valors As String, Dades() As String, Sql As String, Q As rdoQuery, i As Integer, TeNumerics As Boolean, TeEspais As Boolean
   
On Error GoTo nor
   Informa "Important : " & NomTaula
   NomTaula = Car2(NomTaula, ".")
   Primercop = True
   TeNumerics = False
   TeEspais = False
   If NomTaula = "Importat_tarifa" Then TeNumerics = True
   If NomTaula = "Importat_promoc" Then TeNumerics = True
   If NomTaula = "Importat_articu" Then TeEspais = True
   If NomTaula = "Importat_artemp" Then TeEspais = True
   If Left(NomTaula, 19) = "Importat_Plantilla_" Then TeNumerics = True
      
   f = FreeFile
   K = 0
   
   Open Fil For Input As f
   While Not EOF(f)
      Line Input #f, lin
      lin = Normalitza2(lin, TeNumerics, False)
      Dades = Split(lin, "|")
            
      If Primercop Then
         Creataula = ""
         Valors = ""
         For i = 0 To UBound(Dades)
            If Not Creataula = "" Then Creataula = Creataula & ","
            Creataula = Creataula & "Param_" & i & " [nvarchar] (255) NULL "
            If Not Valors = "" Then Valors = Valors & ","
            Valors = Valors & "?"
         Next
         
         ExecutaComandaSql "Drop Table [" & NomTaula & "]"
         ExecutaComandaSql "CREATE TABLE [" & NomTaula & "] (" & Creataula & ") ON [PRIMARY]"
         Set Q = Db.CreateQuery("", "Insert into [" & NomTaula & "] Values (" & Valors & ") ")
         Primercop = False
      End If
      For i = 0 To UBound(Dades)
         If NomTaula = "Importat_articu" And i = 3 Then Dades(i) = Normalitza2(Dades(i), False, True)
         If NomTaula = "Importat_artemp" And i = 2 Then Dades(i) = Normalitza2(Dades(i), False, True)
         Q.rdoParameters(i) = Dades(i)
      Next
      Q.Execute
      
      DoEvents
   Wend
   
   Close f
nor:
End Sub
Sub FtpExternImportaCsv(Fil As String, NomTaula As String)
   Dim f, lin As String, Creataula As String, Primercop As Boolean, K As Integer, Valors As String, Dades() As String, Sql As String, Q As rdoQuery, i As Integer, TeNumerics As Boolean, TeEspais As Boolean
   
   
   Informa "Important : " & NomTaula
   NomTaula = Car2(NomTaula, ".")
   Primercop = True
   TeNumerics = False
   TeEspais = False
   If NomTaula = "Importat_tarifa" Then TeNumerics = True
   If NomTaula = "Importat_articu" Then TeEspais = True
   If NomTaula = "Importat_artemp" Then TeEspais = True
   
   f = FreeFile
   K = 0
   
   Open Fil For Input As f
   While Not EOF(f)
      Line Input #f, lin
      lin = Normalitza2(lin, TeNumerics, False)
      Dades = Split(lin, ";")
            
      If Primercop Then
         Creataula = ""
         Valors = ""
         For i = 0 To UBound(Dades)
            If Not Creataula = "" Then Creataula = Creataula & ","
            Creataula = Creataula & "Param_" & i & " [nvarchar] (255) NULL "
            If Not Valors = "" Then Valors = Valors & ","
            Valors = Valors & "?"
         Next
         
         ExecutaComandaSql "Drop Table [" & NomTaula & "]"
         ExecutaComandaSql "CREATE TABLE [" & NomTaula & "] (" & Creataula & ") ON [PRIMARY]"
         Set Q = Db.CreateQuery("", "Insert into [" & NomTaula & "] Values (" & Valors & ") ")
         Primercop = False
      End If
      For i = 0 To UBound(Dades)
         If NomTaula = "Importat_articu" And i = 3 Then Dades(i) = Normalitza2(Dades(i), False, True)
         If NomTaula = "Importat_artemp" And i = 2 Then Dades(i) = Normalitza2(Dades(i), False, True)
         If i < Q.rdoParameters.Count Then Q.rdoParameters(i) = Dades(i)
      Next
      Q.Execute
      
      DoEvents
   Wend
   Close f

End Sub

Function Normalitza(s) As String
   Dim Ss As String, P As Integer
   
   Ss = ""
   If Not IsNull(s) Then Ss = s
   
   Normalitza = Ss
   P = 0
   Do
      P = InStr(P + 1, Ss, "#")
      If P > 0 Then
         Ss = Left(Ss, P) & "#" & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
   
   P = 0
   Do
      P = InStr(P + 1, Ss, Chr(13) & Chr(10))
      If P > 0 Then
         Ss = Left(Ss, P - 1) & "#R" & Right(Ss, Len(Ss) - P - 1)
         P = P + 1
      End If
   Loop While P > 0
   
   P = 0
   Do
      P = InStr(P + 1, Ss, "\|")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & "  " & Right(Ss, Len(Ss) - P - 1)
         P = P + 1
      End If
   Loop While P > 0
   
   P = 0
   Do
      P = InStr(P + 1, Ss, "'")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & " " & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
   
   P = 0
   Do
      P = InStr(P + 1, Ss, "¦")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & " " & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
   
   Normalitza = Ss

End Function



Function Normalitza2(s, TreureComes As Boolean, TeEspais As Boolean) As String
   Dim Ss As String, P As Integer
   
   Ss = ""
   If Not IsNull(s) Then Ss = s
   
   Normalitza2 = Ss
   P = 0
   Do
      P = InStr(P + 1, Ss, "¤")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & "ñ" & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
   
   P = 0
   Do
      P = InStr(P + 1, Ss, "#")
      If P > 0 Then
         Ss = Left(Ss, P) & "#" & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
   
   
   P = 0
   Do
      P = InStr(P + 1, Ss, Chr(13) & Chr(10))
      If P > 0 Then
         Ss = Left(Ss, P - 1) & "#R" & Right(Ss, Len(Ss) - P - 1)
         P = P + 1
      End If
   Loop While P > 0
   
   P = 0
   Do
      P = InStr(P + 1, Ss, "\|")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & "  " & Right(Ss, Len(Ss) - P - 1)
         P = P + 1
      End If
   Loop While P > 0
   
   P = 0
   Do
      P = InStr(P + 1, Ss, "'")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & " " & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
   
   If TreureComes Then
      P = 0
      Do
         P = InStr(P + 1, Ss, ",")
         If P > 0 Then
            Ss = Left(Ss, P - 1) & "." & Right(Ss, Len(Ss) - P)
            P = P + 1
         End If
       Loop While P > 0
   End If
   
   If TeEspais Then
      P = 0
      Do
         P = InStr(P + 1, Ss, "  ")
         If P > 0 Then
            Ss = Left(Ss, P - 1) & Right(Ss, Len(Ss) - P)
            P = P - 1
         End If
       Loop While P > 0
   End If
      
   P = 0
   Do
      P = InStr(P + 1, Ss, "¦")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & " " & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
   
   Normalitza2 = Ss

End Function




Function Car2(ByRef s As String, Tag As String) As String
   Dim P As Integer
   P = InStr(s, Tag)
   If P > 0 Then
      Car2 = Left(s, P - 1)
      s = Right(s, Len(s) - P)
   Else
      Car2 = s
      s = ""
   End If
   
End Function




Sub SincronitzaExterns(Empresa As String, Accio As String)
   Dim rs As rdoResultset, ElPath As String
   
   ExecutaComandaSql "CREATE TABLE [Importat_Params] ([Clau]    [nvarchar] (255) NULL ,[Valor_1] [nvarchar] (255) NULL ,[Valor_2] [nvarchar] (255) NULL ,[Valor_3] [nvarchar] (255) NULL ,[Valor_4] [nvarchar] (255) NULL ,[Valor_5] [nvarchar] (255) NULL ) ON [PRIMARY]"
   ElPath = ""
   Set rs = Db.OpenResultset("select valor_1 from Importat_Params where Clau = 'PathFtp'")
   If Not rs.EOF Then If Not IsNull(rs(0)) Then ElPath = rs(0)
   rs.Close
   
   If Not ElPath = "" Then
      Select Case Accio
          Case "Envia"
              SincronitzaIntegracionesReb ElPath
          Case "Reb"
              
          Case Else
              SincronitzaIntegracionesReb ElPath
      End Select
   End If
    
End Sub

Function tipusFile(s As String) As String
   Dim P As Integer
   tipusFile = ""
   
   P = 0
   While InStr(P + 1, s, ".")
      P = InStr(P + 1, s, ".")
   Wend
   
   If P > 0 Then tipusFile = Right(s, Len(s) - P)

End Function


Function EsNumero(s) As String
   Dim Ss As String, P As Integer
   Dim i
   
   
   Ss = ""
   If Not IsNull(s) Then Ss = s
   
   EsNumero = Ss
'   For i = 1 To Len(EsNumero)
'      If Mid(EsNumero, i, 1) = "." And i > 1 Then
'        If (Mid(EsNumero, i - 1, 1) >= "0" And Mid(EsNumero, i - 1, 1) <= "9") And (Mid(EsNumero, i + 1, 1) >= "0" And Mid(EsNumero, i + 1, 1) <= "9") Then
'        Else
'            EsNumero = Left(EsNumero, i - 1) & " " & Right(EsNumero, Len(EsNumero) - i)
'        End If
'      Else
'        If (Mid(EsNumero, i, 1) >= "0" And Mid(EsNumero, i, 1) <= "9") Or Mid(EsNumero, i, 1) = " " Or Mid(EsNumero, i, 1) = "," Or Mid(EsNumero, i, 1) = " " Then
'        Else
'            EsNumero = Left(EsNumero, i - 1) & " " & Right(EsNumero, Len(EsNumero) - i)
'        End If
'      End If
'   Next
   
   Ss = EsNumero
   
   P = 0
   Do
      P = InStr(P + 1, Ss, ".")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
  
   P = 0
   Do
      P = InStr(P + 1, Ss, ",")
      If P > 0 Then
         Ss = Left(Ss, P - 1) & "." & Right(Ss, Len(Ss) - P)
         P = P + 1
      End If
   Loop While P > 0
  
   EsNumero = Ss

End Function
Sub TraspasaHuellas()
    Dim biFir() As Byte, nDataSize As Double, aa, rs As rdoResultset, RsAdo As New ADODB.Recordset
    Dim objNBioBSP As NBioBSPCOMLib.NBioBSP
    Dim objDevice As IDevice                ' Device object
    Dim objExtraction As IExtraction        ' Extraction object
    Dim objMatching As IMatching            ' Matching object
    Dim objFPData As IFPData                ' FPData object
    Dim objFPImage As IFPImage              ' FPImage object
Exit Sub
    Set objNBioBSP = New NBioBSPCOMLib.NBioBSP
    Set objDevice = objNBioBSP.Device
    Set objExtraction = objNBioBSP.Extraction
    Set objMatching = objNBioBSP.Matching
    Set objFPData = objNBioBSP.FPData
    Set objFPImage = objNBioBSP.FPImage


    Set rs = Db.OpenResultset("select * from Dedos where usuario = 1246 ", rdConcurRowVer)
    While Not rs.EOF
        aa = DameValor(rs, "Fir")    ' "AQAAABQAAABUAQAAAQASAAMAZAAAAAAAUAEAAOtQzci2ZciRZwNk8dP*59LMzOOt/N7I9CifWE8dNlhkt6X2/opSq3leDSZLvACnIEJ1pQGtWNkxOibvOB38AZs7SAfgm9pWoL9U*Zav9h*HvZ4XnuSNbDzB*l6us8t2hjeISMRNxO2u8NyAJeXDVJ1CoV9MlE527MfYe8lRHeXsYF3l*XZ5*2gruhkYzXBpZzVbhVYA1yb0h1L33tZ/R0f9APP/5eoBilx9cGtrUlKPbNOy2qAGuToQun5b3VqZ4ZiGvy*Tqj5fIHOgMuN2/TnmKQ6B1mdmcQtUOyIw/nYXuexATBVluh7ytwiXYXzSIUzNUR31IAnaY0dQdoed01Y5Uib8kBj61OkB0/zmQ4tXv*Qsn/ZsD4e8BCwQQhyfOah5CSI*D1ANFIJiMoVu/YiRe9gw2GVhgrfqX6Ue*dggzP8SOoZWuB3ms618z7g6qQ"
        nDataSize = Len(aa)
        ReDim biFir(nDataSize)
        biFir = aa
    
        objFPData.import 1, 0, 1, 0, nDataSize, biFir
        ' select * from [NitgenAccessManager].[dbo].[NGAC_USERFIR]
        
                Set db2 = New ADODB.Connection
        db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"
   
        Set RsAdo = rec("delete [NitgenAccessManager].[dbo].[NGAC_USERFIR] where userIdindex = 3 ", True)
        Set RsAdo = rec("select * from [NitgenAccessManager].[dbo].[NGAC_USERFIR] where userIdindex = 3 ", True)
        RsAdo.AddNew
        RsAdo("userIdindex").Value = 3
        RsAdo("FIR").Value = objFPData.FIR
        RsAdo.Update
        RsAdo.Close
        
        
        
        MsgBox Len(objFPData.FIR)
        rs.MoveNext
    Wend
  
    
    
End Sub


