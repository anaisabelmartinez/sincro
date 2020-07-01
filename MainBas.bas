Attribute VB_Name = "MainBas"
Option Explicit

'tiempo de espera en segundos
Public Const SESSION_TIME = 120

'Puertos por defecto
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080

'Tipos de servicios
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

'Tipos de conexión
Private Const INTERNET_CONNECTION_LAN As Long = &H2
Private Const INTERNET_CONNECTION_MODEM As Long = &H1

'API para detectar conexión
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpSFlags As Long, ByVal dwReserved As Long) As Long

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Dim Bandera ' Handel al fitcher obert mentres l'aplicació encesa.

Public Cfg_Llicencia  As String
Public Cfg_Server As String
Public Cfg_Database As String
Public Db As New rdoConnection
Public db2 As New ADODB.Connection
Public db2MyId As String
Public db2User As String
Public db2Psw As String
Public db2NomDb As String, DbCone As String
Public db2Server As String

Public NomServerInternet As String
Dim UltimaHoraFeinaHorariaFeta As Integer, FeinesAFerUnCopCadaHora As Boolean

Dim Q_Insert As rdoQuery
Dim Q_Delete As rdoQuery

Type TipFeina
   empresa As String
   Db As String
   Path As String
   llicencia As String
   Server As String
   Tipus As Integer
   EscoltaLlicencies() As String
   Ftp_Server As String
   Ftp_User As String
   Ftp_Pssw As String
End Type

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long


Global feina() As TipFeina, AppPath As String, UltimaAccio As Date, Velocitat As Double, EmpresaActualNum As Integer, EmailGuardia As String
Global LastServer As String, LastDatabase As String, SistemaMud As Boolean, SistemaObert As Boolean, FeinaAfer As String, Sempre As Boolean, EmpresaActual As String, LastLlicencia As String, ServerActual As String, PdaEnviaTot As Boolean, EsDispacher As Boolean
Dim slo_connexio, SGS_ND, SGS_REMOTEADDR, SGC_DBUSER, SGC_DBPASSWORD, SGC_DBGDR, SGC_SERVER, SGS_CONGDR, SGS_CONUSER
Dim SGD_Literal
Dim ultimaSql
Dim slo_fs, slo_fname As Object


'-------------------------------------------------------------------------------------------------------------------

' Constantes para las funciones Api
Const scUserAgent = "API-Guide test program"
Const INTERNET_OPEN_TYPE_DIRECT = 1
Const INTERNET_OPEN_TYPE_PROXY = 3
Const INTERNET_FLAG_RELOAD = &H80000000
Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
  
' Esta función crea una conexión a internet
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" ( _
    ByVal sAgent As String, _
    ByVal lAccessType As Long, _
    ByVal sProxyName As String, _
    ByVal sProxyBypass As String, _
    ByVal lFlags As Long) As Long
  
' Esta Api abre un Url
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" ( _
    ByVal hInternetSession As Long, _
    ByVal lpszUrl As String, _
    ByVal lpszHeaders As String, _
    ByVal dwHeadersLength As Long, _
    ByVal dwFlags As Long, _
    ByVal dwContext As Long) As Long
  
' Esta cierra la conexión pasandole el Handle que habíamos obtenido antes
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
  
' Esta Api lee el contenido y lo devuelve en un Buffer que _
    contendrá el contenido del fichero
Private Declare Function InternetReadFile Lib "wininet" ( _
    ByVal hFile As Long, _
    ByVal sBuffer As String, _
    ByVal lNumBytesToRead As Long, _
    lNumberOfBytesRead As Long) As Integer
    

Function DbNameEsBuit() As Boolean
On Error GoTo nor
DbNameEsBuit = False
If Db.Name = "" Then
    DbNameEsBuit = True
    
End If
Exit Function

nor:
    tancasipots
    DbNameEsBuit = True
End Function

Sub descargaSFTP(empresa)
    Dim rsLicencias As rdoResultset, rsVpnToc As rdoResultset, rsHit As rdoResultset
    Dim ip As String, User As String, sftp_path As String, Lic As String, cn As String
    Dim idShell As Double, Fils As Integer
    Dim sArchivo As String
    Dim tempsDescarga As Date
    Dim pathSincro As String
    
    pathSincro = "..\..\DadesSincro\Empreses\" & empresa
    Set rsHit = Db.OpenResultset("select Path from hit.dbo.Web_Empreses where nom = '" & empresa & "'")
    If Not rsHit.EOF Then pathSincro = Mid(rsHit("path"), 2, Len(rsHit("path")) - 1)
    
    'Vaciamos directorios de descargas
    sArchivo = Dir(App.Path & "\descargas\" & empresa & "\*.*")
    Do While sArchivo <> ""
        Kill App.Path & "\descargas\" & empresa & "\" & sArchivo
        sArchivo = Dir
    Loop
   
    sArchivo = Dir(App.Path & "\descargas\" & empresa & "\cfg\*.*")
    Do While sArchivo <> ""
        Kill App.Path & "\descargas\" & empresa & "\cfg\" & sArchivo
        sArchivo = Dir
    Loop
   
   'Por cada licencia de la empresa descargamos lo que hay pendiente en el TOC y cargamos los ficheros necesarios
    Set rsLicencias = Db.OpenResultset("select * from paramshw where valor2<>'NO ACTIVA' order by codi")
    While Not rsLicencias.EOF
On Error GoTo nextBot
        Lic = rsLicencias("codi")
        If Lic <> "" Then
        'If Lic = "178" Then 'Or Lic = "115" Or Lic = "110" Or Lic = "129" Then
            Set rsVpnToc = Db.OpenResultset("Select [user], ip, sftp_path, Llicencia, cn From [Hit].[dbo].[vpntoc] Where activo = 1 And llicencia = " & Lic)
            If Not rsVpnToc.EOF Then
                User = rsVpnToc("user")
                sftp_path = rsVpnToc("sftp_path")
                ip = rsVpnToc("ip")
                cn = rsVpnToc("cn")
                
                'DESCARGA (TOC -->> SINCRO) ----------------------------------------------------------------------------------------------------------------------
                InformaMiss "DESCARGANDO " & cn
                
                tempsDescarga = DateAdd("n", 3, Now())
                idShell = Shell(App.Path & "\sftp_op2.cmd " & User & " " & ip & " """ & sftp_path & """ """ & Lic & """ """ & cn & """ """ & empresa & """", vbHide)
                If idShell > 0 Then
                    'Esperar a que finalice
                     WaitForTerm idShell
                End If

                'Una vez descargado lo borramos del directorio de descargas
                Fils = 0
                sArchivo = Dir(App.Path & "\descargas\" & empresa & "\*.*")
                Do While sArchivo <> ""
On Error GoTo siguiente
                    FileCopy App.Path & "\descargas\" & empresa & "\" & sArchivo, pathSincro & "\" & sArchivo
                    If (ExisteixTaula("SFPT_DEBUG")) Then
                        ExecutaComandaSql "insert into SFPT_DEBUG values (getdate(), 'DESCARGA', '" & Lic & "', '" & User & "', '" & sftp_path & "', '" & ip & "', '" & cn & "', '" & sArchivo & "')"
                    End If
                        
                    'FileCopy App.Path & "\descargas\" & Empresa & "\" & sArchivo, App.Path & "\descargas\" & Empresa & "\Bak\" & sArchivo
                    Kill App.Path & "\descargas\" & empresa & "\" & sArchivo
siguiente:
                    
                    sArchivo = Dir
                    Fils = Fils + 1
                Loop
                
                '-------------------------------------------------------------------------------------------------------------------------------------------
                'DE MOMENTO NO ME FIO DE DESCARGAR FICHEROS DE CONFIGURACIÓN (YA SE HA LIADO PARDA CON UNO DE DEPENDIENTAS QUE HA ENTRADO DE NO SÉ DONDE)
                '-------------------------------------------------------------------------------------------------------------------------------------------
                'sArchivo = Dir(App.Path & "\descargas\" & Empresa & "\cfg\*.*")
                'Do While sArchivo <> ""
                '    FileCopy App.Path & "\descargas\" & Empresa & "\cfg\" & sArchivo, pathSincro & "\" & sArchivo
                '    Kill App.Path & "\descargas\" & Empresa & "\cfg\" & sArchivo
                '    sArchivo = Dir
                '    Fils = Fils + 1
                'Loop
                
                InformaMiss Fils & " FICHEROS DESCARGADOS en " & cn, True
                
                'CARGA (SINCRO -->> TOC) ----------------------------------------------------------------------------------------------------------------------
                InformaMiss "CARGANDO " & cn
                
                Fils = 0
                sArchivo = Dir(App.Path & "\cargas\" & empresa & "\" & Lic & "\*.*")
                If sArchivo <> "" Then
                    idShell = Shell(App.Path & "\sftp_subida.cmd " & User & " " & ip & " """ & App.Path & "\cargas\" & empresa & "\" & Lic & "\"" """ & sftp_path & """ ", vbHide)
                    If idShell > 0 Then
                        'Esperar a que finalice
                         WaitForTerm idShell
                    End If
                    
                    
                    Do While sArchivo <> ""
                        If (ExisteixTaula("SFPT_DEBUG")) Then ExecutaComandaSql "insert into SFPT_DEBUG values (getdate(), 'CARGA', '" & Lic & "', '" & User & "', '" & sftp_path & "', '" & ip & "', '" & cn & "', '" & sArchivo & "')"
                        
                        Kill App.Path & "\cargas\" & empresa & "\" & Lic & "\" & sArchivo
                        sArchivo = Dir
                        Fils = Fils + 1
                    Loop
                End If
                
                InformaMiss Fils & " FICHEROS CARGADOS en " & cn, True
            End If
        'End If
        End If
nextBot:
        rsLicencias.MoveNext
    Wend
    

End Sub

Sub WaitForTerm(ByVal PID As Long)
On Error GoTo Gestion_Error

    'Variables locales
    Dim phnd As Long

    phnd = OpenProcess(SYNCHRONIZE, 0, PID)
    If phnd > 0 Then
        Call WaitForSingleObject(phnd, 60000)
        Call CloseHandle(phnd)
    End If
    Exit Sub
Gestion_Error:
    
End Sub
Function sf_recGdr(ByVal sls_sql)
    Dim slo_rs
    Dim slo_comm
    
    ultimaSql = sls_sql

    Set slo_rs = CreateObject("ADODB.Recordset")
    Set slo_comm = CreateObject("ADODB.Command")
    
    slo_comm.ActiveConnection = SGS_CONGDR
    slo_comm.CommandTimeout = 1500
    slo_comm.CommandText = sls_sql
    ' AQUI HAY UN ERROR
    On Error Resume Next
        
    Set slo_rs = slo_comm.Execute
    ' aqui había errores
    Set sf_recGdr = slo_rs
End Function
Sub ActualtizaUrls()
    Dim Rs As ADODB.Recordset, iD, Resultat As String, K
        
    Resultat = llegeigHtml("http://www.minimalia.net/frase.asp?id=" & K + DatePart("y", Now))
    If Len(Resultat) > 0 Then
        Set Rs = rec("select newid() i")
        iD = Rs("i")
        
        Set Rs = rec("select Top 1 * from " & WwwCache() & " ", True)
        Rs.AddNew
        Rs("Id").Value = iD
        Rs("NomFile").Value = "LaFrase"
        Rs("TimeStamp").Value = Now
        Rs("Original").Value = "XLS"
        Rs("Resum").Value = "application/vnd.ms-excel"
        
        Rs.Update
        Rs.Close
        
   End If



End Sub

Function ArticlesPropietats(Codi, Variable, Defecte)
    Dim Rs As rdoResultset
       
    ArticlesPropietats = Defecte
    Set Rs = Db.OpenResultset("Select cast(Valor as nvarchar(255)) From ArticlesPropietats with (nolock) Where CodiArticle = " & Codi & "  and Variable = '" & Variable & "' ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then ArticlesPropietats = Rs(0)
    Rs.Close

End Function

Sub AvisaGosPaso2(Estat)
    Dim Rs As rdoResultset
On Error Resume Next
'    ConnectaSqlServer LastServer, "Hit"
'    Set Rs = Db.OpenResultset("Update Hit.Dbo.GosDeTura Set TsVista = getdate(),UltimaFraseObella = '" & Estat & "' Where NomObella = '" & FeinaAfer & "'")

End Sub

Sub CalculsEdiPut(iD As String)
    Dim f, a, Rs
    
    Set Rs = Db.OpenResultset("select * from FilesEdi Where id = '" & iD & "' Order by NumLinea ")
    f = FreeFile
    
'    Open "C:\eDiversa\" & UCase(EmpresaActual) & "\" & rs("NomFile") For Output As #f
' ***************************************************************************************
' COMENTADA LA LINEA SIGUIENTE 07/05/2012 FALLO EN ENDIVERSA QUE SOLO COGE EL DIRECTORIO
' POR DEFECTO c:\eDiversa\Send\Planos y si entras a configuracion se borra lo de c:\Ediversa\Daza
'    Open "\\5.82.233.79\c$\eDiversa\" & UCase(EmpresaActual) & "\" & Rs("NomFile") For Output As #f
' ****************************************************************************************
    Open "\\5.82.233.79\c$\eDiversa\Send\Planos\" & Rs("NomFile") For Output As #f
    While Not Rs.EOF
        Print #f, Rs("LinModificada")
        Rs.MoveNext
    Wend
    Rs.Close
    Close f

End Sub

Function EnviaFTPSerunion(fname As String, ftype As String) As String
    Dim localPath As String
    localPath = "c:\" & fname
    Dim remoteFilename As String
    remoteFilename = "/export/" & ftype & fname
    
    Dim IpFtp As FTP
   
    If frmSplash.IpConexio Is Nothing Then Set frmSplash.IpConexio = New ConexioIp
    
    Set IpFtp = Nothing
    Set IpFtp = frmSplash.FTP1
   
    IpFtp.RemotePort = 21
    IpFtp.WinsockLoaded = True
    IpFtp.RemoteHost = "ftp.spairal.com"
    IpFtp.User = "ftp_7571_ESPAI_POSTR"
    IpFtp.Password = "Sefigece$72"
    IpFtp.Passive = True
   
    Informa "Intentant Logon."
    If frmSplash.IpConexio.ExecutaAction(2) Then
        frmSplash.IpConexio.ExecutaAction 0, "RemoteFile", remoteFilename
        frmSplash.IpConexio.ExecutaAction 0, "LocalFile", localPath
        frmSplash.IpConexio.ExecutaAction 0, "TransferMode", 2

        If frmSplash.IpConexio.ExecutaAction(a_Upload) Then
            Informa "Transfer ok."
        End If
    End If
    frmSplash.IpConexio.Desconecta
    
    EnviaFTPSerunion = "OK"
End Function
Sub ExecutaCalculPuntual(FeinaAfer, Optional CalPlegar As Boolean = True)
    Dim Rs As rdoResultset, rsEmail As rdoResultset, P, iD As String, IdT As String, emp, MLlicencia As String, MServer As String
    Dim MDb As String, Tipus As String, p1 As String, P2 As String, P3 As String, P4 As String, P5 As String, pt1, pt2, pt3
    Dim D As Date
 
    If CalPlegar Then ConnectaSqlServer LastServer, "Hit"
    P = InStr(FeinaAfer, ":")
    Set Rs = Db.OpenResultset("select * from hit.dbo.CalculsEspecials Where IdFeinaAFer = '" & Right(FeinaAfer, Len(FeinaAfer) - P) & "' ")
'    Set Rs = Db.OpenResultset("sp_who '" & User & "'")
   
    If Not Rs.EOF Then
        iD = Rs("Id")
        IdT = Rs("IdFeinaAFer")
        emp = Rs("Empresa")
        Tipus = NoNull(Rs("Tipus"))
        p1 = NoNull(Rs("Param1"))
        P2 = NoNull(Rs("Param2"))
        P3 = NoNull(Rs("Param3"))
        P4 = NoNull(Rs("Param4"))
        P5 = NoNull(Rs("Param5"))
        ExecutaComandaSql "Update " & TaulaCalculsEspecials & " Set Estat = 'Calculant' Where Id = '" & iD & "' "
        Set Rs = Db.OpenResultset("Select Llicencia,Db,Path,Db_Server As Servidor,Nom From hit.dbo.Web_Empreses Where Nom = '" & emp & " ' ")
        If Not Rs.EOF Then
            MDb = Rs("Db")
            MLlicencia = Rs("Llicencia")
            MServer = Rs("Servidor")
            InformaEmpresa Rs("Nom")
            If ConnectaSqlServer(MServer, MDb) Then
                    InformaMiss Tipus
                    Select Case Tipus
                        Case "PedidoCalculaExtern":
                             'PedidoCalcula p1, CDbl(P2), CDbl(P3)
                             PedidoCalcula_91 p1, CDbl(P2), CDbl(P3)
                             PedidoCalcula_V2 p1, CDbl(P2), CDbl(P3)
                        Case "ExportaCp":
                           ExportaCp p1, P2, P3, P4, P5, iD
                        Case "SincroDbExternaSp":
                           SincroDbExternaSp p1, P2, P3, P4
                           ExecutaComandaSql "Update FeinesAFer Set Param2 = getdate()  Where Id = '" & IdT & "' "
                        Case "SincroDbExternaBdp":
                           SincroDbExternaBdp p1, P2, P3, P4
                           ExecutaComandaSql "Update FeinesAFer Set Param2 = getdate()  Where Id = '" & IdT & "' "
                         Case "ExcelEmail":
                            GuardaHistoric Cnf.llicencia, Now, "Peticio->" & Tipus, p1, P2, P3, P4
                            pt1 = Car(p1)
                            pt2 = Car(p1)
                            If Not InStr(pt2, "@") Then
                                Set rsEmail = Db.OpenResultset("select email from hit.dbo.secretaria where Empresa='" & LastDatabase & "' and Usuario='" & pt2 & "'")
                                If Not rsEmail.EOF Then pt2 = rsEmail("email")
                            End If
                            pt3 = Car(p1)
                            p1 = pt1
                            ExecutaComandaSql "Update FeinesAFer Set Param1 = '[" & pt1 & "][" & pt2 & "][Calculando ... " & Now & "]'  Where Id = '" & iD & "' "
                            EnviaEmailAdjunto pt2, "Calcul Excel " & p1, CalculaExcel(p1, P2, P3, P4, P5)
                        Case "SincronitzaTangram":
                            SincronitzaTangram p1, P2, P3, P4
                        Case "SincronitzaFornsEnrich":
                            SincronitzaFornsEnrich
                            ExecutaComandaSql "Delete FeinesAFer Where Tipus = 'SincronitzaFornsEnrich' "
                            ExecutaComandaSql "Insert Into FeinesAFer (Tipus,Ciclica,Param1,Param2,Param3,Param4) Values ('SincronitzaFornsEnrich',0,'[" & DateAdd("n", 10, Now) & "]','Si " & Now() & "','','') "
                        Case "SincronitzaComandaFornsEnrich":
                            SincronitzaComandaFornsEnrich
                            ExecutaComandaSql "Delete FeinesAFer Where Tipus = 'SincronitzaComandaFornsEnrich' "
                            p1 = Day(DateAdd("d", 1, Now)) & "/" & Month(DateAdd("d", 1, Now)) & "/" & Year(DateAdd("d", 1, Now)) & " 14:00:00"
                            ExecutaComandaSql "Insert Into FeinesAFer (Tipus,Ciclica,Param1,Param2,Param3,Param4) Values ('SincronitzaComandaFornsEnrich',0,'[" & p1 & "]','Si " & Now() & "','','') "
                        Case "SincronitzaComandaXDiaFornsEnrich":
                            SincronitzaComandaXDiaFornsEnrich p1
                        Case "SUBCUENTAS":
                            ExportaMURANO "SUBCUENTAS", p1, P2, P3, P4, P5, iD
                        Case "SincronitzaCaixesMURANO":
                            GuardaHistoric Cnf.llicencia, Now, Tipus, p1, P2, P3, P4
                            D = Car(p1)
                            SincronitzaCaixesMURANO DateAdd("d", -1, D), iD 'Se exporta lo del dia anterior
                            ExecutaComandaSql "Delete FeinesAFer Where Tipus = 'SincronitzaCaixesMURANO' "
                            ExecutaComandaSql "Insert Into FeinesAFer (Tipus,Ciclica,Param1,Param2,Param3,Param4) Values ('SincronitzaCaixesMURANO',0,'[" & Format(DateAdd("d", 1, D), "dd-mm-yy") & "]','Si " & Now() & "','','') "
                        Case "CalculaDadesFichador":
                            CalculaDadesFichador
                            ExecutaComandaSql "Delete FeinesAFer Where Tipus = 'CalculaDadesFichador' "
                            p1 = Day(DateAdd("d", 1, Now)) & "/" & Month(DateAdd("d", 1, Now)) & "/" & Year(DateAdd("d", 1, Now)) & " 22:00:00"
                            ExecutaComandaSql "Insert Into FeinesAFer (Tipus,Ciclica,Param1,Param2,Param3,Param4) Values ('CalculaDadesFichador',0,'[" & p1 & "]','Si " & Now() & "','','') "
                        Case "PrevisionsVendesSetmanal":
                            PrevisionsVendesSetmanal
                            ExecutaComandaSql "Delete FeinesAFer Where Tipus = 'PrevisionsVendesSetmanal' "
                            ExecutaComandaSql "Insert Into FeinesAFer (Tipus,Ciclica,Param1,Param2,Param3,Param4) Values ('PrevisionsVendesSetmanal',0,'[" & DateAdd("d", 7, Now) & "]','Si " & Now() & "','','') "
                        Case "OptimizaSerieOracul"
                            OptimizaSerieOracul
                            p1 = Day(DateAdd("d", 1, Now)) & "/" & Month(DateAdd("d", 1, Now)) & "/" & Year(DateAdd("d", 1, Now)) & " 03:00:00"
                            ExecutaComandaSql "Delete FeinesAFer Where Tipus = 'OptimizaSerieOracul' "
                            ExecutaComandaSql "Insert Into FeinesAFer (Tipus,Ciclica,Param1,Param2,Param3,Param4) Values ('OptimizaSerieOracul',0,'[" & p1 & "]','Si " & Now() & "','','') "
                        Case "SincroANALITICA_SEMANAL"
                            ExportaANALITICA p1, P2, P3, iD
                    End Select
                    ExecutaComandaSql "Delete FeinesAFer Where Id = '" & IdT & "' And Ciclica = 0"
'                End If
            End If
        End If
    End If
    
    ExecutaComandaSql "Delete " & TaulaCalculsEspecials & " Where Id = '" & iD & "' "
    ExecutaComandaSql "Delete hit.dbo.gosdetura Where NomObella = '" & FeinaAfer & "' "
    If CalPlegar Then End

End Sub

Sub FesLaFeina()
    Dim i As Integer, Nexti As Integer
                  
    AvisaAlGos "Inici De Feina "
    FesElConnect
    CarregaLlistaDeFeines feina
    If FeinaAfer = "Tot" Or FeinaAfer = "Sincro" Or FeinaAfer = "Envia" Or FeinaAfer = "Reb" Then
        For i = 1 To UBound(feina)
            If Nexti > 0 Then i = Nexti
            Nexti = 0
            If feina(i).Tipus = 1 Or feina(i).Tipus = 2 Or feina(i).Tipus = 4 Then
                AppPath = App.Path
                If Len(feina(i).Path) > 0 Then AppPath = App.Path & feina(i).Path
                    InformaEmpresa feina(i).empresa, False
                    Debug.Print feina(i).empresa
                    Informa "Testejant llicencia per : " & feina(i).empresa, False
                    If Connecta(feina(i).Tipus, feina(i).llicencia, feina(i).Server, feina(i).Db, i) Then
                        Informa feina(i).empresa, False
                        EmpresaActual = feina(i).empresa
                        ServerActual = feina(i).Server
                        My_DoEvents
                        If feina(i).Tipus = 1 Then SincronitzaEmpresa feina(i).empresa, FeinaAfer
                        ExecutaComandaSql "Delete paramsTPV where (variable like 'DosNivells%' or variable like 'lliure%' or variable like 'Desco%' or variable like 'Capselera%' or variable like 'Lliure%' ) and  valor = '0' "
                        My_DoEvents
                       'If Feina(i).Tipus = 4 Then SincronitzaExterns Feina(i).Empresa, FeinaAfer
                        Nexti = CarregaEnchufat()
                        Db.Close
                    End If
                TancaDb
                My_DoEvents
            End If
        Next
    End If
         
    If FeinaAfer = "Tot" Or FeinaAfer = "Calculs" Or FeinaAfer = "CalculsLlargs" Or FeinaAfer = "CalculsResi" Or FeinaAfer = "CalculsCurts" Or FeinaAfer = "Emails" Or FeinaAfer = "SFtp" Then
        For i = 1 To UBound(feina)
            If Nexti > 0 Then i = Nexti
            Nexti = 0
            If feina(i).Tipus = 5 And FeinaAfer = "CalculsLlargs" Then
                EmpresaActual = feina(i).empresa
                ServerActual = feina(i).Server
                If Connecta(feina(i).Tipus, feina(i).llicencia, feina(i).Server, feina(i).Db, 1) Then
                    EmpresaActual = feina(i).empresa
                    Informa feina(i).empresa
                    CalculsWebKiosk
                    'CalculsEdi
                End If
            End If
            If feina(i).Tipus = 3 And Not FeinaAfer = "CalculsResi" Then
                AppPath = App.Path
                If Len(feina(i).Path) > 0 Then AppPath = App.Path & feina(i).Path
                InformaEmpresa feina(i).empresa
                EmpresaActual = feina(i).empresa
                ServerActual = feina(i).Server
                'Debug.Print Feina(i).Empresa
'            If UCase(Feina(i).Empresa) = UCase("daunis") Then
                If Connecta(feina(i).Tipus, feina(i).llicencia, feina(i).Server, feina(i).Db, i) Then
                    EmpresaActual = feina(i).empresa
                    Informa feina(i).empresa
                    'If UCase(Feina(i).Empresa) = UCase("capdelavila") Then CarregaClientsXls "c:\CarregaClients.xls"
                    RealitzaCalculs feina(i).empresa, ""
'                    Db.Close
                    If FeinesAFerUnCopCadaHora And FeinaAfer = "CalculsCurts" And UCase(feina(i).empresa) = UCase("integraciones") Then
                        FeinaAfer = "Envia"
                        If Connecta(1, feina(i).llicencia, feina(i).Server, feina(i).Db, i) Then SincronitzaEmpresa feina(i).empresa, "Envia"
                        FeinaAfer = "CalculsCurts"
                    End If
'                End If
                Nexti = CarregaEnchufat
                TancaDb
                End If
                My_DoEvents
            End If

        Next
    End If
    
    frmSplash.reord = Format(Now, "hh:mm") & " - " & DateDiff("n", UltimaAccio, Now) & " n."
    FeinesAFerUnCopCadaHora = False
    If Hour(Now) <> UltimaHoraFeinaHorariaFeta Then
        FeinesAFerUnCopCadaHora = True
        UltimaHoraFeinaHorariaFeta = Hour(Now)
        
'        If FeinaAfer = "CalculsLlargs" Then
'           InformaMiss "Feines Horaries Globals "
'           'FrmMain2.RebEmail "SecreHit@gmail.com", "secrehit2130"
'           'TradueixTot
'        End If
    End If
    DoEvents
    
    If InStr(UCase(FeinaAfer), "IDF:") > 0 Then ExecutaCalculPuntual FeinaAfer
    AvisaAlGos "Fi De Feina"
    

End Sub

Sub AvisaAlGos(Estat)
    Dim Rs As rdoResultset
On Error GoTo nor
    Set Rs = Db.OpenResultset("Update Hit.Dbo.GosDeTura Set TsVista = getdate(),UltimaFraseObella = '" & Estat & "' Where NomObella = '" & FeinaAfer & "'")
    Exit Sub
nor:
    AvisaGosPaso2 (Estat)
End Sub





Sub SincronitzaPda()
    Dim Files() As String, i, gCurrMsg, nom As String
    
 Exit Sub
    If Not UCase(EmpresaActual) = UCase("Barcos") Then Exit Sub
    
    InformaMiss "Sincro PDA"
    'PdaEnviaTot = True
    ReDim Files(0)
    If PdaEnviaTot Then GemeraConfiguracioPda Files

    If frmSplash.IpConexio.ConnectaFtpUn("217.116.0.172", "ftp2.sicom397", "barcos1234", False) Then
        For i = 1 To UBound(Files)
            frmSplash.IpConexio.CarregaFileFtp Files(i), AppPath & "\" & Files(i)
        Next
        frmSplash.IpConexio.CarregaDirectoriFtp "/"
        For gCurrMsg = 1 To UBound(DirectoriDesti_Tot)
            nom = DirectoriDesti_NomFile(gCurrMsg)
            If Not Left(nom, 3) = "DAT" Then
                If frmSplash.IpConexio.DescarregaFileFtp(nom, Cnf.AppPath & "\" & nom, 2) Then
                    frmSplash.IpConexio.FtpDeleteFile "", nom
                End If
           End If
        Next
        frmSplash.IpConexio.Desconecta
    End If
    
    For i = 1 To UBound(Files)
        MyKill AppPath & "\" & Files(i)
    Next
    
    Interpreta_SqlTrans frmSplash.Estat
    
    CarregaDir Files, AppPath & "\*.Zip"
    For i = 1 To UBound(Files)
       FitcherProcesat Files(i)
    Next
    CarregaDir Files, AppPath & "\*.Des"
    For i = 1 To UBound(Files)
       FitcherProcesat Files(i)
    Next
    
    
End Sub

Sub CalculaExcelWk(p1 As String, P2 As String, P3 As String, P4 As String, P5 As String)
   Dim MsExcel As Excel.Application, iD As String, Rs As ADODB.Recordset, i As Integer
   Dim Libro As Excel.Workbook, a As New Stream, s() As Byte
    Dim sql As String
    
    On Error Resume Next
        db2.Close
    On Error GoTo noRRRR
    Set db2 = New ADODB.Connection
    db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"

On Error GoTo nok

  InformaMiss "Calculs Excel"
  Set MsExcel = CreateObject("Excel.Application")
  Set Libro = MsExcel.Workbooks.Add
  MsExcel.Visible = frmSplash.Debugant
On Error GoTo 0
    sql = "select s.conductor cod,c.nombre nom,s.fecha,s.resultado,s.codigo from wkservicio s "
    sql = sql & "left join wkconductor c on c.codigo = s.conductor where s.codigo > -1 "
'    Sql = Sql & " and  s.conductor in ('" & P1 & "') "
    sql = sql & " and s.fecha between '" & P2 & "' and '" & P3 & "' + "
    sql = sql & "convert(datetime,'23:59:59',8) "
    sql = sql & " order by nom,fecha desc"
    
    
    Libro.Sheets(1).Name = "Exportacio "
'    Libro.Sheets(1).Range(Libro.Sheets(1).Cells(1, 1), Libro.Sheets(1).Cells(UBound(data, 1), 7)).Value = data
    Set Rs = rec(sql)
    i = 1
    While Not Rs.EOF
        Libro.Sheets(1).Cells(i, 1).Value = Rs(0)
        Libro.Sheets(1).Cells(i, 2).Value = Rs(1)
        Libro.Sheets(1).Cells(i, 3).Value = Rs(2)
        Libro.Sheets(1).Cells(i, 4).Value = Rs(3)
        
        i = i + 1
        Rs.MoveNext
    Wend
    
  Set Rs = rec("select newid() i")
  iD = Rs("i")
    Libro.SaveAs "c:\" & iD & ".xls"
    Libro.Close
  
    Set Libro = Nothing
    Set MsExcel = Nothing
  
    a.Open
    a.LoadFromFile "c:\" & iD & ".xls"
    s = a.ReadText()
  
    Set Rs = rec("select * from " & tablaArchivo() & " where fecha=(select max(fecha) from " & tablaArchivo() & ")", True)
    Rs.AddNew
    Rs("id").Value = iD
  
    Rs("nombre").Value = Left("Anual " & Month(Now) & " " & Year(Now), 20)
    Rs("descripcion").Value = Left(Now, 250)
    Rs("extension").Value = "XLS"
    Rs("mime").Value = "application/vnd.ms-excel"
    Rs("propietario").Value = P3
    Rs("archivo").Value = s
    Rs("fecha").Value = Now
    Rs("tmp").Value = 0
    Rs("down").Value = 1
    Rs.Update
    Rs.Close
    a.Close
  
    
nok:
noRRRR:
  TancaExcel MsExcel, Libro
   

End Sub

Sub BlatPaCreacodis()
   Dim Rs As rdoResultset, i As Integer, K As Integer, Rs2 As rdoResultset, s As String, Ss As String, c As String
   
    ExecutaComandaSql "delete Con "
    Set Rs = Db.OpenResultset("select codistr,nom from blatpapreus")
    While Not Rs.EOF
       s = Rs("Nom")
       K = 0
       c = "a"
       Ss = ""
       For i = 1 To Len(s)
          If c <> " " And Mid(s, i, 1) = " " Then Ss = Ss & "%"
          c = Mid(s, i, 1)
          If c <> " " Then Ss = Ss & c
       Next
       Set Rs2 = Db.OpenResultset("select codi from articles where nom like '" & Ss & "'")
       If Not Rs2.EOF Then ExecutaComandaSql "Insert into Con (CodiStr,Codi) Values ('" & Rs("codistr") & "'," & Rs2("codi") & ") "
       Rs2.Close
       Rs.MoveNext
    Wend
    Rs.Close
End Sub

Sub CalculsWebKiosk()
    Dim Rs As rdoResultset, p1 As String, P2 As String, P3 As String, P4 As String, P5 As String
    
    Set Rs = Db.OpenResultset("Select * From FeinesAFer where Tipus = 'Exportaexcel' ")
    
    While Not Rs.EOF
        p1 = Rs("Param1")
        P2 = Rs("Param2")
        P3 = Rs("Param3")
'        P4 = Rs("Param4")
'        P5 = Rs("Param5")
        CalculaExcelWk p1, P2, P3, P4, P5
        Rs.MoveNext
    Wend
    Set Rs = Db.OpenResultset("delete FeinesAFer where Tipus = 'Exportaexcel' ")
    
End Sub

Sub CalculsEdi()
   Dim Files() As String, i, sql As String, Qi, f, K, a, Ff, Idf As String, Rs, P, PathRemot

On Error Resume Next
   'If Not TePing("5.82.233.79") Then Exit Sub
   PathRemot = "\\5.82.233.79\f$\Users Shared Folders\Alicia\Edi\"

   MkDir PathRemot & "\Bak"
On Error GoTo norrr:
    If Not ExisteixTaula("FilesEdi") Then
      sql = "CREATE TABLE FilesEdi ( "
      sql = sql & " [Id]     [nvarchar] (255) NULL ,"
      sql = sql & " [NomFile]        [nvarchar] (255) NULL ,"
      sql = sql & " [TimeStamp]      [datetime] NULL ,"
      sql = sql & " [LinOriginal]    [nvarchar] (255) NULL ,"
      sql = sql & " [LinModificada]  [nvarchar] (255) NULL ,"
      sql = sql & " [NumLinea]       [float] NULL)"
      ExecutaComandaSql sql
    End If
    Set Qi = Db.CreateQuery("", "Insert into [FilesEdi] (Id,NomFile,TimeStamp,LinOriginal,LinModificada,NumLinea) Values (?,?,?,?,?,?) ")
    
    CarregaDir Files, PathRemot & "\*.*"
    For i = 1 To UBound(Files)
        Set Rs = Db.OpenResultset("select newid() as a")
        Idf = Rs("a")
        Qi(0) = Idf
        Qi(1) = Files(i)
        Qi(2) = Now
        
        f = FreeFile
        Open PathRemot & "\" & Files(i) For Input As #f
        K = 0
        While Not EOF(f)
           Line Input #f, a
           Qi(3) = a
           K = K + 1
           P = InStr(a, "|PCE")
           If P > 0 Then a = Left(a, P - 1)
'           P = InStr(a, "VAT|8")
'           If P > 0 Then Mid(a, P + 4, 1) = "7"
           
           Qi(4) = a
           Qi(5) = K
           Qi.Execute
           DoEvents
       Wend
       Close f
       DoEvents
       InformaMiss "Edi:" & Files(i)
       Name PathRemot & "\" & Files(i) As PathRemot & "\Bak\[Processat#" & Format(Now, "yyyymmddhhnnss") & "]" & Files(i)
       DoEvents
       CalculsEdiPut Idf
       DoEvents
    Next

norrr:

End Sub


Function CarregaEnchufat() As Integer
   Dim K As Integer, Instruccio As String, Rs4 As rdoResultset
   
On Error Resume Next
   CarregaEnchufat = -1
   
   If Not frmSplash.Enchufat.ListIndex = 0 Then
      For K = 1 To UBound(feina)
         If frmSplash.Enchufat.List(frmSplash.Enchufat.ListIndex) = feina(K).empresa Then
            If FeinaAfer = "Tot" Or FeinaAfer = "Sincro" Or FeinaAfer = "Envia" Or FeinaAfer = "Reb" Then
               If feina(K).Tipus = 1 Or feina(K).Tipus = 2 Or feina(K).Tipus = 4 Then CarregaEnchufat = K
            End If
            If FeinaAfer = "Tot" Or FeinaAfer = "Calculs" Or FeinaAfer = "CalculsLlargs" Or FeinaAfer = "CalculsCurts" Or FeinaAfer = "CalculsResi" Then
               If feina(K).Tipus = 3 Then CarregaEnchufat = K
            End If
         End If
      Next
      frmSplash.Enchufat.ListIndex = 0
   End If
   
   Set Rs4 = Db.OpenResultset("Select instruccioPerObella From  hit.dbo.GosDeTura Where nomObella = '" & FeinaAfer & "' ")
   If Not Rs4.EOF Then
        If Not IsNull(Rs4("instruccioPerObella")) Then
            Instruccio = Rs4("instruccioPerObella")
            
            Db.Execute "Update hit.dbo.GosDeTura Set instruccioPerObella = null Where nomObella = '" & FeinaAfer & "'  "
            
            If Instruccio = "STOP" Then TheEnd
            If InStr(Instruccio, "Cuela:") > 0 Then
                Instruccio = Split(Instruccio, ":")(1)
                For K = 1 To UBound(feina)
                    If Instruccio = feina(K).empresa Then
                        If FeinaAfer = "Sincro" And (feina(K).Tipus = 1 Or feina(K).Tipus = 2 Or feina(K).Tipus = 4) Then CarregaEnchufat = K
                        If FeinaAfer = "CalculsLlargs" And feina(K).Tipus = 3 Then CarregaEnchufat = K
                        If FeinaAfer = "CalculsCurts" And feina(K).Tipus = 3 Then CarregaEnchufat = K
                    End If
                 Next
            End If
        End If
   End If
   
End Function

Function Fixe(st As String, Tamany As Integer, Optional AlineatDret As Boolean = False) As String
    
    If AlineatDret Then
        Fixe = Left(st & Space(Tamany), Tamany)
    Else
        Fixe = Right(Space(Tamany) & st, Tamany)
    End If
    
    
End Function

Sub GemeraConfiguracioPda(Files() As String)
    Dim Rs As rdoResultset, f, Zp As New cZip, Str, i As Integer, FilesTmp() As String, UnidadVenta As String, sql As String, FilesTmpK As Integer
    
    FilesTmpK = 1
    ReDim FilesTmp(FilesTmpK)
    f = FreeFile
        
    ExecutaComandaSql "drop table TmpOrdre<r "
    ExecutaComandaSql "SELECT distinct IDENTITY(int, 1,1) AS ID_Num ,f.pare into TmpOrdreFamiliar from families f join articles a on a.familia = f.nom order by f.pare "
   
    ReDim Preserve FilesTmp(FilesTmpK)
    FilesTmp(FilesTmpK) = AppPath & "\FAM01001.DAT"
    MyKill FilesTmp(FilesTmpK)
    Open FilesTmp(FilesTmpK) For Append As f
    
    Set Rs = Db.OpenResultset("Select distinct ID_Num Codigo,Pare Nombre from TmpOrdreFamiliar order by Pare ")
    While Not Rs.EOF
        Str = ""
        Str = Str & Fixe(Rs("Codigo"), 7)
        Str = Str & Fixe(Rs("Nombre"), 30, True)
        Print #f, Str & Chr(13) & Chr(10);
        Rs.MoveNext
    Wend
    Close f
    
    FilesTmpK = FilesTmpK + 1
    ReDim Preserve FilesTmp(FilesTmpK)
    FilesTmp(FilesTmpK) = AppPath & "\ART01001.DAT"
    MyKill FilesTmp(FilesTmpK)
    Open FilesTmp(FilesTmpK) For Append As f
    
    sql = ""
    sql = sql & "Select isnull(cast(p.Valor as nvarchar(255)),0)  Codigo,"
    sql = sql & "Codi ,"
    sql = sql & "isnull(ID_Num,1) CodFamilia,"
    sql = sql & "a.nom Nombre,"
    sql = sql & "preuMajor Tarifa1,"
    sql = sql & "0 Tarifa2,"
    sql = sql & "0 Tarifa3,"
    sql = sql & "0 Tarifa4,"
    sql = sql & "0 Tarifa5,"
    sql = sql & "0 Tarifa6,"
    sql = sql & "0 Tarifa7,"
    sql = sql & "0 Tarifa8,"
    sql = sql & "0 Tarifa9,"
    sql = sql & "0 Tarifa10,"
    sql = sql & "0 PrecioCoste,"
    sql = sql & "TipoIva CodTipoIva,"
    sql = sql & "'0' Lote1,"
    sql = sql & "'0' Lote2,"
    sql = sql & "'M' ModPrecio,"
    sql = sql & "'M' ModDesc1,"
    sql = sql & "'M' ModDesc2,"
    sql = sql & "'M' ModDesc3,"
    sql = sql & "0 Envase,"
    sql = sql & "0 Trazabilidad "
    sql = sql & "from Articles A "
    sql = sql & "join ArticlesPropietats P  on P.CodiArticle  = a.codi and P.Variable  = 'CODI_PROD'  "
    sql = sql & " left Join Families F1 on a.familia = f1.nom left join TmpOrdreFamiliar f on f.Pare = f1.pare "
    sql = sql & "where not isnull(P.valor,'')=''  "
    sql = sql & "Order By codi  "
    
    Set Rs = Db.OpenResultset(sql)
    While Not Rs.EOF
        Str = ""
        Str = Str & Fixe(Rs("Codigo"), 18)
        Str = Str & Fixe(Rs("CodFamilia"), 7)
        Str = Str & Fixe(UCase(Rs("Nombre")), 30, True)
        Str = Str & Fixe(Rs("Tarifa1"), 7)
        Str = Str & Fixe(Rs("Tarifa2"), 7)
        Str = Str & Fixe(Rs("Tarifa3"), 7)
        Str = Str & Fixe(Rs("Tarifa4"), 7)
        Str = Str & Fixe(Rs("Tarifa5"), 7)
        Str = Str & Fixe(Rs("Tarifa6"), 7)
        Str = Str & Fixe(Rs("Tarifa7"), 7)
        Str = Str & Fixe(Rs("Tarifa8"), 7)
        Str = Str & Fixe(Rs("Tarifa9"), 7)
        Str = Str & Fixe(Rs("Tarifa10"), 7)
        Str = Str & Fixe(Rs("PrecioCoste"), 7)
        Str = Str & Fixe(Rs("CodTipoIva"), 1)
        Str = Str & Fixe(Rs("Lote1"), 6)
        Str = Str & Fixe(Rs("Lote2"), 6)
        Str = Str & Fixe(Rs("ModPrecio"), 1)
        Str = Str & Fixe(Rs("ModDesc1"), 1)
        Str = Str & Fixe(Rs("ModDesc2"), 1)
        Str = Str & Fixe(Rs("ModDesc3"), 1)
        UnidadVenta = ArticlesPropietats(Rs("Codi"), "Unidades", 1)
       
        If UnidadVenta = "" Or UnidadVenta = "1" Then
            Str = Str & Fixe(2, 1)
            Str = Str & Fixe(1, 18)
        Else
            Str = Str & Fixe(1, 1)
            Str = Str & Fixe(UnidadVenta, 18)
        End If
        Str = Str & Fixe(Rs("Trazabilidad"), 1)
        
        Print #f, Str & Chr(13) & Chr(10);
        
        Rs.MoveNext
    Wend
    Close f
    
    FilesTmpK = FilesTmpK + 1
    ReDim Preserve FilesTmp(FilesTmpK)
    FilesTmp(FilesTmpK) = AppPath & "\CLI01001.DAT"
    MyKill FilesTmp(FilesTmpK)
    Open FilesTmp(FilesTmpK) For Append As f
    
    Str = ""
    Str = Str & "select "
    Str = Str & "c.codi Codigo,"
    Str = Str & "Nom NombreComercial,"
    Str = Str & "[Nom Llarg] RazonSocial,"
    Str = Str & "adresa Direccion,"
    Str = Str & "ciutat + ' ' + Cp Poblacion,"
    Str = Str & "nif Cif,"
    Str = Str & "isnull(cTel.Valor,'')  Telefonos,"
    Str = Str & "isnull(cOnt.Valor,'')  Contacto,"
    Str = Str & "[Tipus Iva] -1  Impuestos,"
    Str = Str & "0 ImpuestosEspeciales,"
    Str = Str & "[Desconte 5] Tarifa,"
    Str = Str & "case AlbaraValorat when 0 then 7 else '0' end ValoracionNota,"
    Str = Str & "[Desconte 1] Desc1,"
    Str = Str & "[Desconte 2] Desc2,"
    Str = Str & "[Desconte 3] Desc3,"
    Str = Str & "isnull(cGru.Valor,'')   Grupo,"
    Str = Str & "0 Riesgo,"
    Str = Str & "0 TipoRiesgo,"
    Str = Str & "0 CodTipoNota,"
    Str = Str & "0 CTipoNota,"
    Str = Str & "0 CodFormaPago,"
    Str = Str & "'D' CformaPago,"
    Str = Str & "'D' Exclusividad,"
    Str = Str & "'D' ModDesc1,"
    Str = Str & "'D' ModDesc2,"
    Str = Str & "'D' ModDesc3,"
    Str = Str & "'D' Alternativo,"
    Str = Str & "'' SuProveedor "
    Str = Str & "From Clients c left join ConstantsClient cTel on c.codi = cTel.codi and cTel.variable ='Tel' left join ConstantsClient cOnt on c.codi = cOnt.codi and cOnt.variable ='P_Contacte'  left join ConstantsClient cGru on c.codi = cGru.codi and cGru.variable ='Grup_client'  Order by c.codi  "
    
    Set Rs = Db.OpenResultset(Str)
    While Not Rs.EOF
        Str = ""
        Str = Str & Fixe(Rs("Codigo"), 10)
        Str = Str & Fixe(Rs("NombreComercial"), 30, True)
        Str = Str & Fixe(Rs("RazonSocial"), 30, True)
        Str = Str & Fixe(Rs("Direccion"), 30, True)
        Str = Str & Fixe(Rs("Poblacion"), 30, True)
        Str = Str & Fixe(Rs("Cif"), 14, True)
        Str = Str & Fixe(Rs("Telefonos"), 20, True)
        Str = Str & Fixe(Rs("Contacto"), 20, True)
        Str = Str & Fixe(Rs("Impuestos"), 1)
        If Rs("ImpuestosEspeciales") >= 10 Then
            Str = Str & Fixe(0, 1)
        Else
            Str = Str & Fixe(Rs("ImpuestosEspeciales"), 1)
        End If
        Str = Str & Fixe(Rs("Tarifa"), 2)
        Str = Str & Fixe(Rs("ValoracionNota"), 1)
        If Rs("Desc1") = 100 Then
            Str = Str & Fixe(99.99, 5)
        Else
            Str = Str & Fixe(Rs("Desc1"), 5)
        End If
        If Rs("Desc2") = 100 Then
            Str = Str & Fixe(99.99, 5)
        Else
            Str = Str & Fixe(Rs("Desc2"), 5)
        End If
        If Rs("Desc3") = 100 Then
            Str = Str & Fixe(99.99, 5)
        Else
            Str = Str & Fixe(Rs("Desc3"), 5)
        End If
        Str = Str & Fixe(Rs("Grupo"), 10)
        Str = Str & Fixe(Rs("Riesgo"), 8)
        Str = Str & Fixe(Rs("TipoRiesgo"), 1)
        Str = Str & Fixe(Rs("CodTipoNota"), 1)
        Str = Str & Fixe(Rs("CTipoNota"), 1)
        Str = Str & Fixe(Rs("CodFormaPago"), 1)
        Str = Str & Fixe(Rs("CformaPago"), 1)
        Str = Str & Fixe(Rs("Exclusividad"), 1)
        Str = Str & Fixe(Rs("ModDesc1"), 1)
        Str = Str & Fixe(Rs("ModDesc2"), 1)
        Str = Str & Fixe(Rs("ModDesc3"), 1)
        Str = Str & Fixe(Rs("Alternativo"), 1)
        Str = Str & Fixe(Rs("SuProveedor"), 10)
        
        Print #f, Str & Chr(13) & Chr(10);
        
        Rs.MoveNext
    Wend
    Close f
    
'    FilesTmpK = FilesTmpK +1
'    ReDim Preserve FilesTmp(FilesTmpK)
'    FilesTmp(FilesTmpK) = AppPath & "\COD01001.DAT"
'    MyKill FilesTmp(FilesTmpK)
'    Open FilesTmp(FilesTmpK) For Append As f
'    Set Rs = Db.OpenResultset("select CodiArticle,Valor From ArticlesPropietats where variable = 'CODI_PROD' and not valor = ''  Order by CodiArticle ")
'    While Not Rs.EOF
'        Str = ""
'        Str = Str & Fixe(Rs("Valor"), 18)
'        Str = Str & Fixe(Rs("CodiArticle"), 18)
'        Print #f, Str & Chr(13) & Chr(10);
'        Rs.MoveNext
'    Wend
'    Close f
    
    FilesTmpK = FilesTmpK + 1
    ReDim Preserve FilesTmp(FilesTmpK)
    FilesTmp(FilesTmpK) = AppPath & "\PRE01001.DAT"
    MyKill FilesTmp(FilesTmpK)
    Open FilesTmp(FilesTmpK) For Append As f
    Set Rs = Db.OpenResultset("select  c.codi ClientCodi,t.codi ArticleCodi ,t.preumajor Preu From Clients c  join tarifesespecials t on t.tarifacodi=c.[Desconte 5] order by c.codi ,t.codi ")
    While Not Rs.EOF
        Str = ""
        Str = Str & Fixe(Rs("ClientCodi"), 10)
        Str = Str & Fixe(Rs("ArticleCodi"), 18)
        Str = Str & Fixe(Rs("preu"), 7)
        Str = Str & Fixe(0, 7)
        Str = Str & Fixe(0, 5)
        Str = Str & Fixe(0, 5)
        Print #f, Str & Chr(13) & Chr(10);
        Rs.MoveNext
    Wend
    Close f

    Zp.Encrypt = False
    Zp.AddComment = False
    Zp.ZipFile = AppPath & "\DAT01001.ZIP"
    Zp.StoreFolderNames = False
    Zp.RecurseSubDirs = False
    Zp.ClearFileSpecs
    For i = 1 To UBound(FilesTmp)
        Zp.AddFileSpec FilesTmp(i)
    Next
    Zp.Zip
    ReDim Files(1)
    Files(1) = "DAT01001.ZIP"
    
    For i = 1 To UBound(FilesTmp)
        MyKill FilesTmp(i)
    Next
    
End Sub


Sub ConverteixEmpresa()
   Dim D As Date, UltimMes As Integer, UltimAny As Integer, Path As String, empresa As String, EmpresaY As String, Rs As rdoResultset, Recs As Double
      
   Path = "E:\Data\"
   D = Now
   UltimMes = Month(D)
   UltimAny = Year(D)
   empresa = LastDatabase
   Recs = 0
   
   While Year(D) > 1998
      If Year(D) < Year(Now) Then
         EmpresaY = empresa & "_" & Year(D)
         If Not UltimAny = Year(D) Then
            Recs = 0
On Error Resume Next   ' les taules poden no existir i no les volem crear
            Set Rs = Db.OpenResultset("Select Count(*) from [Servit-" & Format(D, "yy-mm-dd") & "] ")
            If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Recs = Rs(0)
            If Recs = 0 Then
               Set Rs = Db.OpenResultset("Select Count(*) from [" & NomTaulaVentas(D) & "] ")
               If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Recs = Rs(0)
            End If
On Error GoTo 0
            If Recs > 0 Then
               ExecutaComandaSql "CREATE DATABASE [" & EmpresaY & "]  ON (NAME = N'" & EmpresaY & "_dat', FILENAME = N'" & Path & "" & EmpresaY & ".mdf' , SIZE = 1, FILEGROWTH = 10%) LOG ON (NAME = N'" & EmpresaY & "_log', FILENAME = N'" & Path & "" & EmpresaY & ".ldf' , SIZE = 1, FILEGROWTH = 10%)"
               UltimAny = Year(D)
            End If
         End If
         
         If Recs > 0 And Not UltimMes = Month(D) Then
            ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[" & NomTaulaAlbarans(D) & "] from [" & NomTaulaAlbarans(D) & "] "
            ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[" & NomTaulaAnulats(D) & "] from [" & NomTaulaAnulats(D) & "] "
            ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[" & NomTaulaDevol(D) & "] from [" & NomTaulaDevol(D) & "] "
            ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[" & NomTaulaFacturaData(D) & "] from [" & NomTaulaFacturaData(D) & "] "
            ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[" & NomTaulaFacturaIva(D) & "] from [" & NomTaulaFacturaIva(D) & "] "
            ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[" & NomTaulaHoraris(D) & "] from [" & NomTaulaHoraris(D) & "] "
            ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[" & NomTaulaInventari(D) & "] from [" & NomTaulaInventari(D) & "] "
            ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[" & NomTaulaLog(D) & "] from [" & NomTaulaLog(D) & "] "
            ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[" & NomTaulaMovi(D) & "] from [" & NomTaulaMovi(D) & "] "
            ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[" & NomTaulaVentas(D) & "] from [" & NomTaulaVentas(D) & "] "
            UltimMes = Month(D)
         End If
         
         If Recs > 0 Then ExecutaComandaSql "Select * Into [" & EmpresaY & "].[Dbo].[Servit-" & Format(D, "yy-mm-dd") & "] from [Servit-" & Format(D, "yy-mm-dd") & "] "
      End If
      D = DateAdd("d", -1, D)
      Debug.Print Format(D, "dddd dd-mm-yy")
      DoEvents
   Wend
   
   
End Sub

Sub FuncioAuxiliar()
   Dim D As Date, Rs As rdoResultset, b As Double
   
   D = DateSerial(2005, 1, 1)
   
   While D < Now
      Set Rs = Db.OpenResultset("Select distinct botiga from [" & NomTaulaDevol(D) & " ] ")
      While Not Rs.EOF
         b = Rs(0)
         ActualizaPanSeco D, b
         Rs.MoveNext
Debug.Print b & " --> " & D
      Wend
      Rs.Close
      
      
      D = DateAdd("d", 1, D)
      
      
   Wend
   
   
   
   
End Sub

Sub SincronitzaEmpresaEnvia(empresa As String)
   Dim CalBorrarArticles As Boolean, Files() As String, Rs As rdoResultset, RsF As rdoResultset, Botis() As String, K As Integer, p1 As String, P2 As String, AccionsEnviat As Boolean, i As Integer, Tipus() As String, Param() As String, Contingut As String, j As Integer, LaAgafem As Boolean, CalTrucar As Boolean, nom As String, Interesa As Integer, Esborrem As Integer, EsCfg As Boolean, ClientsGenerats As Boolean, EnviaArticles As Boolean, B_Origen, D_Origen, B_Desti, D_Desti, Grup, sql As String, LlistaNegra() As String, Files2() As String
   Dim llicencia As String
   Dim P As Integer
   Dim sqlPromo As String
   
   llicencia = ""
    
   For i = 1 To UBound(feina)
       If feina(i).empresa = empresa And Not feina(i).EscoltaLlicencies(0) = "" Then Exit Sub
   Next
   
'If UCase(EmpresaActual) = UCase("LaForneria") Then
'EmpresaActual = EmpresaActual
'Missatges_CalEnviar "Resum De Punts", ""
'End If
   
   LlistaNegra = Split(UCase("DiccionariTot,ComandesPlantilles,Santoral,TpvVellsCodis,Dependentes,Memotecnics,Atributs,Promocions,ConstantsClient,Clients,Viatges,Equips,CodisBarres,ArticlesPropietats,PreusArticles,ProductesPromocionats,FamiliesArticles,Punts,Facturacio,FeinaFeta,CaixesCongelador,ClientsFinalsAcumulat,ClientsFinals,Clients,tarifesespecialsclients,ARTICLES"), ",")
   InformaMiss "Configurant", False
   CalTrucar = False
   frmSplash.IpConexio.InteresPerContingutReset
   If ExisteixTaula("InteresaContingut") Then
      Set Rs = Db.OpenResultset("Select distinct * From InteresaContingut ")
      While Not Rs.EOF
         LaAgafem = True
         For i = 0 To UBound(LlistaNegra)
            If UCase(Rs("Nom")) = LlistaNegra(i) Then
                LaAgafem = False
                Exit For
            End If
         Next
         
         If Left(UCase(Rs("Nom")), 14) = "TARIFAESPECIAL" Then LaAgafem = False
         If Left(UCase(Rs("Nom")), 7) = "COMANDA" Then LaAgafem = False
         frmSplash.IpConexio.InteresPerContingut Rs("Nom"), LaAgafem, Rs("LaEsborrem") = 1
         Rs.MoveNext
      Wend
      Rs.Close
   End If
   
   EnviaArticles = True
   If ExisteixTaula("QueTinc") Then
      Set Rs = Db.OpenResultset("Select * From QueTinc Where QueEs='EnviarArticles'  ")
      If Not Rs.EOF Then If Not IsNull(Rs(0)) Then If UCase(Rs("QuinEs")) = UCase("No") Then EnviaArticles = False
      Rs.Close
   End If
  
  
'   If UCase(EmpresaActual) = UCase("Forn Del Passeig") Then
'      Missatges_CalEnviar "DeutesAnticips", "4"
'      Missatges_CalEnviar "Santoral", ""
'   End If
'   Missatges_CalEnviar "Facturacio", ""
'   Missatges_CalEnviar "FeinaFeta", ""
'   Missatges_CalEnviar "ComandesPlantilles", ""
'   Missatges_CalEnviar "CaixesCongelador", ""
   
   Missatges_CalEnviar "Missatges", ""
   If FeinesAFerUnCopCadaHora Then
      InformaMiss "Preparant Comandes", False
      PreparaComandes
   End If
   

   frmSplash.IpConexio.LabelEstat = frmSplash.Estat
   frmSplash.IpConexio.LabelEstatDbg = frmSplash.lblVersion
   
   Set Rs = Db.OpenResultset("Select * From MissatgesAEnviar ")
   ReDim Tipus(0)
   ReDim Param(0)
   While Not Rs.EOF
      ReDim Preserve Tipus(UBound(Tipus) + 1)
      ReDim Preserve Param(UBound(Param) + 1)
      Tipus(UBound(Tipus)) = NoNull(Rs("Tipus"))
      Param(UBound(Param)) = NoNull(Rs("Param"))
      If Tipus(UBound(Tipus)) = "Clients" Then
         ReDim Preserve Tipus(UBound(Tipus) + 1)
         ReDim Preserve Param(UBound(Param) + 1)
         Tipus(UBound(Tipus)) = "ConstantsClient"
         Param(UBound(Param)) = ""
      End If
      If Tipus(UBound(Tipus)) = "Articles" Then
         ReDim Preserve Tipus(UBound(Tipus) + 1)
         ReDim Preserve Param(UBound(Param) + 1)
         Tipus(UBound(Tipus)) = "ArticlesPropietats"
         Param(UBound(Param)) = ""
      End If
      If Tipus(UBound(Tipus)) = "Resum De Punts" Then
         ReDim Preserve Tipus(UBound(Tipus) + 1)
         ReDim Preserve Param(UBound(Param) + 1)
         Tipus(UBound(Tipus)) = "ClientsFinalsAcumulat"
         Param(UBound(Param)) = ""
      End If
      Rs.MoveNext
   Wend
   Rs.Close
   
   If EnviaArticles Then
      Db.Execute "Delete MissatgesAEnviar "
   Else
      Db.Execute "Delete MissatgesAEnviar Where not (Tipus = 'Articles' Or Tipus = 'Tarifa' )"
   End If
   AccionsEnviat = False
   ClientsGenerats = False
   
   For i = 1 To UBound(Tipus)
      ReDim Files(0)
      Contingut = ""
      EsCfg = False
      InformaMiss "Preparant " & Tipus(i), False
      Select Case Tipus(i)
         Case "IntegracionesEnviaFacturacio"
            IntegracionesEnviaFacturacio Files
            If UBound(Files) > 0 Then Contingut = "IntegracionesEnviaFacturacio"
         Case "EnviaDiccionari"
            EsCfg = True
            If Param(i) = "" Then Param(i) = "TOC"
            GemeraSqlTrans "Diccionari", Files, "Select * from hit.dbo.Diccionari Where App ='" & Param(i) & "' "
            If UBound(Files) > 0 Then Contingut = "DiccionariTot"
         Case "IntegracionesEnviaValidacions"
            IntegracionesEnviaValidacions Files
            If UBound(Files) > 0 Then Contingut = "IntegracionesEnviaValidacions"
         Case "EnviarTeclat"  '[Contingut#TeclatToc_00305]PreferenciasTeclat.SqlTrans
            EsCfg = True
            B_Origen = Car(Param(i))
            D_Origen = Car(Param(i))
            D_Desti = Car(Param(i))
            B_Desti = Car(Param(i))
            
            llicencia = B_Desti
            
            GemeraSqlTrans "PreferenciasTeclat", Files, "select [Maquina],[Dependenta],[Ambient],[Article],[Pos],[Color] from teclatstpv  where  llicencia = " & B_Origen & " and data = '" & D_Origen & "' order by ambient,pos "
            If UBound(Files) > 0 Then Contingut = "TeclatToc_" & Format(B_Desti, "00000")
         Case "EnviarTeclatGrup"  '[Contingut#TeclatToc_00305]PreferenciasTeclat.SqlTrans
            EsCfg = True
            B_Origen = Car(Param(i))
            D_Origen = Car(Param(i))
            D_Desti = Car(Param(i))
            B_Desti = Car(Param(i))
            Grup = Car(Param(i))
            
            llicencia = B_Desti
            
            sql = "Select [Maquina],[Dependenta],[Ambient],[Article],[Pos],[Color] from teclatstpv where  "
            sql = sql & " llicencia = " & B_Origen & " and data = '" & D_Origen & "' And Ambient = '" & Grup & "' "
            
            sql = sql & " Union "
            
            sql = sql & "Select [Maquina],[Dependenta],[Ambient],[Article],[Pos],[Color] from teclatstpv where  "
            sql = sql & " llicencia = " & B_Desti & " and data = (Select max(data) from teclatstpv where  llicencia = " & B_Desti & " ) and not ambient = '" & Grup & "' "
            sql = sql & " order by ambient,pos "
            
            GemeraSqlTrans "PreferenciasTeclat", Files, sql
            If UBound(Files) > 0 Then Contingut = "TeclatToc_" & Format(B_Desti, "00000")
         Case "ComandesPlantilles"
            EsCfg = True
            GemeraSqlTrans "ComandesPlantilles", Files
            If UBound(Files) > 0 Then Contingut = "ComandesPlantilles"
         Case "DeutesAnticips", "DeutesABotiga"
            EsCfg = True
            GemeraSqlTrans "DeutesAnticips", Files, "Select * from DeutesAnticips Where estat = 0 and Botiga = " & Param(i)
            If UBound(Files) > 0 Then Contingut = "Tpv_Configuracio_" & Param(i)
         Case "DeutesAnticips", "DeutesABotiga", "EnviaDeutesBotiga"
            EsCfg = False
            GemeraSqlTransDestesBotiga Param(i), Files
            If UBound(Files) > 0 Then Contingut = "Tpv_Configuracio_" & CodiClientLlicencia(Param(i))
         Case "Santoral"
            EsCfg = True
            GemeraSqlTrans "Santoral", Files, "Select * from hit.dbo.Santoral "
            If UBound(Files) > 0 Then Contingut = "Santoral"
         Case "TpvVellsCodis"
            EsCfg = True
            GemeraSqlTransTpvVellsCodis Files
            If UBound(Files) > 0 Then Contingut = "TpvVellsCodis"
         Case "Dependentes"
            EsCfg = True
            GemeraSqlTrans "Dependentes", Files
            If UBound(Files) > 0 Then Contingut = "Dependentes"
            ReDim Files2(0)
            
            GemeraSqlTrans "DependentesExtes", Files2, "select e.* from dependentesextes e join Dependentes d on e.id = d.codi where e.nom = 'CODICFINAL' or  e.nom = 'TIPUSTREBALLADOR' or  e.nom = 'PASSWORD'  or  e.nom = 'CODI_TARGETA' or e.nom = 'CODI_DEP' "
            If UBound(Files2) > 0 Then
                ReDim Preserve Files(UBound(Files) + 1)
                Files(UBound(Files)) = Files2(1)
            End If
         Case "Media"
            EsCfg = True
            GemeraSqlTrans "Select * from Archivo where id = '" & Param(i) & "' ", Files
            If UBound(Files) > 0 Then Contingut = "Media"
         Case "Dedos"
            EsCfg = True
            GemeraSqlTrans "Dedos", Files
            If UBound(Files) > 0 Then Contingut = "Dedos"
         Case "Memotecnics"
            EsCfg = True
            GemeraSqlTrans "Memotecnics", Files
            If UBound(Files) > 0 Then Contingut = "Memotecnics"
         Case "Atributs"
            EsCfg = True
            GemeraSqlTrans "Atributs", Files
            If UBound(Files) > 0 Then Contingut = "Atributs"
         Case "Promocions"
            EsCfg = True
            GemeraSqlTrans "Promocions", Files
            If UBound(Files) > 0 Then Contingut = "Promocions"
         Case "ConstantsClient"
            EsCfg = True
            GemeraSqlTrans "ConstantsClient", Files
            If UBound(Files) > 0 Then Contingut = "ConstantsClient"
         Case "Clients"
            EsCfg = True
            GemeraSqlTrans "Clients", Files
            If UBound(Files) > 0 Then Contingut = "Clients"
            PdaEnviaTot = True
         Case "Viatges"
            EsCfg = True
            GemeraSqlTrans "Viatges", Files
            If UBound(Files) > 0 Then Contingut = "Viatges"
            PdaEnviaTot = True
         Case "Equips"
            EsCfg = True
            GemeraSqlTrans "EquipsDeTreball", Files
            If UBound(Files) > 0 Then Contingut = "Equips"
         Case "CodisBarres"
            EsCfg = True
            GemeraSqlTrans "CodisBarres", Files
            If UBound(Files) > 0 Then Contingut = "CodisBarres"
         Case "ArticlesPropietats"
            If EnviaArticles Then
               EsCfg = True
               GemeraSqlTrans "ArticlesPropietats", Files, "select p.* from ArticlesPropietats p join Articles a on p.CodiArticle = a.codi where p.Variable in ('OBSERVACIONES','DESCRIPCIoN','SUGERENCIAS','NOTAS','SEGONIDIOMA','CODI_PROD','ES_SUPLEMENT','IMPRESORA','NoSumaPunts') and not isnull(p.valor,'') = ''"
               If UBound(Files) > 0 Then Contingut = "ArticlesPropietats"
            End If
         Case "Articles"
            If EnviaArticles Then
               EsCfg = True
               GemeraSqlTrans "Articles", Files
               If UBound(Files) > 0 Then Contingut = "PreusArticles"
            End If
            PdaEnviaTot = True
         Case "Tpv_Configuracio_"
            If Not UCase(EmpresaActual) = UCase("Barcos") Then
               EsCfg = True
               GemeraConfiguracio Param(i), Files
               If UBound(Files) > 0 Then Contingut = "Tpv_Configuracio_" & Param(i)
            End If
        'FICHERO DE TODAS LAS TIENDAS
         Case "ProductesPromocionats"
            EsCfg = True
            
'D:Producto - D:Producto
            sqlPromo = "select pp.id, pp.Di, pp.Df,pp.D_producte as D_producte, pp.d_quantitat, cast(pp.s_producte as nvarchar) as s_producte,Pp.s_quantitat , Pp.s_preu, Pp.client "
            sqlPromo = sqlPromo & "from productesPromocionats pp "
            sqlPromo = sqlPromo & "where pp.d_producte not like 'F_%' and pp.s_producte not like 'F_%' "
            sqlPromo = sqlPromo & "Union "
'D:Producto - S:Familia
            sqlPromo = sqlPromo & "select 'F_'+pp.id, pp.Di, pp.Df, pp.d_producte, pp.d_quantitat, cast(a.codi as nvarchar) as s_Producte, pp.s_quantitat , s_preu, client "
            sqlPromo = sqlPromo & "from productesPromocionats pp "
            sqlPromo = sqlPromo & "right join articles a on SUBSTRING(pp.s_producte,3,len(pp.s_producte)-1) = a.familia "
            sqlPromo = sqlPromo & "where pp.d_producte not like 'F_%' and pp.s_producte like 'F_%' "
            sqlPromo = sqlPromo & "Union "
'D:Familia - S:Producto
            sqlPromo = sqlPromo & "select 'F_'+pp.id, pp.Di, pp.Df, a.codi as d_producte, pp.d_quantitat, cast(pp.s_Producte as nvarchar) as s_Producte, pp.s_quantitat , s_preu, client "
            sqlPromo = sqlPromo & "from productesPromocionats pp "
            sqlPromo = sqlPromo & "right join articles a on SUBSTRING(pp.d_producte,3,len(pp.d_producte)-1) = a.familia "
            sqlPromo = sqlPromo & "where pp.d_producte like 'F_%' and pp.s_producte not like 'F_%' "
            sqlPromo = sqlPromo & "Union "
'D:Familia - S:Familia
            sqlPromo = sqlPromo & "select 'F_'+pp.id, pp.Di, pp.Df, a_d.codi, pp.d_quantitat, cast(a_s.codi as nvarchar) s_producte, pp.s_quantitat, Pp.s_preu, Pp.client "
            sqlPromo = sqlPromo & "from productesPromocionats pp "
            sqlPromo = sqlPromo & "right join articles a_d on SUBSTRING(pp.d_producte, 3, len(pp.d_producte)-1) = a_d.familia "
            sqlPromo = sqlPromo & "right join articles a_s on SUBSTRING(pp.s_producte, 3, len(pp.s_producte)-1) = a_s.familia "
            sqlPromo = sqlPromo & "where pp.d_producte like 'F_%' and pp.s_producte like 'F_%' "
            sqlPromo = sqlPromo & "order by client"
            
            GemeraSqlTrans "ProductesPromocionats", Files, sqlPromo
            If UBound(Files) > 0 Then Contingut = "ProductesPromocionats"
            
        'FICHERO TIENDA MODIFICADA
        Case "ProductesPromocionatsBotiga"
            EsCfg = True
            
            B_Origen = Car(Param(i))
            llicencia = B_Origen
            
'D:Producto - D:Producto
            sqlPromo = "select pp.id, pp.Di, pp.Df,pp.D_producte as D_producte, pp.d_quantitat, cast(pp.s_producte as nvarchar) as s_producte,Pp.s_quantitat , Pp.s_preu, Pp.client "
            sqlPromo = sqlPromo & "from productesPromocionats pp "
            sqlPromo = sqlPromo & "where pp.d_producte not like 'F_%' and pp.s_producte not like 'F_%' and pp.client = " & B_Origen & " "
            sqlPromo = sqlPromo & "Union "
'D:Producto - S:Familia
            sqlPromo = sqlPromo & "select 'F_'+pp.id, pp.Di, pp.Df, pp.d_producte, pp.d_quantitat, cast(a.codi as nvarchar) as s_Producte, pp.s_quantitat , s_preu, client "
            sqlPromo = sqlPromo & "from productesPromocionats pp "
            sqlPromo = sqlPromo & "right join articles a on SUBSTRING(pp.s_producte,3,len(pp.s_producte)-1) = a.familia "
            sqlPromo = sqlPromo & "where pp.d_producte not like 'F_%' and pp.s_producte like 'F_%' and pp.client = " & B_Origen & " "
            sqlPromo = sqlPromo & "Union "
'D:Familia - S:Producto
            sqlPromo = sqlPromo & "select 'F_'+pp.id, pp.Di, pp.Df, a.codi as d_producte, pp.d_quantitat, cast(pp.s_Producte as nvarchar) as s_Producte, pp.s_quantitat , s_preu, client "
            sqlPromo = sqlPromo & "from productesPromocionats pp "
            sqlPromo = sqlPromo & "right join articles a on SUBSTRING(pp.d_producte,3,len(pp.d_producte)-1) = a.familia "
            sqlPromo = sqlPromo & "where pp.d_producte like 'F_%' and pp.s_producte not like 'F_%' and pp.client = " & B_Origen & " "
            sqlPromo = sqlPromo & "Union "
'D:Familia - S:Familia
            sqlPromo = sqlPromo & "select 'F_'+pp.id, pp.Di, pp.Df, a_d.codi, pp.d_quantitat, cast(a_s.codi as nvarchar) s_producte, pp.s_quantitat, Pp.s_preu, Pp.client "
            sqlPromo = sqlPromo & "from productesPromocionats pp "
            sqlPromo = sqlPromo & "right join articles a_d on SUBSTRING(pp.d_producte, 3, len(pp.d_producte)-1) = a_d.familia "
            sqlPromo = sqlPromo & "right join articles a_s on SUBSTRING(pp.s_producte, 3, len(pp.s_producte)-1) = a_s.familia "
            sqlPromo = sqlPromo & "where pp.d_producte like 'F_%' and pp.s_producte like 'F_%'  and pp.client = " & B_Origen & " "
                        
            GemeraSqlTrans "ProductesPromocionats_" & Format(B_Origen, "00000"), Files, sqlPromo
            If UBound(Files) > 0 Then Contingut = "ProductesPromocionats_" & Format(B_Origen, "00000")
            
         Case "Tpv_Families_"
            EsCfg = True
            GemeraSqlTrans "Families", Files
            If UBound(Files) > 0 Then Contingut = "FamiliesArticles"
         Case "Resum De Punts"
            EsCfg = True
            If Day(Now) = 10 Then ' Un cop al mes l enviem sencer !!
                GemeraSqlTrans "Punts", Files
            Else
                GemeraSqlTrans "Punts", Files, "select * from Punts where DATEDIFF(M,data2,GETDATE ()) < 6"
            End If
            If UBound(Files) > 0 Then Contingut = "Punts"
         Case "Facturacio"
            EnviaFacturacio Files
            If UBound(Files) > 0 Then Contingut = "Facturacio"
         Case "FeinaFeta"
            GemeraSqlTrans "FeinaFeta", Files
            If UBound(Files) > 0 Then Contingut = "FeinaFeta"
         Case "Missatges"
            GemeraSqlTrans "Missatges", Files
            If UBound(Files) > 0 Then Contingut = "Missatges"
         Case "CaixesCongelador"
            GemeraSqlTrans "Creades", Files
            If UBound(Files) > 0 Then Contingut = "CaixesCongelador"
         Case "ClientsFinalsAcumulat"
               EsCfg = True
               GeneraClients Tipus, Param, Files, True
               If UBound(Files) > 0 Then Contingut = "ClientsFinalsAcumulat"
         Case "ClientsFinals"
            If Not ClientsGenerats Then
               GeneraClients Tipus, Param, Files
               ClientsGenerats = True
               If UBound(Files) > 0 Then Contingut = "ClientsFinals"
            End If
'         Case "ClientsFinalsPropietats"
'            EsCfg = True
'            GemeraSqlTrans "ClientsFinalsPropietats", Files, "select * from ClientsFinalsPropietats "
'            If UBound(Files) > 0 Then Contingut = "ClientsFinalsPropietats"
         Case "ClientsFinalsTots"
               GeneraClients Tipus, Param, Files, True
               If UBound(Files) > 0 Then Contingut = "ClientsFinals"
               ClientsGenerats = True
         Case "Versio"
            GemeraSqlTransFiles "Versio", Param(i), Files, Contingut
         Case "reenviadia"
            P = InStr(Param(i), ":")
            If P > 0 Then llicencia = Left(Param(i), P - 1)
            GemeraSqlAccio Param(i), Files, Contingut
         Case "ClientFinal_Esborrat", "Deute_Cambiat"
            If Not AccionsEnviat Then GemeraSqlAccioTots Tipus, Param, Files, Contingut
            AccionsEnviat = True
         Case "Tarifa"
            If EnviaArticles Then
               GemeraSqlTrans "TarifaEspecial", Files, "Select TarifaCodi,TarifaNom,Articles.CodiGenetic as Codi,TarifesEspecials.Preu,TarifesEspecials.PreuMajor from TarifesEspecials join Articles on TarifesEspecials.Codi = Articles.Codi Where TarifaCodi = " & Param(i)
               If UBound(Files) > 0 Then Contingut = "TarifaEspecial_" & Param(i)
               EsCfg = True
            End If
         Case "tarifesespecialsclients"
               GemeraSqlTrans "tarifesespecialsclients", Files, "Select * From tarifesespecialsclients"
               If UBound(Files) > 0 Then
                Contingut = "tarifesespecialsclients"
                EsCfg = True
               End If
         Case "recursosExtes"
               GemeraSqlTrans "recursosextes", Files, , True
               If UBound(Files) > 0 Then
                Contingut = "recursosextes"
                EsCfg = True
               End If
         Case "recursos"
               GemeraSqlTrans "recursos", Files, , True
               If UBound(Files) > 0 Then
                Contingut = "recursos"
                EsCfg = True
               End If
         Case "appccComo"
               GemeraSqlTrans "appccComo", Files, , True
               If UBound(Files) > 0 Then
                Contingut = "appccComo"
                EsCfg = True
               End If
         Case "appccCuando"
               GemeraSqlTrans "appccCuando", Files, , True
               If UBound(Files) > 0 Then
                Contingut = "appccCuando"
                EsCfg = True
               End If
         Case "appccTareas"
               GemeraSqlTrans "appccTareas", Files, , True
               If UBound(Files) > 0 Then
                Contingut = "appccTareas"
                EsCfg = True
               End If
         Case "appccTareasAsignadas"
               GemeraSqlTrans "appccTareasAsignadas", Files, , True
               If UBound(Files) > 0 Then
                Contingut = "appccTareasAsignadas"
                EsCfg = True
               End If
         Case "appccTareasResueltas"
               GemeraSqlTrans "appccTareasResueltas", Files, "Select top 10 * from appccTareasResueltas "
               If UBound(Files) > 0 Then
                Contingut = "appccTareasResueltas"
                EsCfg = True
               End If
         Case "Comandes"
            p1 = Car(Param(i))
            EnviaComandesBotiga Files, p1
            If UBound(Files) > 0 Then
               Contingut = "Comanda_" & p1
            End If
            llicencia = p1
         Case "Imatgesdependentes", "ImatgesDependentes"
'            EsCfg = True  ' Eze 18-10-2018  por volumen de datos no subimos las fotos al ftp
'            ReDim Files(0)
'            'Nomes envia fotos tamany TPV i SCREEN
'            Set RsF = Db.OpenResultset("select e.id as id,Archivo,Nombre,Extension from archivo a join DependentesExtes e on e.valor  = a.id where a.nombre in ('SCREEN','TPV')  and fecha > isnull((select max([TimeStamp]) from records where concepte = 'DependentesImatgesEnviades'),dateadd(yy,-5,getdate())) ")
'            While Not RsF.EOF
'                If Not IsNull(RsF.rdoColumns("Archivo")) Then
'                ReDim Preserve Files(UBound(Files) + 1)
'                Files(UBound(Files)) = AppPath & "\Treb_" & RsF("Id") & "_" & RsF("Nombre") & "." & RsF("Extension")
'                ColumnToFile RsF.rdoColumns("Archivo"), Files(UBound(Files)), 102400, RsF("Archivo").ColumnSize
'                End If
'                RsF.MoveNext
'            Wend
'            RsF.Close
'            If UBound(Files) > 0 Then Contingut = "ImatgesDependentes"
         Case "Imatgesarticles", "ImatgesArticles"
'            EsCfg = True   ' Eze 18-10-2018  por volumen de datos no subimos las fotos al ftp
'            ReDim Files(0)
'            Set RsF = Db.OpenResultset("select timestamp from records where concepte='ArticlesImatgesEnviades' ")
'            If RsF.EOF Then ExecutaComandaSql "INSERT into records values (dateAdd(d,-1,getdate()),'ArticlesImatgesEnviades')"
'            Set RsF = Db.OpenResultset("select e.id as id,Archivo,Nombre,Extension from archivo a join articlesextes e on e.valor  = a.id where not archivo is null and a.nombre in ('SCREEN','TPV') and fecha > isnull((select max([TimeStamp]) from records where concepte = 'ArticlesImatgesEnviades'),dateadd(yy,-5,getdate())) ")
'            While Not RsF.EOF
'                If Not IsNull(RsF.rdoColumns("Archivo")) Then
'                ReDim Preserve Files(UBound(Files) + 1)
'                Files(UBound(Files)) = AppPath & "\Art_" & RsF("Id") & "_" & RsF("Nombre") & "." & RsF("Extension")
'                ColumnToFile RsF.rdoColumns("Archivo"), Files(UBound(Files)), 102400, RsF("Archivo").ColumnSize
'                End If
'                RsF.MoveNext
'            Wend
'            RsF.Close
'            ExecutaComandaSql "update records set timestamp = getdate() where concepte = 'ArticlesImatgesEnviades'"
'            If UBound(Files) > 0 Then Contingut = "ImatgesArticles"
        Case "ImatgesLogos"
            EsCfg = True
            ReDim Files(0)
            Set RsF = Db.OpenResultset("select timestamp from records where concepte='LogosImatgesEnviades' ")
            If RsF.EOF Then ExecutaComandaSql "INSERT into records values (dateAdd(d,-1,getdate()),'LogosImatgesEnviades')"
            Set RsF = Db.OpenResultset("select Archivo,Nombre,Descripcion,Extension,DATALENGTH(archivo) size from archivo a where not archivo is null and a.nombre='LOGO' and (a.descripcion like  '%IMP>' or a.descripcion like '%TPV>') and fecha > isnull((select max([TimeStamp]) from records where concepte = 'LogosImatgesEnviades'),dateadd(yy,-5,getdate())) ")
            While Not RsF.EOF
                If Not IsNull(RsF.rdoColumns("Archivo")) Then
                ReDim Preserve Files(UBound(Files) + 1)
                If Right(RsF("descripcion"), 4) = "IMP>" Then
                    Files(UBound(Files)) = AppPath & "\[PackNom#LogoImpresora][PackNum#1][Size#" & RsF("size") & "][Nom#LogoImpresora,gif][Contingut#File_Toc]VersioTocAsist.sqltrans"
                    'Files(UBound(Files)) = AppPath & "\[PackNom#LogoImpresora" & EmpresaActual & "][PackNum#1][Size#" & RsF("size") & "][Nom#LogoImpresora,gif][Contingut#File_Toc]VersioTocAsist.sqltrans"
                Else
                    Files(UBound(Files)) = AppPath & "\[PackNom#LogoColor][PackNum#1][Size#" & RsF("size") & "][Nom#LogoColor,gif][Contingut#File_Toc]VersioTocAsist.sqltrans"
                    'Files(UBound(Files)) = AppPath & "\[PackNom#LogoColor" & EmpresaActual & "][PackNum#1][Size#" & RsF("size") & "][Nom#LogoColor,gif][Contingut#File_Toc]VersioTocAsist.sqltrans"
                End If
                ColumnToFile RsF.rdoColumns("Archivo"), Files(UBound(Files)), 102400, RsF("Archivo").ColumnSize
                End If
                RsF.MoveNext
            Wend
            RsF.Close
            ExecutaComandaSql "update records set timestamp = getdate() where concepte = 'LogosImatgesEnviades'"
            If UBound(Files) > 0 Then Contingut = "ImatgesLogos"
'         Case "ClientsFinalsTarifa"   ' taula obsoleta la tarifa esta a clientsfinalspropietats
'            EsCfg = True
'            GemeraSqlTrans "ClientsFinalsTarifes", Files
'            If UBound(Files) > 0 Then Contingut = "ClientsFinalsTarifes"
         Case Else
            Contingut = ""
      End Select
      
      If Not SistemaMud And Not Contingut = "" Then
         frmSplash.IpConexio.AddMisatgeEnviar Contingut, Files, EsCfg, llicencia
         CalTrucar = True
      End If
      
      On Error Resume Next
      For j = 1 To UBound(Files)
         FitcherProcesat Files(j), True, True
         Kill Files(j)
      Next
      On Error GoTo 0
      My_DoEvents
   Next
   
   ExecutaComandaSql "Drop table Servit_tmp"
   
   If Not CalTrucar Then
      Dim UnFile As String
      
      
      
      UnFile = Dir(Cnf.AppPath & "\Msg\")
      If Len(UnFile) > 0 Then CalTrucar = True
      
      UnFile = Dir(Cnf.AppPath & "\Msg\Cfg\")
      If Len(UnFile) > 0 Then CalTrucar = True
   End If
   
   If CalTrucar Then
      FtpCopy
      InformaMiss "Connectant Enviar ... "
      frmSplash.IpConexio.CarterUltimaData = BuscaLastData
      frmSplash.IpConexio.CarterUltimaData = DateAdd("m", -3, Now)
      frmSplash.IpConexio.Cfg_ConnexioString = NomServerInternet
      frmSplash.IpConexio.Cfg_AppPath = AppPath
      frmSplash.IpConexio.EnviaIReb "Envia"
   End If
   My_DoEvents
   
   InformaMiss ""

End Sub

Sub GemeraSqlTransDestesBotiga(botiga, Files() As String)
    Dim sql As String
    
    
    ExecutaComandaSql "CREATE INDEX CertificatDeutesAnticips_Idx ON CertificatDeutesAnticips (IdDeute,Accio,Params1,Params2,Client)    "
    
    sql = "select a.IdDeute As Id,a.Dependenta,a.Client,a.DataDaute As Data,0 as Estat,1 as Tipus,a.Import,a.Params2 As Botiga,'[NumTick:'+ltrim(a.Params1)+']' Detall "
    sql = sql & "from CertificatDeutesAnticips a "
    sql = sql & "left join CertificatDeutesAnticips b on  a.iddeute = b.iddeute and b.accio = 'Pagat' "
    sql = sql & "Where "
    sql = sql & "a.params2 ='" & botiga & "' And "
    sql = sql & "a.accio = 'Asumit' And "
    sql = sql & "b.accio is null "
    ReDim Files(0)
    
    GemeraSqlTrans "DeutesAnticipsv2", Files, sql

End Sub

Sub PreparaComandes()
   Dim D As Date, sql As String, Rs As rdoResultset
   Dim algun As Boolean
   
'   If UCase(EmpresaActual) = UCase("Pa Natural") _
'      Or UCase(EmpresaActual) = UCase("Tena") _
'      Or UCase(EmpresaActual) = UCase("Enrich") _
'      Or UCase(EmpresaActual) = UCase("Daunis") _
'      Or UCase(EmpresaActual) = UCase("villena") _
'      Or UCase(EmpresaActual) = UCase("sistare") _
'      Or UCase(EmpresaActual) = UCase("Carne") _
'      Or UCase(EmpresaActual) = UCase("Saborit") _
'      Or UCase(EmpresaActual) = UCase("Armengol") _
'      Or UCase(EmpresaActual) = UCase("Demo") _
'      Or UCase(EmpresaActual) = UCase("DemoForn") _
'      Or UCase(EmpresaActual) = UCase("Vilapan") _
'      Or UCase(EmpresaActual) = UCase("Cuinem") _
'      Or UCase(EmpresaActual) = UCase("PaDeCava") Then
      
      
   If UCase(EmpresaActual) = UCase("Carne") _
      Or UCase(EmpresaActual) = UCase("Armengol") _
      Or UCase(EmpresaActual) = UCase("LaPanera") _
      Or UCase(EmpresaActual) = UCase("PanAbad") _
      Or UCase(EmpresaActual) = UCase("sistare") _
      Or UCase(EmpresaActual) = UCase("Saborit") _
      Or UCase(EmpresaActual) = UCase("PaNatural") _
      Or UCase(EmpresaActual) = UCase("Pa Natural") Then
  
      ExecutaComandaSql "Drop table Servit_Tmp"
      
      Set Rs = Db.OpenResultset("Select * From Records Where Concepte = 'ComandaBotiga'")
      If Rs.EOF Then ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('ComandaBotiga',DATEADD(Day, -1, GetDate()))"

      D = DateAdd("d", -7, Now)
      'D = DateAdd("d", -20, Now)
      ExecutaComandaSql "Drop table Servit_Tmp "
      sql = "CREATE TABLE [Servit_Tmp]([DiaDesti] [varchar](8) NOT NULL,[Id] [uniqueidentifier] NULL,[TimeStamp] [datetime] NULL,    [QuiStamp] [nvarchar](255) NULL,[Client] [float] NULL,[CodiArticle] [int] NULL,[PluUtilitzat] [nvarchar](255) NULL,[Viatge] [nvarchar](255) NULL,[Equip] [nvarchar](255) NULL,[QuantitatDemanada] [float] NULL,[QuantitatTornada] [float] NULL,[QuantitatServida] [float] NULL,[MotiuModificacio] [nvarchar](255) NULL,"
      sql = sql & " [Hora] [float] NULL,[TipusComanda] [float] NULL,    [Comentari] [nvarchar](255) NULL,  [ComentariPer] [nvarchar](255) NULL,  [Atribut] [int] NULL,  [CitaDemanada] [nvarchar](255) NULL,  [CitaServida] [nvarchar](255) NULL,  [CitaTornada] [nvarchar](255) NULL )"
      ExecutaComandaSql sql
      
      ' TABLA PARA RELLENARLO TODO
      ExecutaComandaSql "Drop table Servit_Tmp2 "
      sql = "CREATE TABLE [Servit_Tmp2]([DiaDesti] [varchar](8) NOT NULL,[Id] [uniqueidentifier] NULL,[TimeStamp] [datetime] NULL,    [QuiStamp] [nvarchar](255) NULL,[Client] [float] NULL,[CodiArticle] [int] NULL,[PluUtilitzat] [nvarchar](255) NULL,[Viatge] [nvarchar](255) NULL,[Equip] [nvarchar](255) NULL,[QuantitatDemanada] [float] NULL,[QuantitatTornada] [float] NULL,[QuantitatServida] [float] NULL,[MotiuModificacio] [nvarchar](255) NULL,"
      sql = sql & " [Hora] [float] NULL,[TipusComanda] [float] NULL,    [Comentari] [nvarchar](255) NULL,  [ComentariPer] [nvarchar](255) NULL,  [Atribut] [int] NULL,  [CitaDemanada] [nvarchar](255) NULL,  [CitaServida] [nvarchar](255) NULL,  [CitaTornada] [nvarchar](255) NULL )"
      ExecutaComandaSql sql
      
      While D < DateAdd("d", 60, Now)
          sql = "Select '" & Format(D, "yy-mm-dd") & "' as DiaDesti,* from [" & DonamNomTaulaServit(D) & "] where timestamp > (select timestamp from records where concepte = 'ComandaBotiga') "
          ExecutaComandaSql "insert into Servit_Tmp2  select * from (" & sql & ") t "
          D = DateAdd("d", 1, D)
          InformaMiss "Preparant Comandes Dia " & D
      Wend
      
      Set Rs = Db.OpenResultset("Select distinct Client,DiaDesti from Servit_Tmp2 join paramshw on Servit_Tmp2.client = paramshw.valor1")
      While Not Rs.EOF
        sql = "Select '" & Format(DateSerial(Mid(Rs("diadesti"), 1, 2), Mid(Rs("diadesti"), 4, 2), Mid(Rs("diadesti"), 7, 2)), "yy-mm-dd") & "' as DiaDesti,* from [" & DonamNomTaulaServit(DateSerial(Mid(Rs("diadesti"), 1, 2), Mid(Rs("diadesti"), 4, 2), Mid(Rs("diadesti"), 7, 2))) & "] where client = " & Rs("Client")
        ExecutaComandaSql "insert into Servit_Tmp select * from (" & sql & ") t "
        Rs.MoveNext
      Wend
      Rs.Close
            
      Set Rs = Db.OpenResultset("Select distinct Client from Servit_Tmp join paramshw on Servit_Tmp.client = paramshw.valor1")
      While Not Rs.EOF
         If Not IsNull(Rs("Client")) Then If IsNumeric(Rs("Client")) Then Missatges_CalEnviar "Comandes", "[" & Rs("Client") & "]"
         Rs.MoveNext
      Wend
      Rs.Close
      
      Set Rs = Db.OpenResultset("Select Top 1 * from Servit_Tmp ")
      If Not Rs.EOF Then
         ExecutaComandaSql "delete Records Where Concepte = 'ComandaBotiga'"
         ExecutaComandaSql "insert into Records  (timestamp,concepte) Select max([TimeStamp]) as [TimeStamp], 'ComandaBotiga' as concepte From Servit_Tmp"
      End If
      
   Else
   End If
   
   ExecutaComandaSql "Delete ComandesModificades"

End Sub


Sub SincronitzaEmpresaReb(empresa)
   Dim CalBorrarArticles As Boolean, Files() As String, Rs As rdoResultset, Botis() As String, K As Integer, p1 As String, P2 As String
   Dim i As Integer, Tipus() As String, Param() As String, Contingut As String, j As Integer, LaAgafem As Boolean, nom As String, Interesa As Integer, Esborrem As Integer, EsCfg As Boolean, ClientsGenerats As Boolean, EnviaArticles As Boolean, LlistaNegra() As String, Salve As Boolean, AgafaTot As Boolean
   
   AgafaTot = False
   For i = 1 To UBound(feina)
      If feina(i).empresa = empresa And Not feina(i).EscoltaLlicencies(0) = "" Then AgafaTot = True
   Next

   LlistaNegra = Split(UCase("DiccionariTot,ComandesPlantilles,Santoral,TpvVellsCodis,Dependentes,Memotecnics,Atributs,Promocions,ConstantsClient,Clients,Viatges,Equips,CodisBarres,ArticlesPropietats,PreusArticles,ProductesPromocionats,FamiliesArticles,Punts,Facturacio,FeinaFeta,CaixesCongelador,ClientsFinalsAcumulat,Clients,tarifesespecialsclients,ARTICLES"), ",")
   InformaMiss "Configurant"
   
   frmSplash.IpConexio.InteresPerContingutReset
   If ExisteixTaula("InteresaContingut") Then
      Set Rs = Db.OpenResultset("Select distinct * From InteresaContingut ")
      While Not Rs.EOF
         LaAgafem = True
         For i = 0 To UBound(LlistaNegra)
            If UCase(Rs("Nom")) = LlistaNegra(i) Then
                LaAgafem = False
                Exit For
            End If
         Next
         
         If AgafaTot Then LaAgafem = True
         If Left(UCase(Rs("Nom")), 14) = "TARIFAESPECIAL" Then LaAgafem = False
         frmSplash.IpConexio.InteresPerContingut Rs("Nom"), LaAgafem, Rs("LaEsborrem") = 1
         Rs.MoveNext
      Wend
      Rs.Close
   End If
   
   frmSplash.IpConexio.LabelEstat = frmSplash.Estat
   frmSplash.IpConexio.LabelEstatDbg = frmSplash.lblVersion
   
   InformaMiss "Connectant ... "
   frmSplash.IpConexio.CarterUltimaData = BuscaLastData
   frmSplash.IpConexio.CarterUltimaData = DateAdd("m", -3, Now)
   frmSplash.IpConexio.Cfg_ConnexioString = NomServerInternet
   frmSplash.IpConexio.Cfg_AppPath = AppPath
'ExecutaComandaSql "insert into recordsfilesBak select * from recordsfiles where data <  DATEADD(day, -15, getdate())"
'ExecutaComandaSql "delete recordsfiles where data <  DATEADD(day, -15, getdate())"
   
'frmSplash.Enabled = False
   If frmSplash.IpConexio.EnviaIReb("Reb") Then SetLastData DateAdd("s", -3, frmSplash.IpConexio.CarterUltimaData)
   'frmSplash.EsperaCues
   
'frmSplash.Enabled = True
   My_DoEvents
   i = 1
   If ExisteixTaula("InteresaContingut") Then
      ExecutaComandaSql "Delete InteresaContingut"
      While frmSplash.IpConexio.InteresPerContingutGet(i, nom, Interesa, Esborrem)
         Db.Execute ("Insert Into InteresaContingut (Nom,LaAgafem,LaEsborrem) Values ('" & nom & "' , " & Interesa & ", " & Esborrem & ")")
         i = i + 1
      Wend
   End If
  
   For i = 1 To UBound(feina)
       If feina(i).empresa = empresa And Not feina(i).EscoltaLlicencies(0) = "" Then
          CarregaDir Files, AppPath & "\*.*"
          For K = 1 To UBound(Files)
             My_DoEvents
             Salve = False
             For j = 0 To UBound(feina(i).EscoltaLlicencies)
                If InStr(Files(K), "[Maquina#" & Format(feina(i).EscoltaLlicencies(j), "00000") & "]") > 0 Then Salve = True
             Next
             If InStr(Files(K), "[maquina#" & Mid(Format(feina(i).llicencia / 3, "0000000000"), 3, 5) & "]") > 0 Then Salve = True
             If InStr(Files(K), "[Maquina#") = 0 Then Salve = True
             If Not Salve Then
                Kill AppPath & "\" & Files(K)
             End If
          Next
       End If
   Next

   InformaMiss "Interpretant Fitchers Rebuts ... "
   If UCase(EmpresaActual) = UCase("integraciones") Then
      SincronitzaIntegracionesRebPas2 AppPath & "\tmp"
   Else
      Interpreta_SqlTrans frmSplash.Estat
   End If
   
   For i = 1 To UBound(feina)
       If feina(i).empresa = empresa And Not feina(i).EscoltaLlicencies(0) = "" Then
           ExecutaComandaSql "update dependentes set nom = '" & empresa & "' ,memo = '" & empresa & "_WEB' where codi = 1"
       End If
   Next
   InformaMiss ""

End Sub


Sub CarregaBotiguesOnEnviarComandes(Botis() As String)
    Dim Rs As rdoResultset
    
    ReDim Botis(0)
On Error GoTo nor
    Set Rs = Db.OpenResultset("SELECT Camp From ComandesParams WHERE Tipus = 1 AND (Valor = 4 OR Valor = 3 OR Valor = 2)")
    
    While Not Rs.EOF
       ReDim Preserve Botis(UBound(Botis) + 1)
       Botis(UBound(Botis)) = Rs(0)
       Rs.MoveNext
    Wend
    Rs.Close
nor:
End Sub

Function LlistaBotigues() As String
    Dim Rs As rdoResultset
    
On Error GoTo nor
    LlistaBotigues = ""
    Set Rs = Db.OpenResultset("select c.codi from paramshw w join clients c on c.codi = w.valor1 order by c.nom ")
    
    While Not Rs.EOF
        If Not LlistaBotigues = "" Then LlistaBotigues = LlistaBotigues & ","
        LlistaBotigues = LlistaBotigues & Rs(0)
        Rs.MoveNext
    Wend
    Rs.Close
nor:

End Function

Function Connecta(Tipus As Integer, llicencia As String, Server As String, Database As String, i) As Boolean

   EmpresaActualNum = i
   Connecta = False
   If InStr(UCase(FeinaAfer), "CALCULS") = 0 And InStr(UCase(FeinaAfer), "IDF:") = 0 Then ValidaLlicencia llicencia
   LastServer = "tcp:silema.hiterp.com" ' Server
   LastDatabase = Database
   LastLlicencia = llicencia
'If UCase(Server) = UCase("Titan") Then

   Connecta = ConnectaSqlServer(LastServer, LastDatabase)
   
End Function

Sub CarregaLlistaDeFeines(feina() As TipFeina)
   Dim Rs As rdoResultset, Rs2 As rdoResultset, Tip As Integer, i As Integer, CondicioExtra  As String, Cada As String
   Dim sql, Rs3 As ADODB.Recordset
   
   ReDim feina(0)
   CondicioExtra = ""
'   CondicioExtra = CondicioExtra & " And ( "
'   CondicioExtra = CondicioExtra & " Empresa = 'Integraciones' "
'   CondicioExtra = CondicioExtra & " Or Empresa = 'iartpa' "
'   CondicioExtra = CondicioExtra & " Or Empresa = 'iblatpa' "
'   CondicioExtra = CondicioExtra & ")"
   
   If ExisteixTaula("Web_ServeisComuns") Then
      Set Rs = Db.OpenResultset("Select Empresa,Tipus From Web_ServeisComuns Where Actiu = 1 " & CondicioExtra & " Order By Empresa  ")
      While Not Rs.EOF
         If Rs("Tipus") = 10 Then 'Residencies
            sql = "Select id,codigo,nombre,password,webserver,[database] Db,dbServer From gdrEmpresas "
            sql = sql & "Where codigo = '" & Rs("Empresa") & "' and regEstado='A' "
            Set Rs3 = sf_recGdr(sql)
            If Not Rs3.EOF Then
               ReDim Preserve feina(UBound(feina) + 1)
               feina(UBound(feina)).Tipus = Rs("Tipus")
               feina(UBound(feina)).empresa = Rs("Empresa")
               feina(UBound(feina)).Db = Rs3("Db")
               feina(UBound(feina)).Path = Rs3("webserver")
               feina(UBound(feina)).llicencia = ""
               feina(UBound(feina)).Ftp_Server = ""
               feina(UBound(feina)).Ftp_User = ""
               feina(UBound(feina)).Ftp_Pssw = ""
               
               feina(UBound(feina)).Server = Rs3("dbServer")
               ReDim feina(UBound(feina)).EscoltaLlicencies(0)
               feina(UBound(feina)).EscoltaLlicencies(0) = ""
            End If
         Else
            Set Rs2 = Db.OpenResultset("Select  isnull(Ftp_Server,'') Ftp_Server ,isnull(Ftp_User,'') Ftp_User ,isnull(Ftp_Pssw,'') Ftp_Pssw , isnull(EscoltaLlicencies,'') EscoltaLlicencies,Llicencia,Db,Path,Db_Server As Servidor From Web_Empreses Where Nom = '" & Rs("Empresa") & " ' ")
            If Not Rs2.EOF Then
               ReDim Preserve feina(UBound(feina) + 1)
               feina(UBound(feina)).Tipus = Rs("Tipus")
               feina(UBound(feina)).empresa = Rs("Empresa")
               feina(UBound(feina)).Db = Rs2("Db")
               feina(UBound(feina)).Path = Rs2("Path")
               feina(UBound(feina)).llicencia = GeneraClau(DiscSerialNumber(), Mid(Format(Rs2("Llicencia") / 3, "0000000000"), 3, 5))
               feina(UBound(feina)).Server = Rs2("Servidor")
               feina(UBound(feina)).Ftp_Server = Rs2("Ftp_Server")
               feina(UBound(feina)).Ftp_User = Rs2("Ftp_User")
               feina(UBound(feina)).Ftp_Pssw = Rs2("Ftp_Pssw")
               
               ReDim feina(UBound(feina)).EscoltaLlicencies(0)
               feina(UBound(feina)).EscoltaLlicencies(0) = ""
               If Len(Rs2("EscoltaLlicencies")) > 0 Then
                   ReDim feina(UBound(feina)).EscoltaLlicencies(UBound(Split(Rs2("EscoltaLlicencies"))))
                   feina(UBound(feina)).EscoltaLlicencies = Split(Rs2("EscoltaLlicencies"), ",")
               End If
            Else
             
            End If
            Rs2.Close
         End If
         Rs.MoveNext
      Wend
      Rs.Close
   Else
      ReDim Preserve feina(1)
      feina(1).Tipus = 1
      feina(1).empresa = Cfg_Database
      feina(1).Db = Cfg_Database
      feina(1).llicencia = Cfg_Llicencia
      feina(1).Server = Cfg_Server
   End If
      
   frmSplash.Enchufat.Clear
   frmSplash.Enchufat.AddItem "Seguent"
   For i = 1 To UBound(feina)
      If frmSplash.Enchufat.List(frmSplash.Enchufat.ListCount - 1) <> feina(i).empresa Then frmSplash.Enchufat.AddItem feina(i).empresa
   Next
   frmSplash.Enchufat.ListIndex = 0
   
   Cada = "Cada 10 n"
   If FeinaAfer = "CalculsLlargs" Or FeinaAfer = "CalculsResi" Then
        ' ExecutaComandaSql "Delete hit.dbo.calculsespecials"
        'Set Rs = Db.OpenResultset("Select Empresa From hit.dbo.CalculsEspecials Where DATEDIFF(n, ti,GETDATE()) > 60 ")
        Set Rs = Db.OpenResultset("Select Empresa From hit.dbo.CalculsEspecials Where DATEDIFF(n, ti,GETDATE()) > 180 ")
        If Not Rs.EOF Then
            If Not IsNull(Rs(0)) Then
                'sf_enviarMail "Secrehit@gmail.com", EmailGuardia , "Error calcul Puntual parat desde mes de 1 hora a empresa " & Rs(0), "", "", ""
                'Set Rs = Db.OpenResultset("Delete from hit.dbo.CalculsEspecials Where DATEDIFF(n, ti,GETDATE()) > 60 ")
                Set Rs = Db.OpenResultset("Delete from hit.dbo.CalculsEspecials Where DATEDIFF(n, ti,GETDATE()) > 180 ")
            End If
        End If
        
        Cada = "Cada 30 n"
   End If
   
   If Left(UCase(FeinaAfer), 2) = "ID" Then
        Cada = "Cada 60 n"
   End If
   
   ExecutaComandaSql "Delete hit.dbo.gosdetura Where NomObella = '" & FeinaAfer & "' "
   ExecutaComandaSql "insert into hit.dbo.gosdetura   (id     ,NomObella      ,NomGos     ,NomPasto       ,ObellaPeriodicitat,TsVista  ,TsRevisada,TsAvisat ,UltimaFraseObella,UltimaFraseGosDeTura) values  (newId(),'" & FeinaAfer & "','GosDeTura','sat@HitSystems.Es','" & Cada & "',getDate(),getDate() ,getDate() ,'Inici'        ,'Tot Ok') "

End Sub

Function ConnectaSqlServer(Server As String, NomDb As String) As Boolean
   Dim User As String, Psw As String, MyId As String, Rs As rdoResultset, Modus
   ConnectaSqlServer = False

On Error Resume Next
'   Db.Close
On Error GoTo nor
   If InStr(Command, "192.9.199.202") > 0 And UCase(Server) = UCase("juliet") Then Server = "192.9.199.202"
   'Server = "172.26.0.101"
   User = "sa"
   Psw = "LOperas93786"
   
   
   If UCase(Server) = "NEPTU" Then
      User = "sa"
      Psw = "adminhit"
   End If
   
   If UCase(Server) = "SERVIDORNT" Then
      User = "IIS_ElFornet"
      Psw = "adminjuliet"
   End If
   
   If UCase(Server) = "SERVIDOR" Then
      User = "user"
      Psw = "resu"
   End If
   
   If UCase(Server) = "JULIET" Then
      User = "IIS_ElFornet"
      Psw = "Iis2003"
   End If
   
   If UCase(Server) = "TITAN" Then
      User = "IIS_ElFornet"
      Psw = "Iis2003"
   End If
   
   If UCase(Server) = "192.9.199.202" Or Server = "172.26.0.101" Then
      User = "sa"
      Psw = "margarita"
   End If
   
   If UCase(Server) = "86.109.98.189" Then
      User = "SQL_Admin"
      Psw = "sql_admin4071"
   End If
   
   If UCase(Server) = "10.1.2.16" Then
     ' Server = "10.1.3.16"
      User = "SQL_Admin"
      Psw = "sql_admin4071"
   End If
   
   If UCase(Server) = UCase("SERVERCLOUD") Then
     ' Server = "10.1.3.16"
      User = "SQL_Admin"
      Psw = "sql_admin4071"
   End If
   
    If UCase(Server) = UCase("inc.hiterp.com") Then
     ' Server = "10.1.3.16"
      User = "SQL_Admin"
      Psw = "sql_admin4071"
   End If
   
   If UCase(Server) = UCase("silema.hiterp.com") Or UCase(Server) = UCase("WEB") Then
      User = "sa"
      Psw = "LOperas93786"
   End If
   
   
   If UCase(Server) = UCase("hitsystems2") Then Server = "192.168.0.1"
   If UCase(Server) = UCase("hitsys") Then Server = "192.168.0.1"
   If Server = "86.109.96.151:81" Then Server = "86.109.96.151"
   If UCase(Server) = "86.109.96.151" Then
      User = "SQL_Admin"
      Psw = "SQL_Admin4071"
   End If
   
   If UCase(Server) = UCase("86.109.96.151") Then Server = "192.168.0.1"
   
   If UCase(Server) = "192.168.0.1" Then
      User = "SQL_Admin"
      Psw = "SQL_Admin4071"
   End If
   
   
'      User = "sa"
'      Psw = "LOperas93786"
      
   Modus = "Exclusiu_"
   If FeinaAfer = "CalculsCurts" Or InStr(FeinaAfer, "Idf:") > 0 Then Modus = "NoExlusiu_"
   MyId = Modus & GetNomMaquinaSql() & "_" & App.Title & "_" & FeinaAfer
        
   If UCase(NomDb) = "HIT" Then MyId = "Global_" & GetNomMaquinaSql() & "_" & App.Title & "_" & FeinaAfer
   
   db2MyId = "SegonaConnexio_" & MyId
   db2User = "sa" ' User
   db2Psw = "LOperas93786" ' Psw
   db2NomDb = NomDb
   db2Server = Server
   
'   Db.Connect = "WSID=" & MyId & ";UID=" & User & ";PWD=" & Psw & ";Database=" & NomDb & ";Server=" & Server & ";Driver={SQL Server};DSN='';"
   If DbNameEsBuit() Then
        User = "sa"
        Psw = "LOperas93786"
       Db.Connect = "WSID=" & MyId & ";UID=" & User & ";PWD=" & Psw & ";Database=" & NomDb & ";Server=" & Server & ";Driver={SQL Server};DSN='';"
       Db.EstablishConnection rdDriverNoPrompt
   End If
   
   Db.Execute "use " & NomDb
   
Debug.Print "Conectant a ... " & NomDb

'
   
''   Set Rs = Db.OpenResultset("Sp_Who")
   Set Rs = Db.OpenResultset("sp_who '" & User & "'")
'
   ConnectaSqlServer = True
    If Not (FeinaAfer = "CalculsCurts" Or InStr(FeinaAfer, "Idf:") > 0) Then
        While Not Rs.EOF
            If Left(Trim(Rs("HostName")) & "aaaaaaaaaa", 9) = "Exclusiu_" And UCase(Trim(Rs("DbName"))) = UCase(NomDb) And Not UCase(Trim(Rs("HostName"))) = UCase(MyId) Then
                ConnectaSqlServer = False
                Rs.Close
'                Db.Close
                Exit Function
            End If
            Rs.MoveNext
        Wend
        Rs.Close
    End If
    
    
Exit Function

nor:
    If InStr(err.Description, "No se puede abrir la base de datos") > 0 Then
        ConnectaSqlServer = False
        Exit Function
    End If
    Db.Connect = "WSID=" & MyId & ";UID=" & User & ";PWD=" & Psw & ";Database=" & NomDb & ";Server=" & Server & ";Driver={SQL Server};DSN='';"
    Db.EstablishConnection rdDriverNoPrompt
    Db.Execute "use " & NomDb
Debug.Print err.Description
Resume Next
End Function





Sub DescarregaLog(NomEmpresa As String)
    
    Informa "Descarregant Llicencies .. "
    My_DoEvents
    Download
    Informa "Interpretant Llicencies.. "
    My_DoEvents
    Interpreta
   
End Sub

Sub Interpreta()
   Dim Fil As String, Files() As String, i As Integer, sql As String, Rs As rdoResultset
   
   AseguraExisteixtaula
   
   CarregaDir Files, AppPath & "\*.log"
   Set Q_Insert = Db.CreateQuery("", "Insert Into LogConnexio ([Llicencia],[Hi],[Hf],[Client],[NumDisc],[Fils_Rebudes],[Fils_Enviades],[Sw_Nom],[Sw_Versio],[Tot]) Values (?,?,?,?,?,?,?,?,?,?)")
   For i = 1 To UBound(Files)
      Informa2 Files(i)
      My_DoEvents
      FitcherProcesat Files(i), ParsejaLog(Files(i))
   Next
   Q_Insert.Close
   ExecutaComandaSql "Update LogConnexio Set LogConnexio.Client = Llicencies.Empresa From LogConnexio  Inner Join  Llicencies On LogConnexio.Llicencia = Llicencies.Llicencia Where LogConnexio.Client is null or LogConnexio.Client = '' Or LogConnexio.Client = '0' "
   
   CarregaDir Files, AppPath & "\?????.Txt"
   Set Q_Insert = Db.CreateQuery("", "Insert Into Llicencies ([Llicencia],[Empresa],[Tot]) Values (?,?,?)")
   Set Q_Delete = Db.CreateQuery("", "Delete Llicencies Where [Llicencia] = ? ")
   For i = 1 To UBound(Files)
      FitcherProcesat Files(i), ParsejaLic(Files(i))
   Next
   Q_Insert.Close
   Q_Delete.Close
   
   CarregaDir Files, AppPath & "\empresa#*.Txt"
   Set Q_Insert = Db.CreateQuery("", "Insert Into Empresas ([Nom],[DiesMemoria],[Tot]) Values (?,?,?)")
   Set Q_Delete = Db.CreateQuery("", "Delete Empresas Where [Nom] = ? ")
   For i = 1 To UBound(Files)
      Informa Files(i)
      My_DoEvents
      FitcherProcesat Files(i), ParsejaEmp(Files(i))
   Next
   Q_Insert.Close
   Q_Delete.Close
   
   Set Rs = Db.OpenResultset("Select distinct web_empreses.Db empresa from web_empreses ")
   While Not Rs.EOF
      If Not IsNull(Rs("empresa")) Then ExecutaComandaSql "update l set l.nom = c.nom,l.adresa=c.adresa,l.ciutat=c.ciutat from  llicencies l join " & Rs("Empresa") & ".dbo.paramshw w on l.llicencia =  w.codi join " & Rs("Empresa") & ".dbo.clients c on w.valor1 = c.codi where l.empresa = '" & Rs("Empresa") & "'"
      Rs.MoveNext
   Wend
   Rs.Close

   
End Sub
Sub My_DoEvents()
       DoEvents
End Sub

Function ParsejaEmp(Fil As String) As Boolean
   Dim f, lin As String, Tot As String, Var As String, Valor As String, P As Integer, i As Integer, p1 As Integer
   
   
   ParsejaEmp = False
   For i = 0 To Q_Insert.rdoParameters.Count - 1
      Q_Insert.rdoParameters(i).Value = 0
   Next
   
   f = FreeFile
   Tot = ""
   
   P = InStr(Fil, "#")
   p1 = InStr(Fil, ".")
   Q_Insert.rdoParameters(0).Value = Mid(Fil, P + 1, p1 - P - 1)
   Q_Delete.rdoParameters(0).Value = Q_Insert.rdoParameters(0).Value
   
   Open AppPath & "\" & Fil For Input As f
   While Not EOF(f)
      Line Input #f, lin
      P = InStr(lin, ":")
      If Len(Tot) > 0 Then Tot = Tot & vbCrLf
      Tot = Tot & lin
      If P > 0 Then
         Var = Trim(Left(lin, P - 1))
         Valor = Trim(Right(lin, Len(lin) - P))
         Select Case UCase(Var)
            Case "DIESMEMORIA": Q_Insert.rdoParameters(1).Value = Valor
                                ParsejaEmp = True
         End Select
      End If
   Wend
   Q_Insert.rdoParameters(2).Value = Tot
   Q_Delete.Execute
   Q_Insert.Execute
   Close f
   
End Function





Function ParsejaLic(Fil As String) As Boolean
   Dim f, lin As String, Tot As String, Var As String, Valor As String, P As Integer, i As Integer
   Dim Rs As rdoResultset
   
 
   ParsejaLic = False
   For i = 0 To Q_Insert.rdoParameters.Count - 1
      Q_Insert.rdoParameters(i).Value = 0
   Next
   
   f = FreeFile
   Tot = ""
   
   P = InStr(Fil, ".")
   Q_Insert.rdoParameters(0).Value = Left(Fil, P - 1)
   Q_Delete.rdoParameters(0).Value = Q_Insert.rdoParameters(0).Value
   
   Open AppPath & "\" & Fil For Input As f
   While Not EOF(f)
      Line Input #f, lin
      P = InStr(lin, ":")
      If Len(Tot) > 0 Then Tot = Tot & vbCrLf
      Tot = Tot & lin
      If P > 0 Then
         Var = Trim(Left(lin, P - 1))
         Valor = Trim(Right(lin, Len(lin) - P))
         Select Case UCase(Var)
            Case "EMPRESA": Q_Insert.rdoParameters(1).Value = Trim(Valor)
                            ParsejaLic = True
         End Select
      End If
   Wend
   Q_Insert.rdoParameters(2).Value = Tot
   Q_Delete.Execute
   Q_Insert.Execute
   
   'Set rs = Db.OpenResultset("select * from Fac_HitRs.dbo.recursosExtes where variable='LICENCIA' and valor='" & )
   
   Close f
   
End Function




Sub FitcherProcesat(Fil As String, Optional bo As Boolean = True, Optional Enviat As Boolean = False)
   
On Error Resume Next
   If Enviat Then
      MkDir AppPath & "\Env"
      Name AppPath & "\" & Fil As AppPath & "\Env\[Processat#" & Format(Now, "yyyymmddhhnnss") & "]" & Fil
   Else
      If bo Then
         Name AppPath & "\" & Fil As AppPath & "\Bak\[Processat#" & Format(Now, "yyyymmddhhnnss") & "]" & Fil
      Else
         Name AppPath & "\" & Fil As AppPath & "\Err\[Processat#" & Format(Now, "yyyymmddhhnnss") & "]" & Fil
      End If
   End If
   
   Kill AppPath & "\" & Fil
   
On Error GoTo 0
   
End Sub


Function ParsejaLog(Fil As String) As Boolean
   Dim f, lin As String, Tot As String, Var As String, Valor As String, P As Integer, i As Integer
   
   ParsejaLog = False
   For i = 0 To Q_Insert.rdoParameters.Count - 1
      Q_Insert.rdoParameters(i).Value = 0
   Next
   
   f = FreeFile
   Tot = ""
   
   Open AppPath & "\" & Fil For Input As f
   While Not EOF(f)
      Line Input #f, lin
      P = InStr(lin, ":")
      If Len(Tot) > 0 Then Tot = Tot & Chr(13) & Chr(10)
      Tot = Tot & lin
      If P > 0 Then
         Var = Trim(Left(lin, P - 1))
         Valor = Trim(Right(lin, Len(lin) - P))
         Select Case UCase(Var)
            Case "NUM HI": Q_Insert.rdoParameters(1).Value = HoraNumDate(Valor)
            Case "NUM HF": Q_Insert.rdoParameters(2).Value = HoraNumDate(Valor)
            Case "NUMDISC": Q_Insert.rdoParameters(4).Value = Valor
            Case "NUM LLICENCIA": Q_Insert.rdoParameters(0).Value = Valor
                                  ParsejaLog = True
            Case "FILES AGAFADES": Q_Insert.rdoParameters(5).Value = Valor
            Case "FILES ENVIADES": Q_Insert.rdoParameters(6).Value = Valor
            Case "SW NOM": Q_Insert.rdoParameters(7).Value = Valor
            Case "SW VER": Q_Insert.rdoParameters(8).Value = Valor
         End Select
      End If
   Wend
   
   'Q_Insert.rdoParameters(3).Value = LlicenciaClient(Q_Insert.rdoParameters(0).Value)
   
   Q_Insert.Execute
   Close f
   
End Function

Function HoraNumDate(s As String) As Date
   Dim An As Double, dia As Double, mes As Double, Hora As Double, Minut As Double, Segon As Double
   
   '20010219110050
   
   HoraNumDate = Now
   If Len(s) <> 14 Then Exit Function
   
   An = Mid(s, 1, 4)
   mes = Mid(s, 5, 2)
   dia = Mid(s, 7, 2)
   Hora = Mid(s, 9, 2)
   Minut = Mid(s, 11, 2)
   Segon = Mid(s, 13, 2)
   
   HoraNumDate = DateSerial(An, mes, dia) + TimeSerial(Hora, Minut, Segon)
   
   
End Function



Sub CarregaDir(Files() As String, Patro As String)
   Dim Fil  As String
   
   ReDim Files(0)
   
   Fil = Dir(Patro)
   
   While Len(Fil) > 0
      ReDim Preserve Files(UBound(Files) + 1)
      Files(UBound(Files)) = Fil
      Fil = Dir
   Wend

End Sub



Sub AseguraExisteixtaula()
   Dim sql As String
   
   If Not ExisteixTaula("LogConnexio") Then
      sql = "CREATE TABLE LogConnexio ( "
      sql = sql & " [Llicencia]     [float]    NULL, "
      sql = sql & " [Hi]            [datetime] NULL ,"
      sql = sql & " [Hf]            [datetime] NULL , "
      sql = sql & " [Client]        [float]    NULL , "
      sql = sql & " [NumDisc]       [float]    NULL , "
      sql = sql & " [Fils_Rebudes]  [float]    NULL , "
      sql = sql & " [Fils_Enviades] [float]    NULL , "
      sql = sql & " [Sw_Nom]        [nvarchar] (255) NULL , "
      sql = sql & " [Sw_Versio]     [nvarchar] (255) NULL , "
      sql = sql & " [Tot]           [nvarchar] (255) NULL ) ON [PRIMARY] "
      ExecutaComandaSql sql
   End If
   
   If Not ExisteixTaula("Empresas") Then
      sql = "CREATE TABLE Empresas ( "
      sql = sql & " [Nom]           [nvarchar] (255) NULL , "
      sql = sql & " [DiesMemoria]   [float]    NULL, "
      sql = sql & " [Tot]           [nvarchar] (255) NULL ) ON [PRIMARY] "
      ExecutaComandaSql sql
   End If
   
   If Not ExisteixTaula("Llicencies") Then
      sql = "CREATE TABLE Llicencies ( "
      sql = sql & " [Llicencia]     [float]    NULL, "
      sql = sql & " [Empresa]       [nvarchar] (255) NULL , "
      sql = sql & " [Tot]           [nvarchar] (255) NULL ) ON [PRIMARY] "
      ExecutaComandaSql sql
   End If
   
   If Not ExisteixTaula("Records") Then
      ExecutaComandaSql "CREATE TABLE Records ([TimeStamp] [datetime] Null,[Concepte] [nvarchar] (255) NULL) ON [PRIMARY]"
      ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('Facturacio',DATEADD(Day, -1, GetDate()))"
      ExecutaComandaSql "Insert Into Records (Concepte,[TimeStamp]) Values ('FeinaFeta',DATEADD(Day, -1, GetDate()))"
   End If
   
   If Not ExisteixTaula("ComandesModificades") Then ExecutaComandaSql "CREATE TABLE [ComandesModificades] ([Id] [uniqueidentifier] NULL ,[TimeStamp] [datetime] NULL ,  [TaulaOrigen] [nvarchar] (255) NULL ) ON [PRIMARY] "
   
End Sub



Sub Download()
   
frmSplash.Enabled = False
   If frmSplash.IpConexio.LogDownload("Alicia", NomServerInternet) Then SetLastData frmSplash.IpConexio.CarterUltimaData
frmSplash.Enabled = True
   
End Sub




Function ExecutaComandaSql(MySql As String, Optional StErr As String = "") As Boolean
   DoEvents
   
   On Error GoTo no_Be
   Db.Execute MySql
   On Error Resume Next
   ExecutaComandaSql = True
   Exit Function
   
no_Be:
   StErr = err.Description & "-" & Format(Now, "dd hh:mm:ss") & " * (" & MySql & ")"
   Debuga Format(Now, "dd hh:mm:ss") & " * (" & MySql & ")" & err.Description
   
   If InStr(err.Description, "El objeto no es válido o no está definido.") > 0 Or InStr(err.Description, "Error general de red") > 0 Or InStr(err.Description, "or en el vínculo de comunicació") Then
      Reconecta
      ExecutaComandaSql = ExecutaComandaSqlSegunDaOportunidad(MySql)
   Else
      ExecutaComandaSql = False
   End If
   
End Function

Function ExecutaComandaSqlSegunDaOportunidad(MySql As String) As Boolean
    DoEvents
   
   On Error GoTo no_Be
   Db.Execute MySql
   On Error Resume Next
   ExecutaComandaSqlSegunDaOportunidad = True
   Exit Function
   
no_Be:
   Debuga Format(Now, "dd hh:mm:ss") & " * (" & MySql & ")" & err.Description
      ExecutaComandaSqlSegunDaOportunidad = False
  
End Function


Function DonamSql(MySql As String) As String
    Dim Rs As rdoResultset
   
    DonamSql = ""
    Set Rs = Db.OpenResultset(MySql)
    
    If Not Rs.EOF Then DonamSql = Rs(0)
   
   
   
End Function


Sub Debuga(s As String)
   Static Comprovat As Boolean
   Static Debugant As Boolean
   Dim f
   
   If Not Comprovat Then
      Comprovat = True
   End If
   Debug.Print s
   
End Sub


Sub FesElConnect()
   Dim lin As String, P As Integer
   
   'Lin = Trim(Command)
   lin = "05400004914  10.1.2.16  hit  "
   lin = "05400004914  SERVERCLOUD  hit  "
   lin = "05400004914  silema.hiterp.com  hit  "
   Sempre = False
   If InStr(UCase(lin), UCase(" Hit ")) Then Sempre = True
   
   P = InStr(lin, " ")
   If P > 0 Then
      Cfg_Llicencia = Trim(Left(lin, P))
'      Cfg_Llicencia = GeneraClau(DiscSerialNumber(), Mid(Format(Cfg_Llicencia / 3, "0000000000"), 3, 5))
      lin = Trim(Right(lin, Len(lin) - P)) & " "
   End If
   
   P = InStr(lin, " ")
   If P > 0 Then
      Cfg_Server = Trim(Left(lin, P))
      lin = Trim(Right(lin, Len(lin) - P)) & " "
   End If
   
   P = InStr(lin, " ")
   If P > 0 Then
      Cfg_Database = Trim(Left(lin, P))
      lin = Trim(Right(lin, Len(lin) - P)) & " "
   End If
   
'   If InStr(UCase(Lin), UCase("EsMud")) Then SistemaMud = True
'   If InStr(UCase(Lin), UCase("sistemaobert")) Then
   SistemaObert = True
   EsDispacher = False
   FeinaAfer = "Tot"
   If InStr(UCase(Command), UCase("Sincro")) Then FeinaAfer = "Sincro"
   If InStr(UCase(Command), UCase("Calculs")) Then FeinaAfer = "Calculs"
   If InStr(UCase(Command), UCase("CalculsResi")) Then FeinaAfer = "CalculsResi"
   If InStr(UCase(Command), UCase("CalculsLlargs")) Then FeinaAfer = "CalculsLlargs"
   If InStr(UCase(Command), UCase("CalculsLlarcs")) Then FeinaAfer = "CalculsLlargs"
   If InStr(UCase(Command), UCase("Emails")) Then FeinaAfer = "Emails"
   If InStr(UCase(Command), UCase("RevisaEmails")) Then FeinaAfer = "RevisaEmails"
   If InStr(UCase(Command), UCase("SFtp")) Then FeinaAfer = "SFtp"
   If InStr(UCase(Command), UCase("CalculsCurts")) Then FeinaAfer = "CalculsCurts"
   If InStr(UCase(Command), UCase("Envia")) Then FeinaAfer = "Envia"
   If InStr(UCase(Command), UCase("Reb")) Then FeinaAfer = "Reb"
   If InStr(UCase(Command), UCase("Idf:")) Then FeinaAfer = Trim(Command)
   If InStr(UCase(Command), UCase("Dispacher")) Then EsDispacher = True
   
   
   If Len(Cfg_Server) = 0 Or Len(Cfg_Database) = 0 Or Len(Cfg_Llicencia) = 0 Then End
   
   If SistemaObert Then Cfg_Llicencia = GeneraClau(DiscSerialNumber(), Mid(Format(Cfg_Llicencia / 3, "0000000000"), 3, 5))
   Connecta 4, Cfg_Llicencia, Cfg_Server, Cfg_Database, 0

End Sub
Sub Main()
   Dim i As Integer, Record As String, EmpresaDebug As String
    frmSplash.Debugant = False
    PosaSeparadorDecimal
   Velocitat = 20
   AppPath = App.Path
   EmailGuardia = EmailGuardia
   
   If MesDeUnCop And InStr(UCase(Command), "DEBUG") = 0 Then End
   
   
   frmSplash.Show
   frmSplash.Refresh
   
   Informa "Definint Entorn "
   
   NomServerInternet = "Adsl"
   FesElConnect
   CarregaLlistaDeFeines feina
   While UBound(feina) = 1 And InStr(UCase(Command), "HIT") > 0
      FesElConnect
      CarregaLlistaDeFeines feina
   Wend
         
   UltimaAccio = DateAdd("n", -30, Now)
   Velocitat = 5
   If FeinaAfer = "Sincro" Or FeinaAfer = "Envia" Or FeinaAfer = "Reb" Then Velocitat = 60 * 10
   If FeinaAfer = "CalculsLlargs" Then Velocitat = 20
   If FeinaAfer = "Emails" Then Velocitat = 20
   If FeinaAfer = "RevisaEmails" Then Velocitat = 20
   If FeinaAfer = "SFtp" Then Velocitat = 600
   If FeinaAfer = "CalculsResi" Then Velocitat = 60 * 10
   frmSplash.Metronom.Interval = 1000
   frmSplash.Metronom.Enabled = True
   
End Sub

Function llegeigHtml(Str) As String
    'Dim downHTTP As New HTTP, K, Lafrase As String
    
On Error GoTo err

    'Lafrase = ""
    
    'downHTTP.URL = Str
    'downHTTP.Download
    'If Not downHTTP.HasError Then
     '  Lafrase = Left(downHTTP.DownloadedContent, 32567)
'       Debug.Print Lafrase
    'End If
    
    'llegeigHtml = Lafrase
    
'--------------
Dim hOpen As Long
Dim hFile As Long
Dim sBuffer As String * 128
Dim Ret As Long
Dim str_Total As String
Dim URL As String

URL = Str
 ' Abrimos una conexión a internet
    hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_DIRECT, _
                         vbNullString, vbNullString, 0)
                           
    ' Si devuelve 0 es por que o no hay conexión a internet u otro error
    If hOpen = 0 Then
        Exit Function
    Else
        'Abrimos la url
        hFile = InternetOpenUrl(hOpen, Trim$(URL), vbNullString, _
                            ByVal 0&, INTERNET_FLAG_NO_CACHE_WRITE, ByVal 0&)
    End If
      
    If hFile = 0 Then
       'Error
       Exit Function
    Else
        'Lee una porción del fichero ( 128 bytes )
        Call InternetReadFile(hFile, sBuffer, 128, Ret)
          
        str_Total = sBuffer
          
        While Ret <> 0
            'Lee de 128 bytes. Cuando ret devuelve 0 finalizó
            Call InternetReadFile(hFile, sBuffer, 128, Ret)
              
            'Va acumulando el archivo para luego asignarlo al RichTextBox
            str_Total = str_Total & Mid(sBuffer, 1, Ret)
            
            DoEvents
        Wend
      
    End If
      
    'Cerramos el handle anterior (del archivo y de la conexión a internet )
    Call InternetCloseHandle(hFile)
    Call InternetCloseHandle(hOpen)
    
    llegeigHtml = Trim(str_Total)
'------------
    
    

err:

End Function


Private Function GeneraClau(Maq As Double, Lic As Double) As String
   Dim Ac As Double, crc As Double, Disc As Double, CrcDisc As Double, CodiLlicencia As String, Valor As String, NomEmpresa As String, Llic As String
   
   Disc = Maq Mod 970
   Llic = Format(Lic, "00000") & Format(Disc, "000")
   crc = SumaDeDigits(Llic) Mod 28
   
   GeneraClau = Format(crc, "00") & Format(Lic, "00000") & Format(Disc, "000")

   GeneraClau = Format(GeneraClau * 3, "00000000000")
End Function

    
    
Sub Reconecta()
   Dim i
   On Error Resume Next
   
   While Not Connecta(4, LastLlicencia, LastServer, LastDatabase, EmpresaActualNum)
      If i = 1 Then Informa2 "Reconectant ."
      If i = 2 Then Informa2 "Reconectant .."
      If i = 3 Then Informa2 "Reconectant ..."
      If i = 4 Then Informa2 "Reconectant ...."
      If i = 0 Then Informa2 "Reconectant ....."
      i = i + 1
      i = i Mod 5
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      DoEvents
   Wend
End Sub

Private Function SumaDeDigits(st As String) As Double
   Dim Ac As Double, i As Integer, c As String
   
   Ac = 0
   For i = 1 To Len(st)
      c = Mid(st, i, 1)
      If IsNumeric(c) Then Ac = Ac + Val(c)
   Next
   SumaDeDigits = Ac

End Function







Private Function DiscSerialNumber() As Long
   Dim res As Long
   Dim lpRootPathName As String
   Dim lpVolumeNameBuffer As String
   Dim nVolumeNameSize As Long
   Dim lpVolumeSerialNumber As Long
   Dim lpMaximumComponentLength As Long
   Dim lpFileSystemFlags As Long
   Dim lpFileSystemNameBuffer As String
   Dim nFileSystemNameSize As Long
   
   
   lpRootPathName = "C:\"
   nVolumeNameSize = 100
   
   res = GetVolumeInformation(lpRootPathName, lpVolumeNameBuffer, nVolumeNameSize, lpVolumeSerialNumber, lpMaximumComponentLength, lpFileSystemFlags, lpFileSystemNameBuffer, nFileSystemNameSize)
   
   If res = 1 Then
      DiscSerialNumber = lpVolumeSerialNumber
   Else
      DiscSerialNumber = 0
   End If
   
   If DiscSerialNumber < 0 Then DiscSerialNumber = DiscSerialNumber * -1
   
End Function


Public Sub Informa(s As String, Optional AvisaGos As Boolean = False)
   If AvisaGos Then AvisaAlGos s

   frmSplash.Estat.Caption = s
On Error Resume Next
   frmSplash.Estat.Visible = True
   My_DoEvents
   Debug.Print s
   
End Sub


Public Sub InformaEmpresa(s As String, Optional AvisaGos As Boolean = True)
   If AvisaGos Then AvisaAlGos "Inici De Feina Per " & s
   frmSplash.NomEmpresa.Caption = Format(Now(), "hh:mm") & " " & s
   My_DoEvents
End Sub



Public Sub Informa2(s As String, Optional AvisaGos As Boolean = False)
   
   If AvisaGos Then AvisaAlGos s
   
   frmSplash.lblVersion.Caption = s
   If Not frmSplash.lblVersion.Visible Then frmSplash.lblVersion.Visible = True
   DoEvents
   
   Debug.Print s
   
End Sub



Function MesDeUnCop() As Boolean
   Dim nomfile As String
   
   MesDeUnCop = True
   
   FeinaAfer = "Tot"
   If InStr(UCase(Command), UCase("Sincro")) Then FeinaAfer = "Sincro"
   If InStr(UCase(Command), UCase("Calculs")) Then FeinaAfer = "Calculs"
   If InStr(UCase(Command), UCase("llargs")) Then FeinaAfer = "CalculsLlargs"
   If InStr(UCase(Command), UCase("resi")) Then FeinaAfer = "CalculsResi"
   If InStr(UCase(Command), UCase("llarcs")) Then FeinaAfer = "CalculsLlargs"
   If InStr(UCase(Command), UCase("curts")) Then FeinaAfer = "CalculsCurts"
   If InStr(UCase(Command), UCase("Emails")) Then FeinaAfer = "Emails"
   If InStr(UCase(Command), UCase("RevisaEmails")) Then FeinaAfer = "RevisaEmails"
   If InStr(UCase(Command), UCase("SFtp")) Then FeinaAfer = "SFtp"
   If InStr(UCase(Command), UCase("Envia")) Then FeinaAfer = "Envia"
   If InStr(UCase(Command), UCase("Reb")) Then FeinaAfer = "Reb"
   If InStr(UCase(Command), UCase("IDF:")) Then FeinaAfer = Command
   
   nomfile = App.Path & "\" & GetNomMaquina & "_" & App.EXEName & "_" & Replace(FeinaAfer, ":", "") & ".txt"
   On Error Resume Next
      Kill nomfile
   On Error GoTo algun
   Bandera = FreeFile
   Open nomfile For Output Lock Read Write As #Bandera
    
   MesDeUnCop = False
   
algun:

End Function



Function GetNomMaquinaSql() As String
   Dim nom As String
   
   nom = Space(100)
   GetComputerName nom, Len(nom)
   nom = Trim(nom)
   nom = Left(nom, Len(nom) - 1)
   GetNomMaquinaSql = nom
End Function


Function GetNomMaquina() As String
   Dim nom As String
   
   If Len(Cnf.llicencia) > 0 Then
      nom = Cnf.llicencia
   Else
      nom = Space(100)
      GetComputerName nom, Len(nom)
      nom = Trim(nom)
      nom = Left(nom, Len(nom) - 1)
   End If
   
   GetNomMaquina = nom

End Function



Sub SincronitzaEmpresa(empresa As String, Accio As String)

    PdaEnviaTot = False
    Select Case Accio
        Case "Envia"
            SincronitzaEmpresaEnvia empresa
        Case "Reb"
            SincronitzaEmpresaReb empresa
        Case Else
            SincronitzaEmpresaEnvia empresa
            SincronitzaEmpresaReb empresa
    End Select
    
'    SincronitzaPda
   
End Sub
Sub EsborraTaula(nom As String)
   
   If ExisteixTaula(nom) Then ExecutaComandaSql "Drop Table [" & nom & "]"

End Sub


Function Car(ByRef s As String) As String
   Dim P As Integer, p1 As Integer, P2 As Integer
   Dim c As String, Acc As Integer, Result As String
   Dim Le As Double
   
   Acc = 0
   Result = ""
   Le = Len(s)
   While Le > 0
      c = Left(s, 1)
      Le = Le - 1
      s = Right(s, Le)
      Result = Result & c
      Select Case c
         Case "[": Acc = Acc + 1
         Case "]":
            Acc = Acc - 1
            If Acc <= 0 Then
                If Acc = -1 Then
                   Car = Mid(Result, 2, Len(Result) - 1)
                Else
                   Car = Mid(Result, 2, Len(Result) - 2)
                End If
               Exit Function
            End If
      End Select
   Wend
   Car = ""
   s = ""
   
End Function


Function dataSetmana(fec)
    Dim fec2, temp, lunes, domingo
    If fec = "" Or Not IsDate(fec) Then
        fec2 = Date
    Else
        fec2 = CDate(fec)
    End If
    temp = Weekday(fec2) 'devuelve 1-domingo a 7-sábado
    lunes = temp - 2
    If lunes = -1 Then lunes = 6
    lunes = DateAdd("d", lunes * (-1), fec2)
    domingo = DateAdd("d", 6, lunes)
    dataSetmana = CStr(lunes) & "," & CStr(domingo)
End Function

Sub TancaDb()
On Error Resume Next
'Db.Close
On Error GoTo 0
End Sub


Sub tancasipots()
On Error Resume Next
Db.Close
End Sub

Function WwwCache() As String
    Dim sql As String
    
    If Not ExisteixTaula("WwwCache") Then
      sql = "CREATE TABLE WwwCache ( "
      sql = sql & " [Id]            [nvarchar] (255) NULL ,"
      sql = sql & " [NomFile]       [nvarchar] (255) NULL ,"
      sql = sql & " [TimeStamp]     [datetime] NULL ,"
      sql = sql & " [Original]      [image] (255) NULL ,"
      sql = sql & " [Resum]         [nvarchar] (255) NULL "
      ExecutaComandaSql sql
    End If
    
    WwwCache = "WwwCache"
    
End Function

Sub PosaSeparadorDecimal()


On Error Resume Next
    'Establece a "," el símbolo del separador decimal para números
    SetLocaleInfo &H400, 14, "."
    'Establece a "." el símbolo de separación de miles para números
    SetLocaleInfo &H400, 15, ","
    'Establece a "," el símbolo del separador decimal para moneda
    SetLocaleInfo &H400, 22, "."
    'Establece a "." el símbolo de separación de miles para moneda
    SetLocaleInfo &H400, 23, ","


End Sub

