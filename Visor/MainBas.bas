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
   Empresa As String
   Db As String
   Path As String
   Llicencia As String
   Server As String
   Tipus As Integer
   EscoltaLlicencies() As String
   Ftp_Server As String
   Ftp_User As String
   Ftp_Pssw As String
End Type

Global Feina() As TipFeina, AppPath As String, UltimaAccio As Date, Velocitat As Double, EmpresaActualNum As Integer, FinDelMundo As Boolean
Global LastServer As String, LastDatabase As String, SistemaMud As Boolean, SistemaObert As Boolean, FeinaAfer As String, Sempre As Boolean, EmpresaActual As String, LastLlicencia As String, ServerActual As String, PdaEnviaTot As Boolean, EsDispacher As Boolean, Obelles() As String






Sub FesLaFeina()
    Dim Rs As rdoResultset, i, Limit As Double
      
On Error GoTo nor
    ReDim Obelles(0)
    Db.Execute "delete gosdetura where DATEDIFF(DAY , tsvista  , GETDATE() ) >5 "
    Informa "Vigilant les obelles"
    While Not FinDelMundo
        '
        Set Rs = Db.OpenResultset("Select DATEDIFF(MINUTE,TsVista,GETDATE()) minuts,* from GosDeTura order by NomObella ")
        While Not Rs.EOF
            For i = 0 To UBound(Obelles)
                If Obelles(i) = Rs("NomObella") Then Exit For
            Next
            If i > UBound(Obelles) Then
                ReDim Preserve Obelles(i)
                Obelles(i) = Rs("NomObella")
                Load frmSplash.Obandera(i)
                Load frmSplash.OEstat(i)
                Load frmSplash.oNom(i)
                If i = 1 Then
                    frmSplash.Obandera(i).Top = frmSplash.Estat.Top + frmSplash.Estat.Height + 300
                Else
                    frmSplash.Obandera(i).Top = frmSplash.Obandera(i - 1).Top + frmSplash.Obandera(i - 1).Height + 10
                End If
                
                frmSplash.Obandera(i).Left = frmSplash.Command2.Left
                frmSplash.oNom(i).Top = frmSplash.Obandera(i).Top
                frmSplash.oNom(i).Left = frmSplash.Obandera(i).Left + frmSplash.Obandera(i).Width + 20
                frmSplash.OEstat(i).Top = frmSplash.Obandera(i).Top
                frmSplash.OEstat(i).Left = frmSplash.oNom(i).Left + frmSplash.oNom(i).Width + 20
                frmSplash.OEstat(i).Width = frmSplash.Width - frmSplash.OEstat(i).Left
                
                frmSplash.Obandera(i).Visible = True
                frmSplash.OEstat(i).Visible = True
                frmSplash.oNom(i).Visible = True
                
            End If
            
            frmSplash.OEstat(i).Caption = Rs("ultimafraseobella")
            frmSplash.oNom(i).Caption = Rs("nomobella")
            
            Limit = Split(Rs("ObellaPeriodicitat"), " ")(1)
            
            If Rs("minuts") > Limit Then
                frmSplash.Obandera(i).BackColor = &HFF&       ' super vermell
            Else
                frmSplash.Obandera(i).BackColor = &H80FF80    'verd
            End If
            
            Rs.MoveNext
            DoEvents
            If FinDelMundo Then Exit Sub
        Wend
        
        DoEvents
        If FinDelMundo Then Exit Sub
    Wend
nor:

End Sub

Sub Init()
    Dim Rs
    
    FesElConnect
    
    Set Rs = Db.OpenResultset("select * from web_empreses order by nom ")
    frmSplash.LListaEmpreses.Clear
    While Not Rs.EOF
        frmSplash.LListaEmpreses.AddItem Rs("Nom")
        Rs.MoveNext
    Wend
    
    
End Sub

Sub FesElConnect()
   Dim lin As String, P As Integer
   
   'Lin = Trim(Command)
   lin = "05400004914  10.1.2.16  hit  "
   lin = "05400004914  SERVERCLOUD  hit  "
   
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
   If InStr(UCase(Command), UCase("CalculsCurts")) Then FeinaAfer = "CalculsCurts"
   If InStr(UCase(Command), UCase("Envia")) Then FeinaAfer = "Envia"
   If InStr(UCase(Command), UCase("Reb")) Then FeinaAfer = "Reb"
   If InStr(UCase(Command), UCase("Idf:")) Then FeinaAfer = Trim(Command)
   If InStr(UCase(Command), UCase("Dispacher")) Then EsDispacher = True
   
   
   If Len(Cfg_Server) = 0 Or Len(Cfg_Database) = 0 Or Len(Cfg_Llicencia) = 0 Then End
   
   If SistemaObert Then Cfg_Llicencia = GeneraClau(DiscSerialNumber(), Mid(Format(Cfg_Llicencia / 3, "0000000000"), 3, 5))
   Connecta 4, Cfg_Llicencia, Cfg_Server, Cfg_Database, 0

End Sub
Function Connecta(Tipus As Integer, Llicencia As String, Server As String, Database As String, i) As Boolean

   EmpresaActualNum = i
   Connecta = False
   
   LastServer = Server
   LastDatabase = Database
   LastLlicencia = Llicencia
'If UCase(Server) = UCase("Titan") Then

   Connecta = ConnectaSqlServer(LastServer, LastDatabase)
   
End Function


Function ConnectaSqlServer(Server As String, NomDb As String) As Boolean
   Dim User As String, Psw As String, MyId As String, Rs As rdoResultset, Modus
   ConnectaSqlServer = False

On Error Resume Next
'   Db.Close
On Error GoTo nor
   If InStr(Command, "192.9.199.202") > 0 And UCase(Server) = UCase("juliet") Then Server = "192.9.199.202"
   'Server = "172.26.0.101"
   User = "sa"
   Psw = ""
   
   
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
   
   Modus = "Exclusiu_"
   If FeinaAfer = "CalculsCurts" Or InStr(FeinaAfer, "Idf:") > 0 Then Modus = "NoExlusiu_"
   MyId = Modus & GetNomMaquinaSql() & "_" & App.Title & "_" & FeinaAfer
        
   If UCase(NomDb) = "HIT" Then MyId = "Global_" & GetNomMaquinaSql() & "_" & App.Title & "_" & FeinaAfer
   
   db2MyId = "SegonaConnexio_" & MyId
   db2User = User
   db2Psw = Psw
   db2NomDb = NomDb
   db2Server = Server
   
   'Db.Connect = "WSID=" & MyId & ";UID=" & User & ";PWD=" & Psw & ";Database=" & NomDb & ";Server=" & Server & ";Driver={SQL Server};DSN='';"
   If Db.Name = "" Then
On Error Resume Next
       Db.Close
On Error GoTo nor
       Db.Connect = "WSID=" & MyId & ";UID=" & User & ";PWD=" & Psw & ";Database=" & NomDb & ";Server=" & Server & ";Driver={SQL Server};DSN='';"
       Db.EstablishConnection rdDriverNoPrompt
'       Db.Connect = "WSID=" & MyId & ";UID=" & User & ";PWD=" & Psw & ";Database=" & NomDb & ";Server=" & Server & ";Driver={SQL Server};DSN='';"
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
    frmSplash.Caption = Server
    
Exit Function

nor:

Debug.Print Err.Description
Resume Next
End Function


Function GetNomMaquinaSql() As String
   Dim nom As String
   
   nom = Space(100)
   GetComputerName nom, Len(nom)
   nom = Trim(nom)
   nom = Left(nom, Len(nom) - 1)
   GetNomMaquinaSql = nom
End Function



Private Function GeneraClau(Maq As Double, Lic As Double) As String
   Dim Ac As Double, crc As Double, Disc As Double, CrcDisc As Double, CodiLlicencia As String, Valor As String, NomEmpresa As String, Llic As String
   
   Disc = Maq Mod 970
   Llic = Format(Lic, "00000") & Format(Disc, "000")
   crc = SumaDeDigits(Llic) Mod 28
   
   GeneraClau = Format(crc, "00") & Format(Lic, "00000") & Format(Disc, "000")

   GeneraClau = Format(GeneraClau * 3, "00000000000")
End Function

    
    

Function SumaDeDigits(St As String) As Double
   Dim Ac As Double, i As Integer, c As String
   
   Ac = 0
   For i = 1 To Len(St)
      c = Mid(St, i, 1)
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



Sub Main()
    If MesDeUnCop Then End
    frmSplash.Show
    Informa "Definint Entorn "
    
    Init
    
    FesLaFeina
    
    End
    
End Sub


Public Sub Informa(s As String, Optional AvisaGos As Boolean = False)


   frmSplash.Estat.Caption = s
On Error Resume Next
   frmSplash.Estat.Visible = True
   My_DoEvents
   Debug.Print s
   
End Sub



Sub My_DoEvents()
       DoEvents
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
   If InStr(UCase(Command), UCase("Envia")) Then FeinaAfer = "Envia"
   If InStr(UCase(Command), UCase("Reb")) Then FeinaAfer = "Reb"
   If InStr(UCase(Command), UCase("IDF:")) Then FeinaAfer = Command
   
   nomfile = App.Path & "\" & GetNomMaquina & "_" & App.EXEName & "_" & FeinaAfer & ".txt"
   On Error Resume Next
      Kill nomfile
   On Error GoTo Algun
   Bandera = FreeFile
   Open nomfile For Output Lock Read Write As #Bandera
    
   MesDeUnCop = False
   
Algun:

End Function


Function GetNomMaquina() As String
   Dim nom As String
   
      nom = Space(100)
      GetComputerName nom, Len(nom)
      nom = Trim(nom)
      nom = Left(nom, Len(nom) - 1)

   
   GetNomMaquina = nom

End Function




