Attribute VB_Name = "BasGen"
Option Explicit

'*********** Uff
' Cabecera fichero UFF
Private CRC32Table() As Long
Private BitTable() As Long
Private MaskTable() As Long

Private Type THeaderUFF
    MagicWord As String * 4   'PUFF
    crc32 As Long ' 32 bits   'crc 32 bits Intel-Format
    FormatUFF As Byte         'Formato UFF, de momento solo valido 0
    NumBlocs As Byte          'Numero de bloques en fichero
End Type

Private Const MAGIC_WORD_UFF As String * 4 = "PUFF"
Private Const CRC32_Inicial As Long = &HA5A5A5A5
Private Const FORMAT_UFF_0 As Byte = 0
Private Const FORMAT_UFF_DEFAULT As Byte = FORMAT_UFF_0
Private Const MAX_NumBlocs As Integer = 255
Private Const TD_SinCompresion As Byte = 1
Private Const TD_Compresion1 As Byte = 2



Private Type THeaderBloc1
    LenBloc As Long
    DataType As Byte
    ' Ruta , StringZ
    ' Datos
End Type


' Información para proceso de UFF
Private Type TInfoUff
    canal As Integer
    Formato As Byte
    NumBlocs As Byte
End Type

Private Type TAliasFile
    Temporal As String
    Destino As String
    fDestinoEscrito As Boolean
End Type

Private Enum TErrUFF
    No_ErrUFF = 0
    ErrUFF_Write
    ErrUFF_Read
    ErrUFF_ReadFileBloc
    ErrUFF_WriteFileBloc
    ErrUFF_Corrupted
    ErrUFF_InvalidData
    ErrUFF_Undefined
End Enum

Private Type Typ_Cnf
   llicencia As String
   empresa As String
   Validada As Boolean
   AppName As String
   AppPath As String
   AppDescripcio As String
End Type

Private Type netTime
    Hour As Integer
    Min As Integer
    Sec As Integer
    Day As Integer
    Month As Integer
    Year As Integer
End Type

Dim Hores() As netTime

Private Type Message
    From As String
    Subject As String
    Date As String
    AllHeaders As String
    Size As Long
    Text As String
End Type

Private Type Tip_CalEnviar
    Subject As String
    Missatge As String
    Files() As String
End Type

Dim gMessages(1 To 200) As Message
Dim tMailServer  As String

Dim CalEnviar() As Tip_CalEnviar
Global ClassEstat, ClassEstat2
Global dbSage

Global LlistaContingutsNom() As String
Global LlistaContingutsInteresa() As Boolean
Global LlistaContingutsEsborrem() As Boolean

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long


'***********
Global Const SYNCHRONIZE = &H100000
Global Const INFINITE = &HFFFFFFFF

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
'**************

Const SEC_DAY& = 3600 * 24&
Const SEC_YEAR& = SEC_DAY * 365&   '{1900 was no leap year, therefore one less}
Const SEC_TILL_1968& = ((365& * 4 + 1) * 17 - 1) * 24 * 3600
Dim Clock_HoresTrobades As Integer, Clock_HoresMaxim As Integer


Global FtpServer() As String, FTPUSER() As String, FtpPsw() As String, FtpServerNum As Integer, FtpServerUltimMaster As Integer
Global TimeOutFtp As Date

Global Cnf As Typ_Cnf
Global EstadisticHoraInici As Date, EstadisticFilesAgafades As Integer, EstadisticFilesDeixades As Integer, EstadisticRecarregat As Integer
Global EstatExecutant As Boolean, EstatError As String
Global DiesMemoria As Integer ' Dies que es guarden a internet
Global LastConexio As Date
Global DirectoriDesti_Tot() As String, DirectoriDesti_NomFile()  As String, DirectoriDesti_EsDirectori() As Boolean, DirectoriDesti_Data()  As Date, DirectoriDesti_Kb() As Double, DirectoriDesti_Ok() As Boolean

Global Erronis_Noms() As String
Global Erronis_Intents() As Integer
Global Erronis_Descarregat() As Boolean
Global GlobalConnectStr  As String






Sub CarregaFtpDir(Server, User, Psw)
   
   If Psw = "1" Then Psw = "jordi"
   
   ReDim Preserve FtpServer(UBound(FtpServer) + 1)
   ReDim Preserve FTPUSER(UBound(FTPUSER) + 1)
   ReDim Preserve FtpPsw(UBound(FtpPsw) + 1)
   If EsTelefon(GlobalConnectStr) Or UCase(GlobalConnectStr) = "ADSL" Then
      FtpServer(UBound(FtpServer)) = Server
      FTPUSER(UBound(FTPUSER)) = User
      FtpPsw(UBound(FtpPsw)) = Psw
   Else
      FtpServer(UBound(FtpServer)) = GlobalConnectStr
      FTPUSER(UBound(FTPUSER)) = User & "@" & Server
      FtpPsw(UBound(FtpPsw)) = Psw
   End If

End Sub

Function CarregaCfgLlicencia(NomFil As String) As Boolean
   Dim Valor As String, Tmp_Empresa As String
   Dim f, L As String, OldDir As String
   Dim Hor As String, p1, P2
   Dim Tmp_Internet_NumT As String, Tmp_Internet_User As String, Tmp_Internet_Pswd As String
                 
   CarregaCfgLlicencia = False
   f = FreeFile
   
   If Len(Dir(NomFil)) > 0 Then
      Open NomFil For Input As f
      While Not EOF(f)
         Line Input #f, L
         L = Trim(L)
         p1 = InStr(L, ":")
         If Left(L, 1) = "#" Or Len(L) = 0 Or p1 = 0 Then
         Else
            Valor = ""
            p1 = InStr(L, ":")
            Valor = Trim(Mid(L, p1 + 1, Len(L) - p1))
            L = Left(L, p1 - 1)
            Select Case Trim(UCase(L))
              Case "EMPRESA":
                 Tmp_Empresa = Valor
            End Select
         End If
      Wend
      Close f
      CarregaCfgLlicencia = True
      
      If Len(Tmp_Empresa) > 0 Then
         Cnf.empresa = Tmp_Empresa
      End If
   End If
   
Exit Function

Loga:
   Resume Next

End Function



Function CvFtpData(c As String) As Date
   Dim P As Integer
   
   P = 0
   If P = 0 Then P = InStr(c, "Jan")
   If P = 0 Then P = InStr(c, "Feb")
   If P = 0 Then P = InStr(c, "Mar")
   If P = 0 Then P = InStr(c, "Apr")
   If P = 0 Then P = InStr(c, "May")
   If P = 0 Then P = InStr(c, "Jun")
   If P = 0 Then P = InStr(c, "Jul")
   If P = 0 Then P = InStr(c, "Aug")
   If P = 0 Then P = InStr(c, "Sep")
   If P = 0 Then P = InStr(c, "Oct")
   If P = 0 Then P = InStr(c, "Nov")
   If P = 0 Then P = InStr(c, "Dec")
   
   If P > 0 Then c = Right(c, Len(c) - P + 1) Else c = c
   CvFtpData = TrobaData(c)

End Function

Sub EnviaRecepcio()
    Dim sql As String, Rs As rdoResultset, Q As rdoQuery
    
    If Not UCase(EmpresaActual) = UCase("daza") Then Exit Sub
On Error GoTo nor
    sql = ""
    sql = sql & "Select isnull(m.ubicacion,'000') Seccio,r.fecha ,r.Id As Rid,lote,albaran,r.caducidad,pr.codi,m.nombre,m.codigo,p.precio,cantidad "
    sql = sql & "from ccrecepcion r "
    sql = sql & "join ccproveedores    pr on facturado = 0 And r.proveedor = pr.id  and not pr.codi is null "
    sql = sql & "join ccMateriasPrimas m  on r.MatPrima  = m.id and not m.codigo is null "
    sql = sql & "join ccPedidos        p  on r.pedido  = p.id "
    sql = sql & "order by r.fecha desc "
    
    Set Q = Db.CreateQuery("", "Insert into CalEnviar (Id,Taula,TmSt,Tipo,Param1,Param2,Param3,Param4,Param5,Param6,Param7,Param8,Param9) Values (NewId(),'',getdate(),'Recepcio',?,?,?,?,?,?,?,?,?) ")
    
    Set Rs = Db.OpenResultset(sql)
    While Not Rs.EOF
        If Not (IsNull(Rs("nombre")) Or IsNull(Rs("Codi")) Or IsNull(Rs("Codigo")) Or IsNull(Rs("Cantidad")) Or IsNull(Rs("Precio")) Or IsNull(Rs("Fecha"))) Then
            Q.rdoParameters(0) = Rs("nombre")
            Q.rdoParameters(1) = Rs("Codi")
            Q.rdoParameters(2) = Rs("Codigo")
            Q.rdoParameters(3) = Rs("Cantidad")
            Q.rdoParameters(4) = Rs("Precio")
            Q.rdoParameters(5) = Rs("Fecha")
            Q.rdoParameters(6) = Rs("Lote")
            Q.rdoParameters(7) = Rs("Albaran")
            Q.rdoParameters(8) = Rs("Seccio")
            Q.Execute
            
            ExecutaComandaSql "Update ccrecepcion set Facturado = 1 Where id = '" & Rs("Rid") & "' "
        End If
        
        Rs.MoveNext
    Wend
nor:

End Sub

Sub ExecutaScripts()
   Dim Fil As String, NomSc As String
   
   'NomSc = Cnf.AppPath & "\Tmp\ExeScripts.vbs"
'   On Error Resume Next
'   Kill NomSc
   
'   Fil = Dir(Cnf.AppPath & "\Tmp\*.vbs")
'   If Len(Fil) Then
'      GeneraScript NomSc
'      Shell "wscript.exe " & NomSc, vbNormalFocus
'   End If
      
         
      
End Sub

Sub PreparaVersio()
   Dim Fil As String, NomSc As String, f
  
   If Len(Dir(Cnf.AppPath & "\Tmp\HitConn.Exe")) > 0 Then
   If Len(Dir(Cnf.AppPath & "\Tmp\Vpos.Exe")) > 0 Then
   If Len(Dir(Cnf.AppPath & "\Tmp\Toc.Exe")) > 0 Then
   If Len(Dir(Cnf.AppPath & "\Tmp\FalcoNet.exe")) > 0 Then
      On Error Resume Next
      MkDir Cnf.AppPath & "\Versio"
            
      Kill Cnf.AppPath & "\Versio\HitConn.Exe"
      Name Cnf.AppPath & "\Tmp\HitConn.Exe" As Cnf.AppPath & "\Versio\HitConn.Exe"
      Kill Cnf.AppPath & "\Versio\Vpos.Exe"
      Name Cnf.AppPath & "\Tmp\Vpos.Exe" As Cnf.AppPath & "\Versio\Vpos.Exe"
      Kill Cnf.AppPath & "\Versio\Toc.Exe"
      Name Cnf.AppPath & "\Tmp\Toc.Exe" As Cnf.AppPath & "\Versio\Toc.Exe"
      Kill Cnf.AppPath & "\Versio\FalcoNet.Exe"
      Name Cnf.AppPath & "\Tmp\FalcoNet.Exe" As Cnf.AppPath & "\Versio\FalcoNet.Exe"
      
      Kill Cnf.AppPath & "\Versio.Mk"
      f = FreeFile
      Open Cnf.AppPath & "\Versio.Mk" For Output As #f
           
      Print #f, "File:\HitConn.Exe"
      Print #f, "Accio:Backup"
      Print #f, "File:\Vpos.Exe"
      Print #f, "Accio:Backup"
      Print #f, "File:\Toc.Exe"
      Print #f, "Accio:Backup"
      Print #f, "File:\FalcoNet.Exe"
      Print #f, "Accio:Backup"
                        
      Print #f, "File:\Versio\HitConn.Exe"
      Print #f, "Accio:Copy"
      Print #f, "Accio:Registra"
           
      Print #f, "File:\Versio\Vpos.Exe"
      Print #f, "Accio:Copy"
      Print #f, "Accio:Registra"
            
      Print #f, "File:\Versio\Toc.Exe"
      Print #f, "Accio:Copy"
      
      Print #f, "File:\Versio\FalcoNet.Exe"
      Print #f, "Accio:Copy"
            
      Print #f, "File:\Versio\HitConn.Exe"
      Print #f, "Accio:Delete"
      Print #f, "File:\Versio\Vpos.Exe"
      Print #f, "Accio:Delete"
      Print #f, "File:\Versio\Toc.Exe"
      Print #f, "Accio:Delete"
      Print #f, "File:\Versio\FalcoNet.Exe"
      Print #f, "Accio:Delete"
            
      Print #f, "File:\Versio"
      Print #f, "Accio:DeleteDir"
            
      Close f
   End If
   End If
   End If
   End If
   
   On Error Resume Next
   Kill NomSc
   
   Fil = Dir(Cnf.AppPath & "\Tmp\*.vbs")
   If Len(Fil) Then
      GeneraScript NomSc
      Shell "wscript.exe " & NomSc, vbNormalFocus
   End If
      
End Sub


Sub GeneraScript(nom As String)
   Dim Fil As String, f
   
   f = FreeFile
   Open nom For Output As f
   Print #f, "Function ArrancaScripts()"
   Print #f, "Path = DonamPath()"
   Print #f, "Set fso = CreateObject(""Scripting.FileSystemObject"")"
   Print #f, "Set Folder = fso.GetFolder(Path)"
   Print #f, "Set Files = Folder.Files"
   Print #f, ""
   Print #f, "on error resume next"
   Print #f, "   For Each File In Files"
   Print #f, "      if ucase(right(File.Name,4)) = Ucase("".Vbs"") And not ucase(Wscript.ScriptName) = ucase(File.Name)  then Inicia Path & File.Name"
   Print #f, "   Next"
   Print #f, "End function"
   Print #f, ""
   Print #f, "Sub Inicia(Nom)"
   Print #f, "   on error goto 0"
   Print #f, "   Set WshObj = Wscript.CreateObject(""WScript.Shell"")"
   Print #f, "   WshObj.Run Nom"
   Print #f, "   Espera"
   Print #f, "End Sub"
   Print #f, ""
   Print #f, "   Sub Mata(Nom)"
   Print #f, "   Set fso = CreateObject(""Scripting.FileSystemObject"")"
   Print #f, "   Path = DonamPath()"
   Print #f, "   i=0"
   Print #f, "   Bo=false"
   Print #f, "   Err.clear"
   Print #f, "   on error Resume next"
   Print #f, "   While not Bo"
   Print #f, "      i=i+1"
   Print #f, "      NomDesti = Path & ""PendentDelete_"" & i & "".DeadList"""
   Print #f, "      Set File = fso.CreateTextFile(NomDesti,False)"
   Print #f, "      if Err.Number = 0 then Bo=true"
   Print #f, "      Err.clear"
   Print #f, "   Wend"
   Print #f, "   File.WriteLine(Nom)"
   Print #f, "   File.Close"
   Print #f, "End Sub"
   Print #f, ""
   Print #f, "Sub MataMe"
   Print #f, "   Mata WScript.ScriptFullName "
   Print #f, "End Sub"
   Print #f, ""
   Print #f, "Sub Espera"
   Print #f, "   D = dateadd(""s"",2,now)"
   Print #f, "   while now < d"
   Print #f, "   wend"
   Print #f, "End Sub"
   Print #f, ""
   Print #f, "Function DonamPath()"
   Print #f, "   DonamPath = """""
   Print #f, "   P = instr(Wscript.ScriptFullName,Wscript.ScriptName)"
   Print #f, "   if p>0 then DonamPath = left(Wscript.ScriptFullName,P - 1)"
   Print #f, "End Function"
   Print #f, ""
   Print #f, "   ArrancaScripts"
   Print #f, "   MataMe"
   Print #f, ""
   Print #f, "WScript.Quit()"
   
   Close f

End Sub


Sub FilesPerReintentCarrega()
   Dim Clau As String, Valor As String, i As Integer, MisValores As Variant
   
   ReDim Erronis_Noms(0)
   ReDim Erronis_Intents(0)
   ReDim Erronis_Descarregat(0)
   
   MisValores = GetAllSettings("Hit_" & LCase(Cnf.llicencia), "FilesABaixar")
On Error GoTo cap
   For i = LBound(MisValores, 1) To UBound(MisValores, 1)
      ReDim Preserve Erronis_Noms(UBound(Erronis_Noms) + 1)
      ReDim Preserve Erronis_Intents(UBound(Erronis_Intents) + 1)
      ReDim Preserve Erronis_Descarregat(UBound(Erronis_Descarregat) + 1)
         
      Erronis_Noms(UBound(Erronis_Noms)) = MisValores(i, 1)
      Erronis_Intents(UBound(Erronis_Intents)) = MisValores(i, 0)
      Erronis_Descarregat(UBound(Erronis_Descarregat)) = False
   Next
cap:
   
End Sub




Sub TheEnd()
   
      
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      DoEvents
      End

End Sub

#Const COMPUTE_XOR_PATTERN = False
Private Function Uff_CalcCRC32(Car As Byte, crc32 As Long) As Long
On Error Resume Next
    Uff_CalcCRC32 = CRC32Table((crc32 Xor Car) And &HFF&) Xor ShiftL(crc32, 8)
    If err = 0 Then Exit Function
    Uff_CrearCRC32Table
    Uff_CalcCRC32 = CRC32Table((crc32 Xor Car) And &HFF&) Xor ShiftL(crc32, 8)
End Function


Public Function ShiftL(ByVal L As Long, Optional n As Byte = 1) As Long
    Dim i As Byte
    For i = 1 To n
       L = L And &H7FFFFFFF
       If ((L And &H40000000) <> 0) Then
          L = L And &H3FFFFFFF: L = L * 2: L = L Or &H80000000
       Else
          L = L * 2
       End If
    Next i
    ShiftL = L
End Function




Sub CreaDirs()
  Dim Fil As String, Fils() As String, i As Integer, f, L As String
  On Error Resume Next
     MkDir Cnf.AppPath & "\Msg"
     MkDir Cnf.AppPath & "\Msg\Cfg"
     MkDir Cnf.AppPath & "\Err"
     MkDir Cnf.AppPath & "\Bak"
     MkDir Cnf.AppPath & "\Tmp"
On Error GoTo 0
  
On Error Resume Next
  
  ReDim Fils(0)
  Fil = Dir(Cnf.AppPath & "\Tmp\*.DeadList")
  
  While Len(Fil) > 0
     ReDim Preserve Fils(UBound(Fils) + 1)
     Fils(UBound(Fils)) = Fil
     Fil = Dir
  Wend
  
  For i = 1 To UBound(Fils)
     Fil = Cnf.AppPath & "\Tmp\" & Fils(i)
     f = FreeFile
     Open Fil For Input As f
        While Not EOF(f)
           Line Input #f, L
           Kill L
        Wend
     Close f
     Kill Fil
  Next
  
End Sub


Public Function DonamHoraReal() As Date
'   Dim rnrServers As Integer, rServerAddresses() As String, rServerNames() As String, i As Integer, D1 As Date, Nota() As Double, j As Integer
'
'   ReDim rServerAddresses(200)
'   ReDim rServerNames(200)
'
'   CarregaServers rnrServers, rServerAddresses, rServerNames
'
'   IPClock.WinsockLoaded = True
'   If Not IPClock.Active Then
'      IPClock.RemotePort = 37
'      IPClock.Active = True
'   End If
'
'   Clock_HoresTrobades = 0
'   Clock_HoresMaxim = 15
'   ReDim Hores(Clock_HoresMaxim)
'   ReDim Nota(Clock_HoresMaxim)
'   For i = 1 To rnrServers
'      IPClock.RemoteHost = "servidornt" '& rServerAddresses$(i)
'      IPClock.DataToSend = "hello fron Hit"
'   Next i
'
'   D1 = DateAdd("n", 1, Now)
'   Do
'      DoEvents
'   Loop While Now < D1 And Clock_HoresTrobades < Clock_HoresMaxim
'
'   IPClock.Active = False
'
'   If Clock_HoresTrobades >= Clock_HoresMaxim Then
'      For i = 1 To Clock_HoresTrobades
'         Nota(i) = 0
'         For j = 1 To Clock_HoresTrobades
'            If Hores(i).Day = Hores(j).Day And _
'            Hores(i).Hour = Hores(j).Hour And _
'            Hores(i).Min = Hores(j).Min And _
'            Hores(i).Month = Hores(j).Month And _
'            Hores(i).Year = Hores(j).Year Then Nota(i) = Nota(i) + 1
'         Next
'      Next
'      Dim Maxima  As Integer
'      Dim el As Integer
'      Maxima = 0
'      el = 0
'      For i = 1 To Clock_HoresTrobades
'         If Maxima < Nota(i) Then
'            Maxima = Nota(i)
'            el = i
'         End If
'      Next
'
'      If Maxima > 3 Then
''         Hores(El).Day
'      End If
'   End If
   
End Function



Function EsTelefon(Tel As String) As Boolean
   Dim i As Integer, c As String
   
   EsTelefon = False
   If UCase(Tel) = "ADSL" Then Exit Function
   
   For i = 1 To Len(Tel)
      c = Mid(Tel, i, 1)
      If Not (IsNumeric(c) Or c = "," Or c = "." Or c = " ") Then Exit Function
   Next
   EsTelefon = True
   
End Function





Function ClauData() As String
   
   ClauData = "[Data#" & Format(Now, "yyyymmddhhnnss") & Format(Rnd(100) * 1000, "0000") & "]"

End Function





Sub Descomposa(kItem As Integer, nom As String, data As Date, EsDirectori As Boolean, Contingut As String, Maquina As String)
   Dim P As Integer, p1 As Integer, c As String
   Dim sD As String
      
On Error Resume Next
   nom = DirectoriDesti_NomFile(kItem)
   EsDirectori = DirectoriDesti_EsDirectori(kItem)
   If EsDirectori Then Exit Sub
   
   data = DirectoriDesti_Data(kItem)
    
   Contingut = ""
   P = InStr(DirectoriDesti_NomFile(kItem), "[Contingut#")
   If P > 0 Then
      P = P + 11
      p1 = InStr(P, DirectoriDesti_NomFile(kItem), "]")
      Contingut = Mid(DirectoriDesti_NomFile(kItem), P, p1 - P)
   End If
   
   Maquina = ""
   P = InStr(DirectoriDesti_NomFile(kItem), "[Maquina#")
   If P > 0 Then
      P = P + 9
      p1 = InStr(P, DirectoriDesti_NomFile(kItem), "]")
      Maquina = Mid(DirectoriDesti_NomFile(kItem), P, p1 - P)
   End If
   
End Sub






Function Interesa_EsPerAMi(Ctg As String) As Boolean
   Dim i As Integer
   
   Interesa_EsPerAMi = False
   For i = 1 To UBound(LlistaContingutsNom)
      If LlistaContingutsNom(i) = "Tot" Then
         Interesa_EsPerAMi = True
         Exit Function
      End If
   Next
   
   
   For i = 1 To UBound(LlistaContingutsNom)
      If UCase(LlistaContingutsNom(i)) = UCase(Ctg) Then
      
'If UCase(LlistaContingutsNom(i)) = "DEPENDENTES" Then
'Interesa_EsPerAMi = Interesa_EsPerAMi
'End If

         Interesa_EsPerAMi = LlistaContingutsInteresa(i)
         Exit Function
      End If
      If Left(LlistaContingutsNom(i), 1) = "*" Then
         Interesa_EsPerAMi = LlistaContingutsInteresa(i)
         If UCase(Right(LlistaContingutsNom(i), Len(LlistaContingutsNom(i)) - 1)) = UCase(Right(Ctg, Len(LlistaContingutsNom(i)) - 1)) Then
            Interesa_EsPerAMi = LlistaContingutsInteresa(i)
            Exit Function
         End If
      End If
      If Right(LlistaContingutsNom(i), 1) = "*" Then
         If UCase(Left(LlistaContingutsNom(i), Len(LlistaContingutsNom(i)) - 1)) = UCase(Left(Ctg, Len(LlistaContingutsNom(i)) - 1)) Then
            Interesa_EsPerAMi = LlistaContingutsInteresa(i)
            Exit Function
         End If
      End If
   Next
   
   ReDim Preserve LlistaContingutsNom(UBound(LlistaContingutsNom) + 1)
   LlistaContingutsNom(UBound(LlistaContingutsNom)) = Ctg
         
   ReDim Preserve LlistaContingutsInteresa(UBound(LlistaContingutsInteresa) + 1)
   LlistaContingutsInteresa(UBound(LlistaContingutsInteresa)) = False
   
   ReDim Preserve LlistaContingutsEsborrem(UBound(LlistaContingutsEsborrem) + 1)
   LlistaContingutsEsborrem(UBound(LlistaContingutsEsborrem)) = False
         
End Function

Function Interesa_CalEsborrar(Ctg As String) As Boolean
   Dim i As Integer
   
   Interesa_CalEsborrar = False
   
   For i = 0 To UBound(LlistaContingutsNom)
      If UCase(LlistaContingutsNom(i)) = UCase(Ctg) Then
         'LlistaContingutsInteresa (i)
         Interesa_CalEsborrar = LlistaContingutsEsborrem(i)
         Exit Function
      End If
   Next
   
   ReDim Preserve LlistaContingutsNom(UBound(LlistaContingutsNom) + 1)
   LlistaContingutsNom(UBound(LlistaContingutsNom)) = Ctg
         
   ReDim Preserve LlistaContingutsInteresa(UBound(LlistaContingutsInteresa) + 1)
   LlistaContingutsInteresa(UBound(LlistaContingutsInteresa)) = False
   
   ReDim Preserve LlistaContingutsEsborrem(UBound(LlistaContingutsEsborrem) + 1)
   LlistaContingutsEsborrem(UBound(LlistaContingutsEsborrem)) = False
         
   
End Function


Function ClauOrigen(IdE As String) As String
   
   ClauOrigen = "[Maquina#" & IdE & "]"
   
End Function

Function ClauContingut(IdE As String) As String
   
   ClauContingut = "[Contingut#" & IdE & "]"
   
End Function



Sub MyKill(Fil As String)
On Error Resume Next
   Name Fil As Cnf.AppPath & "\Bak\" & nomfile(Fil)
   Kill Fil

End Sub

Sub MyMkDir(Fil As String)

On Error Resume Next
   MkDir Fil

End Sub


Function nomfile(f As String) As String
   Dim P As Integer, P2 As Integer
   
   P2 = 0
   P = 0
   Do
      P = InStr(P + 1, f, "\")
      If P > 0 Then P2 = P
   Loop While P > 0
   
   nomfile = Right(f, Len(f) - P2)
   
End Function


Function TrobaData(s As String) As String
   Dim Di As String, P As Integer, p1 As Integer, P2 As Integer
   Dim dia As String, mes As String, An As String, Hora As String
   Dim s2 As String
   
On Error Resume Next

   TrobaData = Now
   'Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec
   P = InStr(s, "Jan")
   If P > 0 Then
    s = Left(s, P - 1) & " 01" & Right(s, Len(s) - P - 2)
   Else
    P = InStr(s, "Feb")
    If P > 0 Then
        s = Left(s, P - 1) & " 02" & Right(s, Len(s) - P - 2)
    Else
        P = InStr(s, "Mar")
        If P > 0 Then
            s = Left(s, P - 1) & " 03" & Right(s, Len(s) - P - 2)
        Else
            P = InStr(s, "Apr")
            If P > 0 Then
                s = Left(s, P - 1) & " 04" & Right(s, Len(s) - P - 2)
            Else
                P = InStr(s, "May")
                If P > 0 Then
                    s = Left(s, P - 1) & " 05" & Right(s, Len(s) - P - 2)
                Else
                    P = InStr(s, "Jun")
                If P > 0 Then
                    s = Left(s, P - 1) & " 06" & Right(s, Len(s) - P - 2)
                Else
                    P = InStr(s, "Jul")
                    If P > 0 Then
                        s = Left(s, P - 1) & " 07" & Right(s, Len(s) - P - 2)
                    Else
                        P = InStr(s, "Aug")
                        If P > 0 Then
                            s = Left(s, P - 1) & " 08" & Right(s, Len(s) - P - 2)
                        Else
                            P = InStr(s, "Sep")
                            If P > 0 Then
                                s = Left(s, P - 1) & " 09" & Right(s, Len(s) - P - 2)
                            Else
                                P = InStr(s, "Oct")
                                If P > 0 Then
                                    s = Left(s, P - 1) & " 10" & Right(s, Len(s) - P - 2)
                                Else
                                    P = InStr(s, "Nov")
                                    If P > 0 Then
                                        s = Left(s, P - 1) & " 11" & Right(s, Len(s) - P - 2)
                                    Else
                                        P = InStr(s, "Dec")
                                        If P > 0 Then
                                            s = Left(s, P - 1) & " 12" & Right(s, Len(s) - P - 2)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
   End If
   End If
   
   
   If P > 0 Then
      If P = 1 Then
        TrobaData = Format(s, "mm-dd-yyyy")
      Else
        TrobaData = Format(s, "dd-mm-yyyy")
      End If
   End If
   
   
   s = Trim(s)
   P = InStr(s, " ")
   If P > 0 Then
      p1 = InStr(P + 1, s, " ")
      If p1 > 0 Then
         P2 = InStr(p1 + 1, s, " ")
         If P2 > 0 Then
            mes = Mid(s, 1, P - 1)
            An = Mid(s, P + 1, p1 - P - 1)
            If Len(An) = 0 Then An = "0"
            dia = Mid(s, p1 + 1, P2 - p1 - 1)
            Hora = Mid(s, P2 + 1, Len(s) - P2)
            TrobaData = DateSerial(An, mes, dia) + TimeValue(Hora)
         Else
            mes = Mid(s, 1, P - 1)
            If Val(mes) <= (Month(Now) + 1) Then ' +1, Ens donem un mes de marge, es mes probable un retras en el rellotge que 335 dis un file al ftp.
               An = Year(Now)
            Else
               An = Year(Now) - 1
            End If
            dia = Mid(s, P + 1, p1 - P - 1)
            Hora = Mid(s, p1 + 1, Len(s) - p1)
            TrobaData = DateSerial(An, mes, dia)
            If IsDate(Hora) Then TrobaData = DateSerial(An, mes, dia) + TimeValue(Hora)
            If Not IsDate(Hora) And IsNumeric(Hora) Then TrobaData = DateSerial(Hora, mes, dia)
         End If
      Else
        TrobaData = Format(s, "mm-dd-yyyy hh:nn")
      End If
   End If
   
End Function
Sub MyAccesAlRegistre(Clau As String, Valor As String, Defecte As String, Optional EscriuDefecte As Boolean = False)
   Dim Resultat As String
   
   If Len(Cnf.AppName) = 0 Then Cnf.AppName = App.EXEName
   
   If EscriuDefecte Then SaveSetting "Hit_" & LCase(Cnf.llicencia), Cnf.AppName, Clau, Defecte
   Resultat = GetSetting("Hit_" & LCase(Cnf.llicencia), Cnf.AppName, Clau)

   If Len(Resultat) = 0 Then
      Resultat = Defecte
'      SaveSetting "Hit_" & LCase(Llicencia), cnf.AppName, CLau, Resultat
   End If
   
   Valor = Resultat
   
End Sub




Function TreuPath(AmbPath As String) As String
   Dim Fil As String, i As String
   
   Fil = AmbPath
   If Mid$(Fil, 2, 1) = ":" Then Fil = Mid$(Fil, 3)
   i = InStr(Fil, "/")
   Do While i > 0
      Fil = Mid(Fil, i + 1)
      i = InStr(Fil, "/")
   Loop
   
   TreuPath = Fil
   
End Function

Function DiscSerialNumber() As Long
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


Function SumaDeDigits(st As String) As Double
   Dim Ac As Double, i As Integer, c As String
   
   Ac = 0
   For i = 1 To Len(st)
      c = Mid(st, i, 1)
      If IsNumeric(c) Then Ac = Ac + Val(c)
   Next
   SumaDeDigits = Ac

End Function






Function MissatgeDiscNumber() As String
   MissatgeDiscNumber = Chr(13) & Chr(10) & "Disc Num : [" & DiscSerialNumber() & "]"
End Function



'Private Sub IPClock_DataIn(Datagram As String, SourceAddress As String, SourcePort As Integer)
'   Dim T As netTime, Hi&, Lo&, timeString$
'
'   Hi& = 256& * Asc(Mid(Datagram, 1, 1)) + Asc(Mid(Datagram, 2, 1))
'   Lo& = 256& * Asc(Mid(Datagram, 3, 1)) + Asc(Mid(Datagram, 4, 1))
'
'   If Clock_HoresTrobades >= Clock_HoresMaxim Then Exit Sub
'   Clock_HoresTrobades = Clock_HoresTrobades + 1
'   GetDate Hi&, Lo&, Hores(Clock_HoresTrobades)
'
'End Sub
'
Private Sub GetDate(Hi&, Lo&, tM As netTime)
   Dim i As Integer, t&, NoDays&, DayTime&, LeapYears%, bLeap%
   
   Static DaysYear%(0 To 1, 0 To 13)

   DaysYear%(0, 0) = 0
   DaysYear%(0, 1) = 31
   DaysYear%(0, 2) = 59
   DaysYear%(0, 3) = 90
   DaysYear%(0, 4) = 120
   DaysYear%(0, 5) = 151
   DaysYear%(0, 6) = 181
   DaysYear%(0, 7) = 212
   DaysYear%(0, 8) = 243
   DaysYear%(0, 9) = 273
   DaysYear%(0, 10) = 304
   DaysYear%(0, 11) = 334
   DaysYear%(0, 12) = 365
   DaysYear%(1, 0) = 0
   DaysYear%(1, 1) = 31

   For i = 2 To 12
      DaysYear%(1, i%) = DaysYear(0, i%) + 1
   Next i%
    
   ' subtract the seconds till 00:00 Jan 1 1968}
   'it's only a 31 bit integer :(
   t& = (Hi& - HiWord&(SEC_TILL_1968&) - 1) * &H10000 + (&H10000 + Lo& - LoWord&(SEC_TILL_1968&))

   'adjust here t with +/-GMT * 3600 }

   NoDays& = t& \ SEC_DAY&
   DayTime& = t& Mod SEC_DAY&
   tM.Hour = DayTime& \ 3600
   DayTime& = DayTime& Mod 3600
   tM.Min = DayTime& \ 60
   tM.Sec = DayTime& Mod 60

   tM.Year = NoDays& \ 365 + 1968
   NoDays& = NoDays& Mod 365
   LeapYears% = (tM.Year - 1969) \ 4
   If 0 = (tM.Year \ 4) Then bLeap% = 1 Else bLeap% = 0

   NoDays& = NoDays& - LeapYears%
   If NoDays& < 0 Then
      tM.Year = tM.Year - 1
      NoDays& = NoDays& + 365
   Else
      i% = 1
      Do While NoDays& > DaysYear%(bLeap%, i%)
         i% = i% + 1
      Loop
      tM.Month = i%
      tM.Day = NoDays& - DaysYear(bLeap%, i% - 1)
   End If

End Sub

Private Function HiWord&(L&)
    HiWord& = (L& And &HFFFF0000) \ 65536
End Function


Private Function LoWord&(L&)
    LoWord& = L& Mod 65536
End Function


Public Function Uff_UnPack(NomUff As String, PathDesti As String) As Integer
   Dim Ret As TErrUFF
   
   Uff_UnPack = 1
   
   Ret = Uff_ExtraccionUFF(NomUff, PathDesti)
   If Ret <> No_ErrUFF Then Exit Function
   
   Uff_UnPack = 0
   
End Function

Private Function Uff_CreacionUFF(ByRef InfoUFF As TInfoUff, nombreUFF As String, Optional FormatUFF As Byte = FORMAT_UFF_DEFAULT) As TErrUFF
   Dim canal As Integer, Ret As TErrUFF
   canal = FreeFile
On Error Resume Next
    
    Select Case FormatUFF
        Case FORMAT_UFF_0
        Case Else
            Ret = ErrUFF_InvalidData: GoTo salida
    End Select
    
    ' Crear fichero UFF
    Kill nombreUFF: err.Clear
        
    Open nombreUFF For Binary As canal
    If err Then Ret = ErrUFF_Write: GoTo salida
    
    ' Inicializar Información archivo UFF
    InfoUFF.canal = canal ' Guardar Canal
    InfoUFF.NumBlocs = 0
    InfoUFF.Formato = FormatUFF

    ' Llenar cabecera con valores por defecto y al fichero
    Dim HdrUFF As THeaderUFF
    
    HdrUFF.MagicWord = MAGIC_WORD_UFF
    HdrUFF.FormatUFF = FORMAT_UFF_0
    HdrUFF.crc32 = 0
    HdrUFF.NumBlocs = 0
 
    Put #canal, , HdrUFF
    If err Then Ret = ErrUFF_Write: GoTo salida ' error disco lleno
    
    Ret = No_ErrUFF
    
salida:
    
    If Ret <> No_ErrUFF Then Close #canal
    Uff_CreacionUFF = Ret
    err.Clear

End Function

Private Function Uff_InsertarFicheroEnUFF(ByRef InfoUFF As TInfoUff, _
      FichOrigen As String, RutaDestino As String, _
      TipoDatos As Byte) As TErrUFF
      
    Dim CanalOrigen As Integer, Ret As TErrUFF
    Dim LenOrigen As Long
    
    CanalOrigen = FreeFile
On Error Resume Next

    If InfoUFF.NumBlocs >= MAX_NumBlocs Then Ret = ErrUFF_InvalidData
    ' Abrir fichero a insertar
    Open FichOrigen For Binary As CanalOrigen
    If err Then Ret = ErrUFF_ReadFileBloc: GoTo salida
    
    LenOrigen = LOF(CanalOrigen) ' long. fichero a insertar
    
    Dim OrigenBloque As Long, hdrBloc As THeaderBloc1
    
    OrigenBloque = Seek(InfoUFF.canal) ' posición inicio bloque en UFF
    
    ' Crear cabecera bloque con valores defecto y al UFF
    hdrBloc.LenBloc = 0: hdrBloc.DataType = TipoDatos
    Put #InfoUFF.canal, , hdrBloc
    Put #InfoUFF.canal, , RutaDestino
    Dim hh As Byte: hh = 0: Put #InfoUFF.canal, , hh ' 0 final string
    If err Then Ret = ErrUFF_Write: GoTo salida
    
    Select Case TipoDatos
        Case TD_SinCompresion
            ' Tipo 1 sin compresion
            Dim i As Long, Car As Byte
            
            For i = 1 To LenOrigen
                Get #CanalOrigen, , Car
                Put #InfoUFF.canal, , Car
                If err Then Ret = ErrUFF_Write: GoTo salida ' error disco lleno
            Next i
        ' Case TD_Compresion1
        Case Else
            Ret = ErrUFF_Corrupted: GoTo salida
    End Select
    
    Close #CanalOrigen
    
    Dim FinBloque As Long
    
    FinBloque = Seek(InfoUFF.canal)
    ' long. Bloque = Fin - Inicio en UFF
    hdrBloc.LenBloc = FinBloque - OrigenBloque
    
    ' Guardar Cabecera bloque defintiva
    Seek (InfoUFF.canal), OrigenBloque
    Put #InfoUFF.canal, , hdrBloc
    
    Seek (InfoUFF.canal), FinBloque 'Ir a fin Bloque en UFF
    
    InfoUFF.NumBlocs = InfoUFF.NumBlocs + 1 'NumBloques++
    
    If err Then Ret = ErrUFF_Undefined: GoTo salida ' error anterior no tratado
    
    Ret = No_ErrUFF
    
salida:

    If Ret <> No_ErrUFF Then Close #CanalOrigen
    Uff_InsertarFicheroEnUFF = Ret
    err.Clear
    
End Function
Private Function Uff_CerrarUFF(InfoUFF As TInfoUff) As TErrUFF
    Dim Ret As TErrUFF
    
On Error Resume Next
    Dim HdrUFF As THeaderUFF
    
    ' Guardar cabecera casi definitiva, solo falta CRC para calcular
    Seek (InfoUFF.canal), 1 ' Ir a inicio UFF
    
    HdrUFF.MagicWord = MAGIC_WORD_UFF
    HdrUFF.crc32 = 0
    HdrUFF.FormatUFF = InfoUFF.Formato
    HdrUFF.NumBlocs = InfoUFF.NumBlocs
    Put #InfoUFF.canal, , HdrUFF
       
    ' Calcular CRC
    
    ' Empieza despues de MagicWord y CRC
    Seek (InfoUFF.canal), 1 + Len(HdrUFF.MagicWord) + Len(HdrUFF.crc32)
    
    Dim crc32 As Long, Car As Byte
    
    crc32 = CRC32_Inicial ' semilla inicial
    While (Not (EOF(InfoUFF.canal)))
       Get #InfoUFF.canal, , Car
       crc32 = Uff_CalcCRC32(Car, crc32)
    Wend
    
    ' Guardar cabecera definitiva
    Seek (InfoUFF.canal), 1 ' Ir a inicio UFF
    
    HdrUFF.crc32 = crc32
    Put #InfoUFF.canal, , HdrUFF
    
    Close #InfoUFF.canal
    
    If err Then Ret = ErrUFF_Write: GoTo salida
    Ret = No_ErrUFF

salida:
    
    Uff_CerrarUFF = Ret
    If Ret <> No_ErrUFF Then Close #InfoUFF.canal
    err.Clear
    
End Function







Private Function Uff_ExtraccionUFF(nombreUFF As String, _
       PathExtract As String) As TErrUFF
    Dim canal As Integer, Ret As TErrUFF, InfoUFF As TInfoUff
    
    canal = FreeFile

On Error Resume Next
    Open nombreUFF For Binary As canal  ' Abrir UFF
    If err Then Ret = ErrUFF_Read: GoTo salida

    InfoUFF.canal = canal
    
    Ret = Uff_ExtraerCabeceraUFF(InfoUFF)
    If Ret <> No_ErrUFF Then GoTo salida
    
    Ret = Uff_ExtraerBloquesUFF(InfoUFF, PathExtract)
    If Ret <> No_ErrUFF Then GoTo salida
    
    Close #canal
    If err Then Ret = ErrUFF_Read: GoTo salida
    
    Ret = No_ErrUFF

salida:

    Uff_ExtraccionUFF = Ret
    If Ret <> No_ErrUFF Then Close #canal
    err.Clear

End Function

Private Function Uff_ExtraerCabeceraUFF(InfoUFF As TInfoUff) As TErrUFF
   Dim i As Integer
    Dim Ret As TErrUFF
    
On Error Resume Next

    ' Leer Cabecera UFF
    Dim HdrUFF As THeaderUFF
    Get #InfoUFF.canal, , HdrUFF
    If EOF(InfoUFF.canal) Then Ret = ErrUFF_Read: GoTo salida
    
    ' guardar datos proceso UFF
    InfoUFF.Formato = HdrUFF.FormatUFF
    InfoUFF.NumBlocs = HdrUFF.NumBlocs
    
    ' Verificar MagicWord
    If (HdrUFF.MagicWord <> MAGIC_WORD_UFF) Then Ret = ErrUFF_Corrupted: GoTo salida
    
    ' Verificar Formato UFF
    Select Case InfoUFF.Formato
        Case FORMAT_UFF_0
        Case Else
            Ret = ErrUFF_Corrupted: GoTo salida
    End Select
    
    ' Verificación CRC
    
        ' posición inicio calculo CRC
    Seek (InfoUFF.canal), 1 + Len(HdrUFF.MagicWord) + Len(HdrUFF.crc32)
    
    Dim crc32 As Long, Car As Byte
        ' Calcular CRC
    crc32 = CRC32_Inicial
    i = 0
    While (Not (EOF(InfoUFF.canal)))
       i = i + 1
       Get #InfoUFF.canal, , Car
       crc32 = Uff_CalcCRC32(Car, crc32)
       If i > 10000 Then
          i = 0
          DoEvents
       End If
    Wend
        ' Verificar
    If (crc32 <> HdrUFF.crc32) Then Ret = ErrUFF_Corrupted: GoTo salida
    
    Ret = No_ErrUFF

salida:

    Uff_ExtraerCabeceraUFF = Ret
    err.Clear
        
End Function

Private Function Uff_ExtraerBloquesUFF(InfoUFF As TInfoUff, PathExtract As String) As TErrUFF
    Dim nBloc As Integer, Ret As TErrUFF
    Dim headBloc As THeaderBloc1, headUFF As THeaderUFF
    Dim IniBloc As Long, LenData As Long
    Dim RutaDestino As String
    Dim AliasFile() As TAliasFile
    Dim Car As Byte
    Dim CanalTemp As Integer
    Dim Fso As Object
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    CanalTemp = FreeFile
On Error Resume Next
    If InfoUFF.NumBlocs = 0 Then Ret = No_ErrUFF: GoTo salida
    
    ' Situarse en el inicio del primer bloque
    Seek #InfoUFF.canal, 1 + Len(headUFF)
    
    ReDim AliasFile(1 To InfoUFF.NumBlocs)
    
    For nBloc = 1 To InfoUFF.NumBlocs
        IniBloc = Seek(InfoUFF.canal) ' guardar pos Inicio Bloque
        Get #InfoUFF.canal, , headBloc ' leer cabecera bloque
        If EOF(InfoUFF.canal) Then Ret = ErrUFF_Corrupted: GoTo salida
        
        ' verificar que tipo bloque sea tipo fichero
        Select Case headBloc.DataType
            Case TD_SinCompresion:
            Case Else
                Ret = ErrUFF_Corrupted: GoTo salida
        End Select
        
        ' leer ruta destino
        RutaDestino = ""
        Do While (True)
            Get #InfoUFF.canal, , Car
            If EOF(InfoUFF.canal) Then Ret = ErrUFF_Corrupted: GoTo salida
            If Car = 0 Then Exit Do
            RutaDestino = RutaDestino & Chr(Car)
        Loop
        
        ' si ruta solo nombre fichero crearlo en PathExtract
        AliasFile(nBloc).Destino = Uff_CrearRutaDestinoExtraccion(RutaDestino, PathExtract)
        
        LenData = headBloc.LenBloc + _
                  (IniBloc - Seek(InfoUFF.canal)) ' len cabecera bloque
        
        ' fichero temporal para extracción
        AliasFile(nBloc).Temporal = _
              Fso.BuildPath(Fso.GetSpecialFolder(2), Fso.GetTempName)
        
        Open AliasFile(nBloc).Temporal For Binary As CanalTemp
        If err Then Ret = ErrUFF_WriteFileBloc: GoTo salida

        ' Extraer segun tipo
        Select Case headBloc.DataType
            Case TD_SinCompresion
                Dim i As Long
            
                For i = 1 To LenData
                    Get #InfoUFF.canal, , Car
                    Put #CanalTemp, , Car
                    If err Then Ret = ErrUFF_WriteFileBloc: GoTo salida
                    If (i Mod 10000) = 1 Then DoEvents
                Next i
            Case Else
                Ret = ErrUFF_Corrupted: GoTo salida
        End Select
    
        Close #CanalTemp
        If err Then Ret = ErrUFF_WriteFileBloc: GoTo salida
        
    Next nBloc
    
    ' renombrar temporal a destino
    For nBloc = 1 To InfoUFF.NumBlocs
        Kill AliasFile(nBloc).Destino: err.Clear
        MkDir Fso.GetParentFolderName(AliasFile(nBloc).Destino): err.Clear
        Name AliasFile(nBloc).Temporal As AliasFile(nBloc).Destino
        If err Then Ret = ErrUFF_WriteFileBloc: GoTo salida
        AliasFile(nBloc).Temporal = ""
        AliasFile(nBloc).fDestinoEscrito = True
    Next nBloc

    Set Fso = Nothing
    
    Ret = No_ErrUFF

salida:

    If Ret <> No_ErrUFF Then  ' si error borrar temporales y destinos escritos
        For nBloc = 1 To InfoUFF.NumBlocs
            If AliasFile(nBloc).Temporal <> "" Then
                Kill AliasFile(nBloc).Temporal
            End If
            If AliasFile(nBloc).fDestinoEscrito Then
                Kill AliasFile(nBloc).Destino
            End If
        Next nBloc
        Close #CanalTemp
    End If
    Uff_ExtraerBloquesUFF = Ret
    err.Clear

End Function

Sub CarregaFtpDirs()
   
   ReDim FtpServer(0)
   ReDim FTPUSER(0)
   ReDim FtpPsw(0)
      
'   If UCase(EmpresaActual) = UCase("integraciones") Then
'      CarregaFtpDir "62.151.23.98", "integraciones", "integra2004"
'   Else
    
    
    CarregaFtpDir feina(EmpresaActualNum).Ftp_Server, feina(EmpresaActualNum).Ftp_User, feina(EmpresaActualNum).Ftp_Pssw
    FtpServerUltimMaster = UBound(FtpServer)
    
End Sub

Private Function Uff_CrearRutaDestinoExtraccion(ruta As String, Path As String) As String
    Dim Fso As Object
    
    Set Fso = CreateObject("Scripting.FileSystemObject")
        ' si la unidad esta incluida, no cambia la ruta
        If Fso.GetDriveName(ruta) <> "" Then
            Uff_CrearRutaDestinoExtraccion = ruta: GoTo salida
        End If
        ' si empieza con barra, no añadir el path
        If InStr("\/", Mid(ruta, 1, 1)) <> 0 Then
            Uff_CrearRutaDestinoExtraccion = ruta: GoTo salida
        End If
        ' sino añadir el path delante
        Uff_CrearRutaDestinoExtraccion = Fso.BuildPath(Path, ruta)
salida:
    Set Fso = Nothing
End Function


Private Function Uff_Bit(n As Byte) As Long
On Error Resume Next
    Uff_Bit = BitTable(n)
    If err = 0 Then Exit Function
    If n >= 32 Then Uff_Bit = 0: Exit Function
    ReDim BitTable(0 To 31)
    BitTable(0) = 1
    Dim i As Byte
    For i = 1 To 30
       BitTable(i) = BitTable(i - 1) * 2
    Next i
    BitTable(31) = &H80000000
    Uff_Bit = BitTable(n)
End Function




Private Function Uff_Mask(n As Byte) As Long
On Error Resume Next
    Uff_Mask = MaskTable(n)
    If err = 0 Then Exit Function
    If n >= 32 Then Uff_Mask = &HFFFFFFFF: Exit Function
    ReDim MaskTable(0 To 31)
    MaskTable(0) = Bit(0)
    Dim i As Byte
    For i = 1 To 31
       MaskTable(i) = MaskTable(i - 1) Or Bit(i)
    Next i
    Uff_Mask = MaskTable(n)
End Function





Private Function Uff_ShiftL(ByVal L As Long, Optional n As Byte = 1) As Long
    Dim i As Byte
    For i = 1 To n
       L = L And &H7FFFFFFF
       If ((L And &H40000000) <> 0) Then
          L = L And &H3FFFFFFF: L = L * 2: L = L Or &H80000000
       Else
          L = L * 2
       End If
    Next i
    Uff_ShiftL = L
End Function


Private Function Uff_ShiftR(ByVal L As Long, Optional n As Byte = 1) As Long
    Dim i As Byte
    For i = 1 To n
       If ((L And &H80000000) <> 0) Then
          L = L And &H7FFFFFFF: L = L \ 2: L = L Or &H40000000
       Else
          L = L \ 2
       End If
    Next i
    Uff_ShiftR = L
End Function



Public Function Uff_Pack(NomUff As String, Files() As String) As Integer
   Dim InfoUFF As TInfoUff, Ret As TErrUFF, i As Integer, Destino As String
   
   Uff_Pack = 1
   Ret = Uff_CreacionUFF(InfoUFF, NomUff)
   If Ret <> No_ErrUFF Then Exit Function
    
   For i = LBound(Files) To UBound(Files)
      Destino = Dir(Files(i))
      If Len(Destino) > 0 Then Ret = Uff_InsertarFicheroEnUFF(InfoUFF, Files(i), Destino, TD_SinCompresion)
      If Ret <> No_ErrUFF Then Exit Function
   Next
   
   Ret = Uff_CerrarUFF(InfoUFF)
   If Ret <> No_ErrUFF Then Exit Function
   
   Uff_Pack = 0
   
End Function

Private Sub Uff_CrearCRC32Table()

#If (COMPUTE_XOR_PATTERN) Then
'  /* This piece of code has been left here to explain how the XOR pattern
'   * used in the creation of the crc_table values can be recomputed.
'   * For production versions of this function, it is more efficient to
'   * supply the resultant pattern at compile time.
'   */
   Dim XorPolyn As Long      ' /* polynomial exclusive-or pattern */
'  /* terms of polynomial defining this crc (except x^32): */
   Dim P As Variant, i As Variant
   P = Array(0, 1, 2, 4, 5, 7, 8, 10, 11, 12, 16, 22, 23, 26)
'  /* make exclusive-or pattern from polynomial (0xedb88320L) */
   XorPolyn = 0
   For Each i In P
      XorPolyn = XorPolyn Or Bit(31 - i)
   Next
#Else
   Const XorPolyn As Long = &HEDB88320
#End If
    
    ReDim CRC32Table(0 To 255)
    
    Dim n As Integer, c As Long, K As Byte
    For n = 0 To 255
       c = n
       For K = 1 To 8
          If ((c And 1) = 1) Then
             c = XorPolyn Xor ShiftR(c)
          Else
             c = ShiftR(c)
          End If
       Next K
       CRC32Table(n) = c
    Next n
End Sub
Public Function ShiftR(ByVal L As Long, Optional n As Byte = 1) As Long
    Dim i As Byte
    For i = 1 To n
       If ((L And &H80000000) <> 0) Then
          L = L And &H7FFFFFFF: L = L \ 2: L = L Or &H40000000
       Else
          L = L \ 2
       End If
    Next i
    ShiftR = L
End Function


Public Function Bit(n As Byte) As Long
On Error Resume Next
    Bit = BitTable(n)
    If err = 0 Then Exit Function
    If n >= 32 Then Bit = 0: Exit Function
    ReDim BitTable(0 To 31)
    BitTable(0) = 1
    Dim i As Byte
    For i = 1 To 30
       BitTable(i) = BitTable(i - 1) * 2
    Next i
    BitTable(31) = &H80000000
    Bit = BitTable(n)
End Function





Sub CarregaServers(rnrServers As Integer, rServerAddresses() As String, rServerNames() As String)
   rnrServers = 0
   rnrServers = rnrServers + 1
   rServerAddresses$(rnrServers) = "129.127.28.4"
   rServerNames$(rnrServers) = "augean.eleceng.adelaide.edu.au"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "129.132.98.11"
rServerNames$(rnrServers) = "bernina.ethz.ch"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "193.2.69.11"
rServerNames$(rnrServers) = "biofiz.mf.uni-lj.si"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.173.14.71"
rServerNames$(rnrServers) = "black-ice.cc.vt.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "193.136.54.1"
rServerNames$(rnrServers) = "bug.fe.up.pt"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.100.102.201"
rServerNames$(rnrServers) = "chime.utoronto.ca"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "192.87.106.101"
rServerNames$(rnrServers) = "chime1.surfnet.nl"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.4.1.5"
rServerNames$(rnrServers) = "churchy.udel.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "16.1.0.4"
rServerNames$(rnrServers) = "clepsydra.dec.com"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.118.25.3"
rServerNames$(rnrServers) = "clock.psu.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "192.31.216.5"
rServerNames$(rnrServers) = "clock.tricity.wsu.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "152.2.21.1"
rServerNames$(rnrServers) = "clock1.unc.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.2.250.95"
rServerNames$(rnrServers) = "clock-1.cs.cmu.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.2.222.8"
rServerNames$(rnrServers) = "clock-2.cs.cmu.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "129.15.22.8"
rServerNames$(rnrServers) = "constellation.ecn.uoknor.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "131.216.1.101"
rServerNames$(rnrServers) = "cuckoo.nevada.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "134.111.10.64"
rServerNames$(rnrServers) = "cyclonic.sw.stratus.com"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "131.179.128.36"
rServerNames$(rnrServers) = "delphi.cs.ucla.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "165.91.72.27"
rServerNames$(rnrServers) = "eagle.tamu.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.206.206.12"
rServerNames$(rnrServers) = "everest.cclabs.missouri.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "129.240.102.2"
rServerNames$(rnrServers) = "fartein.ifi.uio.no"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.182.58.100"
rServerNames$(rnrServers) = "fuzz.psc.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "192.80.214.42"
rServerNames$(rnrServers) = "fuzz.sura.net"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.249.2.2"
rServerNames$(rnrServers) = "gazette.bcm.tmc.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.129.93"
rServerNames$(rnrServers) = "gilbreth.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.141.93"
rServerNames$(rnrServers) = "gilbreth.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.147.93"
rServerNames$(rnrServers) = "gilbreth.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.148.93"
rServerNames$(rnrServers) = "gilbreth.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.171.93"
rServerNames$(rnrServers) = "gilbreth.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.128.76"
rServerNames$(rnrServers) = "harbor.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.129.76"
rServerNames$(rnrServers) = "harbor.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.154.76"
rServerNames$(rnrServers) = "harbor.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "149.156.4.11"
rServerNames$(rnrServers) = "info.cyf-kr.edu.pl"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "193.2.208.12"
rServerNames$(rnrServers) = "hmljhp.rzs-hm.si"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.129.95"
rServerNames$(rnrServers) = "molecule.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.132.95"
rServerNames$(rnrServers) = "molecule.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.136.95"
rServerNames$(rnrServers) = "molecule.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.145.95"
rServerNames$(rnrServers) = "molecule.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.167.95"
rServerNames$(rnrServers) = "molecule.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.169.95"
rServerNames$(rnrServers) = "molecule.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.46.181.95"
rServerNames$(rnrServers) = "molecule.ecn.purdue.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.165.196.2"
rServerNames$(rnrServers) = "heechee.esa.lanl.gov"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "129.237.32.1"
rServerNames$(rnrServers) = "finch.cc.ukans.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "129.237.32.2"
rServerNames$(rnrServers) = "kuhub.cc.ukans.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "134.111.10.1"
rServerNames$(rnrServers) = "lectroid.sw.stratus.com"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.42.1.64"
rServerNames$(rnrServers) = "libra.rice.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.175.1.3"
rServerNames$(rnrServers) = "louie.udel.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "165.227.1.1"
rServerNames$(rnrServers) = "ns.scruz.net"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.101.101.101"
rServerNames$(rnrServers) = "ns.nts.umn.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "134.84.84.84"
rServerNames$(rnrServers) = "nss.nts.umn.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "129.127.40.3"
rServerNames$(rnrServers) = "ntp.adelaide.edu.au"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "129.119.80.126"
rServerNames$(rnrServers) = "ntp.cox.smu.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "140.162.1.5"
rServerNames$(rnrServers) = "ntp.css.gov"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.59.64.60"
rServerNames$(rnrServers) = "ntp.ctr.columbia.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "194.207.34.9"
rServerNames$(rnrServers) = "ntp.exnet.com"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.235.20.3"
rServerNames$(rnrServers) = "ntp.lth.se"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "129.189.134.11"
rServerNames$(rnrServers) = "ntp.olivetti.com"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "129.189.134.6"
rServerNames$(rnrServers) = "ntp.olivetti.com"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "203.21.37.18"
rServerNames$(rnrServers) = "ntp.saard.net"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "132.239.51.18"
rServerNames$(rnrServers) = "ntp.ucsd.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "134.214.100.6"
rServerNames$(rnrServers) = "ntp.univ-lyon1.fr"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "192.35.82.50"
rServerNames$(rnrServers) = "ntp0.cornell.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "158.43.128.33"
rServerNames$(rnrServers) = "ntp0.pipex.net"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "206.103.126.59"
rServerNames$(rnrServers) = "ntp1.kansas.net"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "158.43.128.66"
rServerNames$(rnrServers) = "ntp1.pipex.net"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "158.43.192.66"
rServerNames$(rnrServers) = "ntp2.pipex.net"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.159.132.1"
rServerNames$(rnrServers) = "ntp0.strath.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.159.144.62"
rServerNames$(rnrServers) = "ntp0.strath.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.159.132.66"
rServerNames$(rnrServers) = "ntp1.strath.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.159.152.62"
rServerNames$(rnrServers) = "ntp1.strath.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.159.132.6"
rServerNames$(rnrServers) = "ntp2.strath.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.159.132.65"
rServerNames$(rnrServers) = "ntp2.strath.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.159.196.126"
rServerNames$(rnrServers) = "ntp3.strath.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.159.132.2"
rServerNames$(rnrServers) = "ntp3.strath.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.159.132.16"
rServerNames$(rnrServers) = "ntp4.strath.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.159.132.124"
rServerNames$(rnrServers) = "ntp4.strath.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.167.1.222"
rServerNames$(rnrServers) = "ntp1.sura.net"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.167.254.198"
rServerNames$(rnrServers) = "ntp2.sura.net"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "165.91.52.110"
rServerNames$(rnrServers) = "ntp5.tamu.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.2.236.71"
rServerNames$(rnrServers) = "ntp-1.ece.cmu.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.2.25.7"
rServerNames$(rnrServers) = "ntp-2.ece.cmu.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "140.221.10.70"
rServerNames$(rnrServers) = "ntp-1.mcs.anl.gov"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "140.221.9.6"
rServerNames$(rnrServers) = "ntp-2.mcs.anl.gov"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "140.221.10.64"
rServerNames$(rnrServers) = "ntp-2.mcs.anl.gov"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "192.240.4.1"
rServerNames$(rnrServers) = "ntp1.ossi.com"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "192.240.4.50"
rServerNames$(rnrServers) = "ntp2.ossi.com"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.126.24.53"
rServerNames$(rnrServers) = "ntp-0.cso.uiuc.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.126.24.24"
rServerNames$(rnrServers) = "ntp-1.cso.uiuc.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.126.24.44"
rServerNames$(rnrServers) = "ntp-2.cso.uiuc.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "200.34.146.67"
rServerNames$(rnrServers) = "ntp2a.audiotel.com.mx"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "200.34.146.68"
rServerNames$(rnrServers) = "ntp2b.audiotel.com.mx"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "200.34.146.69"
rServerNames$(rnrServers) = "ntp2c.audiotel.com.mx"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.88.200.22"
rServerNames$(rnrServers) = "ntp2a.mcc.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.88.200.6"
rServerNames$(rnrServers) = "ntp2b.mcc.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.88.200.4"
rServerNames$(rnrServers) = "ntp2c.mcc.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.88.203.12"
rServerNames$(rnrServers) = "ntp2d.mcc.ac.uk"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "131.144.4.22"
rServerNames$(rnrServers) = "Rolex.PeachNet.EDU"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "134.226.81.11"
rServerNames$(rnrServers) = "salmon.maths.tcd.ie"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "192.216.191.10"
rServerNames$(rnrServers) = "smart1.svi.org"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.59.35.142"
rServerNames$(rnrServers) = "sundial.columbia.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "200.5.72.1"
rServerNames$(rnrServers) = "tick.anice.net.ar"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "131.216.16.9"
rServerNames$(rnrServers) = "tick.cs.unlv.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "207.48.109.5"
rServerNames$(rnrServers) = "tick.koalas.com"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.100.96.9"
rServerNames$(rnrServers) = "tick.utoronto.ca"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "142.3.100.15"
rServerNames$(rnrServers) = "timelord.uregina.ca"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "150.124.136.4"
rServerNames$(rnrServers) = "ticktock.wang.com"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "207.82.53.2"
rServerNames$(rnrServers) = "time.software.net"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "158.121.104.4"
rServerNames$(rnrServers) = "timeserver.cs.umb.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.59.16.20"
rServerNames$(rnrServers) = "timex.cs.columbia.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "131.144.4.21"
rServerNames$(rnrServers) = "Timex.PeachNet.EDU"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.249.1.1"
rServerNames$(rnrServers) = "tmc.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "200.5.73.1"
rServerNames$(rnrServers) = "tock.anice.net.ar"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "131.216.18.4"
rServerNames$(rnrServers) = "tock.cs.unlv.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.100.100.128"
rServerNames$(rnrServers) = "tock.utoronto.ca"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.217.66.172"
rServerNames$(rnrServers) = "truechimer1.waikato.ac.nz"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.217.66.91"
rServerNames$(rnrServers) = "truechimer2.waikato.ac.nz"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "130.217.96.20"
rServerNames$(rnrServers) = "truechimer3.waikato.ac.nz"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.173.4.6"
rServerNames$(rnrServers) = "vtserf.cc.vt.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.252.135.4"
rServerNames$(rnrServers) = "wuarchive.wustl.edu"

rnrServers = rnrServers + 1
rServerAddresses$(rnrServers) = "128.165.196.1"
rServerNames$(rnrServers) = "xfiles.esa.lanl.gov"

End Sub




Public Function Mask(n As Byte) As Long
On Error Resume Next
    Mask = MaskTable(n)
    If err = 0 Then Exit Function
    If n >= 32 Then Mask = &HFFFFFFFF: Exit Function
    ReDim MaskTable(0 To 31)
    MaskTable(0) = Bit(0)
    Dim i As Byte
    For i = 1 To 31
       MaskTable(i) = MaskTable(i - 1) Or Bit(i)
    Next i
    Mask = MaskTable(n)
End Function

Public Function meses(ByVal mes As Integer) As String
  meses = Split(",GENER,FEBRER,MARÇ,ABRIL,MAIG,JUNY,JULIOL,AGOST,SETEMBRE,OCTUBRE,NOVEMBRE,DESEMBRE", ",")(mes)
End Function

Public Function paramToDate(ByVal D As String) As Date
    Dim aa, ddd, mmm, aaa
    aa = Replace(Mid(D, 2, 8), "-", "/")
    If IsDate(aa) Then
        ddd = Mid(aa, 1, 2)
        If ddd = "00" Then ddd = "01"
        mmm = Mid(aa, 4, 2)
        
        If Len(D) = 12 Then
            aaa = Mid(D, 8, 4)
        Else
           aaa = Mid(aa, 7, 2)
           'If aaa > Mid(Year(Date), 3, 2) Then
           '    aaa = "19" & aaa
           'Else
               aaa = "20" & aaa
           'End If
        End If
        aa = ddd & "/" & mmm & "/" & aaa
        paramToDate = CDate(aa)
    End If
End Function

Public Function dateToParam(ByVal D As Date, ByVal c As String) As String
  If c = "-" Then
    dateToParam = "[" & Right(0 & Day(D), 2) & "-" & Right(0 & Month(D), 2) & "-" & Right(0 & Year(D), 2) & "]"
  Else
    dateToParam = Right(0 & Day(D), 2) & c & Right(0 & Month(D), 2) & c & Year(D)
  End If
End Function

Public Function rec(ByVal sql As String, Optional ByVal upd As Boolean = False) As ADODB.Recordset
  On Error Resume Next
  Dim Rs As New ADODB.Recordset
  If upd Then
    Rs.Open sql, db2, adOpenStatic, adLockOptimistic
  Else
    Rs.Open sql, db2
  End If
  Set rec = Rs
End Function

Public Function tablaArchivo() As String
  On Error Resume Next
  Dim sql As String
  
    If Not ExisteixTaula("Archivo") Then
        sql = "CREATE TABLE [Archivo] ([id] [nvarchar] (255) NULL ,[nombre] [nvarchar] (20) NULL ,"
        sql = sql & "[extension] [nvarchar] (10) NULL ,[descripcion] [nvarchar] (255) NULL ,[mime] [nvarchar] (255) NULL ,"
        sql = sql & "[archivo] [image] NULL,[fecha] [datetime] NULL,[propietario] [nvarchar] (255) NULL,[tmp] bit,[down] bit)"
        ExecutaComandaSql sql
    End If
    
  tablaArchivo = "archivo"
End Function


