Attribute VB_Name = "Email"
Option Explicit

'Private objOLApp As outlook.Application


Type Message
    From As String
    Subject As String
    Date As String
    AllHeaders As String
    Size As Long
    Text As String
End Type

Global gCurrMsg As Integer
Global gMessages(1 To 2000) As Message

Sub ComandaEmail(botiga As Double)
'    Dim loOutlook As Object, objNS As Object, objFolder As Object, objOL As Object, loNameSpace As Object, loMess As Object, loInBox As Object, loMailItem As Object, loSafeItem As Object, Missatge As String, Sql As String, Di As Date, Q As rdoQuery, Df As Date, myattachments As Object, Rs As rdoResultset
''227
'
'    Sql = ""
'    Missatge = ""
'
'MsgBox "Si?"
'
'    Df = Now
'    Di = DateAdd("d", -1, Now)
'
'    Set Rs = Db.OpenResultset("Select * from records where concepte = 'EnviadaComandaDfBotiga" & Botiga & "'  ")
'    If Not Rs.EOF Then If Not IsNull(Rs("TimeStamp")) Then Di = Rs("TimeStamp")
'    Rs.Close
'
'    Set Rs = Db.OpenResultset("Select Max(Data) From [" & NomTaulaVentas(Df) & "] Where Botiga = " & Botiga & "  ")
'    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Da = Rs(0)
'    Rs.Close
'
'    ExecutaComandaSql "Delete Records Where concepte = 'EnviadaComandaDiBotiga" & Botiga & "'  "
'    ExecutaComandaSql "Delete Records Where concepte = 'EnviadaComandaDfBotiga" & Botiga & "'  "
'
'    Df = Now
'    Di = DateAdd("d", -1, Now)
'    Set Q = Db.CreateQuery("", "Insert Into Records (concepte,timestamp) Values ('EnviadaComandaDiBotiga" & Botiga & "' ,?)")
'    Q.rdoParameters(0) = Di
'    Q.Execute
'    Set Q = Db.CreateQuery("", "Insert Into Records (concepte,timestamp) Values ('EnviadaComandaDfBotiga" & Botiga & "' ,?)")
'    Q.rdoParameters(0) = Df
'    Q.Execute
'
'    Sql = ""
'    Sql = Sql & " select a.nom as nom,q.q as q from articles a join "
'    Sql = Sql & " (select sum(quantitat) as q ,Botiga as client ,plu as article from [" & NomTaulaVentas(Df) & "] where botiga = " & Botiga & " "
'    Sql = Sql & " and data >  (Select max(timestamp) From Records where concepte = 'EnviadaComandaDiBotiga" & Botiga & "' ) "
'    Sql = Sql & " and data <= (Select max(timestamp) From Records where concepte = 'EnviadaComandaDfBotiga" & Botiga & "' ) "
'    Sql = Sql & " and plu in (select articulo from ccproveedorpedido where cliente = " & Botiga & ") "
'    Sql = Sql & " group by Botiga,plu) q "
'    Sql = Sql & " on a.codi = q.article "
'
'    Set Rs = Db.OpenResultset(Sql)
'
'    While Not Rs.EOF
'        If Missatge = "" Then
'           Missatge = Missatge & "PEDIDO TIENDA : " & BotigaCodiNom(Botiga) & vbCrLf
'           Missatge = Missatge & "_________________________________" & vbCrLf
'           Missatge = Missatge & "Ventas De " & Format(Di, "dddd dd-mm-yy hh:mm ") & " a " & Format(Df, "dddd dd-mm-yy hh:mm ") & vbCrLf
'           Missatge = Missatge & " " & vbCrLf
'        End If
'
'       Missatge = Missatge & Right(Space(10) & Rs("Q"), 10) & Chr(9) & "  " & Rs("Nom") & vbCrLf
'       Rs.MoveNext
'    Wend
'    Rs.Close
'
'    If Not Missatge = "" Then
'        Missatge = Missatge & "_________________________________" & vbCrLf
'    End If
'
'    If Len(Missatge) > 0 Then
'        Set loOutlook = CreateObject("Outlook.Application")
'        Set loNameSpace = loOutlook.GetNamespace("MAPI")
'        loNameSpace.Logon
'        Set loMailItem = loOutlook.CreateItem(0)    ' This creates the MailItem Object
''        loMailItem.Recipients.Add ("34937263505@efaxsend.com")
'loMailItem.Recipients.Add ("34934326751@efaxsend.com")   'Bellapan
''loMailItem.Recipients.Add ("34973480756@efaxsend.com")   'Ipsic
''        loMailItem.Recipients.Add ("jordi@hitsystems.net")
'        loMailItem.Subject = "Comanda"
'        loMailItem.Body = Missatge
''        Set myattachments = loMailItem.Attachments
''        myattachments.Add "\\sjm-svr\mco\access\Grndata.pgp", olByValue, 2880, "HRH Eligibility Data to Magellan"
'
'        loMailItem.Send   '&& this does not cause the Security Dialog to be displayed.
'        Set loMailItem = Nothing
'        Set loNameSpace = Nothing
'        Set loOutlook = Nothing
'    End If
    
End Sub
Sub EmailFaltaProducte(botiga As String, article As String)
    Dim Rs As rdoResultset, eMailDesti As String, BotigaNom As String, articleNom As String
    
    eMailDesti = ""
    Set Rs = Db.OpenResultset("select * from dependentesextes where id=(select valor from ConstantsEmpresa where camp = 'personaAvisEncarrec') and nom='EMAIL'")
    If Not Rs.EOF Then eMailDesti = Rs("valor")
    Rs.Close

    If eMailDesti = "" Then eMailDesti = "ana@solucionesit365.com"
    
    
    BotigaNom = ""
    Set Rs = Db.OpenResultset("select * from clients where codi = '" & botiga & "'")
    If Not Rs.EOF Then BotigaNom = Rs("nom")
    Rs.Close
    
    articleNom = ""
    Set Rs = Db.OpenResultset("select * from articles where codi = '" & article & "'")
    If Not Rs.EOF Then articleNom = Rs("nom")
    Rs.Close
    
    sf_enviarMail "email@hit.cat", eMailDesti, "FALTA " & articleNom & " A LA BOTIGA " & BotigaNom, "FALTA " & articleNom & " A LA BOTIGA " & BotigaNom, "", ""
    
End Sub
Sub EmailResumBotiga(Desti As String, sDia As String)
    Dim Estat As String, dia As Date, cos As String, Cos2 As String
    Dim Rs As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset, Rs4 As rdoResultset
    Dim Rs5 As rdoResultset, Rs6 As rdoResultset, Rs7 As rdoResultset, vBotiga, vAlbarans, vPing
    Dim Co As String, Es As String, sql, dd, mm, yyyy, dd2, mm2, yyyy2, nBot As Integer, nBal As Integer
    Dim vMysql, DataPing As Date, EurosPing As Double, EsMaquina As String, DataPingActual As Date, DifPing
    
On Error GoTo nor:
    dia = Now
    If IsNumeric(sDia) Then dia = DateAdd("d", -sDia, Now)
    If IsDate(sDia) Then dia = CVDate(sDia)
    Estat = ":-|"
    dd = Day(dia)
    dd2 = Day(DateAdd("d", 1, dia))
    mm = Month(dia)
    mm2 = Month(DateAdd("d", 1, dia))
    If Len(mm) < 2 Then mm = "0" & mm
    If Len(mm2) < 2 Then mm2 = "0" & mm2
    yyyy = Year(dia)
    yyyy2 = Year(DateAdd("d", 1, dia))
    cos = ""
    Set Rs = Db.OpenResultset("select c.Codi,c.nom,w.Codi Wcodi  from ParamsHw w join clients c on w.Valor1 = c.Codi Order by c.nom ")
    While Not Rs.EOF
        If Rs("Codi") = 614 Then Rs.MoveNext 'No volen que surti la balança 4 de Sants (61)
        If Rs("Codi") = 518 Then Rs.MoveNext 'No volen que surti la botiga Forneria Pi i Maragall
        If Rs("Codi") = 602 Then Rs.MoveNext 'No volen que surti la botiga 060 Blanes_HIT Balança 2
        If Rs("Codi") = 712 Then Rs.MoveNext 'No volen que surti la botiga 071 Piera_HIT Balança 2
        If Rs("Codi") = 642 Then Rs.MoveNext 'No volen que surti la botiga 064 Sant Joan Despi_HIT Balança 2

        If Not Rs.EOF Then
            nBot = Left(Rs("codi"), Len(Rs("codi")) - 1)
            nBal = Right(Rs("codi"), 1)
            If Rs("Codi") = 518 Then
                nBal = 1
                nBot = 106
            End If
            Co = "<Td>" & Rs("Nom") & "</Td><Td>"
            'Venut
            vBotiga = 0
            Set Rs2 = Db.OpenResultset("select isnull(MAX(Data),GETDATE()) Df ,isnull( MIN(Data),GETDATE()) di , isnull(count(distinct num_tick),0) tkN ,isnull(min(num_tick),0) Tki, isnull(max(num_tick),0) Tkf , isnull(Sum(import),0) V  from [" & NomTaulaVentas(dia) & "] where day(data) = " & Day(dia) & " and Botiga = " & Rs("Codi"))
            If Not Rs2.EOF Then vBotiga = Rs2("v")
            'Albarans
            vAlbarans = 0
            Set Rs7 = Db.OpenResultset("select isnull(Sum(import),0) A  from [" & NomTaulaAlbarans(dia) & "] where day(data) = " & Day(dia) & " and Botiga = " & Rs("Codi"))
            If Not Rs7.EOF Then vAlbarans = Rs7("A")
            'Quadre
            Set Rs3 = Db.OpenResultset("select isnull(SUM(Import),0) Vr from [" & NomTaulaMovi(dia) & "] where day(data) = " & Day(dia) & " and Tipus_moviment = 'Z'  And Botiga = " & Rs("Codi"))
            'Ultim ping data demanada
            vPing = 0
            Set Rs5 = Db.OpenResultset("select top 1 * from PingMaquina where llicencia = " & Rs("Wcodi") & " and day(tmst) = " & Day(dia) & " and MONTH(tmst) = " & Month(dia) & " order by TmSt desc ")
            'Ultim ping data actual
            Set Rs6 = Db.OpenResultset("select top 1 * from PingMaquina where llicencia = " & Rs("Wcodi") & "  order by TmSt desc ")
            'Dades mysql
            sql = "select * from openquery (AMETLLER,'select sum(importeTotal) vMysql from dat_ticket_cabecera where idEmpresa=1 "
            sql = sql & "and idTienda=" & nBot & " and idBalanzaMaestra=" & nBal & " and idBalanzaEsclava=-1 and Usuario=''Comunicaciones'' and Operacion=''A'' "
            sql = sql & "and nombreBalanzaMaestra like ''-Balan%'' and timestamp>''" & yyyy & "-" & mm & "-" & dd & " 00:00:00.0000000'' "
            sql = sql & "and timestamp<''" & yyyy2 & "-" & mm2 & "-" & dd2 & " 00:00:00.0000000'' and idVendedor<>17  ') "
            Set Rs4 = Db.OpenResultset("Select * from Articles where codi = -13")
            If UCase(EmpresaActual) = UCase("LaForneria") Then Set Rs4 = Db.OpenResultset(sql)
            Es = "<font color='red'>:-(</font>"
            EurosPing = 0
            DifPing = ""
            DataPing = 0
            DataPingActual = 0
            If Not Rs5.EOF Then
                If Not IsNull(Rs5("tmst")) And Not IsNull(Rs5("Param2")) Then
                   EurosPing = Rs5("Param2")
                   DataPing = Rs5("tmst")
                    DifPing = Abs(DateDiff("n", Rs5("Param1"), Rs5("tmst")))
                End If
            End If
            If Not Rs6.EOF Then
                If Not IsNull(Rs6("tmst")) Then
                    DataPingActual = Rs6("tmst")
                End If
            End If
        
        'Co = Co & Format(Rs2("Di"), "hh:nn") & "</Td><Td>"
        'Co = Co & Format(Rs2("Df"), "hh:nn") & "</Td><Td>"
        'Co = Co & Rs2("Tkf") - Rs2("Tki") & "=" & Rs2("tkn") & "</Td><Td>"
        'Co = Co & Round(Rs2("v"), 2) & "=" & Round(Rs3("vr"), 2) & "="
        
        'Igualtats ventes
        'Co = Co & Round(EurosPing, 2) & "=" & Round(Rs2("v"), 2) & "="
        'If Rs("Codi") = "9011" Then
            vBotiga = Round(CDbl(vBotiga) + CDbl(vAlbarans), 2)
            If Round(CDbl(EurosPing) > 0) Then vPing = Round(CDbl(EurosPing) + CDbl(vAlbarans), 2)
        'Else
        '    vBotiga = Round(CDbl(vBotiga), 2)
        '    vPing = Round(CDbl(EurosPing), 2)
        'End If
            Co = Co & vPing & "=" & vBotiga & "="
            If Not Rs4.EOF Then
                If IsNull(Rs4("vMysql")) Or Rs4("vMysql") = "" Then
                    vMysql = 0
                    Co = Co & "0"
                Else
                    vMysql = Round(Rs4("vMysql"), 2)
                    Co = Co & vMysql
                End If
            End If
            Co = Co & "</Td><Td>"
        
            'Ultim Ping del dia demanat
            If IsDate(DataPing) Then
               Co = Co & Format(DataPing, "hh:nn")
               If Hour(DataPing) > 22 Then
                    Es = "<font color='green'>:-)</font>"
               Else
                  Es = "<font color='red'>:-|</font>"
               End If
               If DifPing > 5 Then Co = Co & " Rellotge KO : " & DifPing & " m"
            End If
            Co = Co & "</Td><Td>"
            'Ultim Ping actual
            If IsDate(DataPingActual) Then
               Co = Co & DataPingActual
               nBot = nBot
               If Abs(DateDiff("n", Now, DataPingActual)) > 10 Then
                    EsMaquina = "<font color='red'>NO " & Abs(DateDiff("n", Now, DataPingActual)) & " m</font>"
                    Es = "<font color='red'>:-(</font>"
                Else
                    EsMaquina = "<font color='green'>SI</font>"
                End If
            End If
            Co = Co & "</Td><Td>" & EsMaquina & "</Td>"
        
            'If Not Rs2("Tkf") - Rs2("Tki") = Rs2("tkn") Then Es = ":-("
            'If Abs(Rs2("v") - Rs3("vr")) > 0.01 Then Es = "<font color='red'>:-(</font>"
            If Abs(vBotiga - vMysql) > 0.01 Then Es = "<font color='red'>:-(</font>"
            'If Abs(Rs3("vr") - vMysql) > 0.01 Then Es = "<font color='red'>:-(</font>"
            If vPing > 0 Then If Abs(vBotiga - vPing) > 0.01 Then Es = "<font color='red'>:-(</font>"
            If Es = ":-(" Then Estat = Es
        
            cos = cos & "<Tr><Td>" & Es & "</Td>" & Co & "</Td></Tr><Br>"
        
            Rs.MoveNext
        End If
    Wend
    Rs.Close
    'Cos = "<Table><Tr><Td><b>Ok</b></Td><Td><b>Maquina</b></Td><Td><b>Di</Td><Td><b>Df</b></Td><Td><b>Num Ticks</b></Td><Td><b>Vendes</b></Td></Tr> " & Cos & "</Table>"
    Cos2 = cos
    cos = "<Table><Tr><Td style='width:10%;'><b>Ok</b></Td><Td style='width:30%;'><b>Maquina</b></Td>"
    'Cos = Cos & "<Td style='width:10%;'><b>Di</b></Td><Td style='width:10%;'><b>Df</b></Td>"
    'Cos = Cos & "<Td style='width:20%;'><b>Num Ticks</b></Td>"
    cos = cos & "<Td style='width:30%;'><b>Botiga=Hit=Client</b></Td>"
    cos = cos & "<Td style='width:10%;'><b>Ping</b></Td><Td style='width:20%;'><b>Ping Actual</b></Td>"
    cos = cos & "<Td style='width:30%;'><b>Maq. encesa</b></Td></Tr> " & Cos2 & "</Table>"
    
    sql = "SELECT object_id FROM sys.objects with (nolock) WHERE name='TicketsTemporalsAccions' AND type='U'"
    Set Rs = Db.OpenResultset(sql)
    If Not Rs.EOF Then
        sql = "select a.Tmst,cl.Nom cli,d.NOM Dep ,isnull(ap.valor,'') v,ar.nom Art ,a.Preu from TicketsTemporalsAccions a "
        sql = sql & "join Dependentes d on d.CODI = a.dependenta join paramshw w on w.Codi = a.Botiga join clients cl on cl.Codi = w.Valor1 "
        sql = sql & "join articles ar on ar.Codi = a.cd join ArticlesPropietats ap on ar.Codi =ap.CodiArticle "
        sql = sql & "and ap.Variable = 'CODI_PROD' where Comentari = 'ModificaPreu' and day(a.Tmst)='" & dd & "' "
        sql = sql & "and month(a.Tmst)='" & Month(dia) & "' and year(a.Tmst)='" & yyyy & "' "
        Set Rs = Db.OpenResultset(sql)
    End If
    
    cos = cos & "<H><b> Preus Modificats. </b></H>"
    cos = cos & "<Table><Tr><Td style='width:20%;'><b>Data</b></Td><Td style='width:30%;'><b>Maquina</b></Td>"
    cos = cos & "<Td style='width:20%;'><b>Dependenta</b></Td><Td style='width:20%;'><b>Article</b></Td>"
    cos = cos & "<Td style='width:10%;'><b>Preu</b></Td></Tr>"
    While Not Rs.EOF
        cos = cos & "<Tr><Td>" & Rs("TmSt") & "</Td>" & "<Td>" & Rs("Cli") & "</Td>" & "<Td>" & Rs("Dep") & "</Td>" & "<Td>" & Rs("Art") & "</Td>" & "<Td>" & Rs("Preu") & "</Td>" & "</Tr>"
        Rs.MoveNext
    Wend
nor:
    cos = cos & "</Table>"
    sf_enviarMail "secrehit@hit.cat", Desti, Estat & " Resum Comunicacions Dia " & dia, cos, "", ""
    'Desti = EmailGuardia
    'sf_enviarMail "secrehit@hit.cat", Desti, Estat & " Resum Comunicacions Dia " & Dia, Cos, "", ""
    
End Sub

Sub EmailPedidoTVoice(Proveedor As String, almacen As String, fPedidoStr As String, repartidor As String)
    Dim Rs As rdoResultset
    Dim diaPedido As String, horaPedido As String
    Dim pNom As String, pTlf1 As String, pTlf2 As String, pFax As String, pMail As String, pDto As Double
    Dim nEmp As String, eNom As String, eNif As String, eTel As String, rMail As String
    Dim sql As String
    Dim baseIva1 As Double, baseIva2 As Double, baseIva3 As Double, totalSinIVA As Double, totalIVA As Double, dtoMat As Double
    Dim preuDte As Double, totalDescuento As Double
    Dim BASEIVA3b As Double, iva3 As Double, BASEIVA2b As Double, iva2 As Double, BASEIVA1b As Double, iva1 As Double
    Dim total1 As Double, total2 As Double, total3 As Double, TotalIva2 As Double, Dpp As Double
    Dim cos As String
    
On Error GoTo nor:

    rMail = ""
    Set Rs = Db.OpenResultset("select * from dependentesextes where ID='" & repartidor & "' and nom='EMAIL'")
    If Not Rs.EOF Then rMail = Rs("valor")

    diaPedido = Split(fPedidoStr, " ")(0)
    horaPedido = Split(fPedidoStr, " ")(1)
    
    Set Rs = Db.OpenResultset("select * from ccproveedores where id = '" & Proveedor & "'")
    If Not Rs.EOF Then
        pNom = Rs("Nombre")
        pTlf1 = Rs("Tlf1")
        pTlf2 = Rs("Tlf2")
        pFax = Rs("Fax")
        pMail = Rs("Email")
        pDto = Rs("descuento")
    End If
    
    Dpp = 0
    Set Rs = Db.OpenResultset("select isnull(valor, 0) dpp from ccproveedoresextes where nom='DPP' and id='" & Proveedor & "'")
    If Not Rs.EOF Then Dpp = Rs("dpp")

    nEmp = "Camp"
    Set Rs = Db.OpenResultset("select * from constantsempresa where camp = 'predeterminada'")
    If Not Rs.EOF Then nEmp = Rs("valor")
    
    Set Rs = Db.OpenResultset("select a1.valor as nom, a2.valor as nif ,a3.valor as tel  from constantsempresa a1 left join constantsempresa  A2 on a2.camp = '" & nEmp & "nif' left join constantsempresa  A3 On a3.camp = '" & nEmp & "Tel' where a1.camp = '" & nEmp & "Nom' ")
    If Not Rs.EOF Then
        eNom = Rs("nom")
        eNif = Rs("nif")
        eTel = Rs("tel")
    End If

    sql = "select cc.proveedor as Proveedor, fecha, recepcion, m.nombre MateriaNombre, m.codigo MateriaCodigo, "
    sql = sql & " isnull(cc.precio, 0) precioFormato, isnull(a.nombre,isnull(c.nom, 'No Definido')) Almacen, cantidad, "
    sql = sql & "isnull(nv.valor,'') formato, isnull(m.unidades, 1) unidades, isnull(m.iva, 2) iva, isnull(nv2.valor, 0) dtoMat, isnull(nv3.valor , 'No Definido') refinterna  "
    sql = sql & "from ccpedidos cc "
    sql = sql & "join ccmateriasprimas m on m.id = cc.MateriaPrima "
    sql = sql & "left join ccnombrevalor nv2 on m.id=nv2.id and nv2.nombre = 'DTO' "
    sql = sql & "join ccProveedores p on p.id = cc.proveedor "
    sql = sql & "left join CCMatprop mpo on mpo.id=m.id "
    sql = sql & "left join ccalmacenes a on a.id= cc.almacen "
    sql = sql & "left join clients c on cast(c.codi as nvarchar) = case when len(cc.almacen)>6 then substring(cc.almacen, 8, len(cc.almacen)-6) else '' end "
    sql = sql & "left join ccnombrevalor nv on m.id=nv.id and nv.nombre = 'Formato' "
    sql = sql & "left join ccnombrevalor nv3 on m.id=nv3.id and nv3.nombre = 'Refinterna' "
    sql = sql & "Where cc.proveedor = '" & Proveedor & "' and cc.activo = 1 and "
    sql = sql & "cc.fecha between CONVERT(datetime, '" & diaPedido & "', 103)+CONVERT(datetime,'00:00:00', 108) and CONVERT(datetime, '" & diaPedido & "', 103)+CONVERT(datetime,'" & horaPedido & "', 108) "
    If almacen <> "" Then
        sql = sql & " and cc.Almacen = '" & almacen & "' "
    End If
    sql = sql & "order by CASE isnull(mpo.orden,'') when '' THEN '9999' ELSE mpo.orden END "
    Set Rs = Db.OpenResultset(sql)
    
    baseIva1 = 0
    baseIva2 = 0
    baseIva3 = 0
    
    If Not Rs.EOF Then
        cos = "<TABLE BORDER=""0"" CELLPADDING=""2"">"
        cos = cos & "  <TR><TD>Proveedor: " & pNom & "</TD></TR>"
        cos = cos & "  <TR><TD>Fecha Recepcion: " & Rs("recepcion") & "</TD></TR>"
        cos = cos & "  <TR><TD>&nbsp;</TD></TR>"
        cos = cos & "  <TR><TD>PEDIDO PARA: " & eNom & " " & eNif & " " & eTel & "</TD></TR>"
        cos = cos & "  <TR><TD>&nbsp;</TD></TR>"
        cos = cos & "  <TR><TD>Proveedor: " & pNom & " Tel :" & pTlf1 & " Fax :" & pFax & " Email :" & pMail & " Fecha Pedido " & Rs("fecha") & "</TD></TR>"
        cos = cos & "</TABLE>"
        
        cos = cos & "<TABLE BORDER=""0"" CELLPADDING=""2"">"
        cos = cos & "  <TR>"
        cos = cos & "    <TD><B>Ref.prov</B></TD>"
        cos = cos & "    <TD><B>Producto</B></TD>"
        cos = cos & "    <TD><B>Cantidad</B></TD>"
        cos = cos & "    <TD><B>Formato</B></TD>"
        cos = cos & "    <TD><B>Precio</B></TD>"
        cos = cos & "    <TD><B>Dto</B></TD>"
        cos = cos & "    <TD><B>Iva</B></TD>"
        cos = cos & "    <TD><B>Importe</B></TD>"
        cos = cos & "  </TR>"

        totalSinIVA = 0
        totalIVA = 0
        dtoMat = 0
        While Not Rs.EOF
            If Rs("dtoMat") = "" Then
                dtoMat = 0
            Else
                dtoMat = CDbl(Rs("dtoMat"))
            End If
        
            cos = cos & "  <TR>"
            cos = cos & "    <TD>" & Rs("refinterna") & "</TD>"
            cos = cos & "    <TD>" & Rs("MateriaNombre") & "</TD>"
            cos = cos & "    <TD>" & Rs("Cantidad") & " " & Rs("Formato") & "</TD>"
            cos = cos & "    <TD>   1x" & Rs("unidades") & "</TD>"
            cos = cos & "    <TD>" & CDbl(Rs("precioFormato")) & " &euro;" & "</TD>"
            
            preuDte = CDbl(Rs("precioFormato")) - (CDbl(Rs("precioFormato")) * (dtoMat / 100))
            totalSinIVA = totalSinIVA + preuDte * (CDbl(Rs("Cantidad")))
            totalDescuento = totalSinIVA - (totalSinIVA * (CDbl(pDto) / 100))
            
            cos = cos & "    <TD>" & dtoMat & " %</TD>"
            
            
            If Rs("iva") = 1 Then
                baseIva1 = baseIva1 + preuDte * (CDbl(Rs("Cantidad")))
                cos = cos & "    <TD>4 %</TD>"
            ElseIf Rs("iva") = 2 Then
                baseIva2 = baseIva2 + preuDte * (CDbl(Rs("Cantidad")))
                cos = cos & "    <TD>10 %</TD>"
            ElseIf Rs("iva") = 3 Then
                baseIva3 = baseIva3 + preuDte * (CDbl(Rs("Cantidad")))
                cos = cos & "    <TD>21 %</TD>"
            Else
                baseIva1 = baseIva1 + preuDte * (CDbl(Rs("Cantidad")))
                cos = cos & "    <TD>0 %</TD>"
            End If
            
            cos = cos & "    <TD>" & preuDte * CDbl(Rs("cantidad")) & " &euro;</TD>"
            
            If CDbl(pDto) > 0 Then
                BASEIVA1b = baseIva1 - (baseIva1 * (CDbl(pDto) / 100))
                BASEIVA2b = baseIva2 - (baseIva2 * (CDbl(pDto) / 100))
                BASEIVA3b = baseIva3 - (baseIva3 * (CDbl(pDto) / 100))
                iva1 = BASEIVA1b * (4 / 100)
                iva2 = BASEIVA2b * (10 / 100)
                iva3 = BASEIVA3b * (21 / 100)
                total1 = (baseIva1 - (baseIva1 * (CDbl(pDto) / 100))) + iva1
                total2 = (baseIva2 - (baseIva2 * (CDbl(pDto) / 100))) + iva2
                total3 = (baseIva3 - (baseIva3 * (CDbl(pDto) / 100))) + iva3
                TotalIva2 = iva1 + iva2 + iva3
                totalIVA = totalDescuento + TotalIva2
            Else
                BASEIVA1b = baseIva1
                BASEIVA2b = baseIva2
                BASEIVA3b = baseIva3
                iva1 = BASEIVA1b * (4 / 100)
                iva2 = BASEIVA2b * (10 / 100)
                iva3 = BASEIVA3b * (21 / 100)
                total1 = baseIva1 + iva1
                total2 = baseIva2 + iva2
                total3 = baseIva3 + iva3
                TotalIva2 = iva1 + iva2 + iva3
                totalIVA = totalSinIVA + TotalIva2
            End If
        
            totalIVA = total1 + total2 + total3
                
            cos = cos & "  </TR>"
        
            Rs.MoveNext
        Wend
        
        cos = cos & "<TR><TD COLSPAN=""2""><B>TOTAL BRUTO</B></TD><TD COLSPAN=""6"" ALIGN=""RIGHT""> " & FormatNumber(totalSinIVA, 2) & " &euro;</TD></TR>"
        If CDbl(pDto) > 0 Then
            cos = cos & "<TR><TD COLSPAN=""2""><B>DTO " & pDto & " %</B></TD><TD COLSPAN=""6"" ALIGN=""RIGHT""> " & FormatNumber(totalDescuento, 2) & " &euro;</TD></TR>"
        End If
        cos = cos & "</TABLE>"
        
        cos = cos & "<BR>"
        
        cos = cos & "<TABLE BORDER=""0"" CELLPADDING=""2"">"
        
        cos = cos & "<TR><TD><B>BASE IVA 21 % </B> = " & FormatNumber(BASEIVA3b, 2) & " &euro;</TD><TD><B>IVA 21 %</B> = " & FormatNumber(iva3, 2) & " &euro;</TD></TR>"
        cos = cos & "<TR><TD><B>BASE IVA 10 % </B> = " & FormatNumber(BASEIVA2b, 2) & " &euro;</TD><TD><B>IVA 10 %</B> = " & FormatNumber(iva2, 2) & " &euro;</TD></TR>"
        cos = cos & "<TR><TD><B>BASE IVA 4 % </B> = " & FormatNumber(BASEIVA1b, 2) & " &euro;</TD><TD><B>IVA 4 %</B> = " & FormatNumber(iva1, 2) & " &euro;</TD></TR>"
        
        cos = cos & "<TR><TD COLSPAN=""2""><B>TOTAL IMPUESTO BASE </B>" & FormatNumber(TotalIva2, 2) & " &euro;</TD></TR>"
        cos = cos & "<TR><TD COLSPAN=""2""><B>TOTAL </B>" & FormatNumber(totalIVA, 2) & " &euro;</TD></TR>"
        
    End If
    Rs.Close
    
                

   '         If CDbl(Dpp) > 0 Then
   '            lin = lin + 1
   '           Page.Canvas.DrawText "TOTAL CON DPP " & Dpp & " %", "x=300; y=" & Header - salto * lin & "; width=" & Page.Width - (marg * 2) & "; alignment=left; size=10; color=&H000000", VERDANA
   '          Page.Canvas.DrawText FormatNumber(totalIVA - (totalIVA * (CDbl(Dpp) / 100)), 2) & " €", "x=500; y=" & Header - salto * lin & "; width=75; alignment=right; size=8; color=&H000000", VERDANAB
   '      End If
    
nor:
    sf_enviarMail "email@hit.cat", pMail, " PEDIDO de " & eNom, cos, "", ""
    sf_enviarMail "email@hit.cat", rMail, " PEDIDO de " & eNom, cos, "", ""
    
End Sub


Sub EmailResumBotigaEstat(Desti As String, sDia As String)
    Dim Estat As String, dia As Date, cos As String, Cos2 As String, Rs As rdoResultset, Rs5 As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset, Rs4 As rdoResultset, Co As String, Es As String
    Dim sql, dd, mm, yyyy, dd2, mm2, yyyy2, nBot As Integer, nBal As Integer, vMysql, DataPing As Date, Di As Date, Df As Date
    Dim EurosPing As Double
    
    'ExecutaComandaSql "CREATE TABLE [PingMaquina](   [Id] [nvarchar](255) NULL,   [Llicencia] [float] NULL,   [TmSt] [datetime] NULL,   [Param1] [nvarchar](255) NULL,   [Param2] [nvarchar](255) NULL,    [Param3] [nvarchar](255) NULL ) ON [PRIMARY]"
    'ExecutaComandaSql "delete from hit.dbo.PingMaquina where DATEDIFF(d,tmst,getdate()) > 5 and TmSt not in (select MAX(tmst) tmst from Hit.dbo.PingMaquina Where Llicencia = hit.dbo.PingMaquina.Llicencia group by convert(datetime,CONVERT(nvarchar(10),tmst,101)),llicencia)"
    'ExecutaComandaSql "insert into PingMaquina select * from  hit.dbo.PingMaquina where DATEDIFF(d,tmst,getdate()) > 5 and llicencia in(select distinct llicencia from hit.dbo.llicencies where Empresa='" & EmpresaActual & "') "
    'ExecutaComandaSql "Delete hit.dbo.PingMaquina where DATEDIFF(d,tmst,getdate()) > 5 and llicencia in(select distinct llicencia from hit.dbo.llicencies where Empresa='" & EmpresaActual & "') "
        
    dia = Now
    If IsNumeric(sDia) Then dia = DateAdd("d", -sDia, Now)
    If IsDate(sDia) Then dia = CVDate(sDia)
    Estat = ":-|"
    Di = DateSerial(Year(dia), Month(dia), 1)
    Df = DateAdd("m", 1, Di)
    Df = DateAdd("d", -1, Df)
    cos = ""
'    While Not Di > Df
        Set Rs = Db.OpenResultset("select c.Codi,c.nom,w.Codi Wcodi  from ParamsHw w join clients c on w.Valor1 = c.Codi Order by c.nom ")
        While Not Rs.EOF
            nBot = Left(Rs("codi"), 2)
            nBal = Right(Rs("codi"), 1)
            If Rs("Codi") = 518 Then
                nBal = 1
                nBot = 106
            End If
            Co = "<Td>" & Rs("Nom") & "</Td><Td>"
            Set Rs2 = Db.OpenResultset("select CASE WHEN Tipus='' THEN 'Antiga' ELSE CONVERT(NVARCHAR(255),ISNULL(Tipus ,'Antiga')) END tipus from hit.dbo.llicencies where Llicencia= " & Rs("WCodi"))
            If Not Rs2.EOF Then Co = Co & "<Td>" & DameValor(Rs2, "Tipus") & "</Td><Td>"
            cos = cos & "<Tr><Td>" & Es & "</Td>" & Co & "</Td></Tr><Br>"
            Rs.MoveNext
        Wend
        Rs.Close
'        Di = DateAdd("D", 1, Di)
'    Wend
    
    cos = cos & "</Table>"
    sf_enviarMail "secrehit@hit.cat", Desti, Estat & " Resum Versions ", cos, "", ""
    'Desti = EmailGuardia
    'sf_enviarMail "secrehit@hit.cat", Desti, Estat & " Resum Comunicacions Dia " & Dia, Cos, "", ""
    
End Sub

Sub EmailRecepcio(nAlbara As String, idProveedor As String)
    Dim Rs As rdoResultset, rsRec As rdoResultset
    Dim sql As String, textEmail As String, Desti As String
    
    Informa "Enviant emails Recepcio"
    
    Set Rs = Db.OpenResultset("select * From dependentesextes where ID IN (select id from dependentesextes where nom = 'COMPRES' and valor='1') AND nom = 'EMAIL' and valor<>''")
    If Not Rs.EOF Then
        textEmail = "<TABLE CELLPADDING='2' BORDER='1'>"
        textEmail = textEmail & "<TR><TD>DATA</TD><TD>PROVEEDOR</TD><TD>ALBARÀ</TD><TD>LOT</TD><TD>PRODUCTE</TD><TD>QTAT REBUDA</TD></TR>"
        
        sql = "select r.id idRecepcion, r.caducidad Caducidad, isnull(r.pedido,'') pedido, r.fecha, p.nombre pro, "
        sql = sql & "r.albaran, r.lote, m.nombre mat, caract, envas, usuario, isnull(d.cantidad,0) cant, d.fecha fped "
        sql = sql & "from ccrecepcion r "
        sql = sql & "left join ccproveedores p on p.id=r.proveedor "
        sql = sql & "left join ccMateriasPrimas m on m.id=r.matPrima "
        sql = sql & "left join ccpedidos d on d.id=r.pedido "
        sql = sql & "where r.albaran ='" & nAlbara & "' and r.proveedor='" & idProveedor & "' "
        sql = sql & "order by r.fecha"
        
        Set rsRec = Db.OpenResultset(sql)
        While Not rsRec.EOF
            textEmail = textEmail & "<TR>"
            textEmail = textEmail & "<TD>" & rsRec("fecha") & "</TD>"
            textEmail = textEmail & "<TD>" & rsRec("pro") & "</TD>"
            textEmail = textEmail & "<TD>" & rsRec("albaran") & "</TD>"
            textEmail = textEmail & "<TD>" & rsRec("lote") & "</TD>"
            textEmail = textEmail & "<TD>" & rsRec("mat") & "</TD>"
            textEmail = textEmail & "<TD>" & rsRec("cant") & "</TD>"
            textEmail = textEmail & "</TR>"
            rsRec.MoveNext
        Wend
        
        textEmail = textEmail & "</TABLE>"
    End If
        
    While Not Rs.EOF
        Informa2 "Enviem a :" & Desti
        Desti = Rs("valor")
        sf_enviarMail "secrehit@hit.cat", Desti, " Recepció de materia ", textEmail, "", ""
        Rs.MoveNext
        Informa2 "Enviat a :" & Desti
    Wend
    
End Sub


Sub empresaEmail(sls_de, SLS_SMTPSERVER, SLS_SMTPUSERNAME, SLS_SMTPPASSWORD, SLN_SMTPSERVERPORT)
    Dim rsx As rdoResultset, Ps, Pu, Pp, Port
    Dim Rs As rdoResultset
    Dim sql As String
    
On Error GoTo nor
    sql = "select isnull(valor,'') valor  from constantsempresa where camp = 'CampServidorSMTP'"
    Set Rs = Db.OpenResultset("select isnull(valor,'') valor  from constantsempresa where camp = 'CampServidorSMTP'")
    If Not Rs.EOF Then
        Ps = Rs("Valor")
        Set Rs = Db.OpenResultset("select isnull(valor,'') valor  from constantsempresa where camp = 'CampUsuariSMTP'")
        If Not Rs.EOF Then
            Pu = Rs("Valor")
            Set Rs = Db.OpenResultset("select isnull(valor,'') valor  from constantsempresa where camp = 'CampContrasenyaSMTP'")
            If Not Rs.EOF Then
                Pp = Rs("Valor")
                Set Rs = Db.OpenResultset("select isnull(valor,'') valor  from constantsempresa where camp = 'CampPORT'")
                If Not Rs.EOF Then
                    Port = Rs("Valor")
                End If
            End If
        End If
    End If
nor:
   If Len(Ps) > 0 And Len(Pu) > 0 And Len(Pp) > 0 Then
       SLS_SMTPSERVER = Ps
        SLS_SMTPUSERNAME = Pu
        SLS_SMTPPASSWORD = Pp
    End If
    If Len(Port) > 0 Then
        SLN_SMTPSERVERPORT = Port
    End If
End Sub

Sub EnviaPressupostEmail(client, IdFile, numPress, tabPress)
    Dim Texte As String, empCodi As Double, EmpSerie As String, empNom, empNif, empAdresa, empCp, EmpLliure, empTel, empFax, empEMail, ClientCampMercantil, empCiutat, empresa, EmailClient, Rs, TenimFile As Boolean, IdiomaClient As String
    Dim rsCom As rdoResultset, sql As String
    Dim idPress As String
        
    CarregaDadesEmpresa client, empCodi, EmpSerie, empNom, empNif, empAdresa, empCp, EmpLliure, empTel, empFax, empEMail, ClientCampMercantil, empCiutat, empresa
    
    EmailClient = ""
    Set Rs = Db.OpenResultset("select isnull(eMail,'') eMail, idFactura from " & tabPress & " where numFactura='" & numPress & "'")
    If Not Rs.EOF Then
        EmailClient = Rs("eMail")
        idPress = Rs("idFactura")
    End If
    If EmailClient = "" Then
       Set Rs = Db.OpenResultset("Select isnull(valor,'') valor from constantsClient  where variable = 'eMail' and codi = " & client)
        If Not Rs.EOF Then EmailClient = Rs("Valor")
    End If
    
    IdiomaClient = ""
    Set Rs = Db.OpenResultset("select isnull(valor,'') valor from constantsclient where variable='IDIOMA' and codi = " & client)
    If Not Rs.EOF Then IdiomaClient = Rs("Valor")
        
    TenimFile = False
    Set Rs = Db.OpenResultset("select Count(*) from Archivo Where id  = '" & IdFile & "' ")
    If Not Rs.EOF Then TenimFile = True
    
'EmailClient = "jordi.bosch.maso@gmail.com"
    If Not (Not TenimFile Or empEMail = "" Or EmailClient = "") Then
        If IdiomaClient = "ES" Then
            Texte = "Estimado Cliente," & "<Br>"
            Texte = Texte & "Le adjuntamos Presupuesto, correspondiente a los Servicios solicitados a nuestra entidad." & "<Br>"
            Texte = Texte & "Atentamente," & "<Br>"
            Texte = Texte & "Dpto. admón." & "<Br>"
            Texte = Texte & empNom & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Oficinas Centrales :" & "<Br>"
            Texte = Texte & empAdresa & " (" & empCp & "), " & empCiutat & "<Br>"
            Texte = Texte & "Teléfono: " & empTel & "<Br>"
            Texte = Texte & "Fax: " & empFax & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
            Texte = Texte & "Los datos de carácter personal que constan en el presente Presupuesto son tratados e incorporados en un fichero responsabilidad de " & empNom & ".  Conforme a lo dispuesto en los artículos 15 y 16 de la Ley Orgánica 15/1999, de 13 de diciembre, de Protección de Datos de Carácter Personal, le informamos que puede ejercitar los derechos de acceso, rectificación, cancelación y oposición en: " & empAdresa & ", " & empCp & " (" & empCiutat & "), o bien, enviar un correo electrónico a: " & empEMail & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Piense en el medio ambiente antes de imprimir este e-mail" & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
        Else
            Texte = "Apreciat Client," & "<Br>"
            Texte = Texte & "Li adjuntem Pressupost, corresponent als Serveis sol·licitats a la nostra entitat." & "<Br>"
            Texte = Texte & "Atentament," & "<Br>"
            Texte = Texte & "Dpt. adm." & "<Br>"
            Texte = Texte & empNom & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Oficines Centrals :" & "<Br>"
            Texte = Texte & empAdresa & " (" & empCp & "), " & empCiutat & "<Br>"
            Texte = Texte & "Telèfon: " & empTel & "<Br>"
            Texte = Texte & "Fax: " & empFax & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
            Texte = Texte & "Les dades de caràcter personal que consten al present Pressupost son tractats i incorporats en un fitxer responsabilitat de " & empNom & ".  Conforme al disposat als artícles 15 y 16 de la Llei Orgànica 15/1999, de 13 de desembre, de Protecció de Dades de Caràcter Personal, l'informem que pot exercir els drets d'accès, rectificació, cancelació i oposició en: " & empAdresa & ", " & empCp & " (" & empCiutat & "), o bé, enviar un correu electrònic a: " & empEMail & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Pensi en el medi ambient abans d'imprimir aquest e-mail" & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
        End If
        sf_enviarMail empEMail, EmailClient, empNom & " ,pressupost ", Texte, IdFile, "c:\Pressupost.Pdf"
        sf_enviarMail "", empEMail, "Enviat Pressupost a " & EmailClient & " " & empNom & " ", Texte, IdFile, "c:\Pressupost.Pdf"
        
        Set rsCom = Db.OpenResultset("select idFactura from FacturacioComentaris where idFactura='" & idPress & "'")
        If Not rsCom.EOF Then
            sql = "update FacturacioComentaris set Enviada='" & Now() & " OK' where idFactura='" & idPress & "'"
        Else
            sql = "insert into FacturacioComentaris (IdFactura, Comentari, Cobrat, Data, Enviada) values('" & idPress & "', '', '', getdate(), '" & Now() & " OK')"
        End If
        ExecutaComandaSql sql

        
    End If
    
    ExecutaComandaSql "Delete archivo Where id  = '" & IdFile & "' "

End Sub

Sub FacturaSms(mov, Miss)
    Dim sql As String
    
      sql = "CREATE TABLE Hit.[Dbo].SmsEnviats( "
      sql = sql & " [Id]        [nvarchar] (255) NULL ,"
      sql = sql & " [TimeStamp] [datetime] NULL ,"
      sql = sql & " [Client]    [nvarchar] (255) NULL ,"
      sql = sql & " [NumTel]    [nvarchar] (255) NULL ,"
      sql = sql & " [Texte]     [nvarchar] (255) NULL ,"
      sql = sql & " [Desde]     [nvarchar] (255) NULL ,"
      sql = sql & " [Facturat]  [nvarchar] (255) NULL ,"
      sql = sql & " [Aux1]      [nvarchar] (255) NULL ,"
      sql = sql & " [Aux2]      [nvarchar] (255) NULL ,"
      sql = sql & " [Aux3]      [float] NULL)"
      ExecutaComandaSql sql
    
      ExecutaComandaSql "Insert into Hit.Dbo.SmsEnviats (Id,TimeStamp,Client,NumTel,Texte,Desde,Facturat) Values (newId(),getdate(),'" & EmpresaActual & "','" & mov & "','" & Miss & "','','1')"
    
    
End Sub

Sub GetEmail()
   Dim P As Integer
   Dim Pp As Integer, Frase As String
   
   
'Automatic@HitSystems.Net
frmSplash.Show
frmSplash.POP1.User = "secrehit@gmail.com"
'frmSplash.POP1.Password = "secrehit2130"
' CONTRASEÑA APLICACION
frmSplash.POP1.Password = "njnservqauapdyvd"
frmSplash.POP1.MailServer = "pop.gmail.com"
frmSplash.POP1.MailPort = 995
frmSplash.POP1.Action = 1   'Connect
    
For gCurrMsg = 1 To frmSplash.POP1.MessageCount
    frmSplash.POP1.MessageNumber = gCurrMsg
    frmSplash.POP1.MaxLines = 0 'all lines
    frmSplash.POP1.Action = 3 'Retrieve Message
    
    Debug.Print gMessages(gCurrMsg).Text
    P = InStr(gMessages(gCurrMsg).Text, "· · · · · · · · · · · · · · · · · · · ·")
    If P > 0 Then
       Pp = InStr(P, gMessages(gCurrMsg).Text, Chr(13) & Chr(10))
       
    End If
' · · · · · · · · · · · · · · · · · · · ·
' El gran moment de l'amor és el d'abans de començar. La prehistòria del sentiment.
' Josep Pla(1897 - 1981)
    
'    frmSplash.POP1.Action = a_Delete
    
    
Next gCurrMsg



End Sub
Sub EnviaEmailAdjunto(Desti, Caption, FileId)
    
    sf_enviarMail "Secrehit@gmail.com", Desti, Caption, "Te adjunto documento Solicitado. " & Chr(13) & Chr(10) & "Atentamente Joana.", FileId, "c:\Excel.Xls"
    ExecutaComandaSql "Delete archivo Where id  = '" & FileId & "' "
    
End Sub


Sub EnviaEmail(Desti, Caption)
    
    sf_enviarMail "Secrehit@gmail.com", Desti, Caption, "", "", ""
    
End Sub
Sub EnviaFacturaEmail(client, IdFile, numFactura As String, tabFactura, idFactura As String)
    Dim Texte As String, empCodi As Double, EmpSerie As String, empNom, empNif, empAdresa, empCp, EmpLliure, empTel, empFax, empEMail, ClientCampMercantil, empCiutat, empresa, EmailClient, TenimFile As Boolean, IdiomaClient As String
    Dim URL, resposta
    Dim rsCom As rdoResultset
    Dim rs1 As rdoResultset
    Dim Rs2 As rdoResultset
    Dim sql As String
    Dim para() As String, P As Integer
    Dim dataFactura As Date
    
    On Error GoTo ERR_EMAIL
    
    empCodi = "0"
    Set rs1 = Db.OpenResultset("select isnull(empresaCodi,'') empCodi, dataFactura from " & tabFactura & " where numFactura='" & numFactura & "' and IdFactura='" & idFactura & "'")
    If Not rs1.EOF Then
        empCodi = rs1("empCodi")
        dataFactura = CDate(rs1("dataFactura"))
    End If
    
    CarregaDadesEmpresa client, empCodi, EmpSerie, empNom, empNif, empAdresa, empCp, EmpLliure, empTel, empFax, empEMail, ClientCampMercantil, empCiutat, empCodi

    If tabFactura = "" Then
        sf_enviarMail "Secrehit@gmail.com", EmailGuardia, "Error en calculs llargs enviada factura a cliente con tabla factura vacio :( ", "", "", ""
        Exit Sub
    End If
    
    EmailClient = ""
    Set rs1 = Db.OpenResultset("select isnull(eMail,'') eMail from " & tabFactura & " where numFactura='" & numFactura & "' and IdFactura='" & idFactura & "'")
    If Not rs1.EOF Then EmailClient = rs1("eMail")
    If EmailClient = "" Then
       Set rs1 = Db.OpenResultset("Select isnull(valor,'') valor from constantsClient  where variable = 'eMail' and codi = " & client)
        If Not rs1.EOF Then EmailClient = rs1("Valor")
    End If
    'EmailClient = "ana@solucionesit365.com"
    
    IdiomaClient = ""
    Set rs1 = Db.OpenResultset("select isnull(valor,'') valor from constantsclient where variable='IDIOMA' and codi = " & client)
    If Not rs1.EOF Then IdiomaClient = rs1("Valor")
    
    '-------------------------------------------------------------------------
    Dim a As New Stream, s() As Byte, iD As String
    Dim Rs As ADODB.Recordset
    
    On Error Resume Next
    db2.Close
    On Error GoTo 0
   
    On Error GoTo ERR_EMAIL
    
    Set db2 = New ADODB.Connection
    db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"

    FacturaXML idFactura, numFactura, dataFactura
    
    Set Rs = rec("select newid() i")
    iD = Rs("i")
    
    a.Open
    a.LoadFromFile "c:\" & numFactura & ".xml"
    s = a.ReadText()
  
    Set Rs = rec("select * from " & tablaArchivo() & " where fecha=(select max(fecha) from " & tablaArchivo() & ")", True)
    Rs.AddNew
    Rs("id").Value = iD
  
    Rs("nombre").Value = Left("XML Factura " & numFactura, 20)
    Rs("descripcion").Value = Left("XML Factura " & numFactura, 250)
    Rs("extension").Value = "XML"
    Rs("mime").Value = "application/xml"
    Rs("propietario").Value = ""
    Rs("archivo").Value = s
    Rs("fecha").Value = Now
    Rs("tmp").Value = 0
    Rs("down").Value = 1
    Rs.Update
    Rs.Close
    a.Close
    '-------------------------------------------------------------------------
    
    
    'REVISAR!
    TenimFile = False
    If IdFile <> "" Then
        Set rs1 = Db.OpenResultset("select Count(*) from Archivo Where id  = '" & IdFile & "' ")
        If Not rs1.EOF Then TenimFile = True
        'Si no existeix factura
        If rs1.EOF Then
            Set rs1 = Db.OpenResultset("select IdFactura,DataInici,DataFi from " & tabFactura & " where idFactura='" & idFactura & "'")
            If Not rs1.EOF Then
                URL = "http://silema.hiterp.com/Facturacion/ElForn/facturas/imprimirFacturasSINCRO.asp?"
                URL = URL & "id=" & rs1("idFactura")
                URL = URL & "&tab=" & tabFactura
                URL = URL & "&cli=" & client
                URL = URL & "&ini=" & rs1("Datainici")
                URL = URL & "&fin=" & rs1("DataFi")
                URL = URL & "&sav=5"
                resposta = llegeigHtml(URL)
                If resposta <> "" Then IdFile = resposta
                TenimFile = True
            End If
        End If
    Else
        Set rs1 = Db.OpenResultset("select IdFactura,DataInici,DataFi from " & tabFactura & " where idFactura='" & idFactura & "'")
        If Not rs1.EOF Then
            Set Rs2 = Db.OpenResultset("select id from hit.dbo.web_empreses where db = '" & LastDatabase & "'")
            If Not Rs2.EOF Then
                URL = "http://silema.hiterp.com/Facturacion/ElForn/facturas/imprimirFacturasSINCRO.asp?"
                URL = URL & "id=" & rs1("idFactura")
                URL = URL & "&tab=" & Left(tabFactura, 20)
                URL = URL & "&tabIva=" & tabFactura
                URL = URL & "&tabData=" & Left(tabFactura, 20) & "data]"
                URL = URL & "&cli=" & client
                URL = URL & "&ini=" & rs1("Datainici")
                URL = URL & "&fin=" & rs1("DataFi")
                URL = URL & "&sav=4"
                URL = URL & "&empresa=" & Rs2("id")
                resposta = llegeigHtml(URL)
                If resposta <> "" Then
                    If InStr(resposta, "END") Then
                        IdFile = Mid(resposta, 1, InStr(resposta, "END") - 1)
                        TenimFile = True
                    End If
                Else
                    TenimFile = False
                End If
            Else
                TenimFile = False
            End If
        End If
    End If
'EmailClient = "eze@hitsystems.es"
    If TenimFile And (empEMail <> "" Or EmailClient <> "") Then
        If IdiomaClient = "ES" Then
            Texte = "Estimado Cliente," & "<Br>"
            Texte = Texte & "Le adjuntamos Factura." & "<Br>"
            Texte = Texte & "Atentamente," & "<Br>"
            Texte = Texte & "Dpto. admin." & "<Br>"
            Texte = Texte & empNom & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Oficinas Centrales :" & "<Br>"
            Texte = Texte & empAdresa & " (" & empCp & "), " & empCiutat & "<Br>"
            Texte = Texte & "Teléfono: " & empTel & "<Br>"
            Texte = Texte & "Fax: " & empFax & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
            Texte = Texte & "Los datos de carácter personal que constan en la presente Factura son tratados e incorporados en un fichero responsabilidad de " & empNom & ".  Conforme a lo dispuesto en los artículos 15 y 16 de la Ley Orgánica 15/1999, de 13 de diciembre, de Protección de Datos de Carácter Personal, le informamos que puede ejercitar los derechos de acceso, rectificación, cancelación y oposición en: " & empAdresa & ", " & empCp & " (" & empCiutat & "), o bien, enviar un correo electrónico a: " & empEMail & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Piense en el medio ambiente antes de imprimir este e-mail" & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
        Else
            Texte = "Apreciat Client," & "<Br>"
            Texte = Texte & "Li adjuntem Factura." & "<Br>"
            Texte = Texte & "Atentament," & "<Br>"
            Texte = Texte & "Dpt. adm." & "<Br>"
            Texte = Texte & empNom & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Oficines Centrals :" & "<Br>"
            Texte = Texte & empAdresa & " (" & empCp & "), " & empCiutat & "<Br>"
            Texte = Texte & "Telèfon: " & empTel & "<Br>"
            Texte = Texte & "Fax: " & empFax & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
            Texte = Texte & "Les dades de caràcter personal que consten a la present Factura son tractats i incorporats en un fitxer responsabilitat de " & empNom & ".  Conforme al disposat als artícles 15 y 16 de la Llei Orgànica 15/1999, de 13 de desembre, de Protecció de Dades de Caràcter Personal, l'informem que pot exercir els drets d'accès, rectificació, cancelació i oposició en: " & empAdresa & ", " & empCp & " (" & empCiutat & "), o bé, enviar un correu electrònic a: " & empEMail & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Pensi en el medi ambient abans d'imprimir aquest e-mail" & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
        End If
        
        para = Split(EmailClient, ";")
        For P = 0 To UBound(para)
            sf_enviarMail empEMail, Trim(para(P)), "NO CONTESTAR AQUEST EMAIL. " & empNom & " ,factura ", Texte, IdFile, "c:\Factura.Pdf", iD, "c:\" & numFactura & ".xml"
        Next
        
        sf_enviarMail "email@hit.cat", empEMail, "Enviada Factura a " & EmailClient & " " & empNom & " ", Texte, IdFile, "c:\Factura.Pdf", iD, "c:\" & numFactura & ".xml"

        Set rsCom = Db.OpenResultset("select idFactura from FacturacioComentaris where idFactura='" & idFactura & "'")
        If Not rsCom.EOF Then
            sql = "update FacturacioComentaris set Enviada='" & Now() & " OK' where idFactura='" & idFactura & "'"
        Else
            sql = "insert into FacturacioComentaris (IdFactura, Comentari, Cobrat, Data, Enviada) values('" & idFactura & "', '', 'N', getdate(), '" & Now() & " OK')"
        End If
        
        ExecutaComandaSql sql
    Else
        sf_enviarMail "email@hit.cat", empEMail, "Error al Enviar Factura a " & EmailClient & " " & empNom & " ", Texte, IdFile, "c:\Factura.Pdf"
    End If
    
ERR_EMAIL:

    ExecutaComandaSql "Delete archivo Where id  = '" & IdFile & "' "
    ExecutaComandaSql "Delete archivo Where id  = '" & iD & "' "
    
    MyKill "c:\" & numFactura & ".xml"
End Sub
Sub EnviaFacturaEmailSecre(client, IdFile, numFactura As String, tabFactura, idFactura As String)
    Dim Texte As String, empCodi As Double, EmpSerie As String, empNom, empNif, empAdresa, empCp, EmpLliure, empTel, empFax, empEMail, ClientCampMercantil, empCiutat, empresa, EmailClient, TenimFile As Boolean, IdiomaClient As String
    Dim URL, resposta
    Dim rsCom As rdoResultset
    Dim rs1 As rdoResultset
    Dim Rs2 As rdoResultset
    Dim sql As String
    Dim para() As String, P As Integer
    Dim dataFactura As Date
    
    On Error GoTo ERR_EMAIL
    
    empCodi = "0"
    Set rs1 = Db.OpenResultset("select isnull(empresaCodi,'') empCodi, dataFactura from " & tabFactura & " where numFactura='" & numFactura & "' and IdFactura='" & idFactura & "'")
    If Not rs1.EOF Then
        empCodi = rs1("empCodi")
        dataFactura = CDate(rs1("dataFactura"))
    End If
    
    CarregaDadesEmpresa client, empCodi, EmpSerie, empNom, empNif, empAdresa, empCp, EmpLliure, empTel, empFax, empEMail, ClientCampMercantil, empCiutat, empCodi

    If tabFactura = "" Then
        sf_enviarMail "Secrehit@gmail.com", EmailGuardia, "Error en calculs llargs enviada factura a cliente con tabla factura vacio :( ", "", "", ""
        Exit Sub
    End If
    
    EmailClient = "secrehit@hit.cat"
    
    IdiomaClient = ""
    Set rs1 = Db.OpenResultset("select isnull(valor,'') valor from constantsclient where variable='IDIOMA' and codi = " & client)
    If Not rs1.EOF Then IdiomaClient = rs1("Valor")
    
    '-------------------------------------------------------------------------
    Dim a As New Stream, s() As Byte, iD As String
    Dim Rs As ADODB.Recordset
    
    On Error Resume Next
    db2.Close
    On Error GoTo 0
   
    On Error GoTo ERR_EMAIL
    
    Set db2 = New ADODB.Connection
    db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"

    FacturaXML idFactura, numFactura, dataFactura
    
    Set Rs = rec("select newid() i")
    iD = Rs("i")
    
    a.Open
    a.LoadFromFile "c:\" & numFactura & ".xml"
    s = a.ReadText()
  
    Set Rs = rec("select * from " & tablaArchivo() & " where fecha=(select max(fecha) from " & tablaArchivo() & ")", True)
    Rs.AddNew
    Rs("id").Value = iD
  
    Rs("nombre").Value = Left("XML Factura " & numFactura, 20)
    Rs("descripcion").Value = Left("XML Factura " & numFactura, 250)
    Rs("extension").Value = "XML"
    Rs("mime").Value = "application/xml"
    Rs("propietario").Value = ""
    Rs("archivo").Value = s
    Rs("fecha").Value = Now
    Rs("tmp").Value = 0
    Rs("down").Value = 1
    Rs.Update
    Rs.Close
    a.Close
    '-------------------------------------------------------------------------
    
    
    'REVISAR!
    TenimFile = False
    If IdFile <> "" Then
        Set rs1 = Db.OpenResultset("select Count(*) from Archivo Where id  = '" & IdFile & "' ")
        If Not rs1.EOF Then TenimFile = True
        'Si no existeix factura
        If rs1.EOF Then
            Set rs1 = Db.OpenResultset("select IdFactura,DataInici,DataFi from " & tabFactura & " where idFactura='" & idFactura & "'")
            If Not rs1.EOF Then
                URL = "http://silema.hiterp.com/Facturacion/ElForn/facturas/imprimirFacturasSINCRO.asp?"
                URL = URL & "id=" & rs1("idFactura")
                URL = URL & "&tab=" & tabFactura
                URL = URL & "&cli=" & client
                URL = URL & "&ini=" & rs1("Datainici")
                URL = URL & "&fin=" & rs1("DataFi")
                URL = URL & "&sav=5"
                resposta = llegeigHtml(URL)
                If resposta <> "" Then IdFile = resposta
                TenimFile = True
            End If
        End If
    Else
        Set rs1 = Db.OpenResultset("select IdFactura,DataInici,DataFi from " & tabFactura & " where idFactura='" & idFactura & "'")
        If Not rs1.EOF Then
            Set Rs2 = Db.OpenResultset("select id from hit.dbo.web_empreses where db = '" & LastDatabase & "'")
            If Not Rs2.EOF Then
                URL = "http://silema.hiterp.com/Facturacion/ElForn/facturas/imprimirFacturasSINCRO.asp?"
                URL = URL & "id=" & rs1("idFactura")
                URL = URL & "&tab=" & Left(tabFactura, 20)
                URL = URL & "&tabIva=" & tabFactura
                URL = URL & "&tabData=" & Left(tabFactura, 20) & "data]"
                URL = URL & "&cli=" & client
                URL = URL & "&ini=" & rs1("Datainici")
                URL = URL & "&fin=" & rs1("DataFi")
                URL = URL & "&sav=4"
                URL = URL & "&empresa=" & Rs2("id")
                resposta = llegeigHtml(URL)
                If resposta <> "" Then
                    If InStr(resposta, "END") Then
                        IdFile = Mid(resposta, 1, InStr(resposta, "END") - 1)
                        TenimFile = True
                    End If
                Else
                    TenimFile = False
                End If
            Else
                TenimFile = False
            End If
        End If
    End If
'EmailClient = "eze@hitsystems.es"
    If TenimFile And (empEMail <> "" Or EmailClient <> "") Then
        If IdiomaClient = "ES" Then
            Texte = "Estimado Cliente," & "<Br>"
            Texte = Texte & "Le adjuntamos Factura." & "<Br>"
            Texte = Texte & "Atentamente," & "<Br>"
            Texte = Texte & "Dpto. admin." & "<Br>"
            Texte = Texte & empNom & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Oficinas Centrales :" & "<Br>"
            Texte = Texte & empAdresa & " (" & empCp & "), " & empCiutat & "<Br>"
            Texte = Texte & "Teléfono: " & empTel & "<Br>"
            Texte = Texte & "Fax: " & empFax & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
            Texte = Texte & "Los datos de carácter personal que constan en la presente Factura son tratados e incorporados en un fichero responsabilidad de " & empNom & ".  Conforme a lo dispuesto en los artículos 15 y 16 de la Ley Orgánica 15/1999, de 13 de diciembre, de Protección de Datos de Carácter Personal, le informamos que puede ejercitar los derechos de acceso, rectificación, cancelación y oposición en: " & empAdresa & ", " & empCp & " (" & empCiutat & "), o bien, enviar un correo electrónico a: " & empEMail & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Piense en el medio ambiente antes de imprimir este e-mail" & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
        Else
            Texte = "Apreciat Client," & "<Br>"
            Texte = Texte & "Li adjuntem Factura." & "<Br>"
            Texte = Texte & "Atentament," & "<Br>"
            Texte = Texte & "Dpt. adm." & "<Br>"
            Texte = Texte & empNom & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Oficines Centrals :" & "<Br>"
            Texte = Texte & empAdresa & " (" & empCp & "), " & empCiutat & "<Br>"
            Texte = Texte & "Telèfon: " & empTel & "<Br>"
            Texte = Texte & "Fax: " & empFax & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
            Texte = Texte & "Les dades de caràcter personal que consten a la present Factura son tractats i incorporats en un fitxer responsabilitat de " & empNom & ".  Conforme al disposat als artícles 15 y 16 de la Llei Orgànica 15/1999, de 13 de desembre, de Protecció de Dades de Caràcter Personal, l'informem que pot exercir els drets d'accès, rectificació, cancelació i oposició en: " & empAdresa & ", " & empCp & " (" & empCiutat & "), o bé, enviar un correu electrònic a: " & empEMail & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "Pensi en el medi ambient abans d'imprimir aquest e-mail" & "<Br>"
            Texte = Texte & "----------------------------" & "<Br>"
            Texte = Texte & "" & "<Br>"
        End If
        
        para = Split(EmailClient, ";")
        For P = 0 To UBound(para)
            sf_enviarMail empEMail, Trim(para(P)), "NO CONTESTAR AQUEST EMAIL. " & empNom & " ,factura ", Texte, IdFile, "c:\Factura.Pdf", iD, "c:\" & numFactura & ".xml"
        Next
        
        sf_enviarMail "email@hit.cat", empEMail, "Enviada Factura a " & EmailClient & " " & empNom & " ", Texte, IdFile, "c:\Factura.Pdf", iD, "c:\" & numFactura & ".xml"

        Set rsCom = Db.OpenResultset("select idFactura from FacturacioComentaris where idFactura='" & idFactura & "'")
        If Not rsCom.EOF Then
            sql = "update FacturacioComentaris set Enviada='" & Now() & " OK' where idFactura='" & idFactura & "'"
        Else
            sql = "insert into FacturacioComentaris (IdFactura, Comentari, Cobrat, Data, Enviada) values('" & idFactura & "', '', 'N', getdate(), '" & Now() & " OK')"
        End If
        
        ExecutaComandaSql sql
    Else
        sf_enviarMail "email@hit.cat", empEMail, "Error al Enviar Factura a " & EmailClient & " " & empNom & " ", Texte, IdFile, "c:\Factura.Pdf"
    End If
    
ERR_EMAIL:

    ExecutaComandaSql "Delete archivo Where id  = '" & IdFile & "' "
    ExecutaComandaSql "Delete archivo Where id  = '" & iD & "' "
    
    MyKill "c:\" & numFactura & ".xml"
End Sub


Sub EnviaSms(Desti, Caption)
On Error Resume Next
    FacturaSms Desti, Caption
    sf_enviarMail "info@hit.cat", Desti & "@sms.popfax.com", "EnviaFax", Caption, "", ""
    
End Sub

Function HTMLDecode(s) As String
    Dim Str As String
    
    Str = s
    
    Str = Replace(Str, "á", "&aacute;")
    Str = Replace(Str, "é", "&eacute;")
    Str = Replace(Str, "í", "&iacute;")
    Str = Replace(Str, "ó", "&oacute;")
    Str = Replace(Str, "ú", "&uacute;")
    Str = Replace(Str, "ñ", "&ntilde;")
    Str = Replace(Str, "à", "&agrave;")
    Str = Replace(Str, "è", "&egrave;")
    Str = Replace(Str, "ò", "&ograve;")
    
    HTMLDecode = Str
End Function

Function sf_enviarMail(ByVal sls_de, ByVal sls_a, ByVal sls_asunto, ByVal sls_cuerpo, ByVal sls_adjunto As String, NomFileDesti As String, Optional ByVal sls_adjunto2 As String, Optional NomFileDesti2 As String)
    Dim Rs   As rdoResultset, SLN_SENDUSING, SLS_SMTPSERVER, SLN_SMTPSERVERPORT, SLB_SMTPUSESSL, SLN_SMTPCONNTIMEOUT, SLN_SMTPAUTENTICATE, SLS_SMTPUSERNAME, SLS_SMTPPASSWORD, sls_aList() As String, e As Integer
    Dim sls_reply As String
    
    Informa "enviant email a : " & sls_a
'**** Constantes
    If IsNull(sls_a) Then Exit Function
    If InStr(sls_a, "@") = 0 Then Exit Function

    sls_a = Trim(sls_a)
    
    SLN_SENDUSING = 2
    'SLS_SMTPSERVER = "smtp.gmail.com"
    'SLS_SMTPSERVER = "email-smtp.eu-west-1.amazonaws.com"
    'SLN_SMTPSERVERPORT = "465"
    'SLB_SMTPUSESSL = True
    SLN_SMTPCONNTIMEOUT = 60
    SLN_SMTPAUTENTICATE = 1
    'SLS_SMTPUSERNAME = "email@hit.cat"
    'SLS_SMTPUSERNAME = "AKIAJ4QSTGHHR2JL5A3A"
    'SLS_SMTPPASSWORD = "emailhit"
    'SLS_SMTPPASSWORD = "And2QQwBrwdfZFbAm83akSs8MD+9CwpfQJ149aJmk7fO"
    'Const SLS_DEFAULTDE = "NoContestar@hit.cat"
    Const SLS_DEFAULTDE = "email@hit.cat"
    
'From: Email@ hit.cat
'Servidor SMTP: email-smtp.eu-west-1.amazonaws.com
'Seguridad: SSL
'Puerto SMTP: 465
'Usuario: AKIAJ4QSTGHHR2JL5A3A
'Contraseña: And2QQwBrwdfZFbAm83akSs8MD+9CwpfQJ149aJmk7fO
SLS_SMTPSERVER = "email-smtp.eu-west-1.amazonaws.com"
SLN_SMTPSERVERPORT = "465"
SLB_SMTPUSESSL = True
SLS_SMTPUSERNAME = "AKIAJT4WEFTQR7NHKSCA"
SLS_SMTPPASSWORD = "Ap2WrHnaraCPtAPcmlQjE7JBzkFDuXfDTBeTt4R3VczB"

    empresaEmail sls_de, SLS_SMTPSERVER, SLS_SMTPUSERNAME, SLS_SMTPPASSWORD, SLN_SMTPSERVERPORT
    
'**** Valores por defecto
    sls_reply = sls_de
    
    sls_de = SLS_DEFAULTDE
    'If sls_de = "" Then sls_de = SLS_DEFAULTDE
    'If sls_de = "email@hit.cat" Then sls_de = "NoContestar@hit.cat"
    
    'If SLS_SMTPUSERNAME <> sls_de Then sls_de = SLS_SMTPUSERNAME
    'If SLS_SMTPUSERNAME <> sls_de Then sls_de = SLS_DEFAULTDE
    
    If sls_asunto = "" Then
        sls_asunto = "Sin Asunto"
    End If
    If sls_cuerpo = "" Then
        sls_cuerpo = "Sin texto"
    End If
'**** validamos email a
    Dim sls_regEx
    Set sls_regEx = New RegExp
    sls_regEx.Pattern = "^[a-z][\w\.-]*[\w-]@[\da-z][\w-]*[\.]?[\w-]*[\w]\.[a-z]{2,3}$"
    sls_regEx.IgnoreCase = True
    
    If Not sls_regEx.Test(sls_reply) Then sls_reply = SLS_DEFAULTDE
    If Not sls_regEx.Test(sls_a) Then
        sf_enviarMail = False
        Exit Function
    End If
    
'**** Validamos e-mail de
    If Not sls_regEx.Test(sls_de) Then
        sf_enviarMail = False
        Exit Function
    End If
    If Len(sls_adjunto) > 0 Then
        Set Rs = Db.OpenResultset("select * from Archivo Where id  = '" & sls_adjunto & "' ", rdOpenKeyset)
        If Not Rs.EOF Then
        ' ALGU HA CAMBIAT AIXÓ I LO HA FIJADO A EXCEL.XLS
        'sls_adjunto = "c:\excel.xls" '"c:\" & rs.rdoColumns("Nombre") & "." & rs.rdoColumns("Extension")
            sls_adjunto = NomFileDesti
            sls_cuerpo = sls_cuerpo & Chr(13) & Chr(10) & "Tema : " & Rs.rdoColumns("descripcion")
            MyKill sls_adjunto
            ColumnToFile Rs.rdoColumns("Archivo"), sls_adjunto, 102400, Rs("Archivo").ColumnSize
        Else
            If Len(Dir(sls_adjunto)) = 0 Then
                sls_adjunto = ""
            Else
                
            End If
        End If
        Rs.Close
    End If
    
    
    If Len(sls_adjunto2) > 0 Then
        Set Rs = Db.OpenResultset("select * from Archivo Where id  = '" & sls_adjunto2 & "' ", rdOpenKeyset)
        If Not Rs.EOF Then
        ' ALGU HA CAMBIAT AIXÓ I LO HA FIJADO A EXCEL.XLS
        'sls_adjunto = "c:\excel.xls" '"c:\" & rs.rdoColumns("Nombre") & "." & rs.rdoColumns("Extension")
            sls_adjunto2 = NomFileDesti2
            sls_cuerpo = sls_cuerpo & Chr(13) & Chr(10) & "Tema : " & Rs.rdoColumns("descripcion")
            MyKill sls_adjunto2
            ColumnToFile Rs.rdoColumns("Archivo"), sls_adjunto2, 102400, Rs("Archivo").ColumnSize
        Else
            If Len(Dir(sls_adjunto2)) = 0 Then
                sls_adjunto2 = ""
            Else
                
            End If
        End If
        Rs.Close
    End If
    
'**** Validamos que exista el fichero a adjuntar. Si no existe, no lo adjuntamos
'   if sls_adjunto<>"" then
'       sls_adjunto=Server.MapPath(sls_adjunto)
'       if not sf_ficheroExiste(sls_adjunto) then
'           sls_adjunto=""
        'end if
'   end if

'**** Configuración Objeto eMail
    Dim ObjSendMail
    Set ObjSendMail = CreateObject("CDO.Message")
    
    ObjSendMail.Fields.Item("urn:schemas:mailheader:X-SES-CONFIGURATION-SET") = "monitor1"
    ObjSendMail.Fields.Update
    
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = SLN_SENDUSING
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SLS_SMTPSERVER
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SLN_SMTPSERVERPORT
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = SLB_SMTPUSESSL
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = SLN_SMTPCONNTIMEOUT
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = SLN_SMTPAUTENTICATE
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = SLS_SMTPUSERNAME
    ObjSendMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SLS_SMTPPASSWORD
    ObjSendMail.Configuration.Fields.Update
    ObjSendMail.From = sls_de
    ObjSendMail.ReplyTo = sls_reply
    ObjSendMail.Subject = sls_asunto
    ObjSendMail.HTMLBody = HTMLDecode(sls_cuerpo)
    If Len(sls_adjunto) > 0 Then ObjSendMail.AddAttachment sls_adjunto
    If Len(sls_adjunto2) > 0 Then ObjSendMail.AddAttachment sls_adjunto2

On Error GoTo nor:
    sls_aList = Split(sls_a, ";")
    For e = 0 To UBound(sls_aList)
        ObjSendMail.To = sls_aList(e)
        ObjSendMail.Send
    Next
       
    sf_enviarMail = True
    Set ObjSendMail = Nothing
    Exit Function
nor:

   Set ObjSendMail = Nothing
   sf_enviarMail = False
   
End Function
'************************************************************************
'* sf_ficheroExiste(fichero)
'* Devuelve verdadero si el fichero existe, y falso en caso contrario
'* Espera el camino al fichero ya mapeado
'************************************************************************
Function sf_ficheroExiste(ByVal sls_file)
    If sls_file = "" Then
        sf_ficheroExiste = False
        Exit Function
    End If
    Dim slo_file
    Set slo_file = CreateObject("Scripting.FileSystemObject")
    sf_ficheroExiste = (slo_file.FileExists(sls_file))
    Set slo_file = Nothing
End Function

Function RecurseMKDir(ByVal Path)
    Dim FS
    Set FS = CreateObject("Scripting.FileSystemObject")
    Path = Join(Split(Path, "/"), "\")
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Dim pos, n
    n = 0
    pos = InStr(1, Path, "\")
    Do While pos > 0
        On Error Resume Next
        FS.CreateFolder Left(Path, pos - 1)
        If err = 0 Then n = n + 1
        pos = InStr(pos + 1, Path, "\")
    Loop
    RecurseMKDir = n
End Function
'***********************************************************************
'**** Salva una cadena binaria usando ADODB Stream
'***********************************************************************
Function SaveBinaryDataStream(fileName, ByteArray)
    Dim BinaryStream As ADODB.Stream
    
    BinaryStream.Type = 1 'Binary
    BinaryStream.Open
    If LenB(ByteArray) > 0 Then BinaryStream.Write ByteArray
    On Error Resume Next
    BinaryStream.SaveToFile fileName, 2 'Overwrite
    If err = &HBBC Then '**** No encuentra el camino y lo creará
'        On Error GoTo 0
'        RecurseMKDir GetPath(FileName)
'        On Error Resume Next
'        BinaryStream.SaveToFile FileName, 2 'Overwrite
    End If
    Dim ErrMessage, ErrNumber
    ErrMessage = err.Description
    ErrNumber = err
    On Error GoTo 0
'    If ErrNumber <> 0 Then Err.Raise ErrNumber, "SaveBinaryData", FileName & ":" & ErrMessage
End Function


