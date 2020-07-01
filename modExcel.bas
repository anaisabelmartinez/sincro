Attribute VB_Name = "modExcel"
Option Explicit
Public MsExcel As New Excel.Application


Sub CalculaHistoricSetmana(s, emp, ano)
    Dim sql As String
    
    
    
    ExecutaComandaSql "delete from fichaje_hist where DATEPART(wk,cast(Fecha as datetime)) =  " & s

    sql = ""
    sql = sql & "Insert into fichaje_hist "
    ' 14/02/2011 JORGE: Modificado para incluir el fichaje por equipos
    'sql = sql & "select DependentesExtes.Valor,Dependentes.CODI,Dependentes.Nom,IsNull(DATEPART( weekday, Entrada),1),"
    sql = sql & "select case isnull(equip,'') when '' then dependentesExtes.valor else equip end as Valor,Dependentes.CODI,Dependentes.Nom,IsNull(DATEPART( weekday, Entrada),1),"
    sql = sql & "IsNull(day(Entrada),0),IsNull(left(Entrada,11),cast(dateadd(dd,-(datepart(wk,getdate())-" & s & ")*7,getdate()) as datetime)) as Fecha,IsNull(Sum (DateDiff(Minute, Entrada, salida)),0) "
    sql = sql & "from Dependentes with (nolock) left join "
    sql = sql & "( "
    'sql = sql & "select t1.usuari, t1.tmst as 'Entrada' , "
    sql = sql & "select te1.equip, t1.usuari, t1.tmst as 'Entrada' , "
    sql = sql & "( "
    sql = sql & "select min(tmst) "
    sql = sql & "from cdpDadesFichador t2 with (nolock) "
    sql = sql & "where t2.usuari = t1.usuari and "
    sql = sql & "t2.accio = 2 and "
    sql = sql & "t2.tmst >= t1.tmst and "
    sql = sql & "t2.tmst <= ( "
    sql = sql & "select isnull(min(tmst),'99991231 23:59:59.998') "
    sql = sql & "from cdpDadesFichador with (nolock) "
    sql = sql & "where usuari = t2.usuari and "
    sql = sql & "accio = 1 and "
    sql = sql & "tmst > t1.tmst "
    sql = sql & ") "
    sql = sql & ") as 'Salida' "
    sql = sql & "from cdpDadesFichador t1 with (nolock) "
    sql = sql & "left join cdpdadesfichadorequip te1 with (nolock) on t1.idr = te1.idr "
    sql = sql & "Where T1.Accio = 1 "
    sql = sql & "and DATEPART( wk, tmst)=  " & s & " "
    sql = sql & "and year(tmst) = '" & ano & "'"
    sql = sql & ") as horario "
    sql = sql & "on horario.usuari = Dependentes.Codi "
    sql = sql & "left join DependentesExtes with (nolock) "
    sql = sql & "on Dependentes.Codi = DependentesExtes.id "
    sql = sql & "where DependentesExtes.Nom = 'EQUIPS' "
    sql = sql & "and Dependentes.CODI in (select id from DependentesExtes with (nolock) Where "
    sql = sql & "nom = 'EMPRESA' and valor = '" & emp & "') and "
    sql = sql & "Dependentes.CODI in ( "
    ' cambio para fin de contratos
    sql = sql & "select codi "
    sql = sql & "from dependentes d3 with (nolock) "
    sql = sql & "left join dependentesextes  with (nolock) on d3.codi=dependentesextes.id and dependentesextes.nom = "
    sql = sql & "(select max(nom) from dependentesextes  with (nolock) where nom like 'DATACONTRACTEFIN%' "
    sql = sql & "and dependentesextes.id = d3.CODI) "
    sql = sql & "where case isnull(dependentesextes.valor,'') when '' then " & s & " "
    sql = sql & "else  datepart(wk,convert(smalldatetime,dependentesextes.valor,103)) "
    sql = sql & "end  >= " & s & " "
    sql = sql & "and "
    sql = sql & "case isnull(dependentesextes.valor,'') when '' then 3000 "
    sql = sql & "else  datepart(year,convert(smalldatetime,dependentesextes.valor,103)) "
    sql = sql & "end  >=  " & ano & " "
    'sql = sql & "select id From DependentesExtes with (nolock) Where "
    'sql = sql & "nom like 'DATACONTRACTEFIN%' "
    'sql = sql & "and case isnull(Valor,'') when '' then DATEPART( wk,cast('99991231 23:59:59.998' as datetime)) "
    'sql = sql & "Else "
    'sql = sql & "DATEPART( wk,cast(Valor as datetime)) "
    'sql = sql & "End "
    'sql = sql & "< " & s & "  and case isnull(Valor,'') when '' then year(cast('99991231 23:59:59.998' as datetime)) "
    'sql = sql & "else year(cast(Valor as datetime)) "
    'sql = sql & "End"
    'sql = sql & "<= year(getdate()))"
    sql = sql & ") Group "
    sql = sql & "By Dependentes.Codi, Day(Entrada), Dependentes.nom, equip, DependentesExtes.Valor, DatePart(Weekday, Entrada), Left(Entrada, 11)"
    'sql = sql & "By Dependentes.Codi, Day(Entrada), Dependentes.nom, DependentesExtes.Valor, DatePart(Weekday, Entrada), Left(Entrada, 11)"
    ExecutaComandaSql sql
    
End Sub




Function EnviarReportTasques()
    Dim Rs As ADODB.Recordset, Nom1 As String, Nom2 As String, nom As String, Descripcio As String, D As Date, An, client() As String, pp2 As String, ppp2, Di, Df, LlistaBotiguesPosibles As String
    Dim i As Double, Kk As Integer, iD As String, a As New Stream, s() As Byte, dia As Date, DiaF As Date, sql, K As Integer
    Dim Punts, clients
    Dim Rs2 As rdoResultset
    Dim Rs3 As ADODB.Recordset
  
    Dim MsExcel As Excel.Application
    Dim Libro As Excel.Workbook
                
    Set MsExcel = CreateObject("Excel.Application")
    Set Libro = MsExcel.Workbooks.Add
        
    On Error Resume Next
    db2.Close
    On Error GoTo norR
    
    Set db2 = New ADODB.Connection
    db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"

    MsExcel.DisplayAlerts = False
    MsExcel.Visible = frmSplash.Debugant
    
    Set Rs = rec("select newid() i")
    iD = Rs("i")

    While Libro.Sheets.Count > 1
      Libro.Sheets(1).Delete
    Wend
    
    ExcelReportTasques Libro
    nom = "Trucades" 'Fins " & Now
    Descripcio = "Registre Trucades : " & dia
    Libro.Sheets(1).Select
    Libro.Sheets(1).Cells.EntireColumn.AutoFit
    Libro.Sheets(1).Range("A1").Select
    
    If Not nom = "Cap" Then
        If Excel.Application.Version = 12 Then
            Libro.SaveAs "c:\" & iD & ".xls", xlExcel8
        Else
            Libro.SaveAs "c:\" & iD & ".xls"
        End If
        
        Libro.Close
  
        Set Libro = Nothing
        Set MsExcel = Nothing
  
        a.Open
        a.LoadFromFile "c:\" & iD & ".xls"
        s = a.ReadText()

On Error GoTo norR
        
        Set Rs = rec("select * from " & tablaArchivo() & " where fecha=(select max(fecha) from " & tablaArchivo() & ")", True)
        Rs.AddNew
        Rs("id").Value = iD
        Rs("nombre").Value = nom
        Rs("descripcion").Value = Descripcio
        Rs("extension").Value = "XLS"
        Rs("mime").Value = "application/vnd.ms-excel"
        Rs("propietario").Value = ""
        Rs("archivo").Value = s
        Rs("fecha").Value = Now
        Rs("tmp").Value = 0
        Rs("down").Value = 1
        Rs.Update
        Rs.Close
        a.Close
        EnviarReportTasques = iD
        
        MyKill "c:\" & iD & ".xls"
    End If
    
norR:

    On Error Resume Next
        db2.Close
    On Error GoTo 0
    
End Function



Function EnviarReportBotiga(codiBotiga As Double, Emails As String)
    Dim Estat As String, dia As Date, cos As String, Cos2 As String, article As Double, D2 As Date, DiaInventari As Date, DiaInventari1 As Date, Venuts As Double
    Dim Rs As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset, Rs4 As rdoResultset
    Dim Rs5 As rdoResultset, Rs6 As rdoResultset, Rs7 As rdoResultset, vBotiga, vAlbarans, vPing
    Dim Co As String, Es As String, sql, dd, mm, yyyy, dd2, mm2, yyyy2, nBot As Integer, nBal As Integer
    Dim vMysql, DataPing As Date, EurosPing As Double, EsMaquina As String, DataPingActual As Date, DifPing, sDia
    Dim Caixa As Double, LastQ, H
    
On Error GoTo nor:
    
    
    dia = Now
    
'    dia = DateAdd("d", -4, dia)
'    dia = DateSerial(2015, 12, 8) + TimeSerial(23, 55, 55)
    
    sDia = dia
    If IsNumeric(sDia) Then dia = DateAdd("d", -sDia, Now)
    If IsDate(sDia) Then dia = CVDate(sDia)
    dd = Day(dia)
    dd2 = Day(DateAdd("d", 1, dia))
    mm = Month(dia)
    mm2 = Month(DateAdd("d", 1, dia))
    If Len(mm) < 2 Then mm = "0" & mm
    If Len(mm2) < 2 Then mm2 = "0" & mm2
    yyyy = Year(dia)
    yyyy2 = Year(DateAdd("d", 1, dia))
    H = Minute(dia) + (Hour(dia) * 60)
    
    Co = "<Table border=2>"
    
    Set Rs2 = Db.OpenResultset("Select sum(Import) I from [" & NomTaulaVentas(DateAdd("d", -7, dia)) & "] where botiga =  " & codiBotiga & " and (60 * datepart(hh, Data) + DatePart(n, Data)) <= " & H & " And day(data) = " & Day(DateAdd("d", -7, dia)) & " ")
    LastQ = 0
    If Not Rs2.EOF Then If Not IsNull(Rs2("I")) Then LastQ = Rs2("I")
    
    Set Rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients ,isnull(sum(import),1) as I from [" & NomTaulaVentas(dia) & "] where botiga =  " & codiBotiga & " and  (60 * datepart(hh, Data) + DatePart(n, Data)) <= " & H & " And day(data) = " & Day(dia) & " ")
    Caixa = Rs("i")
    Co = Co & "<Tr><Td>Botiga</Td><Td>Ventas</Td><Td>Clientes</Td><Td>Primera</Td><Td>Ultima</Td><Td>Fa 7 d</Td><Td>Inc</Td></Tr>"
    Co = Co & "<Tr><Td>" & BotigaCodiNom(codiBotiga) & "</Td>"
    Caixa = Rs("i")
    Co = Co & "<Td>" & Caixa & "</Td>"
    Co = Co & "<Td>" & Rs("Clients") & "</Td>"
    
    Set Rs = Db.OpenResultset("Select isnull(min(data), getdate()) hi ,isnull(Max(data), getdate()) hf from [" & NomTaulaVentas(dia) & "] where botiga =  " & codiBotiga & " and  (60 * datepart(hh, Data) + DatePart(n, Data)) <= " & H & " And day(data) = " & Day(dia) & " ")
    Co = Co & "<Td>" & Format(Rs("Hi"), "hh:nn") & "</Td>"
    Co = Co & "    <Td>" & Format(Rs("Hf"), "hh:nn") & "</Td>"
    Co = Co & "    <Td>" & LastQ & "</Td>"
    
    Co = Co & "<Td>"
    If Caixa >= LastQ Then
       Co = Co & "<font color='green'>"
    Else
       Co = Co & "<font color='Red'>"
    End If
    
    If LastQ = 0 Then
        Co = Co & "="
    Else
        Co = Co & Format(Int(100 * (1 - (LastQ / Caixa))), "")
    End If
    Co = Co & "%</font></Td>"
    
    Co = Co & "</Tr>"
    
    Set Rs = Db.OpenResultset("Select plu,min(data) hi,max(data) hf,isnull(a.nom,'No Codificat') Article ,sum(import) Import ,sum(Quantitat) Quantitat from [" & NomTaulaVentas(dia) & "] v left join articles a on a.Codi = v.Plu where botiga =  " & codiBotiga & " and (60 * datepart(hh, Data) + DatePart(n, Data)) <= " & H & " And day(data) = " & Day(dia) & " group by a.nom,plu  order by  a.nom")
    Co = Co & "<Tr><Td>Producto</Td><Td>Importe</Td><Td>Unidades</Td><Td>Primera</Td><Td>Ultima</Td><Td>Fa 7 d</Td><Td>Inc</Td></Tr>"
    While Not Rs.EOF
        Set Rs2 = Db.OpenResultset("Select sum(Quantitat) Q from [" & NomTaulaVentas(DateAdd("d", -7, dia)) & "] where botiga =  " & codiBotiga & " and  (60 * datepart(hh, Data) + DatePart(n, Data)) <= " & H & " And  day(data) = " & Day(DateAdd("d", -7, dia)) & " and plu = " & Rs("Plu"))
        LastQ = 0
        If Not Rs2.EOF Then If Not IsNull(Rs2("Q")) Then LastQ = Rs2("q")
        
        Co = Co & "<Tr><Td>" & Rs("Article") & "</Td><Td>" & Rs("Import") & "</Td><Td>" & Rs("Quantitat") & "</Td><Td>" & Format(Rs("Hi"), "hh:nn") & "</Td><Td>" & Format(Rs("Hf"), "hh:nn") & "</Td>"
        Co = Co & "<Td>" & LastQ & "</Td>"
        Co = Co & "<Td>"
        If Rs("Quantitat") >= LastQ Then
           Co = Co & "<font color='green'>"
        Else
           Co = Co & "<font color='Red'>"
        End If
        
        If LastQ = 0 Then
            Co = Co & "="
        Else
            Co = Co & Format(Int(100 * (1 - (LastQ / Rs("Quantitat")))), "")
        End If
        
        Co = Co & "%</font></Td>"
        Co = Co & "</Tr>"
        Rs.MoveNext
    Wend

    Set Rs = Db.OpenResultset("Select plu,min(data) hi,max(data) hf,isnull(a.nom,'No Codificat') Article ,sum(import) Import ,sum(Quantitat) Quantitat from [" & NomTaulaVentas(DateAdd("d", -7, dia)) & "] v left join articles a on a.Codi = v.Plu where botiga =  " & codiBotiga & " and (60 * datepart(hh, Data) + DatePart(n, Data)) <= " & H & " And day(data) = " & Day(DateAdd("d", -7, dia)) & "  and not Plu in (Select distinct Plu from [" & NomTaulaVentas(dia) & "] where Botiga = " & codiBotiga & " and DAY(data) = " & Day(dia) & ") group by a.nom,plu  order by  a.nom")
    

    While Not Rs.EOF
        Co = Co & "<Tr><Td>" & Rs("Article") & "</Td><Td></Td><Td></Td><Td>" & Format(Rs("Hi"), "hh:nn") & "</Td><Td>" & Format(Rs("Hf"), "hh:nn") & "</Td>"
        Co = Co & "<Td>" & Rs("Quantitat") & "</Td>"
        Co = Co & "<Td>"
        Co = Co & "<font color='Red'>"
        Co = Co & Rs("Import")
        Co = Co & " e</font></Td>"
        Co = Co & "</Tr>"
        Rs.MoveNext
    Wend
        
    Co = Co & "</Table>"
    Co = Co & "<Table border=2>"
    Set Rs = Db.OpenResultset("select * from [" & NomTaulaHoraris(dia) & "] where botiga=" & codiBotiga & "  and DAY(data) = " & Day(dia) & " Order by data ")
    While Not Rs.EOF
        Co = Co & "<Tr><Td>" & DependentaCodiNom(Rs("Dependenta")) & "</Td>"
        Co = Co & "<Td>" & Format(Rs("data"), "hh:nn") & "</Td>"
        If Rs("Operacio") = "E" Then
            Co = Co & "<Td><font color='green' > Arriba </font></Td>"
        Else
            Co = Co & "<Td><font color='Red' > Plega </font></Td>"
        End If
        Co = Co & "</Tr>"
        Rs.MoveNext
    Wend
    
'    Set rs = Db.OpenResultset("select import,motiu from [" & NomTaulaMovi(dia) & "] where botiga=" & codiBotiga & "  and DAY(data) = " & Day(dia) & " ")
'    While Not rs.EOF
'        Co = Co & "<Tr><Td>" & rs("Import") / 100 & "</Td>"
'        Co = Co & "<Td>" & rs("Motiu") & "</Td>"
'        Co = Co & "</Tr>"
'        rs.MoveNext
'    Wend
    
    Set Rs = Db.OpenResultset("select * from [" & NomTaulaInventari(dia) & "] where botiga=" & codiBotiga & "  and DAY(data) = " & Day(dia) & " ")
    If Not Rs.EOF Then
        Co = Co & "</Table>"
        Co = Co & "<Table border=2>"
        Co = Co & "<Tr><Td>Producto</Td>"
        Co = Co & "<Td>Fecha Inventario</Td>"
        Co = Co & "<Td>Unidades Inventario</Td>"
        Co = Co & "<Td>Fecha PEnultimo Inventario</Td>"
        Co = Co & "<Td>Ventas</Td>"
        Co = Co & "<Td>Cuadre</Td>"
        Co = Co & "</Tr>"
        While Not Rs.EOF
            article = Rs("Plu")
            D2 = dia
                        
            DiaInventari = Rs("Data")
            DiaInventari1 = DiaInventari
            
            Set Rs2 = Db.OpenResultset("select * from [" & NomTaulaInventari(D2) & "] where botiga=" & codiBotiga & " and data < " & SqlDataMinute(DiaInventari) & " and Plu = " & article & " order by Data desc ")
            
            If Not Rs2.EOF Then
                DiaInventari1 = Rs2("Data")
            Else
                D2 = DateAdd("m", -1, D2)
                Set Rs2 = Db.OpenResultset("select * from [" & NomTaulaInventari(D2) & "] where botiga=" & codiBotiga & " and Plu = " & article & " order by Data desc ")
                If Not Rs2.EOF Then
                    DiaInventari1 = Rs2("Data")
                Else
                    D2 = DateAdd("m", -1, D2)
                    Set Rs2 = Db.OpenResultset("select * from [" & NomTaulaInventari(D2) & "] where botiga=" & codiBotiga & " and Plu = " & article & " order by Data desc ")
                    If Not Rs2.EOF Then DiaInventari1 = Rs2("Data")
                End If
            End If
            If Not ExisteixTaula("equivalenciaproductes") Then ExecutaComandaSql "CREATE TABLE [EquivalenciaProductes](   [ProdVenut] [decimal](18, 0) NULL,  [ProdServit] [decimal](18, 0) NULL, [UnitatsEquivalencia] [decimal](5, 2) NULL) ON [PRIMARY]"
            
            Venuts = PedidoCuantosFaltan(article, codiBotiga, DiaInventari1, DiaInventari)
            
            Co = Co & "<Tr><Td>" & ArticleCodiNom(article) & "</Td>"
            Co = Co & "<Td>" & DiaInventari & "</Td>"
            Co = Co & "<Td>" & Rs("Quantitat") & "</Td>"
            Co = Co & "<Td>" & DiaInventari1 & "</Td>"
            Co = Co & "<Td>" & Venuts & "</Td>"
            Co = Co & "<Td>"
            Co = Co & "<font "
            If (Rs("Quantitat") - Venuts) = 0 Then
                Co = Co & " color='green' "
            Else
                Co = Co & " color='Red' "
            End If
            Co = Co & " >"
            If DiaInventari1 = DiaInventari Then
                Co = Co & " - "
            Else
                Co = Co & "" & Rs("Quantitat") - Venuts
            End If
            
            Co = Co & "</font>"
            Co = Co & "</Td>"
            Co = Co & "</Tr>"
            Rs.MoveNext
        Wend
    End If
    
nor:
    Co = Co & "</Table>"
    
    
    sf_enviarMail "secrehit@hit.cat", Emails, Estat & " Resum " & Format(dia, "dddd dd/mm/yy") & " " & Caixa, Co, "", ""
    
End Function

Function EnviarReportBotigaHistoric(codiBotiga As String, Emails As String)
    Dim Estat As String, dia As Date, cos As String, Cos2 As String, article As Double, D2 As Date, DiaInventari As Date, DiaInventari1 As Date, Venuts As Double
    Dim Rs As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset, Rs4 As rdoResultset
    Dim Rs5 As rdoResultset, Rs6 As rdoResultset, Rs7 As rdoResultset, vBotiga, vAlbarans, vPing
    Dim Co As String, Es As String, sql, dd, mm, yyyy, dd2, mm2, yyyy2, nBot As Integer, nBal As Integer
    Dim vMysql, DataPing As Date, EurosPing As Double, EsMaquina As String, DataPingActual As Date, DifPing, sDia
    Dim Caixa As Double, LastQ, H
    
On Error GoTo nor:
    
'562
'719

    dia = Now
    
'    dia = DateAdd("d", -4, dia)
'    dia = DateSerial(2015, 12, 8) + TimeSerial(23, 55, 55)
    
    sDia = dia
    If IsNumeric(sDia) Then dia = DateAdd("d", -sDia, Now)
    If IsDate(sDia) Then dia = CVDate(sDia)
    dd = Day(dia)
    dd2 = Day(DateAdd("d", 1, dia))
    mm = Month(dia)
    mm2 = Month(DateAdd("d", 1, dia))
    If Len(mm) < 2 Then mm = "0" & mm
    If Len(mm2) < 2 Then mm2 = "0" & mm2
    yyyy = Year(dia)
    yyyy2 = Year(DateAdd("d", 1, dia))
    H = Minute(dia) + (Hour(dia) * 60)
    
    Co = "<Table border=2>"
    Set Rs2 = Db.OpenResultset("Select count(distinct num_tick) as C, round(sum(Import),2) I from [" & NomTaulaVentas(DateAdd("d", -7, dia)) & "] where botiga in(" & codiBotiga & " ) and day(data) = " & Day(DateAdd("d", -7, dia)) & " ")
    
    LastQ = 0
    If Not Rs2.EOF Then If Not IsNull(Rs2("I")) Then LastQ = Rs2("I")
    Dim LastC As Double
    LastC = 0
    If Not Rs2.EOF Then If Not IsNull(Rs2("c")) Then LastC = Rs2("c")
    
    Set Rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients ,round(sum(Import),2) as I from [" & NomTaulaVentas(dia) & "] where botiga in(" & codiBotiga & " ) And day(data) = " & Day(dia) & " ")
    Caixa = Rs("i")
    Co = Co & "<Tr><Td>Botiga</Td><Td>Ventas</Td><Td>Clientes</Td><Td>Primera</Td><Td>Ultima</Td><Td>Fa 1 Any</Td><Td>Inc</Td></Tr>"
    Co = Co & "<Tr><Td>" & BotigaCodiNom(codiBotiga) & "</Td>"
    Caixa = Rs("i")
    Co = Co & "<Td>" & Caixa & "</Td>"
    Co = Co & "<Td>" & Rs("Clients") & "</Td>"
    
    Set Rs = Db.OpenResultset("Select min(data) hi ,Max(data) hf from [" & NomTaulaVentas(dia) & "] where  botiga in(" & codiBotiga & " ) And day(data) = " & Day(dia) & " ")
    Co = Co & "<Td>" & Format(Rs("Hi"), "hh:nn") & "</Td>"
    Co = Co & "    <Td>" & Format(Rs("Hf"), "hh:nn") & "</Td>"
    Co = Co & "    <Td>" & LastQ & "</Td>"
    
    Co = Co & "<Td>"
    If Caixa >= LastQ Then
       Co = Co & "<font color='green'>"
    Else
       Co = Co & "<font color='Red'>"
    End If
    
    If LastQ = 0 Then
        Co = Co & "="
    Else
        Co = Co & Format(Int(100 * (1 - (LastQ / Caixa))), "")
    End If
    Co = Co & "%</font></Td>"
    
    Co = Co & "</Tr>"
    
    Co = Co & "</Table>"
    Co = Co & "<Table border=2>"
    Set Rs = Db.OpenResultset("select * from [" & NomTaulaHoraris(dia) & "] where  botiga in(" & codiBotiga & " ) and DAY(data) = " & Day(dia) & " Order by data ")
    While Not Rs.EOF
        Co = Co & "<Tr><Td>" & DependentaCodiNom(Rs("Dependenta")) & "</Td>"
        Co = Co & "<Td>" & Format(Rs("data"), "hh:nn") & "</Td>"
        If Rs("Operacio") = "E" Then
            Co = Co & "<Td><font color='green' > Arriba </font></Td>"
        Else
            Co = Co & "<Td><font color='Red' > Plega </font></Td>"
        End If
        Co = Co & "</Tr>"
        Rs.MoveNext
    Wend
    
nor:
    Co = Co & "</Table>"
    
    
    sf_enviarMail "secrehit@hit.cat", Emails, Estat & " Resum " & Format(dia, "dddd dd/mm/yy") & " " & Caixa, Co, "", ""
    
End Function

Function EnviarResultatBotiga(codiBotiga)
   Dim MsExcel As Excel.Application
   Dim Libro As Excel.Workbook
   EnviarResultatBotiga = ""
    
On Error GoTo 0
frmSplash.Debugant = 1
If Not frmSplash.Debugant Then On Error GoTo nok

    InformaMiss "Calculs Excel"
    Set MsExcel = CreateObject("Excel.Application")
    Set Libro = MsExcel.Workbooks.Add
    
    'EnviarResultatBotiga = CalculaExcelResultatBotiga(MsExcel, Libro, codiBotiga, DateAdd("m", -1, Now))
    EnviarResultatBotiga = CalculaExcelResultatBotiga(MsExcel, Libro, codiBotiga, Now)
    
    TancaExcel MsExcel, Libro
    
    Exit Function
    
nok:
  sf_enviarMail "email@hit.cat", "ana@solucionesit365.com", "Error en excel ", "", "", ""
  TancaExcel MsExcel, Libro
    
End Function





Sub ExcelDevolucionsMensualFranquicia(Libro, fecha, botiga, Usuari)
    Dim RsBot As ADODB.Recordset
    Dim Hoja As Excel.Worksheet
    Dim sql  As String, SqlBot As String
    Dim mes  As Integer, anyo As Integer, iBot As Integer
    Dim fechaIni As Date, fechaAux As Date
        
    mes = Month(fecha)
    anyo = Year(fecha)
    fechaIni = CDate("01/" & mes & "/" & anyo)
    
    SqlBot = "select * from clients "
    If botiga <> "" Then
        SqlBot = SqlBot & " where Codi= '" & botiga & "' "
    Else
        SqlBot = SqlBot & " where Codi in (select Codi from ConstantsClient where Variable='userFranquicia' and Valor='" & Usuari & "')"
    End If
    
    Set RsBot = rec(SqlBot)
  
    iBot = 0
    While Not RsBot.EOF
        If iBot > 0 Then
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            Set Hoja = Libro.Sheets(Libro.Sheets.Count)
        End If
    
        sql = "select a.NOM "
        fechaAux = fechaIni
        While Month(fechaIni) = Month(fechaAux)
            sql = sql & ", isnull(s" & Day(fechaAux) & ".[" & Format(fechaAux, "dd/mm/yyyy") & "], 0) [" & Format(fechaAux, "dd/mm/yyyy") & "] "
            fechaAux = DateAdd("d", 1, fechaAux)
        Wend
        sql = sql & "from articles a "
        
        fechaAux = fechaIni
        While Month(fechaIni) = Month(fechaAux)
            sql = sql & "Left Join "
            sql = sql & "( "
            sql = sql & "select client, codiArticle, sum(s.QuantitatTornada) [" & Format(fechaAux, "dd/mm/yyyy") & "] "
            sql = sql & "from " & DonamTaulaServit(fechaAux) & " s "
            sql = sql & "Where client = " & RsBot("Codi") & " And quantitatTornada > 0 "
            sql = sql & "group by client, codiArticle) s" & Day(fechaAux) & " on s" & Day(fechaAux) & ".codiarticle = a.codi "
    
            fechaAux = DateAdd("d", 1, fechaAux)
        Wend
        
        sql = sql & " Where 1<>1 "
        fechaAux = fechaIni
        While Month(fechaIni) = Month(fechaAux)
            sql = sql & " or [" & Format(fechaAux, "dd/mm/yyyy") & "] Is Not Null "
    
            fechaAux = DateAdd("d", 1, fechaAux)
        Wend
        
        sql = sql & " order by a.NOM "
      
        rellenaHojaSql "Devolucions " & RsBot("Nom"), sql, Libro.Sheets(Libro.Sheets.Count), 0
        
        iBot = iBot + 1
        RsBot.MoveNext
    Wend

End Sub

Sub ExcelConsumPersonalMensualFranquicia(Libro, fecha, botiga, Usuari)
    Dim sql  As String
    Dim mes, anyo As Integer
    Dim fechaIni As Date

    mes = Month(fecha)
    anyo = Year(fecha)
    fechaIni = CDate("01/" & mes & "/" & anyo)
    
    sql = "select d.nom Dependenta, v.Data, c.nom Botiga, a.NOM Article, CAST (v.Quantitat AS nvarchar(10)) Quantitat , CAST (v.Import AS nvarchar(10)) Import, substring(v.Tipus_venta, 6, 3) as Descompte "
    sql = sql & "from [" & NomTaulaVentas(fechaIni) & "] v "
    sql = sql & "left join dependentesExtes de on de.nom='CODICFINAL' and v.Otros  like '%Id:'+ de.valor + '%' "
    sql = sql & "left join dependentes d on de.id=d.CODI "
    sql = sql & "left join clients c on v.Botiga =c.Codi "
    sql = sql & "left join articles a on v.Plu=a.codi "
    sql = sql & "where v.Tipus_venta like 'Desc_%' and v.Otros like '%CliBoti_000_%' and d.NOM is not null and "
    sql = sql & "v.Botiga in (select Codi from ConstantsClient where Variable='userFranquicia' and Valor='" & Usuari & "') "
    sql = sql & "order by d.Nom, v.Data, c.nom"

    rellenaHojaSql "Consum Personal " & mes & "/" & anyo, sql, Libro.Sheets(Libro.Sheets.Count), 0

End Sub


Public Sub ExcelReportTasques(Libro)

    Dim Rs As rdoResultset
    Dim sql As String
    Dim sql2 As String
    Dim dia As Date
    
    dia = Now()
    
    sql = "select convert(varchar(10),isnull(i.TimeStamp,''),105) DATA, isnull(d.NOM, '') RESPONSABLE, isnull(c.nom, '') CLIENT, "
    sql = sql & "isnull(i.incidencia, '') INCIDENCIA, "
    sql = sql & "Case i.Estado "
    sql = sql & "when 'Pendiente' then 'PENDENT' "
    sql = sql & "when 'Curso'     then 'EN CURS' "
    sql = sql & "when 'Resuelta'  then 'FET' "
    sql = sql & "Else '' end ESTAT "
    sql = sql & "from incidencias i "
    sql = sql & "left join dependentes d on i.Tecnico = d.codi "
    sql = sql & "left join clients c on i.Cliente = c.codi "
    sql = sql & "where (i.Estado='Resuelta'  and day(FFinReparacion)= DAY(getdate()) and MONTH(FFinReparacion)=MONTH(getdate()) and YEAR(FFinReparacion)=YEAR(getdate())) "
    sql = sql & "OR i.Estado='Curso' or i.Estado ='Pendiente' "
    sql = sql & "order by TimeStamp"

    sql2 = "select convert(varchar(10),isnull(rt.TimeStamp,''),105) + ' ' + convert(varchar(8), rt.timestamp, 108) DATA, r.Nombre [QUI TRUCA], rt.Telefono TELEFON "
    sql2 = sql2 & "from registretrucades rt "
    sql2 = sql2 & "left join recursos r on rt.idrecurso = r.Id "
    sql2 = sql2 & "where Timestamp between convert(datetime,'" & Format(dia, "dd/mm/yyyy") & "',103)  and convert(datetime,'" & Format(dia, "dd/mm/yyyy") & "',103)+convert(datetime,'23:59:59',8) "
    sql2 = sql2 & "order by Timestamp"

    On Error GoTo 0
    
    rellenaHojaSql "Tasques", sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Trucades", sql2, Libro.Sheets(Libro.Sheets.Count), 0
    
End Sub

Sub ExcelHores2HoresSetmanesAnt(usu, Setmana, H1, H2, H3, ano, equip)
   Dim Rs
   Dim sql As String
sql = "select case"
sql = sql & " isnull(equip,'')"
sql = sql & " when '' then de.valor"
sql = sql & " Else Equip"
sql = sql & " end as equipo,  usuari,"
sql = sql & " IsNull(Sum(DateDiff(Minute, Entrada, salida)), 0) As HorasAcumulado"
sql = sql & " from ("
sql = sql & " select te1.equip, t1.usuari,"
sql = sql & " t1.tmst as 'Entrada' , ( select min(tmst) from cdpDadesFichador t2 with (nolock)"
sql = sql & " Where T2.usuari = T1.usuari And T2.Accio = 2 And T2.tmst >= T1.tmst"
sql = sql & " and t2.usuari = " & usu
sql = sql & " and t2.tmst <= ( select isnull(min(tmst),'99991231 23:59:59.998')"
sql = sql & " from cdpDadesFichador with (nolock) where usuari = t2.usuari"
sql = sql & " and accio = 1 and tmst > t1.tmst ) ) as 'Salida'"
sql = sql & " from cdpDadesFichador t1 with (nolock)"
sql = sql & " left join cdpdadesfichadorequip te1 with (nolock)"
sql = sql & " on t1.idr = te1.idr Where T1.Accio = 1"
sql = sql & " and DATEPART( wk, tmst)= " & Setmana - 3 & " and year(tmst) = '" & ano & "'"
sql = sql & " and t1.usuari = " & usu
sql = sql & " ) as horario"
sql = sql & " left join dependentesExtes de with (nolock)"
sql = sql & " on usuari = de.id and  de.nom = 'EQUIPS'"
sql = sql & " group by case"
sql = sql & " isnull(equip,'')"
sql = sql & " when '' then de.valor"
sql = sql & " Else Equip"
sql = sql & " end, usuari,de.valor"

    
    'ExecutaComandaSql sql
    
   Set Rs = Db.OpenResultset(sql)
   H1 = 0
   If Not Rs.EOF Then If Not IsNull(Rs(2)) Then H1 = Rs(2)
   
sql = "select case"
sql = sql & " isnull(equip,'')"
sql = sql & " when '' then de.valor"
sql = sql & " Else Equip"
sql = sql & " end as equipo,  usuari,"
sql = sql & " IsNull(Sum(DateDiff(Minute, Entrada, salida)), 0) As HorasAcumulado"
sql = sql & " from ("
sql = sql & " select te1.equip, t1.usuari,"
sql = sql & " t1.tmst as 'Entrada' , ( select min(tmst) from cdpDadesFichador t2 with (nolock)"
sql = sql & " Where T2.usuari = T1.usuari And T2.Accio = 2 And T2.tmst >= T1.tmst"
sql = sql & " and t2.usuari = " & usu
sql = sql & " and t2.tmst <= ( select isnull(min(tmst),'99991231 23:59:59.998')"
sql = sql & " from cdpDadesFichador with (nolock) where usuari = t2.usuari"
sql = sql & " and accio = 1 and tmst > t1.tmst ) ) as 'Salida'"
sql = sql & " from cdpDadesFichador t1 with (nolock)"
sql = sql & " left join cdpdadesfichadorequip te1 with (nolock)"
sql = sql & " on t1.idr = te1.idr Where T1.Accio = 1"
sql = sql & " and DATEPART( wk, tmst)= " & Setmana - 2 & " and year(tmst) = '" & ano & "'"
sql = sql & " and t1.usuari = " & usu
sql = sql & " ) as horario"
sql = sql & " left join dependentesExtes de"
sql = sql & " on usuari = de.id and  de.nom = 'EQUIPS'"
sql = sql & " group by case"
sql = sql & " isnull(equip,'')"
sql = sql & " when '' then de.valor"
sql = sql & " Else Equip"
sql = sql & " end, usuari,de.valor"
   
   Set Rs = Db.OpenResultset(sql)
   H2 = 0
   If Not Rs.EOF Then If Not IsNull(Rs(2)) Then H2 = Rs(2)
   
sql = "select case"
sql = sql & " isnull(equip,'')"
sql = sql & " when '' then de.valor"
sql = sql & " Else Equip"
sql = sql & " end as equipo,  usuari,"
sql = sql & " IsNull(Sum(DateDiff(Minute, Entrada, salida)), 0) As HorasAcumulado"
sql = sql & " from ("
sql = sql & " select te1.equip, t1.usuari,"
sql = sql & " t1.tmst as 'Entrada' , ( select min(tmst) from cdpDadesFichador t2 with (nolock)"
sql = sql & " Where T2.usuari = T1.usuari And T2.Accio = 2 And T2.tmst >= T1.tmst"
sql = sql & " and t2.usuari = " & usu
sql = sql & " and t2.tmst <= ( select isnull(min(tmst),'99991231 23:59:59.998')"
sql = sql & " from cdpDadesFichador with (nolock) where usuari = t2.usuari"
sql = sql & " and accio = 1 and tmst > t1.tmst ) ) as 'Salida'"
sql = sql & " from cdpDadesFichador t1 with (nolock)"
sql = sql & " left join cdpdadesfichadorequip te1 with (nolock)"
sql = sql & " on t1.idr = te1.idr Where T1.Accio = 1"
sql = sql & " and DATEPART( wk, tmst)= " & Setmana - 1 & " and year(tmst) = '" & ano & "'"
sql = sql & " and t1.usuari = " & usu
sql = sql & " ) as horario"
sql = sql & " left join dependentesExtes de with (nolock) "
sql = sql & " on usuari = de.id and  de.nom = 'EQUIPS'"
sql = sql & " group by case"
sql = sql & " isnull(equip,'')"
sql = sql & " when '' then de.valor"
sql = sql & " Else Equip"
sql = sql & " end, usuari,de.valor"
   
   Set Rs = Db.OpenResultset(sql)
   H3 = 0
   If Not Rs.EOF Then If Not IsNull(Rs(2)) Then H3 = Rs(2)
   
End Sub

Sub ExcelHoresCreaTaulaResumida(Di, Df)
    
'    ReDim accion(nDias)
'    ReDim Total(nDias)
'    ReDim Phoras(nDias)
'    ReDim PhorasHe(nDias)
'    ReDim PhorasFe(nDias)
'    ReDim PHb(nDias)
'    ReDim PHde(nDias)
'
'    ReDim fechasIntervalo(nDias)
'    For i = 0 To nDias - 1
'        fechasIntervalo(i) = DateAdd("d", i, Di)
'    Next




End Sub

Sub EnviaEmailVell()
    Dim objMessage, Rs
On Error GoTo nor
'BustiaEmails2008
'[Id] [nvarchar] (255) NULL CONSTRAINT [DF_BustiaEmails2008_Id] DEFAULT (newid()),
'[Subject] [nvarchar] (255) Default (''),
'[From] [nvarchar] (255) Default (''),
'[To] [nvarchar] (255) Default (''),
'TextBody [nvarchar] (255) Default (''),
'AddAttachment [nvarchar] (255) Default (''),
'Empresa [nvarchar] (255) Default (''),
'Enviado [nvarchar] (255) Default (''),
'Usuario [nvarchar] (255) Default (''),
'Fecha   [nvarchar] (255) Default ('')
    InformaMiss "Missatgeria"
     
     
    Set objMessage = CreateObject("CDO.Message")
    Set Rs = Db.OpenResultset("Select * from hit.dbo.BustiaEmails2008 Where isnull(enviado,'') = ''  ")

    While Not Rs.EOF
        InformaMiss "Enviant a " & Rs("To")
        objMessage.Subject = Rs("Subject")
        objMessage.From = Rs("From") '"cartero@hit.cat"
        objMessage.To = Rs("To") '"jordi.bosch.maso@gmail.com"
        objMessage.TextBody = Rs("TextBody") '"This is some sample message text."
'        objMessage.AddAttachment = Rs("AddAttachment") '"c:\backup.log"
        objMessage.Send
        
        ExecutaComandaSql "Update  hit.dbo.BustiaEmails2008 Set enviado = getdate() Where id = '" & Rs("Id") & "' "
        Rs.MoveNext
    Wend
nor:
End Sub

Function getDatesComanda(botiga As String, article As String) As Date()
    Dim Rs As ADODB.Recordset
    Dim hoy As Date
    Dim D As Integer
    Dim ArrComandes() As Date
    
    ReDim ArrComandes(0)
    
    hoy = Now
    
    For D = 0 To 90
        Set Rs = rec("select * from " & DonamTaulaServit(DateAdd("d", -D, hoy)) & " where client='" & botiga & "' and CodiArticle='" & article & "'")
        If Not Rs.EOF Then
            ReDim Preserve ArrComandes(UBound(ArrComandes) + 1)
            ArrComandes(UBound(ArrComandes)) = DateAdd("d", -D, hoy)
        End If
        Rs.Close
    Next
    
    getDatesComanda = ArrComandes
End Function

Sub rellenaHojaDiaDeLaSetmanaCreaResumen(Hoja, dia)
    
    Hoja.Select
    Hoja.Name = "Resumen"
    Hoja.Cells(1, 2).Value = ".  " & Format(dia, "dd/mm/yyyy")
    Hoja.Cells(1, 2).Font.Bold = True
    Hoja.Cells(1, 2).Font.Size = 12
    Hoja.Cells(2, 2).Value = Format(dia, "dddd")
    Hoja.Cells(2, 2).Font.Bold = True
    Hoja.Cells(2, 2).Font.Size = 12
    
    Hoja.Cells(4, 1).Value = "Mañana"
    Hoja.Cells(4, 1).Font.Bold = True
    Hoja.Rows("4:4").Select
    Hoja.Rows("4:4").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Rows("4:4").Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Rows("4:4").Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    Hoja.Cells(5, 2).Value = "Recaudacion"
    Hoja.Cells(6, 2).Value = "Clientes"
    Hoja.Cells(7, 2).Value = "Media ticket"
    Hoja.Cells(8, 2).Value = "Horas"
    Hoja.Cells(9, 2).Value = "Rec/Hora"
    Hoja.Cells(10, 2).Value = "Descuadre"

    Hoja.Cells(13, 1).Value = "Tarde"
    Hoja.Cells(13, 1).Font.Bold = True
    Hoja.Rows("13:13").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Rows("13:13").Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Rows("13:13").Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    Hoja.Cells(14, 2).Value = "Recaudacion"
    Hoja.Cells(15, 2).Value = "Clientes"
    Hoja.Cells(16, 2).Value = "Media tiket"
    Hoja.Cells(17, 2).Value = "Horas"
    Hoja.Cells(18, 2).Value = "Rec/Hora"
    Hoja.Cells(19, 2).Value = "Descuadre"
    
    Hoja.Rows("19:19").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(20, 1).Value = "Euros"
    Hoja.Cells(20, 1).Font.Bold = True
    Hoja.Cells(20, 2).Value = "Devolucion"
    Hoja.Cells(21, 2).Value = "Servido"
    Hoja.Cells(22, 2).Value = "Ingreso"
    Hoja.Cells(23, 2).Value = "(Err) Retorno"
    Hoja.Rows("23:23").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    
    Hoja.Cells(24, 1).Value = "Familia"
    Hoja.Cells(24, 1).Font.Bold = True

End Sub

Sub rellenaHojaDiaDeLaSetmanaCreaResumen2(Hoja, dia)
    Dim Rs As rdoResultset, Fila, FilaHoresR, FilaIng, dia2, sql
    Hoja.Select
    Hoja.Name = "Resumen"
    dia2 = dia
    Hoja.Cells(1, 2).Value = Format(dia, "dd/mm/yyyy")
    Hoja.Cells(1, 2).Font.Bold = True
    Hoja.Cells(1, 2).Font.Size = 12
    Hoja.Cells(2, 2).Value = Format(dia, "dddd")
    Hoja.Cells(2, 2).Font.Bold = True
    Hoja.Cells(2, 2).Font.Size = 12
    Fila = 4
    Hoja.Cells(Fila, 1).Value = "Mañana"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Hoja.Rows(Fila & ":" & Fila).Select
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Recaudacion"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Clientes"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Media ticket"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Tickets anulados"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Horas"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Rec/Hora"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Descuadre"
    Fila = Fila + 2

    Hoja.Cells(Fila, 1).Value = "Tarde"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Recaudacion"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Clientes"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Media ticket"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Tickets anulados"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Horas"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Rec/Hora"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Descuadre"
    Fila = Fila + 1
    
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(Fila, 1).Value = "Euros"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Devolucion"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Servido"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Ingreso"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "(Err) Retorno"
    Fila = Fila + 1
    
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(Fila, 1).Value = "Fabrica"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Servido fabrica + IVA"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Devolucion fabrica + IVA"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Neto + % sobre ventas "
    Fila = Fila + 1
    
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(Fila, 1).Value = "Familia"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Set Rs = Db.OpenResultset("Select * from families Where Pare = 'Article' Order by nom ")
    ReDim Families(0)
    ReDim FamiliesPct(0)
    ReDim FamiliesPctAcu(0)
    While Not Rs.EOF
        ReDim Preserve Families(UBound(Families) + 1)
        ReDim Preserve FamiliesPct(UBound(FamiliesPct) + 1)
        ReDim Preserve FamiliesPctAcu(UBound(FamiliesPctAcu) + 1)
        FamiliesPct(UBound(FamiliesPct)) = 0
        FamiliesPctAcu(UBound(FamiliesPctAcu)) = 0
        Families(UBound(Families)) = Rs("Nom")
        Fila = Fila + 1
        Rs.MoveNext
    Wend
    
    Fila = Fila + 2
    'Horas resumen
    'FilaHoresR = 31 + UBound(Families) + 1
    FilaHoresR = Fila
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(Fila, 1).Value = "Resumen Horas"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Horas mañana"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Horas tarde"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Horas totales"
    Fila = Fila + 1
        
    'Ingreso
    'FilaIng = FilaHoresR + 4
    FilaIng = Fila
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(Fila, 1).Value = "Ingreso"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Ingreso mañana"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Ingreso tarde"
    
    

End Sub

Sub CarregaProveedors(empresa, EmpresaNom, dia As Date, Noms() As String, Imports() As String, RecManana, RecTarde)
    Dim i, Rs, Di As String, p, cc
On Error Resume Next
    i = 0
    For i = 0 To UBound(Noms)
        If Not Noms(i) = EmpresaNom Then Imports(i) = 0
    Next
    cc = empresa & "_CampCliente"
    If empresa = 0 Then cc = "CampCliente"
        
    RecManana = 0
    RecTarde = 0
    If empresa = 3 Then
        Set Rs = Db.OpenResultset("select sum(case a.desconte when 1 then QuantitatServida*preumajor*(1-cast(c.[Desconte 1] as real) / 100 ) when 2 then quantitatservida*preumajor*(1-cast(c.[Desconte 2] as real) / 100 ) when 3 then quantitatservida*preumajor*(1-cast(c.[Desconte 3] as real) / 100 ) else quantitatservida*preumajor end) from [" & DonamNomTaulaServit(dia) & "] s join articles a on a.codi = s.codiarticle join Clients  c on c.codi = s.client  where codiarticle in (select codiarticle from articlespropietats   where variable = 'EMP_FACTURA' and valor = " & empresa & " )")
        If Not Rs.EOF Then If Not IsNull(Rs(0)) Then RecTarde = Rs(0)
        
        ' Comanda per T12
        Set Rs = Db.OpenResultset("select sum(case a.desconte when 1 then QuantitatServida*preumajor*(1-cast(c.[Desconte 1] as real) / 100 ) when 2 then quantitatservida*preumajor*(1-cast(c.[Desconte 2] as real) / 100 ) when 3 then quantitatservida*preumajor*(1-cast(c.[Desconte 3] as real) / 100 ) else quantitatservida*preumajor end) from [" & DonamNomTaulaServit(dia) & "] s join articles a on a.codi = s.codiarticle join Clients  c on c.codi = s.client where Client = 1013  ")
        For i = 0 To UBound(Noms)
            If Noms(i) = "IME MIL" Then Exit For
        Next
        If i > UBound(Noms) Then
            ReDim Preserve Noms(i)
            ReDim Preserve Imports(i)
            Noms(i) = "IME MIL"
        End If
        If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Imports(i) = Rs(0)
        
        Set Rs = Db.OpenResultset("SELECT sum(import) / 1.07  FROM [" & NomTaulaMovi(dia) & "]  where  botiga =0  and tipus_moviment = 'O' and day(data) = " & Day(dia))
        If Not Rs.EOF Then If Not IsNull(Rs(0)) Then RecManana = Round(Rs(0), 0)
    Else
        Set Rs = Db.OpenResultset("select sum(case a.desconte when 1 then QuantitatServida*preumajor*(1-cast(c.[Desconte 1] as real) / 100 ) when 2 then quantitatservida*preumajor*(1-cast(c.[Desconte 2] as real) / 100 ) when 3 then quantitatservida*preumajor*(1-cast(c.[Desconte 3] as real) / 100 ) else quantitatservida*preumajor end) from [" & DonamNomTaulaServit(dia) & "] s join articles a on a.codi = s.codiarticle join Clients  c on c.codi = s.client  where codiarticle NOT in (select codiarticle from articlespropietats   where variable = 'EMP_FACTURA' and valor = 3 )")
        If Not Rs.EOF Then If Not IsNull(Rs(0)) Then RecTarde = Rs(0)
        
        ' Comanda Fabrica
        Set Rs = Db.OpenResultset("select sum(case a.desconte when 1 then QuantitatServida*preumajor*(1-cast(c.[Desconte 1] as real) / 100 ) when 2 then quantitatservida*preumajor*(1-cast(c.[Desconte 2] as real) / 100 ) when 3 then quantitatservida*preumajor*(1-cast(c.[Desconte 3] as real) / 100 ) else quantitatservida*preumajor end) from [" & DonamNomTaulaServit(dia) & "] s join articles a on a.codi = s.codiarticle join Clients  c on c.codi = s.client where Client = 1080  ")
        For i = 0 To UBound(Noms)
            If Noms(i) = "PAN BRL S.L" Then Exit For
        Next
        If i > UBound(Noms) Then
            ReDim Preserve Noms(i)
            ReDim Preserve Imports(i)
            Noms(i) = "PAN BRL S.L"
        End If
        If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Imports(i) = Rs(0)
    End If
    
    ' Proveedors Normals
'        Set Rs = Db.OpenResultset("select empnom,sum(BaseIva1+BaseIva2+BaseIva3+BaseIva4) Imp from [ccfacturas_" & Year(Dia) & "_iva]  Where year(datafactura) = " & Year(Dia) & "  and month(datafactura) = " & Month(Dia) & " And day(datafactura) = " & Day(Dia) & " And Clientcodi = '" & Empresa & "' group by empnom  order by empnom ")
        

    If empresa = 3 Then
        Set Rs = Db.OpenResultset("select empnom,sum(BaseIva1+BaseIva2+BaseIva3+BaseIva4) Imp from [ccfacturas_" & Year(dia) & "_iva]  Where year(datafactura) = " & Year(dia) & "  and month(datafactura) = " & Month(dia) & " And day(datafactura) = " & Day(dia) & " And Clientcodi = '0' group by empnom  order by empnom ")
    Else
        Set Rs = Db.OpenResultset("select empnom,sum(BaseIva1+BaseIva2+BaseIva3+BaseIva4) Imp from [ccfacturas_" & Year(dia) & "_iva]  Where year(datafactura) = " & Year(dia) & "  and month(datafactura) = " & Month(dia) & " And day(datafactura) = " & Day(dia) & " And Clientcodi = '99' group by empnom  order by empnom ")
    End If
    
    While Not Rs.EOF
        For i = 0 To UBound(Noms)
            If Noms(i) = Rs("EmpNom") Then Exit For
        Next
        
        If i > UBound(Noms) Then
            ReDim Preserve Noms(i)
            ReDim Preserve Imports(i)
            Noms(i) = Rs("EmpNom")
        End If
        Imports(i) = Rs("Imp")
        
        Rs.MoveNext
    Wend


End Sub

Sub ExcelTarjaClient(Libro)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients
    
    rellenaHojaSql "Clients", "select c.idexterna Tarjeta, c.nom Cliente, cast(sum(punts)+0.01 as int) Puntos ,isnull(b.nom, 'Despacho') [Tienda Alta], max(a.lastData) [Último Uso] from PuntsAcumulatsMensual a join clientsfinals c on c.id = a.client left join clients b on b.codi = substring(c.id,charindex('_',c.id)+1,case when c.id like 'CliBoti%' then charindex('_',c.id,charindex('_',c.id)+1) - charindex('_',c.id) -1 else 0 end) where ISNULL(c.idexterna,'')<>'' group by c.nom,c.id,b.nom,c.idexterna order by c.idexterna", Libro.Sheets(Libro.Sheets.Count), 0
            
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            
            ExecutaComandaSql "Drop Table PuntsTmp1"
            ExecutaComandaSql "select client,sum(punts) P into PuntsTmp1 From  PuntsAcumulatsMensual Group By Client "
            ExecutaComandaSql "drop table PuntsTmp2 "
            K = 1
            ExecutaComandaSql "select " & K & " as k,'< 500                                         ' as Tram , count(*)+0.01 as Clients , Sum(P)+0.01 as Punts , 999.01 as PctClients , 999.01 as PctPunts  into PuntsTmp2 from PuntsTmp1 where p < 500 "
            For i = 500 To 3000 Step 500
                Kk = i + 500
                K = K + 1
                ExecutaComandaSql "Insert Into PuntsTmp2 Select " & K & " as k,'De " & i & " a " & Kk & " ' as Tram , count(*) as Clients , Sum(P) as Punts, 0 as PctClients , 0 as PctPunts from PuntsTmp1 where p>= " & i & " And p < " & Kk & " "
            Next
            For i = 3000 To 10000 Step 1000
                Kk = i + 1000
                K = K + 1
                ExecutaComandaSql "Insert Into PuntsTmp2 Select " & K & " as k,'De " & i & " a " & Kk & " ' as Tram , count(*) as Clients , Sum(P) as Punts , 0 as PctClients , 0 as PctPunts from PuntsTmp1 where p>= " & i & " And p < " & Kk & " "
            Next
            For i = 10000 To 30000 Step 10000
                K = K + 1
                ExecutaComandaSql "Insert Into PuntsTmp2 Select " & K & " as k,'De " & i & " a " & i + 10000 & " ' as Tram , count(*) as Clients , Sum(P) as Punts , 0 as PctClients , 0 as PctPunts from PuntsTmp1 where p>= " & i & " And p < " & i + 10000 & " "
            Next
            K = K + 1
            ExecutaComandaSql "Insert Into PuntsTmp2 Select " & K & " as k,'De " & i & " a 99999999 ' as Tram , count(*) as Clients , Sum(P) as Punts , 0 as PctClients , 0 as PctPunts from PuntsTmp1 where p>= " & i & "  "
            
            Set Rs = Db.OpenResultset("select sum(Punts) as punts , sum(Clients) as clients from PuntsTmp2 ")
            
            Punts = 1
            clients = 1
            If Not Rs.EOF Then If Not IsNull(Rs("Punts")) Then Punts = Rs("Punts")
            If Not Rs.EOF Then If Not IsNull(Rs("Clients")) Then clients = Rs("Clients")
            
            ExecutaComandaSql "update PuntsTmp2 set pctclients = round((clients/" & clients & ") *100 ,2) "
            ExecutaComandaSql "update PuntsTmp2 set pctPunts = round((Punts/" & Punts & ") *100 ,2) "
            
            
            rellenaHojaSql "Trams", "select * from PuntsTmp2 order by k  ", Libro.Sheets(Libro.Sheets.Count), 0
            
            
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            ExecutaComandaSql "Drop Table PuntsTmp1"
            ExecutaComandaSql "select isnull(B.nom,'Desconeguda') Botiga,count(distinct client)+0.01 as ClientsCreats, sum(punts)+0.01 PuntsAcumulats, 999.01 as PctClients , 999.01 as PctPunts  into PuntsTmp1 from  PuntsAcumulatsMensual a left join clients b on b.codi = substring(client,charindex('_',client)+1,charindex('_',client,charindex('_',client)+1) - charindex('_',client) -1)  group by B.nom "
            Set Rs = Db.OpenResultset("select sum(PuntsAcumulats) as punts , sum(ClientsCreats) as clients from PuntsTmp1 ")
            
            Punts = 1
            clients = 1
            If Not Rs.EOF Then If Not IsNull(Rs("Punts")) Then Punts = Rs("Punts")
            If Not Rs.EOF Then If Not IsNull(Rs("Clients")) Then clients = Rs("Clients")
            
            ExecutaComandaSql "update PuntsTmp1 set pctclients = round((ClientsCreats/" & clients & ") *100 ,2) "
            ExecutaComandaSql "update PuntsTmp1 set pctPunts = round((PuntsAcumulats/" & Punts & ") *100 ,2) "
            
            rellenaHojaSql "Botigues", "select * from PuntsTmp1 order by botiga ", Libro.Sheets(Libro.Sheets.Count), 0
            
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            ExecutaComandaSql "Drop Table PuntsTmp1"
            ExecutaComandaSql "select d.nom,count(distinct num_tick)+0.01 ClientsAmbTarjeta , 99999999.01 as ClientsSensaTarjeta, 999.01 as Pct  into PuntsTmp1 from [" & NomTaulaVentas(DateAdd("m", -1, Now)) & "] s join dependentes d on d.codi = s.dependenta  where otros like '%cliboti%' group by d.nom "
            ExecutaComandaSql "Update  PuntsTmp1 set ClientsSensaTarjeta  = 0,Pct = 0  "
            ExecutaComandaSql "Insert Into PuntsTmp1 select d.nom,0 ClientsAmbTarjeta ,count(distinct num_tick) as ClientsSensaTarjeta, 0 as Pct from [" & NomTaulaVentas(DateAdd("m", -1, Now)) & "] s join dependentes d on d.codi = s.dependenta  where otros not like '%cliboti%' group by d.nom "
            ExecutaComandaSql "drop table PuntsTmp2 "
            ExecutaComandaSql "Select nom,sum(ClientsAmbTarjeta) ClientsAmbTarjeta ,sum(clientssensatarjeta) clientssensatarjeta ,999.001 as Pct into PuntsTmp2 from PuntsTmp1 group by nom "

            Set Rs = Db.OpenResultset("select sum(ClientsAmbTarjeta) + sum(ClientsSensaTarjeta) as cli  from PuntsTmp2")
            
            Punts = 1
            If Not Rs.EOF Then If Not IsNull(Rs("cli")) Then Punts = Rs("cli")
            ExecutaComandaSql "update PuntsTmp2 set Pct = round((ClientsAmbTarjeta/(ClientsSensaTarjeta + ClientsAmbTarjeta)) *100 ,2) "
            
            rellenaHojaSql "DepMesAnt", "select * from PuntsTmp2 order by nom  ", Libro.Sheets(Libro.Sheets.Count), 0
            
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            ExecutaComandaSql "Drop Table PuntsTmp1"
            ExecutaComandaSql "select d.nom,count(distinct num_tick)+0.01 ClientsAmbTarjeta , 99999999.01 as ClientsSensaTarjeta, 999.01 as Pct  into PuntsTmp1 from [" & NomTaulaVentas(DateAdd("m", -1, Now)) & "] s join clients d on d.codi = s.botiga  where otros like '%cliboti%' group by d.nom "
            ExecutaComandaSql "Update  PuntsTmp1 set ClientsSensaTarjeta  = 0,Pct = 0  "
            ExecutaComandaSql "Insert Into PuntsTmp1 select d.nom,0 ClientsAmbTarjeta ,count(distinct num_tick) as ClientsSensaTarjeta, 0 as Pct from [" & NomTaulaVentas(DateAdd("m", -1, Now)) & "] s join clients d on d.codi = s.botiga  where otros not like '%cliboti%' group by d.nom "
            ExecutaComandaSql "drop table PuntsTmp2 "
            ExecutaComandaSql "Select nom,sum(ClientsAmbTarjeta) ClientsAmbTarjeta ,sum(clientssensatarjeta) clientssensatarjeta ,999.001 as Pct into PuntsTmp2 from PuntsTmp1 group by nom "

            Set Rs = Db.OpenResultset("select sum(ClientsAmbTarjeta) + sum(ClientsSensaTarjeta) as cli  from PuntsTmp2")
            
            Punts = 1
            If Not Rs.EOF Then If Not IsNull(Rs("cli")) Then Punts = Rs("cli")
            ExecutaComandaSql "update PuntsTmp2 set Pct = round((ClientsAmbTarjeta/(ClientsSensaTarjeta + ClientsAmbTarjeta)) *100 ,2) "
            
            rellenaHojaSql "BotMesAnt.", "select * from PuntsTmp2 order by nom  ", Libro.Sheets(Libro.Sheets.Count), 0
            
            

End Sub

Sub ExcelPalets(Libro)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients
    
    rellenaHojaSql "Resumen", "select a.codi , a.Nom , Count(*) Unidades from palets p join Articles a on a.codi = p.plu and p.estat = 'Etiquetado' group by a.codi , a.Nom order by a.codi , a.Nom ", Libro.Sheets(Libro.Sheets.Count), 0
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Detalle", "select a.codi , a.Nom , CASE Posicion1 WHEN '0' THEN 'Pasillo' ELSE Posicion1 END Estanteria ,Datai from palets p join Articles a on a.codi = p.plu and p.estat = 'Etiquetado' order by a.nom,posicion1", Libro.Sheets(Libro.Sheets.Count), 0
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Historic", "Select a.codi , a.Nom ,p.codi, Datai FechaEntrada, Dataf FechaSalida,estat,Posicion1 Estanteria from palets p join Articles a on a.codi = p.plu and not Estat  = 'Etiquetado' Where DateDiff(m, datai, GetDate()) < 2 order by datai desc ", Libro.Sheets(Libro.Sheets.Count), 0
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Tot", "Select a.nom,p.*  from palets p join Articles a on a.codi = p.plu order by p.Codi Desc ", Libro.Sheets(Libro.Sheets.Count), 0

End Sub


Sub ExcelCaixes(Libro)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients
    
    rellenaHojaSql "Resumen", "select a.codi, a.nom, year(ini),datepart(ww,ini) Setmana ,sum(case facturada when 1 then 1 else 0 end) Facturades ,sum(case facturada when 0 then 1 else 0 end) NoFacturades ,count(*) Produides from cajas c join articles a on a.codi = c.plu group by a.codi, a.nom, year(ini),datepart(ww,ini) order by a.codi, a.nom, year(ini),datepart(ww,ini) ", Libro.Sheets(Libro.Sheets.Count), 0
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    
    rellenaHojaSql "Stock", "select a.codi,a.nom, year(ini) Año ,datepart(ww,ini) Setmana ,count(*) Stock from cajas c join articles a on a.codi = c.plu where fin IS NULL group by a.codi, a.nom, year(ini),datepart(ww,ini) order by year(ini),datepart(ww,ini),a.nom", Libro.Sheets(Libro.Sheets.Count), 0
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    
    rellenaHojaSql "Existencias", "select a.codi, a.nom,count(*) Stock from cajas c join articles a on a.codi = c.plu where fin IS NULL group by a.codi,a.nom order by a.nom", Libro.Sheets(Libro.Sheets.Count), 0
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
End Sub



Sub ExcelTrucades(Libro, Filtre)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients, sql
    
    sql = "select convert(varchar(10),isnull(i.TimeStamp,''),105) Data,convert(varchar(10),isnull(i.TimeStamp,''),108) Hora ,isnull(c.nom,'')Client ,isnull(Nombre,'') Recurs,isnull(i.estado,'') Estado,isnull(i.Incidencia,'') Detall ,isnull(i.contacto,'') Contacto ,isnull(i.Prioridad,0) Prioritat,isnull(i.Observaciones,'') Obs,isnull(d2.nom,'') as Usuario  ,isnull(d1.nom,'') as tecnico ,isnull(cast(ffinreparacion  as nvarchar),'No') FFin "
    sql = sql & "from incidencias i "
    sql = sql & "left join recursos r1 on r1.id = i.recurso "
    sql = sql & "left join recursosextes r2 on r2.id = i.recurso and r2.Variable = 'CLIENTE' "
    sql = sql & "left join Clients C on c.codi = r2.valor "
    sql = sql & "left join recursosextes r3 on r3.id = i.recurso and r3.Variable = 'DESCRIPCION' "
    sql = sql & "left join Dependentes D1 on d1.codi = i.Tecnico "
    sql = sql & "left join Dependentes D2 on d2.codi = i.Usuario "
    Filtre = ""
    If Not Filtre = "" Then
        If Filtre = "1" Then
            sql = sql & "where i.estado not like '%resuelta%' "
        ElseIf Filtre = "2" Then
            sql = sql & "where i.estado like '%resuelta%' "
        ElseIf Filtre <> "0" Then
            sql = sql & "where c.nom like '%" & Filtre & "%' "
            sql = sql & "or nombre like '%" & Filtre & "%' "
            sql = sql & "or r3.valor like '%" & Filtre & "%' "
        End If
    End If
    sql = sql & "order by timestamp desc "

    On Error GoTo 0
    
    rellenaHojaSql "Totes", sql, Libro.Sheets(Libro.Sheets.Count), 0

End Sub



Sub ExcelDevolucions(Libro, Di, Df)
    Dim K, i, Kk, Rs As rdoResultset, sql
    Dim rsvariables As Recordset
    
    Dim pivotar As String
    Dim pivotar2 As String
    Dim D As Date, diff
    Di = Replace(Di, "[", "")
    Di = Replace(Di, "]", "")
    Df = Replace(Df, "[", "")
    Df = Replace(Df, "]", "")
    Di = FormatDateTime(Di, 2)
    Df = FormatDateTime(Df, 2)
    diff = DateDiff("d", Di, Df)
    
    If Abs(diff) > 100 Then Exit Sub
    If diff < 0 Then Exit Sub
    
    sql = " select c.nom Client,f2.pare Familia,a.nom Article, CAST (sum(s.quantitatDemanada) AS FLOAT)Demanat,"
    sql = sql & "CAST(sum(s.quantitatServida) AS FLOAT) Servit,CAST(sum(s.quantitatTornada) AS FLOAT) Tornat,"
    sql = sql & "Percentatge=ROUND(((Sum (s.quantitatTornada)/CASE sum(s.quantitatServida) WHEN 0 THEN 10000 "
    sql = sql & "ELSE sum(s.quantitatServida) END)*100),0) from ( "
    For i = 0 To diff
        D = DateAdd("d", i, Di)
        sql = sql & "select '" & D & "' as fecha, *  from [" & DonamNomTaulaServit(D) & "] "
        If i < diff Then sql = sql & " union "
    Next
    sql = sql & ") s  left join clients c on c.codi = s.client left join Articles a on a.codi = s.codiarticle "
    sql = sql & "left join families f on a.familia = f.nom left join families f2 on f.pare=f2.nom "
    'sql = sql & "group by c.nom,f2.pare,a.familia,a.nom order by c.nom,f2.pare,a.nom "
    sql = sql & "group by c.nom,f2.pare,a.familia,a.nom "
    sql = sql & "having sum(s.quantitatDemanada)>'0' or sum(s.quantitatServida)>'0' or sum(s.quantitatTornada)>'0' "
    sql = sql & "order by c.nom,f2.pare,a.nom "
    
    rellenaHojaSql "Devol. " & Format(Di, "dd mm yy") & " - " & Format(Df, "dd mm yy"), sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    sql = " select c.nom Client,f2.pare Pare,CAST (sum(s.quantitatDemanada) AS FLOAT)Demanat,"
    sql = sql & "CAST(sum(s.quantitatServida) AS FLOAT) Servit,CAST(sum(s.quantitatTornada) AS FLOAT) Tornat,"
    sql = sql & "Percentatge=ROUND(((Sum (s.quantitatTornada)/CASE sum(s.quantitatServida) WHEN 0 THEN 10000 "
    sql = sql & "ELSE sum(s.quantitatServida) END)*100),0) from ( "
    For i = 0 To diff
        D = DateAdd("d", i, Di)
        sql = sql & "select '" & D & "' as fecha, *  from [" & DonamNomTaulaServit(D) & "] "
        If i < diff Then sql = sql & " union "
    Next
    sql = sql & ") s  left join clients c on c.codi = s.client left join Articles a on a.codi = s.codiarticle "
    sql = sql & " left join families f on a.familia = f.nom left join families f2 on f.pare=f2.nom "
    sql = sql & "group by c.nom,f2.pare "
    sql = sql & "having sum(s.quantitatDemanada)>'0' or sum(s.quantitatServida)>'0' or sum(s.quantitatTornada)>'0' "
    sql = sql & "order by c.nom,f2.pare "
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Devol. " & Format(Di, "dd mm yy") & " - " & Format(Df, "dd mm yy"), sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    
    sql = " SELECT COALESCE("
    sql = sql & " '[Semana ' + cast(fecha as varchar) + ']',"
    sql = sql & " '[Semana ' + cast(fecha as varchar )+ ']'"
    sql = sql & " )"
    sql = sql & " from (select distinct(fecha)"
    sql = sql & " From ("
    For i = 0 To diff
        D = DateAdd("d", i, Di)
        sql = sql & "select datepart(wk,convert(smalldatetime,'" & D & "',103)) as fecha  from [" & DonamNomTaulaServit(D) & "] "
        If i < diff Then sql = sql & " union "
    Next
    sql = sql & " ) s  ) s2"
    pivotar = ""
    Set rsvariables = rec(sql)
    While Not rsvariables.EOF
        pivotar = pivotar & rsvariables(0) & ","
        rsvariables.MoveNext
    Wend
    pivotar = Left(pivotar, Len(pivotar) - 1)
    
    sql = "select * from"
    sql = sql & " ( select f2.pare Familia,a.nom Article,"
    'sql = sql & " CAST (sum(s.quantitatDemanada) AS FLOAT)Demanat,"
    'sql = sql & " CAST(sum(s.quantitatServida) AS FLOAT) Servit,"
    'sql = sql & " CAST(sum(s.quantitatTornada) AS FLOAT) Tornat,"
    sql = sql & " Percentatge=ROUND(((Sum (s.quantitatTornada)/CASE sum(s.quantitatServida)"
    sql = sql & " WHEN 0 THEN 10000 ELSE sum(s.quantitatServida) END)*100),2),"
    sql = sql & " 'Semana ' + cast(datepart(wk,convert(smalldatetime,s.fecha,103)) as varchar) Semana"
    sql = sql & " from ( "
    For i = 0 To diff
        D = DateAdd("d", i, Di)
        sql = sql & "select '" & D & "' as fecha,*  from [" & DonamNomTaulaServit(D) & "] "
        If i < diff Then sql = sql & " union "
    Next
    sql = sql & " ) s"
    sql = sql & " left join Articles a on a.codi = s.codiarticle"
    sql = sql & " left join families f on a.familia = f.nom"
    sql = sql & " left join families f2 on f.pare=f2.nom"
    sql = sql & " group by f2.pare,a.familia,a.nom,"
    sql = sql & " datepart(wk,convert(smalldatetime,s.fecha,103))"
    sql = sql & " having sum(s.quantitatDemanada)>'0' or sum(s.quantitatServida)>'0'"
    sql = sql & " ) t"
    sql = sql & " PIVOT(SUM(Percentatge) FOR [Semana] in (" & pivotar & ")) as pivote"
    sql = sql & " order by 1,2"
    
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Devol. " & Format(Di, "dd mm yy") & " - " & Format(Df, "dd mm yy"), sql, Libro.Sheets(Libro.Sheets.Count), 0
    

    
End Sub






Sub ExcelContactes(Libro)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients, sql
    
    sql = "Select nom,nif "
    sql = sql & ",isnull(c6.valor,'') as PersonaContacte "
    sql = sql & ",isnull(c5.valor,'') as Telefon "
    sql = sql & ",isnull(c2.valor,'') as Fax "
    sql = sql & ",isnull(c3.valor,'') as Email "
    sql = sql & ",adresa,ciutat,cp "
    sql = sql & ",isnull(c1.valor,'') as Grup "
    sql = sql & ",isnull(c4.valor,'') as Idioma "
    sql = sql & "from clients c "
    sql = sql & "left join constantsclient c1 on c.codi=c1.codi and c1.variable = 'Grup_client' "
    sql = sql & "left join constantsclient c2 on c.codi=c2.codi and c2.variable = 'Fax' "
    sql = sql & "left join constantsclient c3 on c.codi=c3.codi and c3.variable = 'eMail' "
    sql = sql & "left join constantsclient c4 on c.codi=c4.codi and c4.variable = 'IDIOMA' "
    sql = sql & "left join constantsclient c5 on c.codi=c5.codi and c5.variable = 'Tel' "
    sql = sql & "left join constantsclient c6 on c.codi=c6.codi and c6.variable = 'P_Contacte' "
    sql = sql & "Order by nom "
    rellenaHojaSql "Clients", sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    sql = "select d.Codi,d.Nom,d.Telefon "
    sql = sql & ",D10.valor   as [Tel Movil] "
    sql = sql & ",d.[Adreça] "
    sql = sql & ",D1.valor   as [Adreça 2] "
    sql = sql & ",D2.valor   as [Codi Postal] "
    sql = sql & ",D4.valor   as [Email] "
    sql = sql & ",D5.valor   as [Empresa] "
    sql = sql & ",D3.valor   as [Dni]"
    sql = sql & ",D6.valor   as [Idioma]"
    sql = sql & ",D7.valor   as [Provincia]"
    sql = sql & ",D8.valor   as [Tipus]"
    sql = sql & ",D9.valor   as [Tipus2] "
    sql = sql & "From dependentes d "
    sql = sql & "left join DependentesExtes D1 on d.codi = d1.Id and d1.Nom = 'ADRESA' "
    sql = sql & "left join DependentesExtes D2 on d.codi = d2.Id and d1.Nom = 'CODIGO POSTAL' "
    sql = sql & "left join DependentesExtes D3 on d.codi = d3.Id and d1.Nom = 'DNI' "
    sql = sql & "left join DependentesExtes D4 on d.codi = d4.Id and d1.Nom = 'EMAIL' "
    sql = sql & "left join DependentesExtes D5 on d.codi = d5.Id and d1.Nom = 'EMPRESA' "
    sql = sql & "left join DependentesExtes D6 on d.codi = d6.Id and d1.Nom = 'IDIOMA' "
    sql = sql & "left join DependentesExtes D7 on d.codi = d7.Id and d1.Nom = 'PROVINCIA' "
    sql = sql & "left join DependentesExtes D8 on d.codi = d8.Id and d1.Nom = 'TIPUS' "
    sql = sql & "left join DependentesExtes D9 on d.codi = d9.Id and d1.Nom = 'TIPUSTREBALLADOR' "
    sql = sql & "left join DependentesExtes D10 on d.codi = d10.Id and d1.Nom = 'TLF_MOBIL' "
    sql = sql & " Order By D.Nom "
    rellenaHojaSql "Treballadors", sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    sql = "select  nom as Nom ,Telefon ,Adreca as [Adreça],Emili as Email,Nif ,IdExterna as TarjaClient  from ClientsFinals   order by nom "
    rellenaHojaSql "ClientsFinals", sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    
    sql = "select r.Nombre,Tipo "
    sql = sql & ",r1.valor   as Descripcio "
    sql = sql & ",r2.valor   as Telefon "
    sql = sql & ",r3.valor   as Direccio "
    sql = sql & ",r4.valor   as Contacte "
    sql = sql & "from recursos r "
    sql = sql & "left join recursosextes R1 on r.id = r1.Id and r1.variable = 'DESCRIPCION' "
    sql = sql & "left join recursosextes R2 on r.id = r2.Id and r2.variable = 'TELEFONOS' "
    sql = sql & "left join recursosextes R3 on r.id = r3.Id and r3.variable = 'DIRECCION' "
    sql = sql & "left join recursosextes R4 on r.id = r4.Id and r4.variable = 'CONTACTOS' "
    sql = sql & " Order By r.nombre "

    rellenaHojaSql "Recursos", sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Proveedores", "select nombre,nombrecorto,descripcion,tlf1,tlf2,fax,email from ccproveedores order by nombre", Libro.Sheets(Libro.Sheets.Count), 0
    

End Sub

Sub ExcelAlbaransComercial(Libro, data As String, P3 As String, P4 As String)
    Dim K, i, Kk, Rs As rdoResultset, Data2, sql, tablaServits As String, idclient, codAge, Data1, Data3, Data4
    If P3 = "MIOS" Then 'Agente
        sql = " select valor from dependentesExtes where nom='sincContactosPropis' and id='" & P4 & "'"
        Set Rs = Db.OpenResultset(sql)
        If Not Rs.EOF Then codAge = Rs("valor")
        sql = "select d.cnomcli Cliente,a.nnumalb NumAlb,a.dfecAlb FechaAlb,c.cref Ref,c.cdetalle Detalle,"
        sql = sql & "CAST (c.npreunit AS FLOAT)PrecioUnidad,CAST (c.ndto AS FLOAT) Descuento,CAST (c.niva AS FLOAT) Iva,CAST (c.ncanent AS FLOAT) Cant,"
        sql = sql & "b.ccoment Comentario,CAST (((c.npreunit*c.ncanent)-((c.npreunit*c.ncanent)*(c.ndto/100)))*(1+(c.niva/100)) AS FLOAT) Total "
        sql = sql & "from sp_albclit a left join sp_albclic b on a.nnumalb=b.nnumalb left join sp_albclil c on a.nnumalb=c.nnumalb "
        sql = sql & "left join sp_clientes d on a.ccodcli=d.ccodcli "
        sql = sql & "where a.dfecalb>=dateAdd(m,-1,getDate())  and a.dfecalb<getDate() and c.ccodage='" & codAge & "' order by a.dfecalb desc "
        rellenaHojaSql "Ultimos albaranes agente", sql, Libro.Sheets(Libro.Sheets.Count), 0
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        'Ultims 3 mesos
        data = "15-03-10"
        Data1 = DateAdd("m", -3, data)
        Data1 = "01/" & Month(Data1) & "/" & Year(Data1)
        Data2 = DateAdd("m", -2, data)
        Data2 = "01/" & Month(Data2) & "/" & Year(Data2)
        Data3 = DateAdd("m", -1, data)
        Data3 = "01/" & Month(Data3) & "/" & Year(Data3)
        Data4 = "01/" & Month(Now) & "/" & Year(Now)
        sql = "select c.ccodage Agente,sum(((c.npreunit*c.ncanent)-((c.npreunit*c.ncanent)*(c.ndto/100)))*(1+(c.niva/100))) Total from sp_albclit a left join sp_albclic b on a.nnumalb=b.nnumalb left join sp_albclil c on a.nnumalb=c.nnumalb "
        sql = sql & "where c.ccodage='" & codAge & "' and a.dfecalb>='" & Data1 & "'  and a.dfecalb<'" & Data2 & "' group by c.ccodage"
        rellenaHojaSql "Total Albaranes 3 meses atras", sql, Libro.Sheets(Libro.Sheets.Count), 0
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        sql = "select c.ccodage Agente,sum(((c.npreunit*c.ncanent)-((c.npreunit*c.ncanent)*(c.ndto/100)))*(1+(c.niva/100))) Total from sp_albclit a left join sp_albclic b on a.nnumalb=b.nnumalb left join sp_albclil c on a.nnumalb=c.nnumalb "
        sql = sql & "where c.ccodage='" & codAge & "' and a.dfecalb>='" & Data2 & "'  and a.dfecalb<'" & Data3 & "' group by c.ccodage"
        rellenaHojaSql "Total Albaranes 2 mes atras", sql, Libro.Sheets(Libro.Sheets.Count), 0
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        sql = "select c.ccodage Agente,sum(((c.npreunit*c.ncanent)-((c.npreunit*c.ncanent)*(c.ndto/100)))*(1+(c.niva/100))) Total from sp_albclit a left join sp_albclic b on a.nnumalb=b.nnumalb left join sp_albclil c on a.nnumalb=c.nnumalb "
        sql = sql & "where c.ccodage='" & codAge & "' and a.dfecalb>='" & Data3 & "'  and a.dfecalb<'" & Data3 & "' group by c.ccodage"
        rellenaHojaSql "Total Albaranes 1 mes atras", sql, Libro.Sheets(Libro.Sheets.Count), 0
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        sql = "select c.ccodage Agente,sum(((c.npreunit*c.ncanent)-((c.npreunit*c.ncanent)*(c.ndto/100)))*(1+(c.niva/100))) Total from sp_albclit a left join sp_albclic b on a.nnumalb=b.nnumalb left join sp_albclil c on a.nnumalb=c.nnumalb "
        sql = sql & "where c.ccodage='" & codAge & "' and a.dfecalb>='" & Data4 & "'  and a.dfecalb<getDate() group by c.ccodage"
        rellenaHojaSql "Total Albaranes mes actual", sql, Libro.Sheets(Libro.Sheets.Count), 0
    Else    'Per client
        sql = "select ccodcli codi, cnomcli nom from sp_clientes where  cnomcli like '%" & P3 & "%' "
        sql = sql & " or ccodcli like '%" & P3 & "%' group by ccodcli,cnomcli order by cnomcli "
        Set Rs = Db.OpenResultset(sql)
        If Not Rs.EOF Then
            idclient = Rs("codi")
            sql = " select d.cnomcli Cliente,a.nnumalb NumAlb,a.dfecAlb FechaAlb,c.cref Ref,c.cdetalle Detalle,"
            sql = sql & "CAST (c.npreunit AS FLOAT)PrecioUnidad,CAST (c.ndto AS FLOAT) Descuento,CAST (c.niva AS FLOAT) Iva,CAST (c.ncanent AS FLOAT) Cant,"
            sql = sql & "b.ccoment Comentario,CAST (((c.npreunit*c.ncanent)-((c.npreunit*c.ncanent)*(c.ndto/100)))*(1+(c.niva/100)) AS FLOAT) Total "
            sql = sql & "from sp_albclit a left join sp_albclic b on a.nnumalb=b.nnumalb left join sp_albclil c on a.nnumalb=c.nnumalb "
            sql = sql & "left join sp_clientes d on a.ccodcli=d.ccodcli "
            sql = sql & "where a.ccodcli='" & idclient & "' and a.dfecalb>=dateAdd(m,-1,getDate()) and a.dfecalb<getDate() order by a.dfecalb desc "
            rellenaHojaSql "Ultims albaranes " & Rs("nom"), sql, Libro.Sheets(Libro.Sheets.Count), 0
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            'Ultims 3 mesos
            Data1 = DateAdd("m", -3, data)
            Data1 = "01/" & Month(Data1) & "/" & Year(Data1)
            Data2 = DateAdd("m", -2, data)
            Data2 = "01/" & Month(Data2) & "/" & Year(Data2)
            Data3 = DateAdd("m", -1, data)
            Data3 = "01/" & Month(Data3) & "/" & Year(Data3)
            sql = "select a.ccodcli Agente,sum(((c.npreunit*c.ncanent)-((c.npreunit*c.ncanent)*(c.ndto/100)))*(1+(c.niva/100))) Total from sp_albclit a left join sp_albclic b on a.nnumalb=b.nnumalb left join sp_albclil c on a.nnumalb=c.nnumalb "
            sql = sql & "where a.ccodcli='" & idclient & "' and a.dfecalb>='" & Data1 & "'  and a.dfecalb<'" & Data2 & "' group by a.ccodcli"
            rellenaHojaSql "Total Albaranes 2 meses atras", sql, Libro.Sheets(Libro.Sheets.Count), 0
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            sql = "select a.ccodcli Agente,sum(((c.npreunit*c.ncanent)-((c.npreunit*c.ncanent)*(c.ndto/100)))*(1+(c.niva/100))) Total from sp_albclit a left join sp_albclic b on a.nnumalb=b.nnumalb left join sp_albclil c on a.nnumalb=c.nnumalb "
            sql = sql & "where a.ccodcli='" & idclient & "' and a.dfecalb>='" & Data2 & "'  and a.dfecalb<'" & Data3 & "' group by a.ccodcli"
            rellenaHojaSql "Total Albaranes 1 mes atras", sql, Libro.Sheets(Libro.Sheets.Count), 0
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            sql = "select a.ccodcli Agente,sum(((c.npreunit*c.ncanent)-((c.npreunit*c.ncanent)*(c.ndto/100)))*(1+(c.niva/100))) Total from sp_albclit a left join sp_albclic b on a.nnumalb=b.nnumalb left join sp_albclil c on a.nnumalb=c.nnumalb "
            sql = sql & "where a.ccodcli='" & idclient & "' and a.dfecalb>='" & Data3 & "'  and a.dfecalb<getDate() group by a.ccodcli"
            rellenaHojaSql "Total Albaranes mes actual", sql, Libro.Sheets(Libro.Sheets.Count), 0
        End If
    End If
End Sub

Sub ExcelNota(Libro, data As String, P3 As String)
    Dim K, i, Kk, Rs As rdoResultset, Data2, sql, tablaServits As String
    data = Left(data, Len(data) - 1)
    data = Right(data, Len(data) - 1)
    data = CDate(data)
    data = DateAdd("d", 1, data) 'Sumem 1 dia
    tablaServits = "[Servit-" & Right(Year(data), 2) & "-" & Right("00" & Month(data), 2) & "-" & Right("00" & Day(data), 2) & "]"
    sql = "Select Grup_client,nom Client,CASE WHEN sum(Q)>0 THEN 'Si' ELSE 'No' END AS Comanda,Telefon Anotacio,Nota,max(Param1) Qui from ("
    sql = sql & "select isnull(cg.valor,'') Grup_client,isnull(ct.valor,'') + ' ' + isnull(cp.valor,'') Telefon,'' nota, '' Data ,'' Param1 ,0 q,c.codi ,c.nom from clients c join constantsclient cc on c.codi = cc.codi and Variable = 'TipusComanda' left Join constantsclient " ' and valor = 'Diaria' //Descartat per ensenyar tots els clients
    sql = sql & "ct on ct.codi = c.codi and ct.variable='Tel'  left Join constantsclient  cp on cp.codi = c.codi and cp.variable='P_Contacte'  left join constantsclient cg on cg.codi=c.codi and cg.variable='Grup_client' "
    sql = sql & "union "
    sql = sql & "Select isnull(cg.valor,'') Grup_client,isnull(ct.valor,'') + ' ' + isnull(cp.valor,'') Telefon,Param3,Fecha,Param1,0 q ,c.codi,c.nom from [Agenda_" & Year(data) & "] a join Clients c on a.Concepto = 'ComandaClient_' + cast(c.codi as varchar) and day(a.fecha) = " & Day(data) & " and month(a.fecha) = " & Month(data) & " and year(a.fecha) = " & Year(data) & " left Join constantsclient  ct on ct.codi = c.codi and ct.variable='Tel'   left Join constantsclient  cp on cp.codi = c.codi and cp.variable='P_Contacte' "
    sql = sql & " left Join constantsclient  cg on cg.codi=c.codi and cg.variable='Grup_client' "
    sql = sql & "union "
    sql = sql & "select isnull(cg.valor,'') Grup_client,isnull(max(ct.valor),'') + ' ' + isnull(max(cp.valor),'') Telefon,'' nota, '' Data ,'' Param1 ,round(sum(Quantitatdemanada * a.preu ),2) q,Client,c.nom from  " & tablaServits & "  s join clients c on c.codi = s.client left Join constantsclient  ct on ct.codi = c.codi and ct.variable='Tel' join Articles a on a.codi = s.codiarticle   left Join constantsclient  cp on cp.codi = c.codi and cp.variable='P_Contacte' "
    sql = sql & " left Join constantsclient  cg on cg.codi=c.codi and cg.variable='Grup_client' "
    sql = sql & "group by cg.valor,Client,c.nom) A  "
    If P3 = "1" Then 'Amb anotacio
        sql = sql & "where len(Nota)>0"
    ElseIf P3 = "0" Then 'Sense anotacio
        sql = sql & "where len(Nota)=0"
    End If
    sql = sql & "group by Grup_client,Telefon,codi,nom,nota order by Grup_client,Client,Qui  "
    
    rellenaHojaSql "Nota", sql, Libro.Sheets(Libro.Sheets.Count), 0
    

End Sub


Sub ExcelProduccioEquips(Libro, data As Date)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients, sql
    
    sql = "select equip ,c.nom ,viatge,a.nom,Sum(Quantitatservida) Servit , "
    sql = sql & "isnull(t.preumajor , a.preumajor) Preu , Sum(Quantitatservida) * isnull(t.preumajor , a.preumajor)  Import ,Comentari "
    sql = sql & "from [" & DonamNomTaulaServit(data) & "] s "
    sql = sql & "left join clients c on s.client = c.codi "
    sql = sql & "left join articles a  on s.codiarticle = a.codi "
    sql = sql & "left join Tarifesespecials t on c.[Desconte 5] = t.tarifacodi and t.codi = a.codi "
    sql = sql & "Where Quantitatservida > 0 "
    sql = sql & "group by equip,c.nom ,viatge,a.nom,isnull(t.preumajor , a.preumajor),Comentari "
    sql = sql & "order by equip,c.nom ,viatge,a.nom "
    
    rellenaHojaSql "Equips", sql, Libro.Sheets(Libro.Sheets.Count), 0

End Sub



Sub ExcelMajor(Libro, data As Date)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients, sql
    Dim D As Date


    D = DateAdd("m", -1, Now)
    While D > DateAdd("m", -6, Now)
        sql = "select empnom ,sum(total) from " & tablaFacturaProforma(D) & " "
        sql = sql & "group by empnom "
        sql = sql & "order by sum(total) desc "
        rellenaHojaSql "Proveedores " & Format(D, " mmmm"), sql, Libro.Sheets(Libro.Sheets.Count), 0
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)

        sql = "Select clientnom,Sum(Total) From [" & NomTaulaFacturaIva(D) & "] "
        sql = sql & " group by clientnom order by sum(total) desc ,clientnom "
        rellenaHojaSql "Facturat " & Format(D, " mmmm"), sql, Libro.Sheets(Libro.Sheets.Count), 0
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)

        sql = " select i_brut,treb from sousnominaimportats  where left(data,6) = '" & Format(D, "yyyymm") & "'  order by i_brut desc "
        rellenaHojaSql "Sous " & Format(D, " mmmm"), sql, Libro.Sheets(Libro.Sheets.Count), 0
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        
        D = DateAdd("m", -1, D)
    Wend

End Sub




Sub ExcelMateriasPrimas(Libro, data As Date)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients, sql
    
    sql = "select max(a.nombre) + ' ('+max(a.descripcion)+') '  Almacen ,max(m.nombre)  + ' ('+max(m.descripcion)+') ' MateriaPrima, "
    sql = sql & "convert(varchar , fechaentrada , 105) FechaRecepcion "
    sql = sql & ",sum(s.cantidad) EstockActual "
    sql = sql & "From "
    sql = sql & "ccstock s "
    sql = sql & "join cCmateriasPrimas m on m.id = s.matprima "
    sql = sql & "join ccalmacenes A on a.id = m.almacen "
    sql = sql & "where estado = 'DENTRO' "
    sql = sql & "Group By "
    sql = sql & "M.Id , convert(VarChar, fechaentrada, 105) "
    sql = sql & "Order By "
    sql = sql & "Max (a.nombre), Max(a.DESCRIPCIoN), Max(M.nombre) "
    
    rellenaHojaSql "Tots", sql, Libro.Sheets(Libro.Sheets.Count), 0

End Sub

Sub ExcelModifGraella(Libro, data As Date)
    Dim K, i, Kk, Rs As rdoResultset, sql
    
    sql = "select s.modificat,s.quistamp qui,c.nom client,a.nom article,s.viatge,s.equip,s.quantitatDemanada,"
    sql = sql & " s.quantitatTornada,s.quantitatServida from [" & DonamNomTaulaServit(data) & "trace] s "
    sql = sql & " left join clients c on (s.client=c.codi)"
    sql = sql & " left join articles a on (s.codiArticle=a.codi) order by s.modificat"
    
    rellenaHojaSql "ModifGraella " & data, sql, Libro.Sheets(Libro.Sheets.Count), 0

End Sub


Sub ExcelFacturat(Libro)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients, sql
    Dim D As Date
    D = DateSerial(Year(Now), 1, 1)
    While Year(D) = Year(Now)
        sql = "select clientnom,sum(total) from [" & NomTaulaFacturaIva(D) & "] group by clientnom order by clientnom"
        rellenaHojaSql "Facturat " & Format(D, "mmmm"), "select clientnom,sum(total) from [" & NomTaulaFacturaIva(D) & "] group by clientnom order by clientnom", Libro.Sheets(Libro.Sheets.Count), 0
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        D = DateAdd("m", 1, D)
    Wend

End Sub





Sub ExcelQuotesMensuals(Libro)
    Dim Rs As rdoResultset, rsRec As rdoResultset, rsId As rdoResultset
    Dim sql As String, idRec As String
    Dim D As Date, fecha As Date
    Dim anyo As Integer, mes As Integer
    
    D = Now()
    
    If UCase(EmpresaActual) = UCase("HitRs") Then
        Set Rs = Db.OpenResultset("select distinct nom, db, llicencia from Hit.dbo.web_empreses we left join Hit.dbo.web_serveiscomuns ws on we.nom = ws.empresa where ws.actiu=1 ")
        While Not Rs.EOF
            Set rsRec = Db.OpenResultset("Select * From RecursosExtes where variable='ACCESO_BD' and valor='" & Rs("db") & "'")
            If rsRec.EOF Then
                Set rsId = Db.OpenResultset("select top 1 NEWID() Id from recursos")
                idRec = rsId("Id")
                
                ExecutaComandaSql "Insert Into Recursos values ('" & idRec & "', 'ACCESO BD " & Rs("nom") & "', 'LLOC')"
                ExecutaComandaSql "Insert Into RecursosExtes values ('" & idRec & "', 'ACCESO_BD', '" & Rs("db") & "')"
                ExecutaComandaSql "Insert Into RecursosExtes values ('" & idRec & "', 'LICENCIA', '" & Rs("llicencia") & "')"
                ExecutaComandaSql "Insert Into RecursosExtes values ('" & idRec & "', 'DESCRIPCION', '" & Rs("nom") & "')"
                ExecutaComandaSql "Insert Into RecursosExtes values ('" & idRec & "', 'CLIENTE', '0')"
                ExecutaComandaSql "Insert Into RecursosExtes values ('" & idRec & "', 'PRODUCTO', '1650')"      'Acceso BD
            End If
            Rs.MoveNext
        Wend
    
        sql = "select e2.valor Llicencia,c.codi Codi ,c.nom Client,c.Nif,c.Adresa,c.Ciutat, c.cp,cc1.valor cc,r.nombre as Centre,"
        sql = sql & "a.NOM + ' C. ' + r.nombre  as producte , "
        sql = sql & "round((100.00-right(isnull(d.valor,'0|0'),len(isnull(d.valor,'0|0')) -charindex('|',isnull(d.valor,'0|0'))))/ 100  * a.PreuMajor , 2)  as preu "
        sql = sql & ",cc0.valor as CompteCorrent,cc1.valor DiaPagament,cc2.valor eMail,cc4.valor Venciment,cc5.valor Tel, cc6.valor IDIOMA,cc7.valor Drebuts,cc8.valor Provincia,cc9.valor FormaPago,cc10.valor Contacte "
        sql = sql & "from recursos r "
        sql = sql & "join recursosExtes e2 on e2.id = r.id and e2.variable = 'Licencia' and r.tipo = 'LLOC' "
        sql = sql & "join recursosExtes e on e.id = r.id and e.variable = 'Cliente' and r.tipo = 'LLOC' and ISNUMERIC(e.valor)=1 "
        sql = sql & "left join recursosExtes e3 on e3.id = r.id and e3.Variable = 'PRODUCTO' "
        sql = sql & "left join clients c on c.codi = e.valor and not c.Codi=2421 "
        sql = sql & "left join Articles a on cast(a.codi as varchar) = e3.valor "
        sql = sql & "left join ConstantsClient cc0  on c.codi = cc0.codi  and  cc0.Variable ='CompteCorrent' "
        sql = sql & "left join ConstantsClient cc1  on c.codi = cc1.codi  and  cc1.Variable ='DiaPagament' "
        sql = sql & "left join ConstantsClient cc2  on c.codi = cc2.codi  and  cc2.Variable ='eMail' "
        sql = sql & "left join ConstantsClient cc4  on c.codi = cc4.codi  and  cc4.Variable ='Venciment' "
        sql = sql & "left join ConstantsClient cc5  on c.codi = cc5.codi  and  cc5.Variable ='Tel' "
        sql = sql & "left join ConstantsClient cc6  on c.codi = cc6.codi  and  cc6.Variable ='IDIOMA' "
        sql = sql & "left join ConstantsClient cc7  on c.codi = cc7.codi  and  cc7.Variable ='Drebuts' "
        sql = sql & "left join ConstantsClient cc8  on c.codi = cc8.codi  and  cc8.Variable ='Provincia' "
        sql = sql & "left join ConstantsClient cc9  on c.codi = cc9.codi  and  cc9.Variable ='FormaPago' "
        sql = sql & "left join ConstantsClient cc10 on c.codi = cc10.codi and cc10.Variable ='P_Contacte' "
        sql = sql & "left join constantsclient d    on d.codi = c.codi    and    d.variable ='DtoProducte' and left(d.valor,charindex('|',d.valor)-1) = a.codi "
        sql = sql & "where ISNUMERIC(e2.valor) = 1 "
        sql = sql & "order by c.codi,e2.valor,c.nom,r.nombre,preu "
        rellenaHojaSql "Quotes Mensuals", sql, Libro.Sheets(Libro.Sheets.Count), 0
        
        fecha = DateAdd("M", 1, Now())
        fecha = CDate(DateSerial(Year(fecha), Month(fecha), 1))
        
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        
        
        sql = "select cc1.Valor LLicencia, isnull(c2.Codi, c.codi) codi, isnull(c2.[nom llarg], c.[nom llarg]) Client, isnull(c2.Nif, c.Nif) Nif, isnull(c2.Adresa, c.Adresa) + ' ' + isnull(c2.Ciutat, c.Ciutat) + ' ' + isnull(c2.Cp, c.Cp) + ' ' + isnull(cc3.valor, cc13.valor) Adreça, "
        sql = sql & "isnull(cc4.valor, cc14.Valor) tel, isnull(cc5.Valor, cc15.Valor) email, c.Nom + ' ' + s.comentari centre, a.nom producte, "
        sql = sql & "case "
        sql = sql & "when a.Desconte = 1 then isnull(t.preumajor, a.PreuMajor)-(isnull(t.preumajor, a.PreuMajor))*(c.[Desconte 1]/100.00) "
        sql = sql & "when a.Desconte = 2 then isnull(t.preumajor, a.PreuMajor)-(isnull(t.preumajor, a.PreuMajor))*(c.[Desconte 2]/100.00) "
        sql = sql & "when a.Desconte = 3 then isnull(t.preumajor, a.PreuMajor)-(isnull(t.preumajor, a.PreuMajor))*(c.[Desconte 3]/100.00) "
        sql = sql & "when a.Desconte = 4 then isnull(t.preumajor, a.PreuMajor)-(isnull(t.preumajor, a.PreuMajor))*(c.[Desconte 4]/100.00) "
        sql = sql & "end * S.quantitatServida preu , isnull(cc6.Valor, cc16.Valor) [Compte Corrent] , ISNULL(cc7.valor, cc17.Valor) [Forma Pagament] "
        sql = sql & "from [servit-" & Format(fecha, "yy-mm-dd") & "]  s "
        sql = sql & "left join articles a on s.codiarticle=a.codi "
        sql = sql & "left join clients c on s.Client=c.Codi "
        sql = sql & "left join ConstantsClient cc1 on c.Codi=cc1.Codi and cc1.Variable='OrdreRuta' "
        sql = sql & "left join ConstantsClient cc2 on c.Codi=cc2.Codi and cc2.Variable='empMareFac' "
        sql = sql & "left join ConstantsClient cc13 on c.Codi=cc13.Codi and cc13.Variable='Provincia' "
        sql = sql & "left join ConstantsClient cc14 on c.Codi=cc14.Codi and cc14.Variable='Tel' "
        sql = sql & "left join ConstantsClient cc15 on c.Codi=cc15.Codi and cc15.Variable='eMail' "
        sql = sql & "left join ConstantsClient cc16 on c.Codi=cc16.Codi and cc16.Variable='CompteCorrent' "
        sql = sql & "left join ConstantsClient cc17 on c.Codi=cc17.Codi and cc17.Variable='FormaPago' "
        sql = sql & "left join clients c2 on cc2.Valor  = c2.Codi "
        sql = sql & "left join tarifesEspecials t on a.codi=t.codi and t.tarifaCodi=c2.[Desconte 5] "
        sql = sql & "left join ConstantsClient cc3 on c2.Codi=cc3.Codi and cc3.Variable='Provincia' "
        sql = sql & "left join ConstantsClient cc4 on c2.Codi=cc4.Codi and cc4.Variable='Tel' "
        sql = sql & "left join ConstantsClient cc5 on c2.Codi=cc5.Codi and cc5.Variable='eMail' "
        sql = sql & "left join ConstantsClient cc6 on c2.Codi=cc6.Codi and cc6.Variable='CompteCorrent' "
        sql = sql & "left join ConstantsClient cc7 on c2.Codi=cc7.Codi and cc7.Variable='FormaPago' "
        sql = sql & "Where c.Codi Is Not Null And (tipusComanda = 1 or tipusComanda = 3) "
        sql = sql & "order by isnull(c2.[nom llarg], c.[nom llarg]),c.Nom + ' ' + s.comentari"


        rellenaHojaSql "Quotes Mensuals II", sql, Libro.Sheets(Libro.Sheets.Count), 0
        
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        
        fecha = Now()
        fecha = CDate(DateSerial(Year(fecha), Month(fecha), 1))
        
        sql = "select s.Data, cc1.Valor LLicencia, isnull(c2.Codi, c.codi) codi, isnull(c2.[nom llarg], c.[nom llarg]) Client, isnull(c2.Nif, c.Nif) Nif, isnull(c2.Adresa, c.Adresa) + ' ' + isnull(c2.Ciutat, c.Ciutat) + ' ' + isnull(c2.Cp, c.Cp) + ' ' + isnull(cc3.valor, cc13.valor) Adreça, "
        sql = sql & "isnull(cc4.valor, cc14.Valor) tel, isnull(cc5.Valor, cc15.Valor) email, c.Nom + ' ' + s.comentari centre, a.nom producte, isnull(t.preumajor, a.PreuMajor) preu, isnull(cc6.Valor, cc16.Valor) [Compte Corrent], ISNULL(cc7.valor, cc17.Valor) [Forma Pagament] "
        sql = sql & "from ( "
        While Month(fecha) = Month(Now())
            If Day(fecha) > 1 Then sql = sql & " union all "
            sql = sql & "select '" & Format(fecha, "dd/mm/yyyy") & "' Data, * from " & DonamTaulaServit(fecha) & " "
           
            fecha = DateAdd("D", 1, fecha)
        Wend
        sql = sql & ") s "
        sql = sql & "left join articles a on s.codiarticle=a.codi "
        sql = sql & "left join clients c on s.Client=c.Codi "
        sql = sql & "left join ConstantsClient cc1 on c.Codi=cc1.Codi and cc1.Variable='OrdreRuta' "
        sql = sql & "left join ConstantsClient cc2 on c.Codi=cc2.Codi and cc2.Variable='empMareFac' "
        sql = sql & "left join ConstantsClient cc13 on c.Codi=cc13.Codi and cc13.Variable='Provincia' "
        sql = sql & "left join ConstantsClient cc14 on c.Codi=cc14.Codi and cc14.Variable='Tel' "
        sql = sql & "left join ConstantsClient cc15 on c.Codi=cc15.Codi and cc15.Variable='eMail' "
        sql = sql & "left join ConstantsClient cc16 on c.Codi=cc16.Codi and cc16.Variable='CompteCorrent' "
        sql = sql & "left join ConstantsClient cc17 on c.Codi=cc17.Codi and cc17.Variable='FormaPago' "
        sql = sql & "left join clients c2 on cc2.Valor  = c2.Codi "
        sql = sql & "left join tarifesEspecials t on a.codi=t.codi and t.tarifaCodi=c2.[Desconte 5] "
        sql = sql & "left join ConstantsClient cc3 on c2.Codi=cc3.Codi and cc3.Variable='Provincia' "
        sql = sql & "left join ConstantsClient cc4 on c2.Codi=cc4.Codi and cc4.Variable='Tel' "
        sql = sql & "left join ConstantsClient cc5 on c2.Codi=cc5.Codi and cc5.Variable='eMail' "
        sql = sql & "left join ConstantsClient cc6 on c2.Codi=cc6.Codi and cc6.Variable='CompteCorrent' "
        sql = sql & "left join ConstantsClient cc7 on c2.Codi=cc7.Codi and cc7.Variable='FormaPago' "
        sql = sql & "Where c.Codi Is Not Null And tipusComanda = 2 "
        sql = sql & "order by isnull(c2.[nom llarg], c.[nom llarg]),c.Nom + ' ' + s.comentari"
        rellenaHojaSql "Albarans", sql, Libro.Sheets(Libro.Sheets.Count), 0
        
        
    End If

End Sub
Sub ExcelAlbarans(Libro, Di As Date, Df)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients, sql, equip As String
    Dim D As Date, Que, Dto, Cli
    
    InformaMiss "Calculs Excel Albarans"
    ExecutaComandaSql "Drop Table ExcelAlbarans "
    ExecutaComandaSql "Create Table ExcelAlbarans (Dia datetime,Client float,Article float,Equip [nvarchar] (255), Qs float, Qt float, Preu float, Desconte float) "
    D = Di
    While D < DateAdd("d", 1, Df)
        ExecutaComandaSql "Insert Into ExcelAlbarans (Dia,Client,Article,Equip,Qs,Qt) select '" & D & "' DiaDelMes,Client,CodiArticle,Equip,sum(QuantitatServida) Qs,sum(QuantitatTornada) Qt from [" & DonamNomTaulaServit(D) & "] group by Client,CodiArticle,Equip "
        D = DateAdd("d", 1, D)
        InformaMiss "Calculs Excel Albarans Dia " & D
    Wend
    
' Preus generals
    ExecutaComandaSql "Update e set e.preu = isnull(t.[preu],a.[preu])           from ExcelAlbarans e join articles a on a.codi = e.article join Clients  c on c.codi = e.client left join tarifesespecials t on t.tarifacodi = c.[desconte 5] and t.codi = a.codi Where [preu base] = 1 "
    ExecutaComandaSql "Update e set e.preu = isnull(t.[preumajor],a.[preumajor]) from ExcelAlbarans e join articles a on a.codi = e.article join Clients  c on c.codi = e.client left join tarifesespecials t on t.tarifacodi = c.[desconte 5] and t.codi = a.codi Where [preu base] = 2 "
    
' Historic Preus
    ExecutaComandaSql "update t set t.preu = h.preu      from ExcelAlbarans t join (select t.dia tdia,h.codi hcodi,min(h.fechamodif) modi from ExcelAlbarans t join  articleshistorial h on t.Article = h.codi  where h.codi = t.Article and t.dia < h.fechamodif group by h.codi,t.Dia ) p on t.dia=p.tdia and t.Article  = hcodi join articleshistorial h on h.codi = hcodi and h.fechamodif = modi join Clients c on c.codi = t.client and c.[preu base] = 1 "
    ExecutaComandaSql "update t set t.preu = h.preumajor from ExcelAlbarans t join (select t.dia tdia,h.codi hcodi,min(h.fechamodif) modi from ExcelAlbarans t join  articleshistorial h on t.Article = h.codi  where h.codi = t.Article and t.dia < h.fechamodif group by h.codi,t.Dia ) p on t.dia=p.tdia and t.Article  = hcodi join articleshistorial h on h.codi = hcodi and h.fechamodif = modi join Clients c on c.codi = t.client and c.[preu base] = 2 "
    
' Tarifes Especials
    ExecutaComandaSql "Update e set e.preu = t.[preu]      from ExcelAlbarans e join articles a on a.codi = e.article join Clients  c on c.codi = e.client left join tarifesespecials t on t.tarifacodi = c.[desconte 5] and t.codi = a.codi Where [preu base] = 1 and not t.preu is null "
    ExecutaComandaSql "Update e set e.preu = t.[preumajor] from ExcelAlbarans e join articles a on a.codi = e.article join Clients  c on c.codi = e.client left join tarifesespecials t on t.tarifacodi = c.[desconte 5] and t.codi = a.codi Where [preu base] = 2 and not t.preumajor is null "
    
    
' Preus Especials
    ExecutaComandaSql "Update ExcelAlbarans  set Preu = tarifesespecialsclients.preu      from ExcelAlbarans join tarifesespecialsclients on ExcelAlbarans.Article = tarifesespecialsclients.Codi and tarifesespecialsclients.client = ExcelAlbarans.Client join Clients c on c.codi = ExcelAlbarans.client and c.[preu base] = 1 "
    ExecutaComandaSql "Update ExcelAlbarans  set Preu = tarifesespecialsclients.preumajor from ExcelAlbarans join tarifesespecialsclients on ExcelAlbarans.Article = tarifesespecialsclients.Codi and tarifesespecialsclients.client = ExcelAlbarans.Client join Clients c on c.codi = ExcelAlbarans.client and c.[preu base] = 2 "
    
    ExecutaComandaSql "Update ExcelAlbarans Set Desconte = 0 "
    ExecutaComandaSql "Delete ExcelAlbarans where preu is null "
        
    For i = 1 To 4
        InformaMiss "Calculs Excel Aplicant descomptes Pas " & i
        Set Rs = Db.OpenResultset("select codi,variable,isnull(valor,'') valor from ConstantsClient Where variable = 'DtoProducte' or variable = 'DtoFamilia' ")
        While Not Rs.EOF
            If InStr(Rs("valor"), "|") > 0 Then
                Que = Split(Rs("valor"), "|")(0)
                Dto = Split(Rs("valor"), "|")(1)
                Cli = Rs("Codi")
                    If Rs("Variable") = "DtoFamilia" Then
                        If i = 3 Then ExecutaComandaSql "Update ExcelAlbarans Set Desconte = " & Dto & " From ExcelAlbarans join articles On  ExcelAlbarans.Article = Articles.Codi And Articles.Familia = '" & Que & "' Where ExcelAlbarans.Client = " & Cli
                        If i = 2 Then ExecutaComandaSql "Update ExcelAlbarans Set Desconte = " & Dto & " From ExcelAlbarans Fac join articles A On  Fac.Article = A.Codi join families F3 on A.Familia = F3.nom join families F2 on F3.pare = F2.nom and F2.nom = '" & Que & "' Where Fac.Client = " & Cli
                        If i = 1 Then ExecutaComandaSql "Update ExcelAlbarans Set Desconte = " & Dto & " From ExcelAlbarans Fac join articles A On  Fac.Article = A.Codi join families F3 on A.Familia = F3.nom join families F2 on F3.pare = F2.nom join families F1 on F2.pare = F1.nom and F1.nom = '" & Que & "' Where Fac.Client = " & Cli
                    Else
                        If i = 4 Then ExecutaComandaSql "Update ExcelAlbarans Set Desconte = " & Dto & " Where Article = " & Que & " And  Client = " & Cli
                    End If
            End If
            DoEvents
            Rs.MoveNext
        Wend
    Next
    Rs.Close
    ExecutaComandaSql "Update ExcelAlbarans  Set Preu = round(Preu * ((100 - Desconte) / 100),4) "
        
    rellenaHojaSql "Albarans " & Format(Di, "mmmm yyyy"), "select isnull(cc.valor, '') valor,c.[Nom Llarg] + '(' + nom + ')',sum((Qs-qt) * preu) Import from ExcelAlbarans e join Clients c on c.Codi = e.client left join constantsclient  Cc on c.codi = cc.codi and cc.variable = 'Grup_client' group by cc.valor,c.[Nom Llarg],c.[Nom] order by cc.valor,c.[Nom Llarg]", Libro.Sheets(Libro.Sheets.Count), 0
    
    ExecutaComandaSql "update ExcelAlbarans  set Equip='' where Equip is null"
    Kk = 0
    Set Rs = Db.OpenResultset("Select distinct Equip From ExcelAlbarans order by equip ")
    While Not Rs.EOF
        equip = Rs("Equip")
 '       Set Rs3 = rec("select sum((Qs-qt) * preu) Import from ExcelAlbarans e join Clients c on c.Codi = e.client left join constantsclient  Cc on c.codi = cc.codi and cc.variable = 'Grup_client' And Equip = '" & Equip & "'  group by cc.valor,c.[Nom Llarg],c.[Nom] order by cc.valor,c.[Nom Llarg]")
        sql = "select sum(Import) from "
        sql = sql & "( "
        sql = sql & "select "
        sql = sql & "cc.valor Valor ,c.[Nom Llarg] + '(' + nom + ')' Client, "
        sql = sql & "case  Equip     when  '" & equip & "'  then sum((Qs-qt) * preu)   else 0 end  Import "
        sql = sql & "from ExcelAlbarans e join Clients c on c.Codi = e.client left join constantsclient  Cc on c.codi = cc.codi and cc.variable = 'Grup_client' "
        sql = sql & "group by cc.valor,c.[Nom Llarg],c.[Nom] ,Equip "
        sql = sql & ") a "
        sql = sql & "group by valor,client "
        sql = sql & "order by valor,client "
        rellenaHojaSql equip, sql, Libro.Sheets(Libro.Sheets.Count), Kk + 3
'        Hoja.Range(Hoja.Cells(1, 1), Hoja.Cells(Rs3.Fields.Count, Rs3.Fields.Count) + 500).CopyFromRecordset Rs3
        Kk = Kk + 1
        'Tornat
        sql = "select sum(ImportTornat) from "
        sql = sql & "( "
        sql = sql & "select "
        sql = sql & "cc.valor Valor ,c.[Nom Llarg] + '(' + nom + ')' Client, "
        sql = sql & "case  Equip     when  '" & equip & "'  then sum(qt * preu)   else 0 end  ImportTornat "
        sql = sql & "from ExcelAlbarans e join Clients c on c.Codi = e.client left join constantsclient  Cc on c.codi = cc.codi and cc.variable = 'Grup_client' "
        sql = sql & "group by cc.valor,c.[Nom Llarg],c.[Nom] ,Equip "
        sql = sql & ") a "
        sql = sql & "group by valor,client "
        sql = sql & "order by valor,client "
        rellenaHojaSql equip & " Tornat", sql, Libro.Sheets(Libro.Sheets.Count), Kk + 3
        Kk = Kk + 1
        Rs.MoveNext
    Wend
    
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Alb. Botiga Resum", "select c.nom Client,c.[Nom Llarg] Nom,sum(import) Import  from [" & NomTaulaAlbarans(Di) & "] a join clients c on c.codi = a.otros group by c.nom,c.[Nom Llarg]  Order by c.nom", Libro.Sheets(Libro.Sheets.Count)
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Alb. Botiga Detall", " Select c.nom Client,c.[Nom Llarg] Nom,day(data) Dia,b.nom Botiga ,num_tick NumAlbara,sum(import) Import   from [" & NomTaulaAlbarans(Di) & "] a join clients c on c.codi = a.otros join clients b on a.botiga = b.codi group by b.nom,c.[Nom Llarg] ,c.nom,num_tick,day(data) Order by c.nom,day(data),b.nom,num_tick ", Libro.Sheets(Libro.Sheets.Count)

End Sub


Sub ExcelAlbarans2(Libro, Di As Date, Df)
    Dim rsGrups As rdoResultset
    Dim sql As String
    Dim Grup As String
       
    InformaMiss "Calculs Excel Albarans"
    ExecutaComandaSql "Drop Table ExcelAlbarans2 "
    ExecutaComandaSql "Create Table ExcelAlbarans2 (Dia datetime, Client float, dataFac datetime, nFac nvarchar(40), nAlb nvarchar(40), Article float, Qs float, Qt float, Preu float, Desconte float) "

    sql = "insert into ExcelAlbarans2 (Dia, client, dataFac, nFac, nAlb, Article, Qs, Qt, Preu, Desconte) "
    sql = sql & "select d.data, i.clientCodi, i.dataFActura, i.numfactura, "
    sql = sql & "case when CHARINDEX('IdAlbara:', d.referencia)>0 then "
    sql = sql & " SUBSTRING (d.referencia, CHARINDEX('IdAlbara:', d.referencia)+9, CHARINDEX (']', d.referencia, CHARINDEX('IdAlbara:', d.referencia)+9)-(CHARINDEX('IdAlbara:', d.referencia)+9)) "
    sql = sql & "Else '' end nAlb, d.producte , d.Servit, d.Tornat, d.Import, d.desconte "
    sql = sql & "From [" & NomTaulaFacturaIva(Di) & "] i "
    sql = sql & "left join [" & NomTaulaFacturaData(Di) & "] d on d.idfactura = i.idfactura "
    sql = sql & "Where Servit <> 0 Or Tornat <> 0"
    ExecutaComandaSql sql
       
  '  Sql = "select cc.Valor Tipus_Cli, c.Nom Nom_cli, c.Nif, e.nAlb, e.dia, CAST(sum(E.PREU) AS NVARCHAR(20)) Import "
  '  Sql = Sql & "from excelAlbarans2 e "
  '  Sql = Sql & "left join clients c on e.client=c.Codi "
  '  Sql = Sql & "left join constantsClient cc on c.Codi= cc.Codi and cc.Variable='Grup_client' "
  '  Sql = Sql & "group by cc.Valor, c.Nom, c.Nif, e.nAlb, e.dia "
  '  Sql = Sql & "order by cc.Valor, c.Nom, c.Nif, e.nAlb, e.dia "
    
    
    sql = "select isnull(cc.Valor, cc2.valor) Tipus_Cli, isnull(c.codi, cz.codi) Codi_cli, isnull(c.Nom, cz.nom) Nom_cli, "
    sql = sql & "isnull(c.Nif, cz.Nif) Nif, e.dataFac, e.nFac, e.nAlb, e.dia, CAST(sum(E.PREU) AS numeric(10,3)) Import "
    sql = sql & "from excelAlbarans2 e "
    sql = sql & "left join clients c on e.client=c.Codi "
    sql = sql & "left join Clients_Zombis cz on e.client=cz.Codi "
    sql = sql & "left join constantsClient cc on c.Codi=cc.Codi and cc.Variable='Grup_client' "
    sql = sql & "left join ( "
    sql = sql & "select * "
    sql = sql & "from constantsClientHistorial as a "
    sql = sql & "where variable='Grup_client' and fechamodif = ( "
    sql = sql & "select max(b.fechaModif) "
    sql = sql & "from constantsClientHistorial as b "
    sql = sql & "where b.variable='Grup_client' and b.Codi  = a.Codi ) "
    sql = sql & ") cc2 on cz.Codi = cc2.Codi "
    sql = sql & "group by isnull(c.codi, cz.codi),isnull(cc.Valor, cc2.valor), isnull(c.Nom, cz.nom), isnull(c.Nif, cz.Nif), e.dataFac, e.nFac, e.nAlb, e.dia "
    sql = sql & "order by isnull(cc.Valor, cc2.valor), isnull(c.Nom, cz.nom), isnull(c.Nif, cz.Nif), e.nFac, e.nAlb, e.dia "
    
    rellenaHojaSql "Albarans " & Format(Di, "mmmm yyyy"), sql, Libro.Sheets(Libro.Sheets.Count), 0
    Libro.Sheets(Libro.Sheets.Count).Columns("I:I").NumberFormat = "0.000"
        
        
    Set rsGrups = Db.OpenResultset("select distinct(Valor) Grup from constantsClient where Variable='Grup_client' order by valor")
    While Not rsGrups.EOF
        Grup = Replace(rsGrups("Grup"), "*", "")
        
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        
        sql = "select cc.Valor [Tipus Client], c.Nom [Nom Client], c.NIF, e.nAlb [Albarà], e.dia Data, isnull(ap.valor,a.Codi) [Ref. Article], isnull(a.NOM, az.NOM) [Nom Article], CAST(E.QS-E.QT AS numeric(10,3)) Quantitat, CAST(E.PREU AS numeric(10,3)) Import "
        sql = sql & "from excelAlbarans2 e "
        sql = sql & "left join clients c on e.client=c.Codi "
        sql = sql & "left join constantsClient cc on c.Codi= cc.Codi and cc.Variable='Grup_client' "
        sql = sql & "left join articles a on e.article=a.Codi "
        sql = sql & "left join articles_Zombis az on e.article=aZ.Codi "
        sql = sql & "left join ArticlesPropietats ap on isnull(a.Codi, az.codi) =ap.CodiArticle and ap.Variable='CODI_PROD' "
        sql = sql & "where cc.Valor = '" & rsGrups("Grup") & "' "
        sql = sql & "order by cc.valor, c.nom , e.dia "
            
        rellenaHojaSql Grup, sql, Libro.Sheets(Libro.Sheets.Count), 0
        
        Libro.Sheets(Libro.Sheets.Count).Columns("H:I").NumberFormat = "0.000"
        
        rsGrups.MoveNext
    Wend


End Sub


Sub ExcelDesquadres(Libro, Di As Date, Df As Date)
    Dim rsBotigues As rdoResultset
    Dim sql As String, botiga As String
       
    InformaMiss "Calculs Excel Desquadraments"
    
    sql = "select distinct c.Nom Botiga, c.Codi "
    sql = sql & "from [" & NomTaulaMovi(Di) & "] v "
    sql = sql & "left join clients c on v.Botiga=c.Codi "
    sql = sql & "where Tipus_moviment='J' "
    sql = sql & "order by c.Nom"
    Set rsBotigues = Db.OpenResultset(sql)
    
    While Not rsBotigues.EOF
        botiga = rsBotigues("Botiga")
        
        sql = "select v.Data, d.nom Dependenta, v.Import Desquadrament "
        sql = sql & "from [" & NomTaulaMovi(Di) & "] v "
        sql = sql & "left join clients c on v.Botiga=c.Codi "
        sql = sql & "left join dependentes d on v.Dependenta = d.CODI "
        sql = sql & "where Tipus_moviment='J' and botiga='" & rsBotigues("codi") & "' "
        sql = sql & "order by c.Nom, v.Data"
    
        rellenaHojaSql botiga, sql, Libro.Sheets(Libro.Sheets.Count), 0
        
        Libro.Sheets(Libro.Sheets.Count).Columns("C:C").NumberFormat = "0.000"
        
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        
        rsBotigues.MoveNext
    Wend
        
        
End Sub



Sub ExcelAlbaransFam(Libro, Di As Date, Df)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients, sql As String, equip As String, familia As String
    Dim D As Date, Que, Dto, Cli
    
    InformaMiss "Calculs Excel Albarans"
    ExecutaComandaSql "Drop Table ExcelAlbaransFam "
    ExecutaComandaSql "Create Table ExcelAlbaransFam (Dia datetime,Client float,Article float,Familia [nvarchar] (255), Qs float, Qt float, Preu float, Desconte float) "
    D = Di
    While D < DateAdd("d", 1, Df)
        sql = "Insert Into ExcelAlbaransFam (Dia,Client,Article,Familia,Qs,Qt) "
        sql = sql & "select '" & D & "' DiaDelMes,s.Client,s.CodiArticle,F1.nom Familia,sum(s.QuantitatServida) Qs,"
        sql = sql & "sum(s.QuantitatTornada) Qt from [" & DonamNomTaulaServit(D) & "] s left join articles A on "
        sql = sql & "s.CodiArticle = A.Codi left join families F3 on A.Familia = F3.nom left join families F2 on "
        sql = sql & "F3.pare = F2.nom left join families F1 on F2.pare = F1.nom group by Client,CodiArticle,F1.nom "
        ExecutaComandaSql sql
        D = DateAdd("d", 1, D)
        InformaMiss "Calculs Excel Albarans Dia " & D
    Wend
    
' Preus generals
    ExecutaComandaSql "Update e set e.preu = isnull(t.[preu],a.[preu])           from ExcelAlbaransFam e join articles a on a.codi = e.article join Clients  c on c.codi = e.client left join tarifesespecials t on t.tarifacodi = c.[desconte 5] and t.codi = a.codi Where [preu base] = 1 "
    ExecutaComandaSql "Update e set e.preu = isnull(t.[preumajor],a.[preumajor]) from ExcelAlbaransFam e join articles a on a.codi = e.article join Clients  c on c.codi = e.client left join tarifesespecials t on t.tarifacodi = c.[desconte 5] and t.codi = a.codi Where [preu base] = 2 "
    
' Historic Preus
    ExecutaComandaSql "update t set t.preu = h.preu      from ExcelAlbaransFam t join (select t.dia tdia,h.codi hcodi,min(h.fechamodif) modi from ExcelAlbaransFam t join  articleshistorial h on t.Article = h.codi  where h.codi = t.Article and t.dia < h.fechamodif group by h.codi,t.Dia ) p on t.dia=p.tdia and t.Article  = hcodi join articleshistorial h on h.codi = hcodi and h.fechamodif = modi join Clients c on c.codi = t.client and c.[preu base] = 1 "
    ExecutaComandaSql "update t set t.preu = h.preumajor from ExcelAlbaransFam t join (select t.dia tdia,h.codi hcodi,min(h.fechamodif) modi from ExcelAlbaransFam t join  articleshistorial h on t.Article = h.codi  where h.codi = t.Article and t.dia < h.fechamodif group by h.codi,t.Dia ) p on t.dia=p.tdia and t.Article  = hcodi join articleshistorial h on h.codi = hcodi and h.fechamodif = modi join Clients c on c.codi = t.client and c.[preu base] = 2 "
    
' Tarifes Especials
    ExecutaComandaSql "Update e set e.preu = t.[preu]      from ExcelAlbaransFam e join articles a on a.codi = e.article join Clients  c on c.codi = e.client left join tarifesespecials t on t.tarifacodi = c.[desconte 5] and t.codi = a.codi Where [preu base] = 1 and not t.preu is null "
    ExecutaComandaSql "Update e set e.preu = t.[preumajor] from ExcelAlbaransFam e join articles a on a.codi = e.article join Clients  c on c.codi = e.client left join tarifesespecials t on t.tarifacodi = c.[desconte 5] and t.codi = a.codi Where [preu base] = 2 and not t.preumajor is null "
    
    
' Preus Especials
    ExecutaComandaSql "Update ExcelAlbaransFam  set Preu = tarifesespecialsclients.preu      from ExcelAlbaransFam join tarifesespecialsclients on ExcelAlbaransFam.Article = tarifesespecialsclients.Codi and tarifesespecialsclients.client = ExcelAlbaransFam.Client join Clients c on c.codi = ExcelAlbaransFam.client and c.[preu base] = 1 "
    ExecutaComandaSql "Update ExcelAlbaransFam  set Preu = tarifesespecialsclients.preumajor from ExcelAlbaransFam join tarifesespecialsclients on ExcelAlbaransFam.Article = tarifesespecialsclients.Codi and tarifesespecialsclients.client = ExcelAlbaransFam.Client join Clients c on c.codi = ExcelAlbaransFam.client and c.[preu base] = 2 "
    
    ExecutaComandaSql "Update ExcelAlbaransFam Set Desconte = 0 "
    ExecutaComandaSql "Delete ExcelAlbaransFam where preu is null "
        
    For i = 1 To 4
        InformaMiss "Calculs Excel Aplicant descomptes Pas " & i
        Set Rs = Db.OpenResultset("select codi,variable,isnull(valor,'') valor from ConstantsClient Where variable = 'DtoProducte' or variable = 'DtoFamilia' ")
        While Not Rs.EOF
            If InStr(Rs("valor"), "|") > 0 Then
                Que = Split(Rs("valor"), "|")(0)
                Dto = Split(Rs("valor"), "|")(1)
                Cli = Rs("Codi")
                    If Rs("Variable") = "DtoFamilia" Then
                        If i = 3 Then ExecutaComandaSql "Update ExcelAlbaransFam Set Desconte = " & Dto & " From ExcelAlbaransFam join articles On  ExcelAlbaransFam.Article = Articles.Codi And Articles.Familia = '" & Que & "' Where ExcelAlbaransFam.Client = " & Cli
                        If i = 2 Then ExecutaComandaSql "Update ExcelAlbaransFam Set Desconte = " & Dto & " From ExcelAlbaransFam Fac join articles A On  Fac.Article = A.Codi join families F3 on A.Familia = F3.nom join families F2 on F3.pare = F2.nom and F2.nom = '" & Que & "' Where Fac.Client = " & Cli
                        If i = 1 Then ExecutaComandaSql "Update ExcelAlbaransFam Set Desconte = " & Dto & " From ExcelAlbaransFam Fac join articles A On  Fac.Article = A.Codi join families F3 on A.Familia = F3.nom join families F2 on F3.pare = F2.nom join families F1 on F2.pare = F1.nom and F1.nom = '" & Que & "' Where Fac.Client = " & Cli
                    Else
                        If i = 4 Then ExecutaComandaSql "Update ExcelAlbaransFam Set Desconte = " & Dto & " Where Article = " & Que & " And  Client = " & Cli
                    End If
            End If
            DoEvents
            Rs.MoveNext
        Wend
    Next
    Rs.Close
    ExecutaComandaSql "Update ExcelAlbaransFam  Set Preu = round(Preu * ((100 - Desconte) / 100),4) "
        
    rellenaHojaSql "Albarans " & Format(Di, "mmmm yyyy"), "select cc.valor,c.[Nom Llarg] + '(' + nom + ')',sum((Qs-qt) * preu) Import from ExcelAlbaransFam e join Clients c on c.Codi = e.client left join constantsclient  Cc on c.codi = cc.codi and cc.variable = 'Grup_client' group by cc.valor,c.[Nom Llarg],c.[Nom] order by cc.valor,c.[Nom Llarg]", Libro.Sheets(Libro.Sheets.Count), 0
    
    Kk = 0
    Set Rs = Db.OpenResultset("Select distinct isnull(Familia,'') Familia From ExcelAlbaransFam order by familia ")
    While Not Rs.EOF
        familia = Rs("Familia")
 '       Set Rs3 = rec("select sum((Qs-qt) * preu) Import from ExcelAlbaransFam e join Clients c on c.Codi = e.client left join constantsclient  Cc on c.codi = cc.codi and cc.variable = 'Grup_client' And Equip = '" & Equip & "'  group by cc.valor,c.[Nom Llarg],c.[Nom] order by cc.valor,c.[Nom Llarg]")
        sql = "select sum(Import) from "
        sql = sql & "( "
        sql = sql & "select "
        sql = sql & "cc.valor Valor ,c.[Nom Llarg] + '(' + nom + ')' Client, "
        sql = sql & "case  Familia     when  '" & familia & "'  then sum((Qs-qt) * preu)   else 0 end  Import, "
        sql = sql & "case  Familia     when  '" & familia & "'  then sum(Qt * preu)   else 0 end  ImportTornat "
        sql = sql & "from ExcelAlbaransFam e join Clients c on c.Codi = e.client left join constantsclient  Cc on c.codi = cc.codi and cc.variable = 'Grup_client' "
        sql = sql & "group by cc.valor,c.[Nom Llarg],c.[Nom] ,Familia "
        sql = sql & ") a "
        sql = sql & "group by valor,client "
        sql = sql & "order by valor,client "

                   
        rellenaHojaSql familia, sql, Libro.Sheets(Libro.Sheets.Count), Kk + 3
        Kk = Kk + 1
        
        sql = "select sum(ImportTornat) from "
        sql = sql & "( "
        sql = sql & "select "
        sql = sql & "cc.valor Valor ,c.[Nom Llarg] + '(' + nom + ')' Client, "
        sql = sql & "case  Familia     when  '" & familia & "'  then sum(Qt * preu)   else 0 end  ImportTornat "
        sql = sql & "from ExcelAlbaransFam e join Clients c on c.Codi = e.client left join constantsclient  Cc on c.codi = cc.codi and cc.variable = 'Grup_client' "
        sql = sql & "group by cc.valor,c.[Nom Llarg],c.[Nom] ,Familia "
        sql = sql & ") a "
        sql = sql & "group by valor,client "
        sql = sql & "order by valor,client "

                   
        rellenaHojaSql familia & " Tornat", sql, Libro.Sheets(Libro.Sheets.Count), Kk + 3
        
        
'        Hoja.Range(Hoja.Cells(1, 1), Hoja.Cells(Rs3.Fields.Count, Rs3.Fields.Count) + 500).CopyFromRecordset Rs3
        Kk = Kk + 1
        
        Rs.MoveNext
    Wend
    
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Alb. Botiga Resum", "select c.nom Client,c.[Nom Llarg] Nom,sum(import) Import  from [" & NomTaulaAlbarans(Di) & "] a join clients c on c.codi = a.otros group by c.nom,c.[Nom Llarg]  Order by c.nom", Libro.Sheets(Libro.Sheets.Count)
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Alb. Botiga Detall", " Select c.nom Client,c.[Nom Llarg] Nom,day(data) Dia,b.nom Botiga ,num_tick NumAlbara,sum(import) Import   from [" & NomTaulaAlbarans(Di) & "] a join clients c on c.codi = a.otros join clients b on a.botiga = b.codi group by b.nom,c.[Nom Llarg] ,c.nom,num_tick,day(data) Order by c.nom,day(data),b.nom,num_tick ", Libro.Sheets(Libro.Sheets.Count)
    



End Sub




Sub ExcelVentas(Libro, mes)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients, sql As String
    Dim D As Date, Que, Dto, Cli
    
    D = mes
    
    ExecutaComandaSql "Drop Table ExcelTmp "
    sql = "select c.nom Client ,a.nom Producte,Import,Quantitat Into ExcelTmp  from (select plu,botiga,sum(import) Import,sum(quantitat) Quantitat From  [" & NomTaulaVentas(D) & "] group by plu,botiga)  s "
    sql = sql & "Join articles a on s.plu = a.codi "
    sql = sql & "join clients c on s.botiga = c.codi "
    sql = sql & "order by c.nom,a.nom "
    ExecutaComandaSql sql
    sql = "Select * From ExcelTmp "
    
    rellenaHojaSql "Ventas " & Format(mes, "mmmm yyyy"), sql, Libro.Sheets(Libro.Sheets.Count), 0
    ExecutaComandaSql "Drop Table ExcelTmp "

End Sub


Sub ExcelVentasMensualFranquicia(Libro, mes, botiga, Usuari, importFinal)
    Dim Rs As ADODB.Recordset
    Dim sql As String, SQL1 As String, sqlTot As String, sqlIvas As String, html As String
    Dim D As Date
    Dim taulaVentas
    Dim i As Integer, x As Integer
    
On Error GoTo nor:
    D = mes
    
    If importFinal <> 0 And IsNumeric(importFinal) Then
        taulaVentas = NomTaulaVentasPrevistes(D)
        CalcularPrevisions "Ventes", "[00-" & Month(D) & "-" & Year(D) & "]", "[" & botiga & "]", "[" & importFinal & "]", ""
    Else
        taulaVentas = NomTaulaVentas(D)
    End If

    '******************************************************************************
    'Ventes detallades
    '******************************************************************************
    
    sql = "select c.nom Client, CAST(CONVERT(NVARCHAR, v.Data , 112) AS DATETIME) Data, RIGHT( CONVERT(varchar, v.Data, 108),8) Hora, "
    sql = sql & "cast(v.Num_tick as nvarchar(10)) Ticket, d.NOM Dependenta, a.NOM Producte, cast(v.Import as nvarchar(10)) Import, "
    sql = sql & "cast(v.Quantitat as nvarchar(10)) Quantitat, cast(datePart(hh, data) as nvarchar(2)) Hora, case when DATEPART(hh,data)>14 then 'TARDA' else 'MATI' end Torn, "
    sql = sql & "DAY(data) [Dia num], Case DatePart(WeekDay,data) When 1 Then 'Diumenge' When 2 Then 'Dilluns' when 3 then 'Dimarts' "
    sql = sql & "when 4 then 'Dimecres' when 5 then 'Dijous' when 6 then 'Divendres' when 7 then 'Dissabte' end [Dia setmana], "
    sql = sql & "cast(MONTH(data) as nvarchar(2)) Mes "
    sql = sql & "from [" & taulaVentas & "] v "
    sql = sql & "left join dependentes d on v.Dependenta = d.CODI "
    sql = sql & "left join articles a on v.Plu = a.Codi "
    sql = sql & "left join clients c on v.Botiga = c.codi "
    If botiga <> "" Then
        sql = sql & "Where v.Botiga = '" & botiga & "' "
    Else
        sql = sql & "where v.Botiga in (select Codi from ConstantsClient where Variable='userFranquicia' and Valor='" & Usuari & "')"
    End If
    sql = sql & "order by c.nom,v.data,v.Num_tick,a.nom "
    
    rellenaHojaSql "Ventas " & Format(mes, "mmmm yyyy"), sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    '******************************************************************************
    'Iva desglosat
    '******************************************************************************
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    sql = "Update articles set tipoiva = 4 where not tipoiva in (1,2,3,4)"
    ExecutaComandaSql sql
    
    ReDim Ivas(0)
    Set Rs = rec("select * from " & DonamTaulaTipusIva(D) & " order by tipus")
    i = 0
    While Not Rs.EOF
        ReDim Preserve Ivas(i)
        Ivas(i) = Rs("iva")
        i = i + 1
        Rs.MoveNext
    Wend
    
    sqlTot = ""
    sqlIvas = ""
    For x = 0 To UBound(Ivas)
        If sqlTot <> "" Then sqlTot = sqlTot + "+"
        sqlTot = sqlTot + "isnull([" & Ivas(x) & "],0)"
        
        If sqlIvas <> "" Then sqlIvas = sqlIvas + ", "
        sqlIvas = sqlIvas + "[" & Ivas(x) & "]"
    Next
    
    SQL1 = ""
    For x = 0 To UBound(Ivas)
        SQL1 = SQL1 & "round([" & Ivas(x) & "]/(1+" & Ivas(x) & "*0.01),2)   [Base " & Ivas(x) & "%], round([" & Ivas(x) & "]*100/(" & sqlTot & "),2) [%], "
    Next
    
    sql = "select client as Botiga, Dia, "
    sql = sql & SQL1 & "  "
    sql = sql & "round(" & sqlTot & ",2) Total "
    sql = sql & "from ( "
    sql = sql & "Select Sum(import) import, ti.iva, c.nom Client, Day(data) Dia "
    sql = sql & "from [" & taulaVentas & "] v "
    sql = sql & "left join (select codi, tipoiva, familia from Articles a union all select codi,tipoiva, familia from articles_zombis az) aa on v.Plu = aa.codi "
    sql = sql & "left join clients c on c.Codi = v.botiga "
    sql = sql & "left join " & DonamTaulaTipusIva(D) & " ti on isNull(aa.TipoIva,5) = ti.Tipus "
    If botiga <> "" Then
        sql = sql & "Where v.Botiga = '" & botiga & "' "
    Else
        sql = sql & "where v.Botiga in (select Codi from ConstantsClient where Variable='userFranquicia' and Valor='" & Usuari & "') "
    End If
    sql = sql & "group by v.Botiga, c.nom, day(v.Data), Ti.iva "
    sql = sql & ") DataTable "
    sql = sql & "PIVOT (sum(import) for iva in (" & sqlIvas & ")) PivotTable "
    sql = sql & "order by client, Dia "

    rellenaHojaSql "IVA " & Format(mes, "mmmm yyyy"), sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    'Recarrec equivalencia (Revisar gdt: resultats/calculs3.asp)
    
    
nor:
    If err.Number <> 0 Then
        html = "<p><h3>Error excel </h3></p>"
        html = html & "<p><b>Mes: </b>" & mes & "</p>"
        html = html & "<p><b>Botiga: </b>" & botiga & "</p>"
        html = html & "<p><b>Usuari: </b>" & Usuari & "</p>"
        html = html & "<p><b>ImportFinal: </b>" & importFinal & "</p>"
        html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
        html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
        html = html & "<p><b>ERROR:</b>" & err.Source & "</p>"
        html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"

        sf_enviarMail "secrehit@hit.cat", EmailGuardia, "ERROR! ExcelFR ha fallat", html, "", ""

       Informa "Error : " & err.Description
    End If
End Sub

Sub ExcelQuadreMensualFranquicia(Libro, mes, botiga, Usuari)
    Dim Rs As ADODB.Recordset, sql As String, SqlBot As String, html As String
    Dim D As Date, Hoja As Excel.Worksheet
    Dim taulaMoviments
    
On Error GoTo nor:
    D = mes
    
    taulaMoviments = "[" & NomTaulaMovi(D) & "]"

    If botiga <> "" Then
        SqlBot = SqlBot & " Botiga = '" & botiga & "' "
    Else
        SqlBot = SqlBot & " Botiga in (select Codi from ConstantsClient where Variable='userFranquicia' and Valor='" & Usuari & "')"
    End If

    sql = "select c.nom Botiga, TCalaix.Data, cast(TCalaix.[Calaix Fet] as nvarchar(10)) [Calaix Fet], cast(isnull(TDescuadre.Descuadre, 0) as nvarchar(10)) Descuadre, "
    sql = sql & "cast(TCalaix.[Calaix Fet]+isnull(TDescuadre.Descuadre, 0) as nvarchar(10)) Recaudat, "
    sql = sql & "cast(isnull(TEntrega.[Entrega Diària],0) as nvarchar(10)) [Entrega Diària], cast(isnull(TTargeta.Targeta, 0) as nvarchar(10)) [Pagat Targeta] , cast(isnull(TAltres.Altres,0) as nvarchar(10)) [Altres Pagaments], "
    sql = sql & "cast(isnull(TClients.Clients,0) as nvarchar(10)) [Clients Atesos] "
    sql = sql & "From "
    sql = sql & "( "
    sql = sql & "select Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME) Data, SUM(import) [Calaix Fet] "
    sql = sql & "From " & taulaMoviments & " "
    sql = sql & "where " & SqlBot & " and Tipus_moviment='Z' "
    sql = sql & "group by Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME)) TCalaix "
    sql = sql & "left join ( "
    sql = sql & "select Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME) Data, SUM(import) [Clients] "
    sql = sql & "From " & taulaMoviments & " "
    sql = sql & "where " & SqlBot & " and Tipus_moviment='G' "
    sql = sql & "group by Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME)) TClients on TCalaix.Data=TClients.data "
    sql = sql & "left join ("
    sql = sql & "select Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME) Data, SUM(import) [Descuadre] "
    sql = sql & "From " & taulaMoviments & " "
    sql = sql & "where " & SqlBot & " and Tipus_moviment='J' "
    sql = sql & "group by Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME)) TDescuadre on TCalaix.Data = TDescuadre.Data "
    sql = sql & "left join ( "
    sql = sql & "select Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME) Data, SUM(import)*-1 [Entrega Diària] "
    sql = sql & "From " & taulaMoviments & " "
    sql = sql & "where " & SqlBot & " and Tipus_moviment='O' and Motiu ='Entrega Diària' "
    sql = sql & "group by Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME)) TEntrega on TCalaix.Data = TEntrega.Data "
    sql = sql & "left join ( "
    sql = sql & "select Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME) Data, SUM(import)*-1 [Targeta] "
    sql = sql & "From " & taulaMoviments & " "
    sql = sql & "where " & SqlBot & " and Tipus_moviment='O' and Motiu like 'Pagat Targeta : %' "
    sql = sql & "group by Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME)) TTargeta on TCalaix.Data = TTargeta.Data "
    sql = sql & "left join ( "
    sql = sql & "select Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME) Data, SUM(import)*-1 [Altres] "
    sql = sql & "From " & taulaMoviments & " "
    sql = sql & "where " & SqlBot & " and Tipus_moviment='O' and Motiu not like 'Pagat Targeta : %' and Motiu <> 'Entrega Diària' "
    sql = sql & "group by Botiga, CAST(CONVERT(NVARCHAR, Data , 112) AS DATETIME)) TAltres on TCalaix.Data = TAltres.Data "
    sql = sql & "left join clients c on c.Codi = TCalaix.botiga "
    sql = sql & "order by c.Nom, Data"

    rellenaHojaSql "Quadrar Caixa " & Format(mes, "mmmm yyyy"), sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    sql = "select c.nom Botiga, CAST(CONVERT(NVARCHAR, m.Data , 112) AS DATETIME) Data, m.Motiu, cast(m.Import as nvarchar(10)) Import "
    sql = sql & "from " & taulaMoviments & " m "
    sql = sql & "left join clients c on m.botiga=c.codi "
    sql = sql & "where (m.Tipus_Moviment='O' or m.Tipus_Moviment='A') and " & SqlBot & " "
    sql = sql & "order by m.botiga, m.data "
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Detall ", sql, Libro.Sheets(Libro.Sheets.Count), 0
    
nor:
    If err.Number <> 0 Then
        html = "<p><h3>Error excel </h3></p>"
        html = html & "<p><b>Mes: </b>" & mes & "</p>"
        html = html & "<p><b>Botiga: </b>" & botiga & "</p>"
        html = html & "<p><b>Usuari: </b>" & Usuari & "</p>"
        html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
        html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
        html = html & "<p><b>ERROR:</b>" & err.Source & "</p>"
        html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
            
        sf_enviarMail "secrehit@hit.cat", EmailGuardia, "ERROR! ExcelFR ha fallat", html, "", ""
            
        Informa "Error : " & err.Description
    End If
End Sub


Sub ExcelVentasDetallFranquicia(Libro, mes, botiga, Usuari, importFinal)
    Dim K, i, Kk, Rs As ADODB.Recordset, clients, sql As String
    Dim D As Date, Que, Dto, Cli, Hoja As Excel.Worksheet, Fila, TotalT, TotalT1, TotalT2, TotalT3, TotalT4
    Dim TotalT5, tipoIva, dia, Total, import, perc, botigaAnt, recEquivalencia, campo As String, max, nCols, diaAnt
    Dim taulaVentas, LastRow, CellNumber
        
    D = mes
    taulaVentas = NomTaulaVentas(D)
    '******************************************************************************
    'Ventes detallades
    '******************************************************************************
    sql = "select c.nom Botiga, data, CAST(num_tick as nvarchar(10)) ticket, d.nom dependenta, ISNULL(p.Valor,a.codi) Codi, "
    sql = sql & "a.nom Producte, case when a.nom like '%vari%' then v.otros else '' end [Descripció], CAST(Import as numeric(10,3)) Import,cast(Quantitat as numeric(10,3)) Quantitat "
    sql = sql & "From  [" & taulaVentas & "] v left join articles a on v.plu = a.codi "
    sql = sql & "left join clients c on v.botiga = c.codi "
    sql = sql & "left join Dependentes d on v.Dependenta=d.codi "
    sql = sql & "left join articlespropietats p on p.Variable = 'CODI_PROD' and p.CodiArticle = v.plu "
    If botiga <> "" Then
        sql = sql & "Where v.Botiga = '" & botiga & "' "
    Else
        sql = sql & "where v.Botiga in (select Codi from ConstantsClient where Variable='userFranquicia' and Valor='" & Usuari & "')"
    End If
    
    sql = sql & " and day(Data) = " & Day(mes) & " "
    sql = sql & "order by c.nom,v.data,v.Num_tick,a.nom "
'    ExecutaComandaSql Sql
   'importFinal
    rellenaHojaSql "Vendes Data " & Format(mes, "dd mm yyyy"), sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    Libro.Sheets(Libro.Sheets.Count).Columns("H:H").NumberFormat = "0.000"
    Libro.Sheets(Libro.Sheets.Count).Columns("I:I").NumberFormat = "0.000"
    
    'total
'    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
'    LastRow = (ActiveSheet.UsedRange.Rows.Count)
'    Hoja.Range("F:F").NumberFormat = "0.00"
'    Hoja.Range("F2:F" & LastRow).Select
'    For Each CellNumber In Selection
'        CellNumber.Value = CDbl(CellNumber.Value)
'    Next CellNumber
'    Hoja.Cells(LastRow + 1, 1) = "TOTAL"
'    Hoja.Cells(LastRow + 1, 1).Font.Bold = True
'    Hoja.Cells(LastRow + 1, 6).FormulaR1C1 = "=SUM(R2C6:R" & LastRow & "C6)"
'    Hoja.Cells(LastRow + 1, 6).Font.Bold = True
'    Hoja.Columns("F:F").EntireColumn.AutoFit
'    Hoja.Range("A" & LastRow + 1 & ":X" & LastRow + 1).Borders.Weight = xlThin
'    Hoja.Range("A" & LastRow + 1 & ":X" & LastRow + 1).Borders(xlEdgeTop).Weight = xlMedium
'    Hoja.Range("A" & LastRow + 1 & ":X" & LastRow + 1).Borders(xlEdgeBottom).Weight = xlMedium
'    Hoja.Range("A" & LastRow + 1 & ":X" & LastRow + 1).Interior.ColorIndex = 36
    sql = "select ISNULL(p.Valor,a.codi) Codi,a.NOM Article, SUM(quantitat) Quantitat,SUM(import) Import, COUNT(num_tick) as clients  "
    sql = sql & "From  [" & taulaVentas & "] v left join articles a on v.plu = a.codi "
    sql = sql & "left join articlespropietats p on p.Variable = 'CODI_PROD' and p.CodiArticle = v.plu "
    If botiga <> "" Then
        sql = sql & "Where v.Botiga = '" & botiga & "' "
    Else
        sql = sql & "where v.Botiga in (select Codi from ConstantsClient where Variable='userFranquicia' and Valor='" & Usuari & "')"
    End If
    
    sql = sql & " and day(Data) = " & Day(mes) & " "
    sql = sql & " group by a.NOM, ISNULL(p.Valor,a.codi) order by a.nom  "
'    ExecutaComandaSql Sql
   'importFinal
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Resum ", sql, Libro.Sheets(Libro.Sheets.Count), 0
    'total

    
    
    
    
End Sub


'************************************************************************
' finalMes(fecha)
' Devuelve el último día del mes de una fecha
'************************************************************************
Function finalMes(ByVal sls_fecha)
    Dim sln_mes, sln_anyo
    Dim sln_fm
    sln_mes = Month(sls_fecha)
    Dim sla_dias(12)
    sla_dias(1) = 31
    sla_dias(2) = 28
    sla_dias(3) = 31
    sla_dias(4) = 30
    sla_dias(5) = 31
    sla_dias(6) = 30
    sla_dias(7) = 31
    sla_dias(8) = 31
    sla_dias(9) = 30
    sla_dias(10) = 31
    sla_dias(11) = 30
    sla_dias(12) = 31
    sln_fm = sla_dias(sln_mes)
    If sln_mes = 2 Then
        sln_anyo = Year(sls_fecha)
        If ((sln_anyo Mod 4 = 0) And (sln_anyo Mod 100 <> 0)) Or (sln_anyo Mod 400 = 0) Then
            sln_fm = sln_fm + 1
        End If
    End If
    finalMes = sln_fm
End Function
Sub ExcelServitTornat(ByRef Hoja As Excel.Worksheet, Di, Df)
Dim K, i, j, H, Kk, Rs As ADODB.Recordset, rsClients As ADODB.Recordset, rsDto As ADODB.Recordset, sql, data As Date
    Dim D As Date, diff As Integer, Diff2 As Integer, ArticleAnt As String, DescTe, tipoPreu, Dpp
    Dim dia, mes, anyo, client, nomClient, article, qt, qs, Preu, sumQs As Integer, sumQt As Integer, ArrayDies, Descon
            
    Di = Replace(Di, "[", "")
    Di = Replace(Di, "]", "")
    Df = Replace(Df, "[", "")
    Df = Replace(Df, "]", "")
    Di = FormatDateTime(Di, 2)
    Df = FormatDateTime(Df, 2)
    diff = DateDiff("d", Di, Df)
    'Impressio dies,titols
    Hoja.Name = "ServitTornat"
    ReDim ArrayDies(diff)
    For i = 0 To diff
        D = DateAdd("d", i, Di)
        ArrayDies(i) = D
        Hoja.Cells(1, i + 2).Value = D
        Hoja.Cells(2, i + 2).Value = "S(T)"
    Next
    Hoja.Cells(2, i + 2).Value = "Total"
    Hoja.Cells(2, i + 3).Value = "Preu/Unitat"
    Hoja.Cells(2, i + 4).Value = "Import"
    Hoja.Cells(2, i + 5).Value = "Devolucions"
    j = 3
    'Sql clients
    sql = "select isnull(c.codi,0) as codi, c.nom , cc.valor from clients c with (nolock) "
    sql = sql & "left join constantsclient cc with (nolock) on c.codi=cc.codi and cc.variable='Grup_client' "
    sql = sql & "where cc.valor = '*BOTIGUES PROPIES' "
    sql = sql & "order by c.nom "
    Set rsClients = rec(sql)
    Do While Not rsClients.EOF
        client = rsClients("codi")
        nomClient = rsClients("nom")
        Hoja.Cells(2, 1).Value = nomClient
        'Descomptes
        DescTe = ""
        tipoPreu = ""
        sql = "select nom,nif,adresa,cp,ciutat,[Desconte ProntoPago],[Desconte 1],[Desconte 2],[Desconte 3],[Desconte 4], "
        sql = sql & "(case when [preu base]<2 then '' else 'major' end)as pb "
        sql = sql & "from clients where codi=" & client
        Set rsDto = rec(sql)
        If Not rsDto.EOF Then
            Dpp = rsDto("Desconte ProntoPago")
            ReDim Descon(4)
            Descon(0) = 0#
            Descon(1) = rsDto("Desconte 1")
            Descon(2) = rsDto("Desconte 2")
            Descon(3) = rsDto("Desconte 3")
            Descon(4) = rsDto("Desconte 4")
            tipoPreu = rsDto("pb")
        End If
        
        sql = "select * from constantsclient where variable='descTE' and codi='" & client & "'"
        Set rsDto = rec(sql)
        If Not rsDto.EOF Then DescTe = rsDto("valor")
        'Servit-tornat
        sql = "Select dia,mes,anyo, isnull(f3.nom,'') fam3,isnull(f2.nom,'') fam2,isnull(f1.nom,'') fam1, isnull(a.nom,'') nom,"
        sql = sql & "CodiArticle,isnull(Qtt,0) QT,isnull(Qss,0) QS, isnull(isnull(isnull(te.preumajor, t.preumajor), a.preumajor),0) preu, a.TipoIva iva,"
        If DescTe = "descTE" Then
            sql = sql & "a.desconte as desconte "
        Else
            sql = sql & "(case when isnull(t.preu" & tipoPreu & ",0)=0 and isnull(te.preu" & tipoPreu & ",0)=0 then a.desconte else 0 end) As Desconte "
        End If
        sql = sql & " From ("
        For i = 0 To diff
            D = DateAdd("d", i, Di)
            dia = Day(D)
            mes = Month(D)
            If Len(mes) = 1 Then mes = "0" & mes
            anyo = Year(D)
            sql = sql & "select " & dia & " as dia," & mes & " as mes," & anyo & " as anyo,sum(quantitatservida) as qss,Sum(quantitattornada) as qtt,CodiArticle from [" & DonamNomTaulaServit(data) & "] "
            sql = sql & "Where client=" & client & " and (quantitatservida > 0 or quantitattornada>0) Group By CodiArticle "
            If i < diff Then sql = sql & " union "
        Next
        sql = sql & " ) s join articles a on a.codi=s.codiarticle Left join tarifesEspecials t on a.codi="
        sql = sql & "t.codi and t.tarifaCodi=(select [desconte 5] from clients where codi = " & client & ") "
        sql = sql & " left join tarifesespecialsclients te on a.codi=te.codi and te.client='" & client & "' "
        sql = sql & "Left join families f1 on f1.nom=a.familia left join families f2 on f2.nom=f1.pare left join families f3 on f3.nom=f2.pare "
        sql = sql & "order by fam3,fam2,Fam1,a.nom,anyo,mes,dia"
        Set Rs = rec(sql)
        ArticleAnt = ""
        'SQL
        Do While Not Rs.EOF
            dia = Rs("dia")
            mes = Rs("mes")
            anyo = Rs("anyo")
            article = Rs("nom")
            qt = Rs("qt")
            qs = Rs("qs")
            Preu = Rs("preu")
            If article <> ArticleAnt Then
                Hoja.Cells(j, 1).Value = article
                ArticleAnt = article
                sumQs = 0
                sumQt = 0
            End If
            sumQt = CInt(sumQt + qt)
            sumQs = CInt(sumQs + qs)
            Diff2 = DateDiff("d", D, Df)
            H = CInt(Day(Df) - Diff2) + 1
            Hoja.Cells(j, H).Value = qs & "(" & qt & ")"
            Rs.MoveNext
            'Impressio totals
            If Rs("nom") <> article Then
                Hoja.Cells(j, diff + 1).Value = sumQs & "(" & sumQt & ")"
                Hoja.Cells(j, diff + 2).Value = Preu
                Hoja.Cells(j, diff + 3).Value = CDbl(Preu * CInt(sumQs - sumQt))
            End If
            j = j + 1
            Rs.MoveNext
        Loop
        rsClients.MoveNext
    Loop
End Sub '




Sub ExcelProveedores(Libro)
    Dim K, i, Kk, Rs As rdoResultset, Punts, clients
    
    
On Error GoTo 0
    rellenaHojaSql "Proveedores", "select codi,nombre,nombrecorto,descripcion,tlf1,tlf2,fax,email from ccproveedores order by nombre", Libro.Sheets(Libro.Sheets.Count), 0
    Set Rs = Db.OpenResultset("Select nombre,Id from ccproveedores order by nombre")
    While Not Rs.EOF
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        rellenaHojaSql Rs("nombre"), "select cc.codigo,cc.nombre,cc.unidades Unidades,e.valor Formato,'' as cantidad  ,cc.precio as PrecioFormato ,'' as PrecioTotal from ccmateriasprimas cc join  ccnombrevalor e on e.nombre = 'UnidadesFormato' and e.id = cc.id where proveedor = '" & Rs("Id") & "' Order by cc.nombre ", Libro.Sheets(Libro.Sheets.Count), 0
        Rs.MoveNext
    Wend

End Sub

Sub ExcelRentabilitat(Libro, Di, Df)
    Dim K, Kk, Rs As rdoResultset, Punts, clients, sql
    Dim D As Date, diff As Integer, i As Integer
    Di = Replace(Di, "[", "")
    Di = Replace(Di, "]", "")
    Df = Replace(Df, "[", "")
    Df = Replace(Df, "]", "")
    Di = FormatDateTime(Di, 2)
    Df = FormatDateTime(Df, 2)
    diff = DateDiff("m", Di, Df)
    For i = 0 To diff
        D = DateAdd("m", i, Di)
        
        
        sql = "Select (f.Import-isnull(mp.precio * (f.servit-f.tornat) ,0)) Marge,Pr.nombre  Proveedor,mp.precio Coste ,c.nom Client,a.nom Article ,f.Client ,f.Producte,f.Import,f.Referencia,f.Data,ff.NumFactura,f.Productenom,d.nom Comercial from "
        sql = sql & "[" & NomTaulaFacturaData(D) & "] f "
        sql = sql & "join [" & NomTaulaFacturaIva(D) & "] ff on ff.idFactura = f.idFactura "
        sql = sql & "Left Join clients C on c.codi = f.client "
        sql = sql & "Left Join Articles a on a.codi = f.producte "
        sql = sql & "left join articlespropietats p on a.codi = p.CodiArticle And p.variable = 'MatPri' and not p.valor ='' "
        sql = sql & "Left Join ccmateriasprimas mp on mp.id = p.valor "
        sql = sql & "Left Join ccproveedores Pr on Pr.Id = mp.Proveedor "
        sql = sql & "left join constantsclient cc on c.codi=cc.codi and cc.variable='comercial' "
        sql = sql & "left join dependentes d on cc.valor=d.codi "
        sql = sql & "Order by f.data,ff.NumFactura "
        
'        Sql = "select numFactura NumFactura,clientnom Client,DataFactura,sum(total) Total from "
'        Sql = Sql & "[" & NomTaulaFacturaIva(D) & "] group by numFactura,clientnom,dataFactura order by numFactura,clientnom"
        rellenaHojaSql "Facturat " & Format(D, "mmmm"), sql, Libro.Sheets(Libro.Sheets.Count), 0
        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        D = DateAdd("m", 1, D)
    Next
End Sub


Sub ExcelFacturesRebudes(Libro, Di, Df, idEmpresa)
    Dim K, Kk, Rs As rdoResultset, Punts, clients, sql, Taulas
    Dim D As Date, diff As Integer, i As Integer
        
    sql = "Select "
    sql = sql & "datafactura,empnom,empnif,empadresa, "
    sql = sql & "numfactura,total, "
    sql = sql & "baseiva1,iva1, "
    sql = sql & "baseiva2,iva2, "
    sql = sql & "baseIva3 , Iva3 "
    sql = sql & "From [ccFacturas_" & Year(Di) & "_Iva] "
    
    Taulas = ""
    
    D = Di
    While D < Df Or (Year(D) <= Year(Df))
        If Not Taulas = "" Then Taulas = Taulas & " Union "
        Taulas = Taulas & " Select * from [ccFacturas_" & Year(D) & "_Iva] "
        D = DateAdd("yyyy", 1, D)
    Wend
    
    sql = "Select datafactura DataFactura,empnom NomProv,empnif NifProv,empadresa AdrProv,empcp CpProv,empciutat CiutatProv, "
    sql = sql & "numfactura NumFactura,total Total, "
    sql = sql & "baseiva1 BaseIva1,iva1 QuotaIva1, "
    sql = sql & "baseiva2 BaseIva2,iva2 QuotaIva2, "
    sql = sql & "baseIva3 BaseIva3,Iva3 QuotaIva3 "
    sql = sql & "From (" & Taulas & ") f  "
    sql = sql & "Where "
    sql = sql & "datafactura <= convert(date,'" & Day(Df) & "-" & Month(Df) & "-" & Year(Df) & "',103) and "
    sql = sql & "datafactura >= convert(date,'" & Day(Di) & "-" & Month(Di) & "-" & Year(Di) & "',103) "
    If idEmpresa >= 0 Then sql = sql & "and ClientCodi ='" & idEmpresa & "' "
    sql = sql & "order by datafactura "
    
    rellenaHojaSql "Fac Reb", sql, Libro.Sheets(Libro.Sheets.Count), 0
    Libro.Sheets(Libro.Sheets.Count).Range("H:N").NumberFormat = "0.00"
      
    'PAGARES
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    
    sql = "select c1.DESCRIPCION as pagare, c1.haber ImportPagare, c2.fecha fechaPagare, p.nombre proveedor, f.NumFactura , f.DataFactura, f.Total ImportFactura "
    sql = sql & "From "
    sql = sql & "( "
    sql = sql & "select * from AsientosContables_" & Year(Di) & " where (CONCEPTO like '%pagare%' or CONCEPTO like '%reb%' or CONCEPTO like '%transferencia%' or CONCEPTO  like '%traspas%') and orden=1) c1 "
    sql = sql & "Right Join "
    sql = sql & "(select * from AsientosContables_" & Year(Di) & " where idNorma43 in "
    sql = sql & "(select idNorma43  from AsientosContables_" & Year(Di) & " where CONCEPTO like '%pagare%' or CONCEPTO like '%reb%' or CONCEPTO like '%transferencia%' or CONCEPTO  like '%traspas%') "
    sql = sql & "and orden>1) c2 on c1.idNorma43 =c2.idnorma43 "
    sql = sql & "left join ccproveedores p on p.id=c2.ReferenciaInterna "
    sql = sql & "left join (select * from [ccFacturas_" & Year(DateAdd("yyyy", -1, Df)) & "_Iva] union all select * from [ccFacturas_" & Year(Df) & "_Iva]) f on c2.FacturaId =f.IdFactura "
    sql = sql & "Where "
    sql = sql & "c2.fecha <= convert(date,'" & Day(Df) & "-" & Month(Df) & "-" & Year(Df) & "',103) and "
    sql = sql & "c2.fecha >= convert(date,'" & Day(Di) & "-" & Month(Di) & "-" & Year(Di) & "',103) "
    sql = sql & "order by pagare"
        
    rellenaHojaSql "Pagares", sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    Libro.Sheets(Libro.Sheets.Count).Range("G:G").NumberFormat = "0.00"
    

End Sub



Sub ExcelHores(ByRef Hoja As Excel.Worksheet, Di, Df)
Dim i As Integer, j As Integer, DiaS As Integer, Rs As ADODB.Recordset, rsDades As ADODB.Recordset
    Dim rsId As ADODB.Recordset, rsIns As ADODB.Recordset, data() As Double, mes As Integer
    Dim sql As String, D As Date, diff As Integer, Fila As Integer, fecha, Accio, AccioPrev, equip, Usuari, idTemp, Hores, DiaEnt, DiaSal
    Dim EquipPrev, totalDia, DiaPrev, Col, TotalTreb, Semana, SemanaAnt, SemanaTemp, SemanaTemp2, RestaSem, Seg, Min
    Dim Semanas(5)
    Dim ArraySemana(2)
    ArraySemana(0) = DatePart("ww", Di - 7, vbMonday, vbFirstFourDays)
    ArraySemana(1) = DatePart("ww", Di - 14, vbMonday, vbFirstFourDays)
    ArraySemana(2) = DatePart("ww", Di - 21, vbMonday, vbFirstFourDays)
    Dim ArrayDies(7)
    ArrayDies(0) = "Lunes"
    ArrayDies(1) = "Martes"
    ArrayDies(2) = "Miercoles"
    ArrayDies(3) = "Jueves"
    ArrayDies(4) = "Viernes"
    ArrayDies(5) = "Sabado"
    ArrayDies(6) = "Domingo"
    'Titol
    'Hoja.Name = "Hores " & Di & " - " & Df
    'Creacio taula temporal
    InformaMiss "Calculs Excel Hores"
    ExecutaComandaSql "Drop Table ExcelHores"
    ExecutaComandaSql "Create Table ExcelHores(id [nvarchar] (255),Dia datetime,Equip [nvarchar] (255),Usuari [nvarchar] (255),DiaE datetime,DiaS datetime,Hores float) "
    ExecutaComandaSql "Drop Table ExcelHoresSem"
    ExecutaComandaSql "Create Table ExcelHoresSem(id [nvarchar] (255),Semana [nvarchar] (255),Dia datetime,Equip [nvarchar] (255),Usuari [nvarchar] (255),DiaIni datetime,DiaFi datetime,Hores float) "
    diff = DateDiff("d", Di, Df)
    Hoja.Cells(3, 1).ColumnWidth = 25
    Hoja.Cells(3, 1).Font.Bold = True
    Semana = DatePart("ww", Di, vbMonday, vbFirstFourDays)
    If diff > 7 Then
        For i = 0 To diff
            Semana = Semana & " - " & DatePart("ww", DateAdd("d", i, Di), vbMonday, vbFirstFourDays)
            i = i + 7
        Next
    End If
    Hoja.Cells(3, 1).Value = "Semana " & Semana
    Hoja.Rows(3).Borders(xlEdgeTop).Weight = xlMedium
    Hoja.Rows(3).Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Rows(3).Interior.ColorIndex = 15
    Hoja.Cells(4, 1).Font.Bold = True
    Hoja.Cells(4, 1).Value = "Empleados por seccion"
    Hoja.Rows(4).Borders(xlEdgeBottom).Weight = xlMedium
    'Omplir taula temporal,treballadors
    
    sql = "select d.codi,de.valor equip "
    sql = sql & "from dependentes d with (nolock) "
    sql = sql & "left join dependentesExtes de with (nolock) on d.codi = de.id "
    sql = sql & "where de.nom='TIPUSTREBALLADOR' "
    sql = sql & "order by equip,d.nom"
    Set Rs = rec(sql)
    Do While Not Rs.EOF
        DoEvents
        Df = DateAdd("d", diff, Di)
        For i = 0 To diff
            Informa2 "Hores  " & Rs("equip") & " " & Rs("codi")
            DoEvents
            equip = Rs("equip")
            Usuari = Rs("codi")
            For j = 0 To diff
            'Registre obligatori
                sql = "INSERT into ExcelHores (id,Dia,Equip,Usuari,DiaE,DiaS,Hores) values ("
                sql = sql & " newId(),'" & Format(DateAdd("d", j, Di), "dd/mm/yyyy") & "','" & equip & "','" & Usuari & "',"
                sql = sql & " '" & DateAdd("d", j, Di) & "','','0')"
                Set rsIns = rec(sql)
            Next
            'id nova entrada
            sql = "select newId() id"
            Set rsId = rec(sql)
            idTemp = rsId("id")
            'Dades fichador (accio:1=entrada,2=sortida)
            sql = "select distinct convert(datetime, cast(day(tmst) as varchar) + '/' + cast(month(tmst) as varchar) + '/' + cast(year(tmst) as varchar) + ' ' + cast(datepart(hh, tmst) as varchar) + ':' + cast(datepart(n, tmst) as varchar), 103) as fecha, accio "
            sql = sql & "from cdpDadesFichador f with (nolock) "
            sql = sql & "where  usuari= '" & Usuari & "'  and tmst>= '" & DateAdd("d", i, Di) & "' and "
            'sql = sql & "where  usuari= '" & Usuari & "'  and tmst>= '" & Di & "' and "
            sql = sql & "tmst<=dateadd(d,1, '" & DateAdd("d", i, Di) & "' ) and accio in (1,2) "
            'sql = sql & "tmst<=dateadd(d,1, '" & Df & "' ) and accio in (1,2) "
            sql = sql & "order by fecha"
            Set rsDades = rec(sql)
            AccioPrev = 2
            Do While Not rsDades.EOF
                fecha = rsDades("fecha")
                Accio = rsDades("accio")
                If Accio = 1 Then DiaEnt = fecha
                If Accio = 2 Then DiaSal = fecha
                If Accio = 1 And AccioPrev = 2 Then
                    AccioPrev = Accio
                    sql = "INSERT into ExcelHores (id,Dia,Equip,Usuari,DiaE,DiaS,Hores) values ("
                    sql = sql & " '" & idTemp & "','" & Format(DiaEnt, "dd/mm/yyyy") & "','" & equip & "','" & Usuari & "',"
                    sql = sql & " '" & DiaEnt & "','','')"
                    Set rsIns = rec(sql)
                ElseIf Accio = 2 And AccioPrev = 1 Then
                    AccioPrev = Accio
                    Hores = DateDiff("s", DiaEnt, DiaSal)
                    sql = "UPDATE ExcelHores set DiaS='" & DiaSal & "',Hores='" & Hores & "' "
                    sql = sql & " where id='" & idTemp & "'"
                    Set rsIns = rec(sql)
                    sql = "select newId() id"
                    Set rsId = rec(sql)
                    idTemp = rsId("id")
                End If
                rsDades.MoveNext
            Loop
            If Accio = 1 Then
                'Si ultima accio es una entrada, busquem sortida dia seguent
                sql = "select distinct convert(datetime, cast(day(tmst) as varchar) + '/' + cast(month(tmst) as varchar) + '/' + cast(year(tmst) as varchar) + ' ' + cast(datepart(hh, tmst) as varchar) + ':' + cast(datepart(n, tmst) as varchar), 103) as fecha, accio "
                sql = sql & "from cdpDadesFichador f with (nolock) "
                sql = sql & "where  usuari= '" & Usuari & "'  and tmst>= '" & DateAdd("d", i + 1, Di) & "' and "
                sql = sql & "tmst<=dateadd(d,1, '" & DateAdd("d", i + 1, Di) & "' ) and accio in (1,2) "
                sql = sql & "order by fecha"
                Set rsDades = rec(sql)
                If Not rsDades.EOF Then
                    If rsDades("accio") = 2 Then
                        'Si el primer registre es una sortida
                        DiaSal = rsDades("fecha")
                        Hores = DateDiff("s", DiaEnt, DiaSal)
                        sql = "UPDATE ExcelHores set DiaS='" & DiaSal & "',Hores='" & Hores & "' "
                        sql = sql & " where id='" & idTemp & "'"
                        Set rsIns = rec(sql)
                    End If
                End If
            End If
        Next
        SemanaTemp = dataSetmana(DateAdd("d", -7, Di))
        SemanaTemp2 = Split(SemanaTemp, ",")
        Semanas(0) = SemanaTemp2(0)
        Semanas(1) = SemanaTemp2(1)
        SemanaTemp = dataSetmana(DateAdd("d", -14, Di))
        SemanaTemp2 = Split(SemanaTemp, ",")
        Semanas(2) = SemanaTemp2(0)
        Semanas(3) = SemanaTemp2(1)
        SemanaTemp = dataSetmana(DateAdd("d", -21, Di))
        SemanaTemp2 = Split(SemanaTemp, ",")
        Semanas(4) = SemanaTemp2(0)
        Semanas(5) = SemanaTemp2(1)
        'Calcul semanes anteriors
        j = 0
        Accio = 0
        AccioPrev = 2
        For i = 0 To 2
            sql = "select distinct convert(datetime, cast(day(tmst) as varchar) + '/' + cast(month(tmst) as varchar) + '/' + cast(year(tmst) as varchar) + ' ' + cast(datepart(hh, tmst) as varchar) + ':' + cast(datepart(n, tmst) as varchar), 103) as fecha, accio "
            sql = sql & "from cdpDadesFichador f with (nolock) "
            sql = sql & "where  usuari= '" & Usuari & "'  and tmst>= '" & Semanas(j) & "' and "
            sql = sql & "tmst<=dateadd(d,1, '" & Semanas(j + 1) & "' ) and accio in (1,2) "
            sql = sql & "order by fecha"
            Set rsDades = rec(sql)
            AccioPrev = 2
            'id nova entrada
            sql = "select newId() id"
            Set rsId = rec(sql)
            idTemp = rsId("id")
            Do While Not rsDades.EOF
                'Si ultima accio es una entrada, busquem sortida dia seguent
                If Accio = 1 And rsDades("accio") = 2 Then
                    DiaSal = rsDades("fecha")
                    Hores = DateDiff("s", DiaEnt, DiaSal)
                    sql = "UPDATE ExcelHoresSem set DiaFi='" & DiaSal & "',Hores='" & Hores & "' "
                    sql = sql & " where id='" & idTemp & "'"
                    Set rsIns = rec(sql)
                    sql = "select newId() id"
                    Set rsId = rec(sql)
                    idTemp = rsId("id")
                End If
                fecha = rsDades("fecha")
                Accio = rsDades("accio")
                If Accio = 1 Then DiaEnt = fecha
                If Accio = 2 Then DiaSal = fecha
                If Accio = 1 And AccioPrev = 2 Then
                    AccioPrev = Accio
                    sql = "INSERT into ExcelHoresSem (id,Semana,Dia,Equip,Usuari,DiaIni,DiaFi,Hores) values ("
                    sql = sql & " '" & idTemp & "','" & DatePart("ww", fecha, vbMonday, vbFirstFourDays) & "','" & Format(DiaEnt, "dd/mm/yyyy") & "','" & equip & "','" & Usuari & "',"
                    sql = sql & " '" & DiaEnt & "','','')"
                    Set rsIns = rec(sql)
                ElseIf Accio = 2 And AccioPrev = 1 Then
                    AccioPrev = Accio
                    Hores = DateDiff("s", DiaEnt, DiaSal)
                    sql = "UPDATE ExcelHoresSem set DiaFi='" & DiaSal & "',Hores='" & Hores & "' "
                    sql = sql & " where id='" & idTemp & "'"
                    Set rsIns = rec(sql)
                    sql = "select newId() id"
                    Set rsId = rec(sql)
                    idTemp = rsId("id")
                End If
                rsDades.MoveNext
            Loop
            j = j + 1
        Next
        Rs.MoveNext
    Loop
    'Impressio XLS
    Fila = 4
    Col = 2
    totalDia = 0
    Informa2 "Generant Excel Hores " & Now
    sql = "select e.Dia,e.Equip,d.Nom Usuari,sum(e.Hores) Hores from ExcelHores e with (nolock) left join dependentes d with (nolock) on (e.usuari=d.codi) "
    sql = sql & "group by e.usuari,d.Nom,e.equip,e.dia order by e.dia,e.equip,d.nom"
    Set rsDades = rec(sql)
    If Not rsDades.EOF Then
        EquipPrev = rsDades("equip")
        equip = rsDades("equip")
        Hoja.Cells(Fila, Col) = Format(rsDades("dia"), "mm/dd/yyyy")
        Hoja.Cells(Fila, Col).Font.Bold = True
        Fila = Fila + 1
        Hoja.Cells(Fila, Col) = ArrayDies(Weekday(rsDades("dia")) - 1)
        Hoja.Cells(Fila, Col).Font.Bold = True
        Hoja.Cells(Fila, 1).Font.Bold = True
        Hoja.Cells(Fila, 1) = equip
        Hoja.Rows(Fila).Interior.ColorIndex = 15
        DiaPrev = rsDades("dia")
    End If
    Do While Not rsDades.EOF
        Fila = Fila + 1
        Hoja.Cells(Fila, 1) = rsDades("usuari")
        Hores = rsDades("hores")
        Min = 0
        If Hores > 0 Then
            Min = Hores / 60
            Hores = Int(Min / 60)
            Min = Min Mod 60
            totalDia = totalDia + Hores + (Min / 100)
            totalDia = Round(totalDia, 2)
        End If
        Hoja.Cells(Fila, Col) = Hores + (Min / 100)
        If Hoja.Cells(Fila, Col).Value > 8 Then Hoja.Cells(Fila, Col).Font.ColorIndex = 3
        'Total
        If rsDades("dia") = Df Then
            Hoja.Cells(Fila, Col + 1).Font.Bold = True
            Hoja.Cells(Fila, Col + 1).Interior.ColorIndex = 35
            Hoja.Cells(Fila, Col + 1).FormulaR1C1 = "=SUM(RC[-" & Col - 1 & "]:RC[-1])"
            If Hoja.Cells(Fila, Col + 1).Value > 8 Then Hoja.Cells(Fila, Col + 1).Font.ColorIndex = 3
        End If
        rsDades.MoveNext
        'Totales dia / equipo
        If Not rsDades.EOF Then
            If EquipPrev <> rsDades("Equip") Then
                Fila = Fila + 1
                Hoja.Cells(Fila, 1).Font.Bold = True
                Hoja.Rows(Fila).Borders(xlEdgeTop).Weight = xlMedium
                Hoja.Rows(Fila).Borders(xlEdgeBottom).Weight = xlMedium
                Hoja.Cells(Fila, 1) = "Totales Diarios"
                Hoja.Cells(Fila, Col).Font.Bold = True
                Hoja.Cells(Fila, Col) = totalDia
                Hoja.Cells(Fila, Col).Interior.ColorIndex = 35
                'Total
                If rsDades("dia") = Df Then
                    Hoja.Cells(Fila, Col + 1).Font.Bold = True
                    Hoja.Cells(Fila, Col + 1).Interior.ColorIndex = 35
                    Hoja.Cells(Fila, Col + 1).FormulaR1C1 = "=SUM(RC[-" & Col - 1 & "]:RC[-1])"
                    If Hoja.Cells(Fila, Col + 1).Value > 8 Then Hoja.Cells(Fila, Col + 1).Font.ColorIndex = 3
                End If
                totalDia = 0
                equip = rsDades("equip")
                EquipPrev = equip
                Fila = Fila + 1
                Hoja.Cells(Fila, 1).Font.Bold = True
                Hoja.Cells(Fila, 1) = equip
                Hoja.Rows(Fila).Interior.ColorIndex = 15
            End If
            If DiaPrev <> rsDades("dia") Then
                DiaPrev = rsDades("dia")
                Fila = 4
                Col = Col + 1
                Hoja.Cells(Fila, Col) = Format(rsDades("dia"), "mm/dd/yyyy")
                Hoja.Cells(Fila, Col).Font.Bold = True
                Fila = Fila + 1
                Hoja.Cells(Fila, Col) = ArrayDies(Weekday(rsDades("dia")) - 1)
                Hoja.Cells(Fila, Col).Font.Bold = True
            End If
        End If
    Loop
    Hoja.Cells(5, Col + 1) = "TOTAL"
    Hoja.Cells(5, Col + 1).Font.Bold = True
    Col = Col + 1
    Hoja.Cells(4, Col + 1) = "Resumen Semanal"
    Hoja.Cells(4, Col + 1).Font.Bold = True
    RestaSem = 7
    Semana = DatePart("ww", Di - RestaSem, vbMonday, vbFirstFourDays)
    Hoja.Cells(5, Col + 1) = Semana
    Hoja.Cells(5, Col + 1).Font.Bold = True
    Fila = 6
    'Impresion Semanas
    SemanaAnt = Semana
    sql = "select e.Semana,e.Equip,d.Nom Usuari,sum(e.Hores) Hores from ExcelHoresSem e with (nolock) left join dependentes d with (nolock) on (e.usuari=d.codi) "
    sql = sql & "group by e.usuari,d.Nom,e.equip,e.semana order by e.semana,e.equip,d.nom"
    Set rsDades = rec(sql)
    If Not rsDades.EOF Then
        EquipPrev = rsDades("equip")
        equip = rsDades("equip")
    End If
    Do While Not rsDades.EOF
        If SemanaAnt <> rsDades("semana") Then
            Col = Col + 1
            RestaSem = RestaSem + 7
            Hoja.Cells(5, Col + 1) = DatePart("ww", Di - RestaSem, vbMonday, vbFirstFourDays)
            Hoja.Cells(5, Col + 1).Font.Bold = True
            SemanaAnt = rsDades("semana")
        End If
        Fila = Fila + 1
        Hores = rsDades("hores")
        Min = 0
        If Hores > 0 Then
            Min = Hores / 60
            Hores = CInt(Min / 60)
            Min = Min Mod 60
            totalDia = totalDia + Hores + (Min / 100)
            totalDia = Round(totalDia, 2)
        End If
        Hoja.Cells(Fila, Col).Font.Bold = True
        Hoja.Cells(Fila, Col).Value = Hores + (Min / 100)
        If Hoja.Cells(Fila, Col).Value > 8 Then Hoja.Cells(Fila, Col + 1).Font.ColorIndex = 3
        rsDades.MoveNext
        If Not rsDades.EOF Then
            If EquipPrev <> rsDades("Equip") Then
                Fila = Fila + 1
                equip = rsDades("equip")
                EquipPrev = equip
            End If
        End If
    Loop
    
    'ExecutaComandaSql "Drop Table ExcelHores"
    
    Informa2 ""
End Sub



Sub ExcelHores2(ByRef Hoja As Excel.Worksheet, Di, emp, Eq, ano)
    Dim i As Integer, j As Integer, DiaS As Integer, Rs As rdoResultset, rsDades As ADODB.Recordset, Equipo, usuario, Hacc As Double, HaccEq As Double, HaccEqDia(10) As Double, HaccEqDiaTot(10) As Double, HaccEqNumOp, HaccEqNumOpTot, K, H1, H2, H3, TsA1, TsA2, TsA3, TsAGr1, TsAGr2, TsAGr3
    Dim rsId As ADODB.Recordset, rsIns As rdoResultset, data() As Double, mes As Integer, HEquipSetm As Double, Rs2 As rdoResultset
    Dim sql As String, D As Date, diff As Integer, Fila As Integer, fecha, Accio, AccioPrev, equip, Usuari, idTemp, Hores, DiaEnt, DiaSal, Anoabuscar
    Dim EquipPrev, totalDia, DiaPrev, Col, TotalTreb, Semana, SemanaAnt, SemanaTemp, SemanaTemp2, RestaSem, Seg, Min, Semanas(5), ArraySemana(2)
    Dim CodUsuario As String, equipoAct As String, coma As Integer


ExecutaComandaSql "CREATE TABLE [dbo].[cdpDadesFichadorEquip](    [idr] [nvarchar](255) NULL,    [equip] [nvarchar](255) NULL) ON [PRIMARY] "


    If Not IsNumeric(ano) Then
        ano = Year(Date)
    End If
    Anoabuscar = ano
    If (Anoabuscar = "") Then
        Anoabuscar = Year(Date)
    End If
    If (Anoabuscar = 0) Then
        Anoabuscar = Year(Date)
    End If
    

    'Titol
    'Hoja.Name = "Hores " & Di & " - " & Df
    'Creacio taula temporal
    InformaMiss "Calculs Excel Hores"

    'Vaciado tabla
    ExecutaComandaSql "CREATE TABLE [fichaje_hist](   [Equipo] [varchar](255) NULL,   [CodUsuario] [int] NULL,   [Usuario] [varchar](255) NULL,   [DiaSemana] [int] NULL,   [DiaMes] [int] NULL,   [Fecha] [datetime] NULL,   [HorasAcumulado] [varchar](255) NULL) ON [PRIMARY]"
    ExecutaComandaSql "truncate table fichaje_hist "
        
    'CalculaHistoricSetmana Di, emp
    If Di > 1 Then CalculaHistoricSetmana Di - 1, emp, Anoabuscar
    If Di > 2 Then CalculaHistoricSetmana Di - 2, emp, Anoabuscar
    If Di > 3 Then CalculaHistoricSetmana Di - 3, emp, Anoabuscar
    
    'Set rs = Db.OpenResultset("Select Equipo,Usuario,diasemana,isnull(HorasAcumulado,0) HorasAcumulado,CodUsuario  from fichaje_hist     Where DatePart(wk, Fecha) = " & Di & "     order by Equipo,Usuario ")
    'Set rs = Db.OpenResultset("Select Equipo,Usuario,diasemana,isnull(HorasAcumulado,0) HorasAcumulado,CodUsuario  from fichaje_hist     Where DatePart(wk, Fecha) = " & Di & "     order by Equipo,Usuario ")
    '**************************************************
    ' 14/02/2011 JORGE SIXTO cambiado para coger del campo historial el equipo si tiene mas de uno y lo especifican en el fichaje
    ' sql = "Select Equipo,Usuario,diasemana,isnull(HorasAcumulado,0) HorasAcumulado,CodUsuario  from ("
    '**************************************************
    
    
    
    
    sql = "Select case isnull(equip,'') when '' then Equipo else equip end as Equipo,Usuario,diasemana,isnull(HorasAcumulado,0) HorasAcumulado,CodUsuario  from ("
    'sql = sql & "select DependentesExtes.Valor as Equipo,Dependentes.CODI as CodUsuario,Dependentes.Nom as Usuario,IsNull(DATEPART( weekday, Entrada),1) as diasemana,"
    sql = sql & "select DependentesExtes.Valor as Equipo,equip,Dependentes.CODI as CodUsuario,Dependentes.Nom as Usuario,IsNull(DATEPART( weekday, Entrada),1) as diasemana,"
    sql = sql & "IsNull(day(Entrada),0) as fec1,IsNull(left(Entrada,11),cast(dateadd(dd,-(datepart(wk,getdate())-" & Di & ")*7,getdate()) as datetime)) as Fecha,IsNull(Sum (DateDiff(Minute, Entrada, salida)),0) as HorasAcumulado "
    sql = sql & "from Dependentes with (nolock) FULL OUTER JOIN "
    sql = sql & "( "
    'sql = sql & "select t1.usuari, t1.tmst as 'Entrada' , "
    sql = sql & "select te1.equip, t1.usuari, t1.tmst as 'Entrada' , "
    sql = sql & "( "
    sql = sql & "select min(tmst) "
    sql = sql & "from cdpDadesFichador t2 with (nolock) "
    sql = sql & "where t2.usuari = t1.usuari and "
    sql = sql & "t2.accio = 2 and "
    sql = sql & "t2.tmst >= t1.tmst and "
    sql = sql & "t2.tmst <= ( "
    sql = sql & "select isnull(min(tmst),'99991231 23:59:59.998') "
    sql = sql & "from cdpDadesFichador with (nolock) "
    sql = sql & "where usuari = t2.usuari and "
    sql = sql & "accio = 1 and "
    sql = sql & "tmst > t1.tmst "
    sql = sql & ") "
    sql = sql & ") as 'Salida' "
    sql = sql & "from cdpDadesFichador t1 with (nolock) "
    sql = sql & "left join cdpdadesfichadorequip te1 with (nolock)  on t1.idr = te1.idr "
    sql = sql & "Where T1.Accio = 1 "
    sql = sql & "and DATEPART( wk, tmst)=  " & Di & " "
    sql = sql & "and year(tmst) = '" & Anoabuscar & "'"
    sql = sql & ") as horario "
    sql = sql & "on horario.usuari = Dependentes.Codi "
    sql = sql & "left join DependentesExtes with (nolock) "
    sql = sql & "on Dependentes.Codi = DependentesExtes.id "
    sql = sql & "where DependentesExtes.Nom = 'EQUIPS' "
    If emp <> "" Then
        sql = sql & "and Dependentes.CODI in (select id from DependentesExtes with (nolock) Where "
        sql = sql & "nom = 'EMPRESA' and valor = '" & emp & "') "
    End If
    sql = sql & "and Dependentes.CODI in ( "
    ' cambio para buscar fin de contrato
    sql = sql & "select codi "
    sql = sql & "from dependentes d3 with (nolock) "
    sql = sql & "left join dependentesextes  with (nolock) on d3.codi=dependentesextes.id and dependentesextes.nom = "
    sql = sql & "(select max(nom) from dependentesextes  with (nolock) where nom like 'DATACONTRACTEFIN%' "
    sql = sql & "and dependentesextes.id = d3.CODI) "
    sql = sql & "where case isnull(dependentesextes.valor,'') when '' then " & Di & " "
    sql = sql & "else  datepart(wk,convert(smalldatetime,dependentesextes.valor,103)) "
    sql = sql & "end  >= " & Di & " "
    sql = sql & "and "
    sql = sql & "case isnull(dependentesextes.valor,'') when '' then 3000 "
    sql = sql & "else  datepart(year,convert(smalldatetime,dependentesextes.valor,103)) "
    sql = sql & "end  >= " & Anoabuscar & " "
    'sql = sql & "select id From DependentesExtes with (nolock) Where "
    'sql = sql & "nom like 'DATACONTRACTEFIN%' "
    'sql = sql & "and case isnull(Valor,'') when '' then DATEPART( wk,cast('99991231 23:59:59.998' as datetime)) "
    'sql = sql & "Else "
    'sql = sql & "DATEPART( wk,cast(Valor as datetime)) "
    'sql = sql & "End "
    'sql = sql & "< " & Di & "  and case isnull(Valor,'') when '' then year(cast('99991231 23:59:59.998' as datetime)) "
    'sql = sql & "else year(cast(Valor as datetime)) "
    'sql = sql & "End"
    'sql = sql & "<= year(getdate()))"
    
    
    
    sql = sql & ") Group "
    'sql = sql & "By Dependentes.Codi, Day(Entrada), Dependentes.nom, DependentesExtes.Valor, DatePart(Weekday, Entrada), Left(Entrada, 11)"
    sql = sql & "By Dependentes.Codi, Day(Entrada), Dependentes.nom, DependentesExtes.Valor, equip, DatePart(Weekday, Entrada), Left(Entrada, 11)"
    sql = sql & ") pp Where DatePart(wk, Fecha) = " & Di & " and Equipo<>'' "
    If Eq > "0" Then sql = sql & "and equipo='" & Eq & "' "
    sql = sql & "order by Equipo,Usuario"
    Set Rs = Db.OpenResultset(sql)

    i = 0
    Fila = 1
    Col = 1

   
    'CABECERA ------------------------------------------------------
    'FILA 2
    Fila = 2
    Hoja.Cells(Fila, 2) = "Horas Plantilla"
    Hoja.Cells(Fila, 2).Font.Bold = True
    Hoja.Cells(Fila, 2).Font.Size = 20
                
    If UCase(EmpresaActual) = UCase("Tena") Then
        Hoja.Cells(Fila, 2) = "Horas Plantilla Silemabcn"
        Hoja.Cells(Fila, 2).HorizontalAlignment = xlCenter
        Hoja.Cells(Fila, 10).Select
        Dim RsIm As rdoResultset
        Set RsIm = Db.OpenResultset("Select top 1 archivo,extension from archivo where nombre = 'LOGO' and descripcion = '<0>' ")
''        MyKill "C:\Logo." & RsIm("extension")
'        ColumnToFile RsIm.rdoColumns("Archivo"), "c:\Logo." & RsIm("extension"), 102400, RsIm("Archivo").ColumnSize
 '       Hoja.Pictures.Insert("C:\Logo." & RsIm("extension")).Select
        'Hoja.Pictures.ShapeRange.ScaleWidth 0.65, False, 0
        'Hoja.Pictures.ShapeRange.ScaleHeight 0.65, False, 0
        'Hoja.Pictures.ShapeRange.IncrementTop -70
        'Hoja.Pictures.ShapeRange.IncrementLeft 1050
        ' esto ya estaba comentado 'Hoja.ShapeRange.ScaleHeight 0.4, False, 0
 ''       MyKill "C:\Logo." & RsIm("extension")
    End If
                
    Hoja.Range("B" & Fila & ":T" & Fila).MergeCells = True
    Hoja.Range("B" & Fila & ":T" & Fila).Interior.ColorIndex = 15
    Hoja.Range("B" & Fila & ":T" & Fila).Borders.Weight = xlMedium
                                
    'FILA 3
    Fila = 3
    Hoja.Cells(Fila, 1) = "Semana " & Di & " " & Anoabuscar
    Hoja.Cells(Fila, 1).Font.Bold = True
    Hoja.Cells(Fila, 1).Interior.ColorIndex = 15
    Hoja.Cells(Fila, 1).Borders.Weight = xlMedium
                
    Hoja.Cells(Fila, 2) = "Dias Del Mes"
    Hoja.Cells(Fila, 2).HorizontalAlignment = xlCenter
    Hoja.Cells(Fila, 2).Font.Bold = True
    Hoja.Range("B" & Fila & ":H" & Fila).MergeCells = True
    Hoja.Range("B" & Fila & ":H" & Fila).Borders.Weight = xlThin
                
    Hoja.Cells(3, 14) = "Dias Del Mes"
    Hoja.Cells(3, 14).HorizontalAlignment = xlCenter
    Hoja.Cells(3, 14).Font.Bold = True
    Hoja.Range("N" & Fila & ":T" & Fila).MergeCells = True
    Hoja.Range("N" & Fila & ":T" & Fila).Borders.Weight = xlThin
                
    'FILA 4
    Dim DidInici  As Date
    Fila = 4
    DidInici = DateAdd("ww", Di, DateSerial(Year(Now), 1, 1))
    If Di > 0 Then DidInici = DateAdd("ww", Di - 1, DateSerial(Year(Now), 1, 1))
    Hoja.Cells(Fila, 1) = "Empleados Por Seccion"
    Hoja.Cells(Fila, 1).HorizontalAlignment = xlCenter
    Hoja.Cells(Fila, 1).Font.Bold = True
    Hoja.Cells(Fila, 1).Borders.Weight = xlMedium
    Hoja.Cells(Fila, 2) = Format(DateAdd("d", -5, DidInici), "dd")
    Hoja.Cells(Fila, 2).Font.Bold = True
    Hoja.Cells(Fila, 3) = Format(DateAdd("d", -4, DidInici), "dd")
    Hoja.Cells(Fila, 3).Font.Bold = True
    Hoja.Cells(Fila, 4) = Format(DateAdd("d", -3, DidInici), "dd")
    Hoja.Cells(Fila, 4).Font.Bold = True
    Hoja.Cells(Fila, 5) = Format(DateAdd("d", -2, DidInici), "dd")
    Hoja.Cells(Fila, 5).Font.Bold = True
    Hoja.Cells(Fila, 6) = Format(DateAdd("d", -1, DidInici), "dd")
    Hoja.Cells(Fila, 6).Font.Bold = True
    Hoja.Cells(Fila, 7) = Format(DateAdd("d", 0, DidInici), "dd")
    Hoja.Cells(Fila, 7).Font.Bold = True
    Hoja.Cells(Fila, 8) = Format(DateAdd("d", 1, DidInici), "dd")
    Hoja.Cells(Fila, 8).Font.Bold = True
            
    Hoja.Range("B" & Fila - 1 & ":H" & Fila).Borders.Weight = xlThin
    Hoja.Range("B" & Fila - 1 & ":H" & Fila).Borders(xlEdgeLeft).Weight = xlMedium
    Hoja.Range("B" & Fila - 1 & ":H" & Fila).Borders(xlEdgeRight).Weight = xlMedium
    Hoja.Range("B" & Fila - 1 & ":H" & Fila).Borders(xlEdgeTop).Weight = xlMedium
    Hoja.Range("B" & Fila - 1 & ":H" & Fila).Borders(xlEdgeBottom).Weight = xlMedium
                
    Hoja.Cells(Fila, 14) = Format(DateAdd("d", -5, DidInici), "dd")
    Hoja.Cells(Fila, 14).Font.Bold = True
    Hoja.Cells(Fila, 15) = Format(DateAdd("d", -4, DidInici), "dd")
    Hoja.Cells(Fila, 15).Font.Bold = True
    Hoja.Cells(Fila, 16) = Format(DateAdd("d", -3, DidInici), "dd")
    Hoja.Cells(Fila, 16).Font.Bold = True
    Hoja.Cells(Fila, 17) = Format(DateAdd("d", -2, DidInici), "dd")
    Hoja.Cells(Fila, 17).Font.Bold = True
    Hoja.Cells(Fila, 18) = Format(DateAdd("d", -1, DidInici), "dd")
    Hoja.Cells(Fila, 18).Font.Bold = True
    Hoja.Cells(Fila, 19) = Format(DateAdd("d", 0, DidInici), "dd")
    Hoja.Cells(Fila, 19).Font.Bold = True
    Hoja.Cells(Fila, 20) = Format(DateAdd("d", 1, DidInici), "dd")
    Hoja.Cells(Fila, 20).Font.Bold = True
            
    Hoja.Range("N" & Fila - 1 & ":T" & Fila).Borders.Weight = xlThin
    Hoja.Range("N" & Fila - 1 & ":T" & Fila).Borders(xlEdgeLeft).Weight = xlMedium
    Hoja.Range("N" & Fila - 1 & ":T" & Fila).Borders(xlEdgeRight).Weight = xlMedium
    Hoja.Range("N" & Fila - 1 & ":T" & Fila).Borders(xlEdgeTop).Weight = xlMedium
    Hoja.Range("N" & Fila - 1 & ":T" & Fila).Borders(xlEdgeBottom).Weight = xlMedium
            
    Hoja.Range("I3:I4").MergeCells = True
    Hoja.Range("I3:I4").Borders.Weight = xlMedium
                
    Hoja.Range("M3:M4").MergeCells = True
    Hoja.Range("M3:M4").Borders.Weight = xlMedium

    Hoja.Range("U3:U4").MergeCells = True
    Hoja.Range("U3:U4").Borders.Weight = xlMedium
                
    Hoja.Cells(Fila, 10) = "Resumen Semanal"
    Hoja.Cells(Fila, 10).Font.Bold = True
    Hoja.Range("J3:L4").MergeCells = True
    Hoja.Cells(Fila, 10).HorizontalAlignment = xlCenter
    Hoja.Cells(Fila, 10).VerticalAlignment = xlVAlignCenter
    Hoja.Range("J3:L4").Borders.Weight = xlMedium
                
    Hoja.Cells(Fila, 22) = "Resumen Semanal"
    Hoja.Cells(Fila, 22).Font.Bold = True
    Hoja.Range("V3:X4").MergeCells = True
    Hoja.Cells(Fila, 22).HorizontalAlignment = xlCenter
    Hoja.Cells(Fila, 22).VerticalAlignment = xlVAlignCenter
    Hoja.Range("V3:X4").Borders.Weight = xlMedium
                
    '~CABECERA ------------------------------------------------------
        
    Dim filaIniEquipo As Integer
    Dim filaFinEquipo As Integer
        
    Dim filasTotales() As Integer
    Dim nEquipos As Integer
    nEquipos = 0
    
    Fila = 5
    
    Equipo = ""
    usuario = ""
    filaIniEquipo = 6
    While Not Rs.EOF
        'Si llegan equipos con coma se escoge solo el primero, el resto se descarta
        equipoAct = Rs("Equipo")
        coma = InStr(1, equipoAct, ",")
        If coma > 0 Then
            equipoAct = Trim(Mid(equipoAct, 1, coma - 1))
        End If
        If Equipo <> equipoAct Then
            If Equipo <> "" Then
               'Total por usuario último usuario del equipo
                Hoja.Cells(Fila, 9).FormulaR1C1 = "=SUM(R" & Fila & "C2:R" & Fila & "C8)"
                'Condición de color
                Hoja.Cells(Fila, 9).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=40"
                Hoja.Cells(Fila, 9).FormatConditions(1).Font.ColorIndex = 3
                
                'Coste
                Hoja.Cells(Fila, 13).FormulaR1C1 = 10.5
                'Calculo de coste (valor * coste)
                For Col = 2 To 8
                    Hoja.Cells(Fila, Col + 12).FormulaR1C1 = "=R" & Fila & "C" & Col & "*R" & Fila & "C13"
                    Hoja.Cells(Fila, Col + 12).HorizontalAlignment = xlRight
                    Hoja.Cells(Fila, Col + 12).Font.ColorIndex = 0
                Next
                'Total costes
                Hoja.Cells(Fila, 21).FormulaR1C1 = "=SUM(R" & Fila & "C14:R" & Fila & "C20)"
                
                filaFinEquipo = Fila
                
                Fila = Fila + 1

                usuario = Rs("Usuario")
                CodUsuario = Rs(4)
            
                'Totales por equipo
                Hoja.Cells(Fila, 1) = "Totales Diarios"

                Hoja.Range("A" & Fila & ":X" & Fila & "").Interior.ColorIndex = 35
                Hoja.Range("A" & Fila & ":X" & Fila & "").Font.Bold = True
                
                For Col = 2 To 24
                    Hoja.Cells(Fila, Col).FormulaR1C1 = "=SUM(R" & filaIniEquipo & "C" & Col & ":R" & filaFinEquipo & "C" & Col & ")"
                    Hoja.Cells(Fila, Col + 12).HorizontalAlignment = xlRight
                    Hoja.Cells(Fila, Col + 12).Font.ColorIndex = 0
                Next
                
                ReDim Preserve filasTotales(nEquipos)
                filasTotales(nEquipos) = Fila
                nEquipos = nEquipos + 1
                
                'Formato de los bordes
                Hoja.Range("A" & filaIniEquipo - 1 & ":X" & filaFinEquipo + 1 & "").Borders.Weight = xlThin
                Hoja.Range("A" & filaIniEquipo - 1 & ":H" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
                Hoja.Range("A" & filaIniEquipo - 1 & ":H" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
                Hoja.Range("A" & filaIniEquipo - 1 & ":H" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
                Hoja.Range("A" & filaIniEquipo - 1 & ":H" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
                
                Hoja.Range("I" & filaIniEquipo - 1 & ":I" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
                Hoja.Range("I" & filaIniEquipo - 1 & ":I" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
                Hoja.Range("I" & filaIniEquipo - 1 & ":I" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
                Hoja.Range("I" & filaIniEquipo - 1 & ":I" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
                
                Hoja.Range("J" & filaIniEquipo - 1 & ":L" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
                Hoja.Range("J" & filaIniEquipo - 1 & ":L" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
                Hoja.Range("J" & filaIniEquipo - 1 & ":L" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
                Hoja.Range("J" & filaIniEquipo - 1 & ":L" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
                
                Hoja.Range("M" & filaIniEquipo - 1 & ":M" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
                Hoja.Range("M" & filaIniEquipo - 1 & ":M" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
                Hoja.Range("M" & filaIniEquipo - 1 & ":M" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
                Hoja.Range("M" & filaIniEquipo - 1 & ":M" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
                
                Hoja.Range("N" & filaIniEquipo - 1 & ":T" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
                Hoja.Range("N" & filaIniEquipo - 1 & ":T" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
                Hoja.Range("N" & filaIniEquipo - 1 & ":T" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
                Hoja.Range("N" & filaIniEquipo - 1 & ":T" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
                
                Hoja.Range("U" & filaIniEquipo - 1 & ":U" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
                Hoja.Range("U" & filaIniEquipo - 1 & ":U" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
                Hoja.Range("U" & filaIniEquipo - 1 & ":U" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
                Hoja.Range("U" & filaIniEquipo - 1 & ":U" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium

                Hoja.Range("V" & filaIniEquipo - 1 & ":X" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
                Hoja.Range("V" & filaIniEquipo - 1 & ":X" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
                Hoja.Range("V" & filaIniEquipo - 1 & ":X" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
                Hoja.Range("V" & filaIniEquipo - 1 & ":X" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
                
                Fila = Fila + 1
            End If
            
            'Cabecera de equipo
            Hoja.Cells(Fila, 1) = equipoAct
            Hoja.Cells(Fila, 1).HorizontalAlignment = xlCenter
            Hoja.Cells(Fila, 1).Font.Size = 10
            Hoja.Cells(Fila, 1).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 2) = "Lunes"
            Hoja.Cells(Fila, 2).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 3) = "Martes"
            Hoja.Cells(Fila, 3).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 4) = "Miércoles"
            Hoja.Cells(Fila, 4).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 5) = "Jueves"
            Hoja.Cells(Fila, 5).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 6) = "Viernes"
            Hoja.Cells(Fila, 6).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 7) = "Sábado"
            Hoja.Cells(Fila, 7).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 8) = "Domingo"
            Hoja.Cells(Fila, 8).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 9) = "TOTAL"
            Hoja.Cells(Fila, 9).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 10) = "Sem. " & Di - 3
            Hoja.Cells(Fila, 10).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 11) = "Sem. " & Di - 2
            Hoja.Cells(Fila, 11).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 12) = "Sem. " & Di - 1
            Hoja.Cells(Fila, 12).Interior.ColorIndex = 15
            
            Hoja.Cells(Fila, 13) = "Coste"
            Hoja.Cells(Fila, 13).HorizontalAlignment = xlCenter
            Hoja.Cells(Fila, 13).Font.Size = 10
            Hoja.Cells(Fila, 13).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 14) = "Lunes"
            Hoja.Cells(Fila, 14).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 15) = "Martes"
            Hoja.Cells(Fila, 15).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 16) = "Miércoles"
            Hoja.Cells(Fila, 16).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 17) = "Jueves"
            Hoja.Cells(Fila, 17).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 18) = "Viernes"
            Hoja.Cells(Fila, 18).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 19) = "Sábado"
            Hoja.Cells(Fila, 19).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 20) = "Domingo"
            Hoja.Cells(Fila, 20).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 21) = "TOTAL"
            Hoja.Cells(Fila, 21).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 22) = "Sem. " & Di - 3
            Hoja.Cells(Fila, 22).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 23) = "Sem. " & Di - 2
            Hoja.Cells(Fila, 23).Interior.ColorIndex = 15
            Hoja.Cells(Fila, 24) = "Sem. " & Di - 1
            Hoja.Cells(Fila, 24).Interior.ColorIndex = 15
            
            Hoja.Range("A" & Fila & ":X" & Fila & "").Font.Bold = True
            
            Fila = Fila + 1
            filaIniEquipo = Fila
        End If
        If usuario <> Rs("Usuario") Then
            If usuario <> "" Then
                'Total por usuario
                Hoja.Cells(Fila, 9).FormulaR1C1 = "=SUM(R" & Fila & "C2:R" & Fila & "C8)"
                'Condición de color
                Hoja.Cells(Fila, 9).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=40"
                Hoja.Cells(Fila, 9).FormatConditions(1).Font.ColorIndex = 3
'TOTALS SETMANES
                ExcelHores2HoresSetmanesAnt CodUsuario, Di, H1, H2, H3, Anoabuscar, Equipo
                
                Hoja.Cells(Fila, 10) = digital2(H1)
                Hoja.Cells(Fila, 11) = digital2(H2)
                Hoja.Cells(Fila, 12) = digital2(H3)


                'Coste
                Hoja.Cells(Fila, 13).FormulaR1C1 = 10.5
' TOTAL COSTES SEMANES
                Hoja.Cells(Fila, 22) = "=J" & Fila & "*M" & Fila
                Hoja.Cells(Fila, 23) = "=L" & Fila & "*M" & Fila
                Hoja.Cells(Fila, 24) = "=L" & Fila & "*M" & Fila
' TOTAL COSTES SEMANES
                'Calculo de coste (valor * coste)
                For Col = 2 To 8
                    Hoja.Cells(Fila, Col + 12).FormulaR1C1 = "=R" & Fila & "C" & Col & "*R" & Fila & "C13"
                    Hoja.Cells(Fila, Col + 12).HorizontalAlignment = xlRight
                    Hoja.Cells(Fila, Col + 12).Font.ColorIndex = 0
                Next
                'Total costes
                Hoja.Cells(Fila, 21).FormulaR1C1 = "=SUM(R" & Fila & "C14:R" & Fila & "C20)"
                Fila = Fila + 1
            End If
        End If
        Hoja.Cells(Fila, 1) = Rs("Usuario")
        
        Col = Rs("diasemana") + 1
        
        If Rs("Usuario") = "Concepción Escuder Moya" Or Rs("Usuario") = "Eva María Fernández Llagostera" Then
            Hoja.Cells(Fila, 2).FormulaR1C1 = 8
            Hoja.Cells(Fila, 3).FormulaR1C1 = 8
            Hoja.Cells(Fila, 4).FormulaR1C1 = 8
            Hoja.Cells(Fila, 5).FormulaR1C1 = 8
            Hoja.Cells(Fila, 6).FormulaR1C1 = 8
        Else
            Hoja.Cells(Fila, Col).FormulaR1C1 = digital2(Rs("HorasAcumulado"))
            Hoja.Cells(Fila, Col).Font.ColorIndex = 0
            If Rs("HorasAcumulado") > (60 * 8) Then Hoja.Cells(Fila, Col).Font.ColorIndex = 3
        End If
        Hoja.Cells(Fila, Col).HorizontalAlignment = xlRight
                    
        Equipo = equipoAct
        usuario = Rs("Usuario")
        CodUsuario = Rs(4)
        Rs.MoveNext
    Wend
    
    'Último equipo
    If Equipo <> "" Then
        'Total por usuario último usuario del equipo
        Hoja.Cells(Fila, 9).FormulaR1C1 = "=SUM(R" & Fila & "C2:R" & Fila & "C8)"
        'Condición de color
        Hoja.Cells(Fila, 9).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=40"
        Hoja.Cells(Fila, 9).FormatConditions(1).Font.ColorIndex = 3

        'Coste
        Hoja.Cells(Fila, 13).FormulaR1C1 = 10.5
        'Calculo de coste (valor * coste)
        For Col = 2 To 8
            Hoja.Cells(Fila, Col + 12).FormulaR1C1 = "=R" & Fila & "C" & Col & "*R" & Fila & "C13"
            Hoja.Cells(Fila, Col + 12).HorizontalAlignment = xlRight
            Hoja.Cells(Fila, Col + 12).Font.ColorIndex = 0
        Next
        'Total costes
        Hoja.Cells(Fila, 21).FormulaR1C1 = "=SUM(R" & Fila & "C14:R" & Fila & "C20)"
                
        filaFinEquipo = Fila
                
        Fila = Fila + 1

        'Totales por equipo
        Hoja.Cells(Fila, 1) = "Totales Diarios"

        Hoja.Range("A" & Fila & ":X" & Fila & "").Interior.ColorIndex = 35
        Hoja.Range("A" & Fila & ":X" & Fila & "").Font.Bold = True
                
        For Col = 2 To 24
            Hoja.Cells(Fila, Col).FormulaR1C1 = "=SUM(R" & filaIniEquipo & "C" & Col & ":R" & filaFinEquipo & "C" & Col & ")"
            Hoja.Cells(Fila, Col + 12).HorizontalAlignment = xlRight
            Hoja.Cells(Fila, Col + 12).Font.ColorIndex = 0
        Next
                
        ReDim Preserve filasTotales(nEquipos)
        filasTotales(nEquipos) = Fila
        nEquipos = nEquipos + 1
                
        'Formato de los bordes
        Hoja.Range("A" & filaIniEquipo - 1 & ":X" & filaFinEquipo + 1 & "").Borders.Weight = xlThin
        Hoja.Range("A" & filaIniEquipo - 1 & ":H" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
        Hoja.Range("A" & filaIniEquipo - 1 & ":H" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
        Hoja.Range("A" & filaIniEquipo - 1 & ":H" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
        Hoja.Range("A" & filaIniEquipo - 1 & ":H" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
                
        Hoja.Range("I" & filaIniEquipo - 1 & ":I" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
        Hoja.Range("I" & filaIniEquipo - 1 & ":I" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
        Hoja.Range("I" & filaIniEquipo - 1 & ":I" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
        Hoja.Range("I" & filaIniEquipo - 1 & ":I" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
                
        Hoja.Range("J" & filaIniEquipo - 1 & ":L" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
        Hoja.Range("J" & filaIniEquipo - 1 & ":L" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
        Hoja.Range("J" & filaIniEquipo - 1 & ":L" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
        Hoja.Range("J" & filaIniEquipo - 1 & ":L" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
               
        Hoja.Range("M" & filaIniEquipo - 1 & ":M" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
        Hoja.Range("M" & filaIniEquipo - 1 & ":M" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
        Hoja.Range("M" & filaIniEquipo - 1 & ":M" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
        Hoja.Range("M" & filaIniEquipo - 1 & ":M" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
                
        Hoja.Range("N" & filaIniEquipo - 1 & ":T" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
        Hoja.Range("N" & filaIniEquipo - 1 & ":T" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
        Hoja.Range("N" & filaIniEquipo - 1 & ":T" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
        Hoja.Range("N" & filaIniEquipo - 1 & ":T" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
                
        Hoja.Range("U" & filaIniEquipo - 1 & ":U" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
        Hoja.Range("U" & filaIniEquipo - 1 & ":U" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
        Hoja.Range("U" & filaIniEquipo - 1 & ":U" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
        Hoja.Range("U" & filaIniEquipo - 1 & ":U" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium

        Hoja.Range("V" & filaIniEquipo - 1 & ":X" & filaFinEquipo + 1 & "").Borders(xlEdgeLeft).Weight = xlMedium
        Hoja.Range("V" & filaIniEquipo - 1 & ":X" & filaFinEquipo + 1 & "").Borders(xlEdgeRight).Weight = xlMedium
        Hoja.Range("V" & filaIniEquipo - 1 & ":X" & filaFinEquipo + 1 & "").Borders(xlEdgeTop).Weight = xlMedium
        Hoja.Range("V" & filaIniEquipo - 1 & ":X" & filaFinEquipo + 1 & "").Borders(xlEdgeBottom).Weight = xlMedium
    End If
    
    'Total plantilla
    Fila = Fila + 1
    Hoja.Cells(Fila, 1) = "Total Diario Plantilla"
    Hoja.Cells(Fila, 1).HorizontalAlignment = xlCenter
    Hoja.Cells(Fila, 1).Font.Bold = True
    Hoja.Range("A" & Fila & ":X" & Fila).Interior.ColorIndex = 36
   
    Dim FormulaTotal As String
    For j = 2 To 24
        FormulaTotal = ""
        For i = 0 To nEquipos - 1
            If FormulaTotal <> "" Then FormulaTotal = FormulaTotal & ","
            FormulaTotal = FormulaTotal & "R" & filasTotales(i) & "C" & j
        Next
        If FormulaTotal <> "" Then
            Hoja.Cells(Fila, j).FormulaR1C1 = "=SUM(" & FormulaTotal & ")"
            Hoja.Cells(Fila, j).NumberFormat = "0.00"
            Hoja.Cells(Fila, j).HorizontalAlignment = xlRight
            Hoja.Cells(Fila, j).Font.Bold = True
        End If
    Next
    
    Hoja.Range("A" & Fila & ":X" & Fila).Borders.Weight = xlThin
    Hoja.Range("A" & Fila & ":X" & Fila).Borders(xlEdgeLeft).Weight = xlMedium
    Hoja.Range("A" & Fila & ":X" & Fila).Borders(xlEdgeRight).Weight = xlMedium
    Hoja.Range("A" & Fila & ":X" & Fila).Borders(xlEdgeTop).Weight = xlMedium
    Hoja.Range("A" & Fila & ":X" & Fila).Borders(xlEdgeBottom).Weight = xlMedium
    
    Hoja.Select
    Hoja.Cells.EntireColumn.AutoFit
    Hoja.Range("A1").Select
    
    Informa2 ""
End Sub



Sub ExcelFacturacioAnual(ByRef Hoja As Excel.Worksheet, ano)
    Dim sql As String, article As String, client As String, articleAct As String
    Dim aux As Date
    Dim m As Integer, i As Integer, Fila As Integer, Col As Integer, c As Integer, mesAct As Integer
    Dim rsCli As rdoResultset, rsFact As rdoResultset
    Dim CliArr
    'Dim i As Integer, j As Integer, DiaS As Integer, rs As rdoResultset, rsDades As ADODB.Recordset, Equipo, Usuario, Hacc As Double, HaccEq As Double, HaccEqDia(10) As Double, HaccEqDiaTot(10) As Double, HaccEqNumOp, HaccEqNumOpTot, K, H1, H2, H3, TsA1, TsA2, TsA3, TsAGr1, TsAGr2, TsAGr3
    'Dim rsId As ADODB.Recordset, rsIns As rdoResultset, data() As Double, mes As Integer, HEquipSetm As Double, Rs2 As rdoResultset
    'Dim D As Date, diff As Integer, Fila As Integer, fecha, Accio, AccioPrev, equip, Usuari, idTemp, Hores, DiaEnt, DiaSal, Anoabuscar
    'Dim EquipPrev, TotalDia, DiaPrev, Col, TotalTreb, Semana, SemanaAnt, SemanaTemp, SemanaTemp2, RestaSem, Seg, Min, Semanas(5), ArraySemana(2)
    'Dim CodUsuario As String, equipoAct As String, coma As Integer

    InformaMiss "Calculs Excel Facturacio anual"
    
    ExecutaComandaSql "drop table [GRID_FACTURACIO]"

    sql = "CREATE TABLE [GRID_FACTURACIO]  ( "
    sql = sql & "[ProducteNom] [nvarchar] (255) , "
    sql = sql & "[ClientNom] [nvarchar] (255) , "
    sql = sql & "[Servit] [float] Default (0), "
    sql = sql & "[Import] [float] Default (0), "
    sql = sql & "[Mes] [integer]  "
    sql = sql & ") ON [PRIMARY] "
    ExecutaComandaSql sql

  
    aux = DateSerial(ano, 1, 1)
    m = 1
    While m <= 12 And aux <= Now()
        sql = "insert into [GRID_FACTURACIO] "
        sql = sql & "select ProducteNom, ClientNom, sum(Servit-Tornat) Servit, sum(Import) Import, " & m & " Mes "
        sql = sql & "from [" & NomTaulaFacturaIva(aux) & " ]iva "
        sql = sql & "left join [" & NomTaulaFacturaData(aux) & "] data on iva.IdFactura = data.IdFactura "
        sql = sql & "group by ProducteNom, ClientNom"
        ExecutaComandaSql sql
        
        m = m + 1
        aux = DateSerial(ano, m, 1)
    Wend

    Set rsCli = Db.OpenResultset("select distinct(ClientNom) from [GRID_FACTURACIO] order by ClientNom")
    ReDim CliArr(0)
    CliArr(0) = ""
    i = 1
    While Not rsCli.EOF
        ReDim Preserve CliArr(i)
        CliArr(i) = rsCli("ClientNom")
        i = i + 1
        rsCli.MoveNext
    Wend

    Hoja.Cells(1, 1) = "Productos/Clientes"
    Hoja.Cells(1, 2) = "Mes"
    
    For i = 1 To UBound(CliArr)
        Hoja.Cells(1, i + (i + 1)) = CliArr(i)
        Hoja.Cells(2, i + (i + 1)) = "Servido"
        Hoja.Cells(2, i + (i + 2)) = "Importe"
    Next
    
    Fila = 3
    Col = 1
    
    article = ""
    client = ""
    Set rsFact = Db.OpenResultset("select * from [GRID_FACTURACIO] order by ProducteNom, Mes, ClientNom")
    
    While Not rsFact.EOF
        If rsFact("ProducteNom") <> article Then
            If article <> "" Then Fila = Fila + 12
            
            Hoja.Cells(Fila, 1) = rsFact("ProducteNom")
            Hoja.Cells(Fila, 2) = "Enero"
            Hoja.Cells(Fila + 1, 2) = "Febrero"
            Hoja.Cells(Fila + 2, 2) = "Marzo"
            Hoja.Cells(Fila + 3, 2) = "Abril"
            Hoja.Cells(Fila + 4, 2) = "Mayo"
            Hoja.Cells(Fila + 5, 2) = "Junio"
            Hoja.Cells(Fila + 6, 2) = "Julio"
            Hoja.Cells(Fila + 7, 2) = "Agosto"
            Hoja.Cells(Fila + 8, 2) = "Septiembre"
            Hoja.Cells(Fila + 9, 2) = "Octubre"
            Hoja.Cells(Fila + 10, 2) = "Noviembre"
            Hoja.Cells(Fila + 11, 2) = "Diciembre"
            
        End If
                
        article = rsFact("ProducteNom")
        For c = 1 To UBound(CliArr)
            If Not rsFact.EOF Then
                If rsFact("ClientNom") = CliArr(c) And article = rsFact("ProducteNom") Then
                    If Not rsFact.EOF Then
                        client = rsFact("ClientNom")
                        articleAct = rsFact("ProducteNom")
                        mesAct = rsFact("mes")

                        While client = CliArr(c) And article = articleAct
                            Hoja.Cells(Fila + (mesAct - 1), (c * 2) + 1) = rsFact("Servit")
                            Hoja.Cells(Fila + (mesAct - 1), (c * 2) + 2) = rsFact("Import")
                            article = rsFact("ProducteNom")

                            rsFact.MoveNext
                            If Not rsFact.EOF Then
                                client = rsFact("ClientNom")
                                articleAct = rsFact("ProducteNom")
                                mesAct = rsFact("mes")
                            Else
                                client = ""
                                articleAct = ""
                            End If
                         Wend
                    End If
                End If
            End If
        Next
    Wend
    
    
    
    Hoja.Select
    Hoja.Cells.EntireColumn.AutoFit
    Hoja.Range("A1").Select
    
End Sub


Sub ExcelFacturacioAnualAcc(ByRef Hoja As Excel.Worksheet, ano)
    Dim sql As String, article As String, client As String, articleAct As String
    Dim aux As Date
    Dim m As Integer, i As Integer, Fila As Integer, Col As Integer, c As Integer, mesAct As Integer
    Dim rsCli As rdoResultset, rsFact As rdoResultset
    Dim CliArr, ProducteNom, servit, import
    
    'Dim i As Integer, j As Integer, DiaS As Integer, rs As rdoResultset, rsDades As ADODB.Recordset, Equipo, Usuario, Hacc As Double, HaccEq As Double, HaccEqDia(10) As Double, HaccEqDiaTot(10) As Double, HaccEqNumOp, HaccEqNumOpTot, K, H1, H2, H3, TsA1, TsA2, TsA3, TsAGr1, TsAGr2, TsAGr3
    'Dim rsId As ADODB.Recordset, rsIns As rdoResultset, data() As Double, mes As Integer, HEquipSetm As Double, Rs2 As rdoResultset
    'Dim D As Date, diff As Integer, Fila As Integer, fecha, Accio, AccioPrev, equip, Usuari, idTemp, Hores, DiaEnt, DiaSal, Anoabuscar
    'Dim EquipPrev, TotalDia, DiaPrev, Col, TotalTreb, Semana, SemanaAnt, SemanaTemp, SemanaTemp2, RestaSem, Seg, Min, Semanas(5), ArraySemana(2)
    'Dim CodUsuario As String, equipoAct As String, coma As Integer

    InformaMiss "Calculs Excel Facturacio anual Acc"
    ano = CVDate(ano)
    ano = Year(ano)
    
    ExecutaComandaSql "drop table [GRID_FACTURACIO]"

    sql = "CREATE TABLE [GRID_FACTURACIO]  ( "
    sql = sql & "[ProducteNom] [nvarchar] (255) , "
    sql = sql & "[ClientNom] [nvarchar] (255) , "
    sql = sql & "[Servit] [float] Default (0), "
    sql = sql & "[Import] [float] Default (0), "
    sql = sql & "[Mes] [integer]  "
    sql = sql & ") ON [PRIMARY] "
    ExecutaComandaSql sql

  
    aux = DateSerial(ano, 1, 1)
    m = 1
    While m <= 12 And aux <= Now()
        sql = "insert into [GRID_FACTURACIO] "
        sql = sql & "select ProducteNom, ClientNom, sum(Servit-Tornat) Servit, sum(Import) Import, " & m & " Mes "
        sql = sql & "from [" & NomTaulaFacturaIva(aux) & " ]iva "
        sql = sql & "left join [" & NomTaulaFacturaData(aux) & "] data on iva.IdFactura = data.IdFactura "
        sql = sql & "group by ProducteNom, ClientNom"
        ExecutaComandaSql sql
        
        m = m + 1
        aux = DateSerial(ano, m, 1)
    Wend

    Set rsCli = Db.OpenResultset("select distinct(ClientNom) from [GRID_FACTURACIO] order by ClientNom")
    ReDim CliArr(0)
    CliArr(0) = ""
    i = 1
    While Not rsCli.EOF
        ReDim Preserve CliArr(i)
        CliArr(i) = rsCli("ClientNom")
        i = i + 1
        rsCli.MoveNext
    Wend

    Hoja.Cells(1, 1) = "Productos/Clientes"
    Hoja.Cells(1, 2) = "Mes"
    
    For i = 1 To UBound(CliArr)
        Hoja.Cells(1, i + (i + 1)) = CliArr(i)
        Hoja.Cells(2, i + (i + 1)) = "Servido"
        Hoja.Cells(2, i + (i + 2)) = "Importe"
    Next
    
    Fila = 3
    Col = 1
    
    article = ""
    client = ""
    Set rsFact = Db.OpenResultset("select ProducteNom, ClientNom,SUM(servit) servit ,sum(import) import   from [GRID_FACTURACIO]  group by ProducteNom, ClientNom order by ProducteNom, ClientNom ")
    
    While Not rsFact.EOF
        ProducteNom = Left(rsFact("ProducteNom"), 40)
        
        If ProducteNom <> article Then
            If article <> "" Then Fila = Fila + 1
            Hoja.Cells(Fila, 1) = ProducteNom
            Hoja.Cells(Fila, 2) = "Total"
        End If
        article = ProducteNom
        For c = 1 To UBound(CliArr)
            If Not rsFact.EOF Then
                If rsFact("ClientNom") = CliArr(c) And article = ProducteNom Then
                    If Not rsFact.EOF Then
                        client = rsFact("ClientNom")
                        articleAct = ProducteNom
                        mesAct = 1

                        While client = CliArr(c) And article = articleAct
                            servit = rsFact("Servit")
                            import = rsFact("Import")
                            Hoja.Cells(Fila + (mesAct - 1), (c * 2) + 1) = servit
                            Hoja.Cells(Fila + (mesAct - 1), (c * 2) + 2) = import
                            article = ProducteNom

                            rsFact.MoveNext
                            If Not rsFact.EOF Then
                                client = rsFact("ClientNom")
                                articleAct = ProducteNom
                                mesAct = 1
                            Else
                                client = ""
                                articleAct = ""
                            End If
                         Wend
                    End If
                End If
            End If
        Next
    Wend
    
    
    
    Hoja.Select
    Hoja.Cells.EntireColumn.AutoFit
    Hoja.Range("A1").Select
    
End Sub




Function DonamValorNumeric(sql As String, def As Double) As Double
    Dim Rss As rdoResultset
    DonamValorNumeric = def
On Error GoTo nor
    Set Rss = Db.OpenResultset(sql)
    If Not Rss.EOF Then If Not IsNull(Rss(0)) Then If IsNumeric(Rss(0)) Then DonamValorNumeric = Rss(0)
nor:

End Function

Sub rellenaHojaDiaDeLaSetmanaBuscaDades(D As Date, client, cM, cT, zM, zT, hM, hT, DepsM(), DepsT(), DescM, DescT, Families(), FamiliesPct(), Devol, servit)
Dim Rs As rdoResultset, Dz As Date, Deps(), DepsEstat(), DepsHi(), DepsHm(), DepsHt(), i As Integer, H
    Dim Huv, DepsHistoric() As String
    
    'Devol = 0
'    Set Rs = Db.OpenResultset("Select Sum(Import) From [" & NomTaulaDevol(D) & "] Where Botiga = " & Client & " And day(data) = " & Day(D) & " And tipus_venta = 'S' ")
    'If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Devol = Round(Rs(0), 2)
    
    Devol = 0
    Set Rs = Db.OpenResultset("Select Sum(quantitattornada * preu) From [" & DonamNomTaulaServit(D) & "] S join articles a on s.codiarticle = a.codi  Where client = " & client)
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Devol = Round(Rs(0), 2)
    
    servit = 0
    Set Rs = Db.OpenResultset("Select Sum(quantitatservida * preu) From [" & DonamNomTaulaServit(D) & "] S join articles a on s.codiarticle = a.codi  Where client = " & client)
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then servit = Round(Rs(0), 2)
    
    Dz = DateSerial(Year(D), Month(D), Day(D)) + TimeSerial(23, 55, 55)
    Set Rs = Db.OpenResultset("Select min(data) From [" & NomTaulaMovi(D) & "] Where tipus_moviment = 'Z' And Botiga = " & client & " And day(data) = " & Day(D))
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Dz = Rs(0)
    
    DescM = 0
    Set Rs = Db.OpenResultset("Select sum(import) From [" & NomTaulaMovi(D) & "] Where tipus_moviment = 'J' And Botiga = " & client & " And data <= convert(datetime,'" & Dz & "') and day(data) = " & Day(Dz) & " ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then DescM = Rs(0)
    
    DescT = 0
    Set Rs = Db.OpenResultset("Select sum(import) From [" & NomTaulaMovi(D) & "] Where tipus_moviment = 'J' And Botiga = " & client & " And data > convert(datetime,'" & Dz & "') and day(data) = " & Day(Dz) & " ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then DescT = Rs(0)
    
    Huv = Dz
    Set Rs = Db.OpenResultset("Select max(data) From [" & NomTaulaVentas(D) & "] Where Botiga = " & client & " And day(data) = " & Day(D) & "  ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Huv = Rs(0)
    
    
    Set Rs = Db.OpenResultset("Select count(distinct num_tick) , sum(import) From [" & NomTaulaVentas(D) & "] Where Botiga = " & client & " And day(data) = " & Day(D) & " And data < convert(datetime,'" & Dz & "') ")
    cM = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then cM = Rs(0)
    zM = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then zM = Rs(1)
    
    Set Rs = Db.OpenResultset("Select count(distinct num_tick) , sum(import) From [" & NomTaulaVentas(D) & "] Where Botiga = " & client & " And day(data) = " & Day(D) & " And data > convert(datetime,'" & Dz & "') ")
    cT = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then cT = Rs(0)
    zT = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then zT = Rs(1)
    
    
    Set Rs = Db.OpenResultset("Select Memo Nom,data,operacio from [" & NomTaulaHoraris(D) & "] join dependentes on codi = dependenta where botiga = " & client & " and day(data) = " & Day(D) & " order by data ")
    
    ReDim Deps(0)
    ReDim DepsEstat(0)
    ReDim DepsHi(0)
    ReDim DepsHm(0)
    ReDim DepsHt(0)
    ReDim DepsHistoric(0)
    
    While Not Rs.EOF
        For i = 1 To UBound(Deps)
            If Deps(i) = Rs("Nom") Then Exit For
        Next
        If i > UBound(Deps) Then
            ReDim Preserve Deps(i)
            ReDim Preserve DepsEstat(i)
            ReDim Preserve DepsHi(i)
            ReDim Preserve DepsHm(i)
            ReDim Preserve DepsHt(i)
            ReDim Preserve DepsHistoric(i)
            Deps(i) = Rs("Nom")
        End If
        DepsHistoric(i) = DepsHistoric(i) & " " & UCase(Rs("Operacio")) & "" & Hour(Rs("Data")) & ":" & Minute(Rs("Data"))
        If Rs("Operacio") = "E" Then
           DepsHi(i) = Rs("Data")
        Else
           If DepsHi(i) < Dz Then
              If Rs("Data") < DateAdd("n", 20, Dz) Then
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Rs("Data"))
                 DepsHi(i) = Rs("Data")
              Else
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Dz)
                 DepsHi(i) = Dz
              End If
           End If
           
           If Rs("Data") > Dz Then
              If Rs("Data") > DateAdd("n", 20, Dz) Then
                 DepsHt(i) = DepsHt(i) + DateDiff("n", DepsHi(i), Rs("Data"))
                 DepsHi(i) = Rs("Data")
              End If
           End If
        End If
        
        DepsEstat(i) = Rs("Operacio")
        Rs.MoveNext
    Wend
    For i = 1 To UBound(Deps)
        If DepsEstat(i) = "E" Then
           If DepsHi(i) < Dz Then
              If Huv < DateAdd("n", 20, Dz) Then
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Huv)
                 DepsHi(i) = Huv
              Else
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Dz)
                 DepsHi(i) = Dz
              End If
           End If
           If Huv > Dz Then
              If Huv > DateAdd("n", 20, Dz) Then
                 DepsHt(i) = DepsHt(i) + DateDiff("n", DepsHi(i), Huv)
                 DepsHi(i) = Huv
              End If
           End If
       End If
    Next
    
    
    ReDim DepsT(0)
    ReDim DepsM(0)
    hM = 0
    hT = 0
    For i = 1 To UBound(Deps)
        If DepsHt(i) > 0 Then
            ReDim Preserve DepsT(UBound(DepsT) + 1)
            DepsT(UBound(DepsT)) = Deps(i) & DepsHistoric(i) & "(" & digital(DepsHt(i)) & ")"
            hT = hT + DepsHt(i)
        End If
        If DepsHm(i) > 0 Then
            ReDim Preserve DepsM(UBound(DepsM) + 1)
            DepsM(UBound(DepsM)) = Deps(i) & DepsHistoric(i) & " (" & digital(DepsHm(i)) & ")"
            hM = hM + DepsHm(i)
        End If
    Next
    
    For i = 0 To UBound(Families)
        FamiliesPct(i) = 0
    Next
    
    Set Rs = Db.OpenResultset("Select ff.pare,sum(import) from [" & NomTaulaVentas(D) & "] v join articles a on a.codi = v.plu  join families f on f.nom=a.familia  join families ff on f.pare=ff.nom Where Botiga = " & client & " And day(data) = " & Day(D) & " group by ff.pare order by ff.pare ")
    While Not Rs.EOF
        For i = 0 To UBound(Families)
            If Families(i) = Rs(0) Then Exit For
        Next
        If i > UBound(Families) Then
            i = UBound(Families) + 1
            ReDim Preserve Families(i)
            ReDim Preserve FamiliesPct(i)
            Families(i) = Rs(0)
            FamiliesPct(i) = ""
        End If
        If zM + zT > 0 Then FamiliesPct(i) = Rs(1) / (zM + zT) * 100
        Rs.MoveNext
    Wend

End Sub

Sub rellenaHojaDiaDeLaSetmanaBuscaDades2(D As Date, client, cM, cT, zM, zT, hM, hT, tM, tMI, tT, tTI, Deps(), DepsM(), DepsT(), DescM, DescT, Families(), FamiliesPct(), Devol, servit, DevolF, ServitF, DevolFIVA, ServitFIVA, NetofIVA, NetoPor)
    Dim Rs As rdoResultset, Dz As Date, DepsEstat(), DepsHi(), DepsHm(), DepsHt(), i As Integer, H
    Dim Huv, DepsHistoric() As String
    Dim rsDto As ADODB.Recordset, DesconteTipus, Dpp, tipoPreu, f As Integer, p As Integer, DescTe, descuento
    Dim fam1, fam2, fam3, famD, prodD, import, TImportQT, TImportQS, sql, art
    Dim entra, sale, rsD As ADODB.Recordset, rsH As ADODB.Recordset, rsE As ADODB.Recordset, TTrab, TTipoTrab
    Dim tipoIva, ImpostTornat, ImpoTornat, ImpostServit, ImpoServit
    
    'Devol = 0
    'Set Rs = Db.OpenResultset("Select Sum(Import) From [" & NomTaulaDevol(D) & "] Where Botiga = " & Client & " And day(data) = " & Day(D) & " And tipus_venta = 'S' ")
    'If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Devol = Round(Rs(0), 2)
    
    'Devolucions
    Devol = 0
    Set Rs = Db.OpenResultset("Select Sum(quantitattornada * preu) From [" & DonamNomTaulaServit(D) & "] S join articles a on s.codiarticle = a.codi  Where client = " & client)
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Devol = Round(Rs(0), 2)
    'Servit
    servit = 0
    Set Rs = Db.OpenResultset("Select Sum(quantitatservida * preu) From [" & DonamNomTaulaServit(D) & "] S join articles a on s.codiarticle = a.codi  Where client = " & client)
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then servit = Round(Rs(0), 2)

    Dz = DateSerial(Year(D), Month(D), Day(D)) + TimeSerial(23, 55, 55)
    Set Rs = Db.OpenResultset("Select min(data) From [" & NomTaulaMovi(D) & "] Where tipus_moviment = 'Z' And Botiga = " & client & " And day(data) = " & Day(D))
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Dz = Rs(0)
    
    DescM = 0
    Set Rs = Db.OpenResultset("Select sum(import) From [" & NomTaulaMovi(D) & "] Where tipus_moviment = 'J' And Botiga = " & client & " And data <= convert(datetime,'" & Dz & "') and day(data) = " & Day(Dz) & " ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then DescM = Rs(0)
    
    DescT = 0
    Set Rs = Db.OpenResultset("Select sum(import) From [" & NomTaulaMovi(D) & "] Where tipus_moviment = 'J' And Botiga = " & client & " And data > convert(datetime,'" & Dz & "') and day(data) = " & Day(Dz) & " ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then DescT = Rs(0)
    
    Huv = Dz
    Set Rs = Db.OpenResultset("Select max(data) From [" & NomTaulaVentas(D) & "] Where Botiga = " & client & " And day(data) = " & Day(D) & "  ")
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Huv = Rs(0)
    
    
    Set Rs = Db.OpenResultset("Select count(distinct num_tick) , sum(import) From [" & NomTaulaVentas(D) & "] Where Botiga = " & client & " And day(data) = " & Day(D) & " And data < convert(datetime,'" & Dz & "') ")
    cM = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then cM = Rs(0)
    zM = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then zM = Rs(1)
    
    Set Rs = Db.OpenResultset("Select count(distinct num_tick) , sum(import) From [" & NomTaulaVentas(D) & "] Where Botiga = " & client & " And day(data) = " & Day(D) & " And data > convert(datetime,'" & Dz & "') ")
    cT = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then cT = Rs(0)
    zT = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then zT = Rs(1)
    
    'Tiquets anulats mati
    Set Rs = Db.OpenResultset("Select count(distinct num_tick) , sum(import) From [" & NomTaulaAnulats(D) & "] Where Botiga = " & client & " And day(data) = " & Day(D) & " And data < convert(datetime,'" & Dz & "') ")
    tM = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then tM = Rs(0)
    tMI = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then tMI = Rs(1)
    
    'Tiquets anulats tarda
    Set Rs = Db.OpenResultset("Select count(distinct num_tick) , sum(import) From [" & NomTaulaAnulats(D) & "] Where Botiga = " & client & " And day(data) = " & Day(D) & " And data > convert(datetime,'" & Dz & "') ")
    tT = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then tT = Rs(0)
    tTI = 0
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then tTI = Rs(1)
    
    'Treballadors
    Set Rs = Db.OpenResultset("Select Memo Nom,data,operacio from [" & NomTaulaHoraris(D) & "] join dependentes on codi = dependenta where botiga = " & client & " and day(data) = " & Day(D) & " order by data ")
    
    ReDim Deps(0)
    ReDim DepsEstat(0)
    ReDim DepsHi(0)
    ReDim DepsHm(0)
    ReDim DepsHt(0)
    ReDim DepsHistoric(0)
    
    While Not Rs.EOF
        For i = 1 To UBound(Deps)
            If Deps(i) = Rs("Nom") Then Exit For
        Next
        If i > UBound(Deps) Then
            ReDim Preserve Deps(i)
            ReDim Preserve DepsEstat(i)
            ReDim Preserve DepsHi(i)
            ReDim Preserve DepsHm(i)
            ReDim Preserve DepsHt(i)
            ReDim Preserve DepsHistoric(i)
            Deps(i) = Rs("Nom")
        End If
        'DepsHistoric(i) = DepsHistoric(i) & " " & Hour(rs("Data")) & ":" & Minute(rs("Data")) & "|"
        DepsHistoric(i) = DepsHistoric(i) & " " & UCase(Rs("Operacio")) & " " & Hour(Rs("Data")) & ":" & Minute(Rs("Data")) & "|"
        If Rs("Operacio") = "E" Then
           DepsHi(i) = Rs("Data")
        Else
           If DepsHi(i) < Dz And DepsHi(i) <> "" Then
              If Rs("Data") < DateAdd("n", 20, Dz) Then
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Rs("Data"))
                 DepsHi(i) = Rs("Data")
              Else
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Dz)
                 DepsHi(i) = Dz
              End If
           End If
           
           If Rs("Data") > Dz Then
              If Rs("Data") > DateAdd("n", 20, Dz) Then
                 DepsHt(i) = DepsHt(i) + DateDiff("n", DepsHi(i), Rs("Data"))
                 DepsHi(i) = Rs("Data")
              End If
           End If
        End If
        
        DepsEstat(i) = Rs("Operacio")
        Rs.MoveNext
    Wend
    For i = 1 To UBound(Deps)
        If DepsEstat(i) = "E" Then
           If DepsHi(i) < Dz Then
              If Huv < DateAdd("n", 20, Dz) Then
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Huv)
                 DepsHi(i) = Huv
              Else
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Dz)
                 DepsHi(i) = Dz
              End If
           End If
           If Huv > Dz Then
              If Huv > DateAdd("n", 20, Dz) Then
                 DepsHt(i) = DepsHt(i) + DateDiff("n", DepsHi(i), Huv)
                 DepsHi(i) = Huv
              End If
           End If
       End If
    Next
    
    
    ReDim DepsT(0)
    ReDim DepsM(0)
    hM = 0
    hT = 0
    For i = 1 To UBound(Deps)
        If DepsHt(i) > 0 Then
            ReDim Preserve DepsT(UBound(DepsT) + 1)
            DepsT(UBound(DepsT)) = Deps(i) & "|" & DepsHistoric(i) & "(" & digital(DepsHt(i)) & ")"
            'DepsT(UBound(DepsT)) = DepsHistoric(i) & "(" & digital(DepsHt(i)) & ")"
            hT = hT + DepsHt(i)
        End If
        If DepsHm(i) > 0 Then
            ReDim Preserve DepsM(UBound(DepsM) + 1)
            DepsM(UBound(DepsM)) = Deps(i) & "|" & DepsHistoric(i) & " (" & digital(DepsHm(i)) & ")"
            'DepsM(UBound(DepsM)) = DepsHistoric(i) & "(" & digital(DepsHm(i)) & ")"
            hM = hM + DepsHm(i)
        End If
    Next
    
    For i = 0 To UBound(Families)
        FamiliesPct(i) = 0
    Next
    
    Set Rs = Db.OpenResultset("Select ff.pare,sum(import) from [" & NomTaulaVentas(D) & "] v join articles a on a.codi = v.plu  join families f on f.nom=a.familia  join families ff on f.pare=ff.nom Where Botiga = " & client & " And day(data) = " & Day(D) & " group by ff.pare order by ff.pare ")
    While Not Rs.EOF
        For i = 0 To UBound(Families)
            If Families(i) = Rs(0) Then Exit For
        Next
        If i > UBound(Families) Then
            i = UBound(Families) + 1
            ReDim Preserve Families(i)
            ReDim Preserve FamiliesPct(i)
            Families(i) = Rs(0)
            FamiliesPct(i) = ""
        End If
        If zM + zT > 0 Then FamiliesPct(i) = Rs(1) / (zM + zT) * 100
        Rs.MoveNext
    Wend
            
    'DESCUENTOS ******************************************************************************************************
    sql = "select nom,nif,adresa,cp,ciutat,[Desconte ProntoPago],[Desconte 1],[Desconte 2],[Desconte 3],[Desconte 4], "
    sql = sql & "(case when [preu base]<2 then '' else 'major' end)as pb "
    sql = sql & "from clients where codi=" & client
    Set Rs = Db.OpenResultset(sql)
    ReDim Descon(4)
    If Not Rs.EOF Then
        Dpp = Rs("Desconte ProntoPago")
        Descon(0) = 0
        Descon(1) = Rs("Desconte 1")
        Descon(2) = Rs("Desconte 2")
        Descon(3) = Rs("Desconte 3")
        Descon(4) = Rs("Desconte 4")
        tipoPreu = Rs("pb")
    End If

    f = 0
    sql = "select * from constantsclient where variable = 'DtoFamilia' and codi=" & client
    Set Rs = Db.OpenResultset(sql)
    ReDim descFam(f)
    While Not Rs.EOF
        f = f + 1
        ReDim Preserve descFam(f)
        descFam(f) = Rs("valor")
        Rs.MoveNext
    Wend
    
    p = 0
    sql = "select * from constantsclient where variable = 'DtoProducte' and codi=" & client
    Set Rs = Db.OpenResultset(sql)
    ReDim descProd(p)
    While Not Rs.EOF
        p = p + 1
        ReDim Preserve descProd(p)
        descProd(p) = Rs("valor")
        Rs.MoveNext
    Wend

    DescTe = ""
    Set Rs = Db.OpenResultset("select * from constantsclient where variable='descTE' and codi=" & client)
    If Not Rs.EOF Then DescTe = Rs("valor")
    
    '~DESCUENTOS ******************************************************************************************************
    
    'IVA
    ReDim IvaValor(0)
    IvaValor(0) = 4
    p = 1
    Set Rs = Db.OpenResultset("select tipus, iva, isnull(irpf,0) irpf from tipusIva order by tipus")
    While Not Rs.EOF
        ReDim Preserve IvaValor(p)
        IvaValor(p) = Rs("iva")
        p = p + 1
        Rs.MoveNext
    Wend
    
    'Devolucions produccio i Servit produccio
    DevolF = 0
    TImportQT = 0
    TImportQS = 0
    sql = "Select isnull(f3.nom,'') fam3,isnull(f2.nom,'') fam2,isnull(f1.nom,'') fam1, isnull(a.nom,'') nom, "
    sql = sql & "CodiArticle,isnull(quantitatTornada,0) QT,isnull(quantitatServida,0) QS, "
    sql = sql & "isnull(isnull(isnull(te.preu" & tipoPreu & ", t.preu" & tipoPreu & "), a.preu" & tipoPreu & "),0) preu, "
    sql = sql & "a.TipoIva iva, "
    If DescTe = "descTE" Then
       sql = sql & "a.desconte as desconte "
    Else
       sql = sql & "(case when isnull(t.preu" & tipoPreu & ",0)=0 and isnull(te.preu" & tipoPreu & ",0)=0 then a.desconte else 0 end) As Desconte "
    End If
    
    sql = sql & "From (select quantitatServida,quantitatTornada,CodiArticle "
    sql = sql & "from [" & DonamNomTaulaServit(D) & "] Where client=" & client & " and "
    sql = sql & "(quantitatServida > 0 or quantitatTornada>0) ) s "
    sql = sql & "join articles a on a.codi=s.codiarticle "
    sql = sql & "Left join tarifesEspecials t on a.codi=t.codi and t.tarifaCodi=(select [desconte 5] from clients where codi = " & client & ") "
    sql = sql & "left join tarifesespecialsclients te on a.codi=te.codi and te.client='" & client & "' "
    sql = sql & "Left join families f1 on f1.nom=a.familia left join families f2 on f2.nom=f1.pare "
    sql = sql & "left join families f3 on f3.nom=f2.pare "
    sql = sql & "order by fam3,fam2,Fam1,a.nom "
    Set Rs = Db.OpenResultset(sql)
    ReDim Accu(4)
    Accu(0) = 0
    Accu(1) = 0
    Accu(2) = 0
    Accu(3) = 0
    Accu(4) = 0
    
    ReDim AccuServit(4)
    AccuServit(0) = 0
    AccuServit(1) = 0
    AccuServit(2) = 0
    AccuServit(3) = 0
    AccuServit(4) = 0
    
    ReDim AccuTornat(4)
    AccuTornat(0) = 0
    AccuTornat(1) = 0
    AccuTornat(2) = 0
    AccuTornat(3) = 0
    AccuTornat(4) = 0
    
    Do While Not Rs.EOF
        art = Rs("codiArticle")
        fam1 = Rs("fam1")
        fam2 = Rs("fam2")
        fam3 = Rs("fam3")
        tipoIva = Rs("iva")
        DesconteTipus = Rs("Desconte")
        If DesconteTipus > 0 Then
            descuento = Descon(DesconteTipus)
            If UBound(descFam) > 0 Then
                For f = 1 To UBound(descFam)
                    famD = Split(descFam(f), "|")(0)
                    If famD = fam3 Or famD = fam2 Or famD = fam1 Then
                        descuento = Split(descFam(f), "|")(1)
                    End If
                Next
            End If
            If UBound(descProd) > 0 Then
                For p = 1 To UBound(descProd)
                    prodD = Split(descProd(p), "|")(0)
                    If CStr(prodD) = CStr(art) Then
                        descuento = Split(descProd(p), "|")(1)
                    End If
                Next
            End If
        Else
            descuento = 0
        End If
        import = (Rs("preu") - (Rs("preu") * (descuento / 100)))
        TImportQT = TImportQT + (import * Rs("QT"))
        TImportQS = TImportQS + (import * Rs("QS"))
        AccuServit(tipoIva) = AccuServit(tipoIva) + Rs("QS") * import
        AccuTornat(tipoIva) = AccuTornat(tipoIva) + Rs("QT") * import
        Rs.MoveNext
    Loop
    'Servit/tornat fabrica sense iva
    DevolF = Round(TImportQT, 2)
    ServitF = Round(TImportQS, 2)
    'Servit/tornat fabrica amb iva
    ImpostServit = 0
    ImpostServit = AccuServit(0) + AccuServit(1) + AccuServit(2) + AccuServit(3) + AccuServit(4)
    ImpostTornat = 0
    ImpostTornat = AccuTornat(0) + AccuTornat(1) + AccuTornat(2) + AccuTornat(3) + AccuTornat(4)
    
    If Not tipoPreu = "" Then
        ImpoServit = (AccuServit(1) * (IvaValor(1) / 100)) + (AccuServit(2) * (IvaValor(2) / 100)) + (AccuServit(3) * (IvaValor(3) / 100)) + (AccuServit(4) * 0.32)
        ImpoTornat = (AccuTornat(1) * (IvaValor(1) / 100)) + (AccuTornat(2) * (IvaValor(2) / 100)) + (AccuTornat(3) * (IvaValor(3) / 100)) + (AccuTornat(4) * 0.32)
        If Not Dpp = 0 Then
            ImpoServit = ImpoServit * ((100 - Dpp) / 100)
            ImpoTornat = ImpoTornat * ((100 - Dpp) / 100)
        End If
        ImpostServit = ImpostServit + ImpoServit
        ImpostTornat = ImpostTornat + ImpoTornat
    End If
    DevolFIVA = Round(ImpostTornat, 2)
    ServitFIVA = Round(ImpostServit, 2)
    NetofIVA = Round((ImpostServit - ImpostTornat), 1)
    NetoPor = Round(((NetofIVA / (zM + zT)) * 100), 1)
    '******************************************************************
End Sub

Function NomTaulaTarifesEspecialsClients()
    Dim sql As String
    
    If Not ExisteixTaula("tarifesespecialsclients") Then
        sql = ""
        sql = sql & "CREATE TABLE [dbo].[tarifesespecialsclients]( "
        sql = sql & "[Id] [nvarchar](255) NULL,"
        sql = sql & "[Client] [int] NULL,"
        sql = sql & "[Codi] [int] NULL,"
        sql = sql & "[PREU] [float] NOT NULL,"
        sql = sql & "[PreuMajor] [float] NULL,"
        sql = sql & "[Di] [datetime] NULL,"
        sql = sql & "[Df] [datetime] NULL,"
        sql = sql & "[Qmin] [float] NOT NULL,"
        sql = sql & "[Aux1] [nvarchar](255) NULL"
        sql = sql & ") ON [PRIMARY]"
        ExecutaComandaSql sql
    End If


End Function

Sub rellenaHojaDiaDeLaSetmanaBuscaDades3(D As Date, client, Deps(), DepsM(), DepsT(), DescM, DescT)
    Dim Rs As rdoResultset, Dz As Date, DepsEstat(), DepsHi(), DepsHm(), DepsHt(), i As Integer, H
    Dim Huv, DepsHistoric() As String
    Dim rsDto As ADODB.Recordset, DesconteTipus, Dpp, tipoPreu, f As Integer, p As Integer, DescTe, descuento
    Dim fam1, fam2, fam3, famD, prodD, import, TImportQT, TImportQS, sql, art
    Dim entra, sale, rsD As ADODB.Recordset, rsH As ADODB.Recordset, rsE As ADODB.Recordset, TTrab, TTipoTrab
        
    Dz = DateSerial(Year(D), Month(D), Day(D)) + TimeSerial(23, 55, 55)
    Set Rs = Db.OpenResultset("Select min(data) From [" & NomTaulaMovi(D) & "] Where tipus_moviment = 'Z' And Botiga = " & client & " And day(data) = " & Day(D))
    If Not Rs.EOF Then If Not IsNull(Rs(0)) Then Dz = Rs(0)
    
        'Treballadors
    Set Rs = Db.OpenResultset("Select Memo Nom,data,operacio from [" & NomTaulaHoraris(D) & "] join dependentes on codi = dependenta where botiga = " & client & " and day(data) = " & Day(D) & " order by data ")
    sql = "Select Memo Nom,data,operacio from [" & NomTaulaHoraris(D) & "] join dependentes on codi = dependenta where botiga = " & client & " and day(data) = " & Day(D) & " order by data "
    
    ReDim Deps(0)
    ReDim DepsEstat(0)
    ReDim DepsHi(0)
    ReDim DepsHm(0)
    ReDim DepsHt(0)
    ReDim DepsHistoric(0)
    
    While Not Rs.EOF
        For i = 1 To UBound(Deps)
            If Deps(i) = Rs("Nom") Then Exit For
        Next
        If i > UBound(Deps) Then
            ReDim Preserve Deps(i)
            ReDim Preserve DepsEstat(i)
            ReDim Preserve DepsHi(i)
            ReDim Preserve DepsHm(i)
            ReDim Preserve DepsHt(i)
            ReDim Preserve DepsHistoric(i)
            Deps(i) = Rs("Nom")
        End If
        'DepsHistoric(i) = DepsHistoric(i) & " " & Hour(rs("Data")) & ":" & Minute(rs("Data")) & "|"
        DepsHistoric(i) = DepsHistoric(i) & " " & UCase(Rs("Operacio")) & " " & Hour(Rs("Data")) & ":" & Minute(Rs("Data")) & "|"
        If Rs("Operacio") = "E" Then
           DepsHi(i) = Rs("Data")
        Else
           If Rs("Data") > Dz Then
              If Rs("Data") > DateAdd("n", 20, Dz) Then
                 DepsHt(i) = DepsHt(i) + DateDiff("n", DepsHi(i), Rs("Data"))
                 DepsHi(i) = Rs("Data")
              End If
           End If
           If DepsHi(i) < Dz Then
              If Rs("Data") < DateAdd("n", 20, Dz) Then
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Rs("Data"))
                 DepsHi(i) = Rs("Data")
              Else
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Dz)
                 DepsHi(i) = Dz
              End If
           End If
        End If
        
        DepsEstat(i) = Rs("Operacio")
        Rs.MoveNext
    Wend
    For i = 1 To UBound(Deps)
        If DepsEstat(i) = "E" Then
           If Huv > Dz Then
              If Huv > DateAdd("n", 20, Dz) Then
                 DepsHt(i) = DepsHt(i) + DateDiff("n", DepsHi(i), Huv)
                 DepsHi(i) = Huv
              End If
           End If
           If DepsHi(i) < Dz Then
              If Huv < DateAdd("n", 20, Dz) Then
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Huv)
                 DepsHi(i) = Huv
              Else
                 DepsHm(i) = DepsHm(i) + DateDiff("n", DepsHi(i), Dz)
                 DepsHi(i) = Dz
              End If
           End If
       End If
    Next
    
    
    ReDim DepsT(0)
    ReDim DepsM(0)
    'hM = 0
    'hT = 0
    For i = 1 To UBound(Deps)
        If DepsHt(i) > 0 Then
            ReDim Preserve DepsT(UBound(DepsT) + 1)
            DepsT(UBound(DepsT)) = Deps(i) & "|" & DepsHistoric(i) & " " & digital(DepsHt(i))
            'DepsT(UBound(DepsT)) = DepsHistoric(i) & "(" & digital(DepsHt(i)) & ")"
     '       hT = hT + DepsHt(i)
        End If
        If DepsHm(i) > 0 Then
            ReDim Preserve DepsM(UBound(DepsM) + 1)
            DepsM(UBound(DepsM)) = Deps(i) & "|" & DepsHistoric(i) & " " & digital(DepsHm(i))
            'DepsM(UBound(DepsM)) = DepsHistoric(i) & "(" & digital(DepsHm(i)) & ")"
     '       hM = hM + DepsHm(i)
        End If
    Next
    
End Sub

Sub CarregaHoresEquip(dia, equip)
    Dim sql As String, Rs As rdoResultset
    
    Set Rs = Db.OpenResultset("select Id,d.Nom from dependentesextes e join dependentes d on e.id = d.codi where e.nom = 'EQUIPS'  and valor = '" & equip & "'  ")
    While Not Rs.EOF
    
        Rs.MoveNext
    Wend
    Rs.Close
    
End Sub

Sub rellenaHojaDiaDeLaSetmanaBuscaDadesGrafic(D As Date, client, GraficX(), GraficY())
    Dim i, Rs
    Dim Di As String
    Dim p
    
    i = 0
    ReDim GraficX(i)
    ReDim GraficY(i)
    GraficY(i) = 0
    p = 0
    Di = " round((datepart(hh,data) * 60 + datepart(n,data)) /10,0) "
    Set Rs = Db.OpenResultset("Select min(data) dd, " & Di & " Data ,sum(import)  Import From [" & NomTaulaVentas(D) & "] Where Botiga = " & client & " And day(data) = " & Day(D) & "  group by " & Di & " order by " & Di & " ")
    While Not Rs.EOF
        While i <> 0 And i < Rs("Data")
           i = i + 1
           p = p + 1
           ReDim Preserve GraficX(p)
           ReDim Preserve GraficY(p)
           GraficX(p) = CVDate(Format(Rs("dd"), "hh:mm:ss"))
           GraficY(p) = 0
        Wend
        GraficY(p) = GraficY(p) + Rs("Import")
        i = Rs("Data")
        Rs.MoveNext
    Wend
'    While i < 24
'       i = i + 1
'       ReDim Preserve GraficX(i)
'       ReDim Preserve GraficY(i)
'       GraficX(i) = i
'       GraficY(i) = 0
    'Wend

End Sub


Sub CarregaLlistaBotiguesTotes(Ll As String)
    Dim Rs As rdoResultset
    
   Set Rs = Db.OpenResultset("Select Valor1 from ParamsHw where tipus = 1 ")
   While Not Rs.EOF
      If Not IsNull(Rs("Valor1")) Then
         If IsNumeric(Rs("Valor1")) Then
            If Not Ll = "" Then Ll = Ll & ","
            Ll = Ll & Rs("Valor1")
         End If
      End If
      Rs.MoveNext
   Wend
   Rs.Close
   
End Sub

Public Function CalculaExcel(ByVal p1 As String, ByVal P2 As String, ByVal P3 As String, ByVal P4 As String, ByVal P5 As String) As String
   Dim MsExcel As Excel.Application
   Dim Libro As Excel.Workbook
    CalculaExcel = ""
    
On Error GoTo 0


If Not frmSplash.Debugant Then On Error GoTo nok

    InformaMiss "Calculs Excel"
    Set MsExcel = CreateObject("Excel.Application")
    Set Libro = MsExcel.Workbooks.Add
    
    Select Case p1
        Case "PROD_ANU":
            CalculaExcel = CalculaExcelAnual(P2, P3, P4, P5, MsExcel, Libro)
        Case "ANU":
            CalculaExcel = CalculaExcelCentral(P2, P3, P4, P5, MsExcel, Libro)
        Case Else:
            CalculaExcel = CalculaExcel2(p1, P2, P3, P4, P5, MsExcel, Libro)
    End Select
    
    TancaExcel MsExcel, Libro
    
    Exit Function
    
nok:
  sf_enviarMail "email@hit.cat", "ana@solucionesit365.com", "Error en excel " & p1 & " " & P2 & " " & P3, "", "", ""
  TancaExcel MsExcel, Libro
  
End Function


Public Function CalculaExcelVeureTiquets(ByVal p1 As String, ByVal P2 As String, ByVal P3 As String, ByVal P4 As String, ByVal P5 As String) As String
    Dim MsExcel As Excel.Application, Libro As Excel.Workbook, Hoja As Excel.Worksheet
    Dim params() As String, fecha As Date, botiga As String, botigaCodi As String, sql As String
    Dim Rs As ADODB.Recordset
    Dim rsId As rdoResultset, iD As String
    Dim rsBotiga As rdoResultset
    Dim a As New Stream, s() As Byte
    Dim nom As String, Descripcio As String
    
    CalculaExcelVeureTiquets = ""
    
    params = Split(p1, " ")
    If UBound(params) = 1 Then
        fecha = Now()
        botiga = ""
    ElseIf UBound(params) = 2 Then
        If IsDate(params(2)) Then
            fecha = CDate(params(2))
            botiga = ""
        Else
            fecha = Now()
            botiga = params(2)
        End If
    ElseIf UBound(params) = 3 Then
        If IsDate(params(2)) Then
            fecha = CDate(params(2))
        Else
            fecha = Now()
        End If
        botiga = params(3)
    End If

    If botiga <> "" Then
        Set rsBotiga = Db.OpenResultset("select * from clients where nom = '" & botiga & "'")
        If Not rsBotiga.EOF Then
            botigaCodi = rsBotiga("codi")
        Else
            botiga = ""
        End If
    End If
    
    InformaMiss "Calculs Excel Veure tiquets"
    
    On Error Resume Next
    db2.Close
    On Error GoTo ERR_
    
    Set db2 = New ADODB.Connection
    db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"
    
    Set MsExcel = CreateObject("Excel.Application")
    Set Libro = MsExcel.Workbooks.Add
    
    MsExcel.DisplayAlerts = False
    MsExcel.Visible = frmSplash.Debugant
    
    Set rsId = Db.OpenResultset("select newid() i")
    iD = rsId("i")

    While Libro.Sheets.Count > 1
      Libro.Sheets(1).Delete
    Wend
    
    InformaMiss "ExcelVeureTiquets", True

    sql = "Select n.data [DATA], articles.Nom [ARTICLE], dependentes.nom [DEPENDENTA], clients.nom [BOTIGA], n.num_tick [TIQUET], CAST(n.Total AS nvarchar(10)) [IMPORT], CAST(n.Quantitat AS nvarchar(10)) [QUANTITAT] "
    sql = sql & " From ( "
    sql = sql & " select Import as total, Quantitat, Plu,Num_tick,Dependenta,Botiga,Data,Tipus_Venta, left(Otros,30) as otros, formaMarcar "
    sql = sql & " From [" & NomTaulaVentas(fecha) & "] "
    sql = sql & " where day(Data)=" & Day(fecha) & " "
    If botiga <> "" Then
        sql = sql & " and Botiga = " & botigaCodi & " "
    Else 'SOLO PROPIAS
        sql = sql & " and Botiga in ("
        sql = sql & " select c.Codi "
        sql = sql & " from clients c "
        sql = sql & " join paramshw w on c.Codi = w.Valor1 "
        sql = sql & " where c.nif in (select valor from constantsempresa where camp like '%CampNif%' and isnull(valor, '')<>'') "
        sql = sql & " ) "
    End If
    sql = sql & " Union All "
    sql = sql & " select Import as total, Quantitat, Plu,Num_tick,Dependenta,Botiga,Data,Tipus_Venta, '[A]' + Otros as Otros, formaMarcar "
    sql = sql & " From [" & NomTaulaAnulats(fecha) & "] "
    sql = sql & " where day(data)=" & Day(fecha) & " "
    If botiga <> "" Then
        sql = sql & " and Botiga = " & botigaCodi & " "
    Else 'SOLO PROPIAS
        sql = sql & " and Botiga in ("
        sql = sql & " select c.Codi "
        sql = sql & " from clients c "
        sql = sql & " join paramshw w on c.Codi = w.Valor1 "
        sql = sql & " where c.nif in (select valor from constantsempresa where camp like '%CampNif%' and isnull(valor, '')<>'') "
        sql = sql & " ) "
    End If
    sql = sql & ") n "
    sql = sql & " left join dependentes on n.Dependenta = dependentes.codi "
    sql = sql & " left join clients on n.Botiga = clients.codi "
    sql = sql & " left join articles on n.plu = articles.codi "
    sql = sql & " order by Botiga, Data, Num_Tick"

    rellenaHojaSql "Tiquets (" & Now() & ")", sql, Libro.Sheets(1), 0
    
    
    If Excel.Application.Version >= 12 Then
        Libro.SaveAs Excel.Application.DefaultFilePath & "\" & iD & ".xls", xlExcel8
    Else
        Libro.SaveAs Excel.Application.DefaultFilePath & "\" & iD & ".xls"
    End If
    
    Libro.Close

    Set Libro = Nothing
    Set MsExcel = Nothing

    a.Open
    a.LoadFromFile Excel.Application.DefaultFilePath & "\" & iD & ".xls"
    s = a.ReadText()

        
    Set Rs = rec("SET LANGUAGE Español ;   select * from " & tablaArchivo() & " where fecha=(select max(fecha) from " & tablaArchivo() & ")", True)
    
    Rs.AddNew
    Rs("id").Value = iD
    Rs("nombre").Value = "Tiquets"
    Rs("descripcion").Value = "Tiquets"
    Rs("extension").Value = "XLS"
    Rs("mime").Value = "application/vnd.ms-excel"
    Rs("propietario").Value = ""
    Rs("archivo").Value = s
    Rs("fecha").Value = Now
    Rs("tmp").Value = 0
    Rs("down").Value = 1
    Rs.Update
    Rs.Close
    a.Close

    CalculaExcelVeureTiquets = iD
    
    MyKill Excel.Application.DefaultFilePath & "\" & iD & ".xls"
    

ERR_:

    TancaExcel MsExcel, Libro
    db2.Close
  
End Function


Function CalculaExcel2(ByVal p1 As String, ByVal P2 As String, ByVal P3 As String, ByVal P4 As String, ByVal P5 As String, MsExcel As Object, Libro As Object)
    Dim Rs As ADODB.Recordset, Nom1 As String, Nom2 As String, nom As String, Descripcio As String, D As Date, An, client() As String, pp2 As String, ppp2, Di, Df, LlistaBotiguesPosibles As String
    Dim RsBot As ADODB.Recordset
    Dim i As Double, Kk As Integer, iD As String, a As New Stream, s() As Byte, dia As Date, DiaF As Date, sql, K As Integer
    Dim Punts, clients
    Dim Rs2 As rdoResultset
    Dim Rs3 As ADODB.Recordset
    Dim TMCostH, TTCostH, TCostH, Tz, TzM, TzT, tServit, TDevol, TNeto, TServitIVA, TDevolIVA 'Totals hores,Ingresos i Servit
    
    On Error Resume Next
    db2.Close
    On Error GoTo norR
    
If frmSplash.Debugant Then On Error GoTo 0


    Set db2 = New ADODB.Connection
    db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"

    MsExcel.DisplayAlerts = False
    MsExcel.Visible = frmSplash.Debugant
    
    Set Rs = rec("select newid() i")
    iD = Rs("i")

    While Libro.Sheets.Count > 1
      Libro.Sheets(1).Delete
    Wend
    
    dia = paramToDate(P2)
    If P4 = "" Then
        DiaF = DateAdd("d", 7, dia)
    Else
        DiaF = paramToDate(P4)
    End If
    
  
    Select Case p1
        Case "Horari"
            rellenaHojaVenut Libro.Sheets(1), dia, DiaF, P5
            nom = "Horari " & dateToParam(dia, "_")
            Descripcio = "Detall Horari " & P2 & " a " & DiaF
        Case "MES"
            If Libro.Sheets.Count < 2 Then Libro.Sheets.Add , Libro.Sheets(1)
            If Libro.Sheets.Count < 3 Then Libro.Sheets.Add , Libro.Sheets(2)
            If Libro.Sheets.Count < 4 Then Libro.Sheets.Add , Libro.Sheets(3)
            rellenaHojaMes Libro.Sheets(1), dia, Nom1
            rellenaHojaMes Libro.Sheets(2), DateAdd("yyyy", -1, dia), Nom2
            RellenaHojaMesResumen Libro, dia, Nom1, Nom2
            nom = "28D " & Day(dia) & " " & meses(Month(dia)) & " " & Year(dia)
            Descripcio = "Comparativa mensual de " & meses(Month(dia)) & " " & Year(dia)
        Case "DIA"
            rellenaHojaDia Libro.Sheets(1), dia, "Diari"
            nom = "DIARI " & dateToParam(dia, "_")
            Descripcio = "Vendes diaries de " & P2
        Case "DEM"
            rellenaHojaDem Libro.Sheets(1), dia, "Demanat"
            Set Rs = rec("Select Nom From Viatges Order By Nom")
            Kk = 2
            While Not Rs.EOF
                If Kk > Libro.Sheets.Count Then Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
                rellenaHojaDem Libro.Sheets(Kk), dia, Rs("Nom"), Rs("Nom")
                Rs.MoveNext
                Kk = Kk + 1
            Wend
            nom = "DEMANAT " & dateToParam(dia, "_")
            Descripcio = "Demanat diari de " & P2
        Case "Cuadre"
            rellenaHojaCuadre Libro.Sheets(1), dia, P5
            nom = "Cuadre " & dateToParam(dia, "_") & BotigaCodiNom(P5)
            Descripcio = "Detall Horari " & P2 & " a " & DiaF
        Case "SERMen"
            rellenaHojaSer True, Libro.Sheets(1), dia, "Servit"
            nom = "ServitM" & dateToParam(dia, "_")
            Descripcio = "Demanat diari de " & P2
        Case "SER"
            rellenaHojaSer False, Libro.Sheets(1), dia, "Servit"
            nom = "Servit " & dateToParam(dia, "_")
            Descripcio = "Demanat diari de " & P2
              
    '        Set Rs = rec("Select Nom From Viatges Order By Nom")
    '        kk = 2
    '        While Not Rs.EOF
    '            If kk > Libro.Sheets.Count Then Libro.Sheets.Add , Libro.Sheets(.Sheets.Count)
    '            rellenaHojaSer Libro.Sheets(kk), dia, Rs("Nom"), Rs("Nom")
    '            Rs.MoveNext
    '            kk = kk + 1
    '        Wend
        
        Case "SEM"
            For i = 0 To 6
                If Libro.Sheets.Count < i + 1 Then Libro.Sheets.Add , Libro.Sheets(i)
                rellenaHojaDem Libro.Sheets(i + 1), DateAdd("d", i, dia), "Dia" & i + 1
            Next
            nom = "SETMANA " & dateToParam(dia, "_")
            Descripcio = "Demanat semanal de " & P2 & " a " & dateToParam(DateAdd("d", 6, dia), "-")
            
        Case "HistoricVendes"
            D = dia
            rellenaHojaVenutMes Libro.Sheets(Libro.Sheets.Count), D, P2
            nom = "Ventas " & Month(D) & "/" & Year(D)
            Descripcio = "Ventas de " & Format(D, " mmmm yyyy")
            
        Case "Evolucio Diaria"
            D = dia
            If P5 = "" Then P5 = LlistaBotigues()
            client = Split(P5, ",")
            Descripcio = Format(dia, "dddd") & Day(dia) & " "
            Libro.Styles("Normal").Font.Size = 9
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            rellenaHojaDiaDeLaSetmanaCreaResumen Libro.Sheets(1), D
            Set Rs2 = Db.OpenResultset("Select * from Clients where codi in (" & P5 & ") order by nom ")
            ReDim clients(0)
            While Not Rs2.EOF
                ReDim Preserve clients(UBound(clients) + 1)
                clients(UBound(clients)) = Rs2("Codi")
                Rs2.MoveNext
            Wend
            Rs2.Close
            For i = 1 To UBound(clients)
                rellenaHojaDiaDeLaSetmana Libro.Sheets(Libro.Sheets.Count), D, clients(i), Libro.Sheets(1), (Libro.Sheets.Count - 1) * 2 + 1
                Descripcio = Descripcio & BotigaCodiNom(clients(i)) & " "
                Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            Next
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
            nom = "Tiendas " & dia
        Case "Evolucio Botigues"
            D = dia
            If P4 = "GRUP" Then 'Busquem grup botiga
                sql = "Select c.codi from Clients c with (nolock) left join constantsClient cc with (nolock) on (cc.codi=c.codi) "
                sql = sql & "left join paramshw w on (c.codi = w.valor1) where cc.variable='Grup_client' "
                sql = sql & "and cc.valor like '%" & P5 & "%' order by nom "
                Set Rs3 = rec(sql)
                P5 = ""
                Do While Not Rs3.EOF
                    P5 = P5 & Rs3("codi") & ","
                    Rs3.MoveNext
                Loop
                If P5 <> "" Then P5 = Mid(P5, 1, Len(P5) - 1)
            End If
            If P5 = "" Then P5 = LlistaBotigues()
            client = Split(P5, ",")
            Descripcio = Format(dia, "dddd") & Day(dia) & " "
            Libro.Styles("Normal").Font.Size = 9
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            rellenaHojaDiaDeLaSetmanaCreaResumen2 Libro.Sheets(1), D
            Set Rs3 = rec("Select * from Clients where codi in (" & P5 & ") order by nom ")
            ReDim clients(0)
            Do While Not Rs3.EOF
                ReDim Preserve clients(UBound(clients) + 1)
                clients(UBound(clients)) = Rs3("Codi")
                Rs3.MoveNext
            Loop
            Rs3.Close
            
            TMCostH = 0 'Total cost horari mati
            TTCostH = 0 'Total cost horari tarda
            TCostH = 0 'Total cost horari
            Tz = 0 'Total ingres
            TzM = 0 'Total ingres mati
            TzT = 0 'Total ingres tarda
            tServit = 0 'Total servit
            TDevol = 0 'Total devolucions
            TServitIVA = 0 'Total servit fabrica
            TDevolIVA = 0 'Total devolucions fabrica
            TNeto = 0 'Total neto
            LlistaBotiguesPosibles = LlistaBotigues()
            LlistaBotiguesPosibles = "," & LlistaBotiguesPosibles & ","
            NomTaulaTarifesEspecialsClients
            
            For i = 1 To UBound(clients)
                If InStr(LlistaBotiguesPosibles, "," & clients(i) & ",") > 0 Then
                    If i < UBound(clients) Then
                        rellenaHojaDiaDeLaSetmana3 Libro.Sheets(Libro.Sheets.Count), D, clients(i), Libro.Sheets(1), (Libro.Sheets.Count - 1) * 2 + 1, False, Tz, TzM, TzT, TMCostH, TTCostH, TCostH, tServit, TDevol, TNeto, TServitIVA, TDevolIVA
                    ElseIf i = UBound(clients) Then
                        rellenaHojaDiaDeLaSetmana3 Libro.Sheets(Libro.Sheets.Count), D, clients(i), Libro.Sheets(1), (Libro.Sheets.Count - 1) * 2 + 1, True, Tz, TzM, TzT, TMCostH, TTCostH, TCostH, tServit, TDevol, TNeto, TServitIVA, TDevolIVA
                    End If
                    Descripcio = Descripcio & BotigaCodiNom(clients(i)) & " "
                    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
                End If
            Next
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
            nom = "Tiendas " & dia
            
        Case "Rentabilidad"
            D = dia
            Descripcio = Format(dia, "dddd") & Day(dia) & " "
            
            Set Rs = rec("Select ISNULL(CAMP,'') camp , isnull(valor,'') valor from constantsempresa where camp like '%campnom' and not valor = 'Pomposo' order by camp ")
            While Not Rs.EOF
                rellenaHojaRentabilidad Libro.Sheets(Libro.Sheets.Count), D, Rs("Valor"), Rs("camp")
                Rs.MoveNext
                If Not Rs.EOF Then Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            Wend
            nom = "Rent." & dia
            
        Case "ResumIngredients"
            Dim ArrDatesComanda() As Date
            'Buscar el producto asociado a la materia prima
            Set Rs = rec("select * from ArticlesPropietats where Valor='" & P3 & "'")
            If Not Rs.EOF Then
                'Una hoja por tienda
                Set RsBot = rec("select c.Codi, c.nom from ParamsHw p left join clients c on p.Valor1=c.Codi Where c.Codi Is Not Null order by c.nom")
                While Not RsBot.EOF
                    'Revisar si hay servit tres meses atras
                    ArrDatesComanda = getDatesComanda(RsBot("codi"), Rs("codiArticle"))
                    If UBound(ArrDatesComanda) > 0 Then
                        rellenaHojaIngredients Libro.Sheets(Libro.Sheets.Count), P3, RsBot("Codi"), RsBot("nom"), Rs("codiArticle"), ArrDatesComanda
                        RsBot.MoveNext
                        If Not RsBot.EOF Then Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
                    Else
                        RsBot.MoveNext
                    End If
                Wend
            End If
            Libro.Sheets(1).Select
            nom = "Ingredients" & dia
        Case "Hores Per Botiga"
            D = dia
            client = Split(P5, ",")
            Descripcio = Format(dia, "dddd") & Day(dia) & " "
            For i = 0 To UBound(client)
                Informa " Expportant Excel " & BotigaCodiNom(client(i))
                rellenaHojaHoresPerBotiga Libro.Sheets(Libro.Sheets.Count), D, client(i)
                Descripcio = Descripcio & BotigaCodiNom(client(i)) & " "
                Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            Next
            nom = "Hores " & dia
        
        Case "Control Cobraments"
            D = dia
            Descripcio = "Cobraments Fins " & Format(Now, "dddd") & Day(Now) & " "
            Informa " Expportant Excel Cobraments "
            rellenaCobraments Libro, D
            nom = "Cobraments" ' & Year(dia) & "_" & Format(Now, "ddmmyy")
        Case "Clients"
            D = dia
            Descripcio = "Clients Modificables " & Format(Now, "dddd") & Day(Now) & " "
            Informa " Exportant Excel Clients"
            rellenaClients Libro
            nom = "Clients_" & Year(Now) & "_" & Format(Now, "ddmmyy")
        Case "Preus"
            D = dia
            Descripcio = "Preus Modificables " & Format(Now, "dddd") & Day(Now) & " "
            Informa " Expportant Excel Preus "
            rellenaPreus Libro
            nom = "Preus_" & Year(Now) & "_" & Format(Now, "ddmmyy")
        Case "Preus Per Clients"
            D = dia
            Descripcio = "Preus Modificables " & Format(Now, "dddd") & Day(Now) & " "
            Informa " Expportant Excel Preus Per Clients"
            rellenaPreusPerClients Libro
            nom = "PreusCli_" & Year(Now) & "_" & Format(Now, "ddmmyy")
        Case "FeinaSetmanal"
            D = dia
            While Not Weekday(D) = vbMonday
              D = DateAdd("d", -1, D)
            Wend
            
            Descripcio = "Feina Administracio Setmana " & Format(Now, "dd mmmm yyyy ")
            Informa " Expportant Feina Setmanal"
            FeinaSetmanal Libro, dia
            nom = "Comandes " & Year(Now) & "_" & Format(D, "ddmmyy")
        Case "ExcelInventari"
            Descripcio = "Inventaris Botigues Data " & Format(Now, "dd-mm-yy- hh:nn")
            Informa " Expportant Excel Inventari"
            ExcelInventari Libro, dia
            nom = "Inventaris " & Year(Now) & "_" & Format(D, "ddmmyy")
        Case "Rendiment Equips"
            D = dia
            Descripcio = "Equips " & Format(Now, "dddd") & Day(Now) & " "
            Informa " Expportant Rendiment Equips"
            RendimentEquips Libro, dia
            nom = "Equips" & Year(Now) & "_" & Format(Now, "ddmmyy")
        Case "Preus Especials"
            D = dia
            Descripcio = "Condicions Especials " & Format(Now, "dddd") & Day(Now) & " "
            Informa " Expportant Preus Especials "
            CondicionsEspecials Libro
            nom = "Espe_" & Year(Now) & "_" & Format(Now, "ddmmyy")
        Case "TarjetaClient"
            ExcelTarjaClient Libro
            nom = "TarjetaClient" 'Fins " & Now
            Descripcio = "TarjetaClient" & Now
        Case "QuotesMensuals"
            ExcelQuotesMensuals Libro
            nom = "QuotesMensuals" 'Fins " & Now
            Descripcio = "QuotesMensuals" & Now
        Case "PROVEEDORES"
            ExcelProveedores Libro
            nom = "PROVEEDORES" 'Fins " & Now
            Descripcio = "PROVEEDORES I Materias a " & Now
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Nota"
            ExcelNota Libro, P2, P3
            nom = "Notes" 'Fins " & Now
            Descripcio = "Notes" & Now
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Caixes"
            ExcelCaixes Libro
            nom = "Caixes" 'Fins " & Now
            Descripcio = "Caixes Fabricadas Hasta " & Now
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Palets"
            ExcelPalets Libro
            nom = "Palets" 'Fins " & Now
            Descripcio = "Palets En Conjelador a " & Now
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Trucades"
            ExcelTrucades Libro, P2
            nom = "Trucades" 'Fins " & Now
            Descripcio = "Registre Trucades : " & dia
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Contactess", "Contactes"
            ExcelContactes Libro
            nom = "Contactes" 'Fins " & Now
            Descripcio = "Contactes Crm a " & Now
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Albarans"
            If P2 = "" Then P2 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            pp2 = Car(P2)
            ppp2 = Split(pp2, "-")
            If ppp2(0) = 0 Then
                Di = DateSerial(ppp2(2), ppp2(1), 1)
                Df = DateAdd("m", 1, Di)
                Df = DateAdd("d", -1, Df)
            Else
                Di = DateSerial(ppp2(2), ppp2(1), ppp2(0))
                Df = DateSerial(ppp2(2), ppp2(1), ppp2(0))
            End If
                            
            ExcelAlbarans Libro, CDate(Di), Df
            
            nom = "Albarans" & Format(dia, "mmmm yyyy")
            Descripcio = "Albarans de " & Format(Di, "dd-mm-yyyy") & " a " & Format(Df, "dd-mm-yyyy")
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Albarans2"
            If P2 = "" Then P2 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            pp2 = Car(P2)
            ppp2 = Split(pp2, "-")
            If ppp2(0) = 0 Then
                Di = DateSerial(ppp2(2), ppp2(1), 1)
                Df = DateAdd("m", 1, Di)
                Df = DateAdd("d", -1, Df)
            Else
                Di = DateSerial(ppp2(2), ppp2(1), ppp2(0))
                Df = DateSerial(ppp2(2), ppp2(1), ppp2(0))
            End If
                            
            ExcelAlbarans2 Libro, CDate(Di), Df
            
            nom = "Albarans" & Format(dia, "mmmm yyyy")
            Descripcio = "Albarans de " & Format(Di, "dd-mm-yyyy") & " a " & Format(Df, "dd-mm-yyyy")
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Desquadres"
            If P2 = "" Then P2 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            pp2 = Car(P2)
            ppp2 = Split(pp2, "-")
            If ppp2(0) = 0 Then
                Di = DateSerial(ppp2(2), ppp2(1), 1)
                Df = DateAdd("m", 1, Di)
                Df = DateAdd("d", -1, Df)
            Else
                Di = DateSerial(ppp2(2), ppp2(1), ppp2(0))
                Df = DateSerial(ppp2(2), ppp2(1), ppp2(0))
            End If
                            
            ExcelDesquadres Libro, CDate(Di), CDate(Df)
            
            nom = "Desquadraments botigues " & Format(dia, "mmmm yyyy")
            Descripcio = "Desquadraments de botigues del " & Format(Di, "dd-mm-yyyy") & " a " & Format(Df, "dd-mm-yyyy")
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "AlbaransFam"
            If P2 = "" Then P2 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            pp2 = Car(P2)
            ppp2 = Split(pp2, "-")
            If ppp2(0) = 0 Then
                Di = DateSerial(ppp2(2), ppp2(1), 1)
                Df = DateAdd("m", 1, Di)
                Df = DateAdd("d", -1, Df)
            Else
                Di = DateSerial(ppp2(2), ppp2(1), ppp2(0))
                Df = DateSerial(ppp2(2), ppp2(1), ppp2(0))
            End If
                            
            ExcelAlbaransFam Libro, CDate(Di), Df
            
            nom = "Albarans" & Format(dia, "mmmm yyyy")
            Descripcio = "Albarans de " & Format(Di, "dd-mm-yyyy") & " a " & Format(Df, "dd-mm-yyyy")
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "AlbaransComercial"
            P2 = Format(Now, "dd-mm-yyyy")
            ExcelAlbaransComercial Libro, P2, P3, P4
            nom = "Albarans comercial" & P2
            Descripcio = "Albarans "
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Ventas"
            ExcelVentas Libro, dia
            nom = "Ventas" & Format(dia, "mmmm yyyy")
            Descripcio = "Ventas" & Format(dia, "mmmm yyyy")
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Ventas Detall Franquicia"
            ExcelVentasDetallFranquicia Libro, dia, P3, P4, P5
            nom = "Ventas" & Format(dia, "mmmm yyyy")
            Descripcio = "Ventas " & Format(dia, "mmmm yyyy")
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Ventas Mensual Franquicia"
            ExcelVentasMensualFranquicia Libro, dia, P3, P4, P5
            nom = "Ventas " & Format(dia, "mmmm yyyy")
            Descripcio = "Ventas " & Format(dia, "mmmm yyyy")
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Quadre Mensual Franquicia"
            ExcelQuadreMensualFranquicia Libro, dia, P3, P4
            nom = "Quadrar Caixa " & Format(dia, "mmmm yyyy")
            Descripcio = "Quadrar Caixa " & Format(dia, "mmmm yyyy")
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Devolucions Mensual Franquicia"
            ExcelDevolucionsMensualFranquicia Libro, dia, P3, P4
            nom = "Devolucions " & Format(dia, "mmmm yyyy")
            Descripcio = "Devolucions " & Format(dia, "mmmm yyyy")
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Consum Personal Mensual Franquicia"
            ExcelConsumPersonalMensualFranquicia Libro, dia, P3, P4
            nom = "Consum Personal " & Format(dia, "mmmm yyyy")
            Descripcio = "Consum Personal " & Format(dia, "mmmm yyyy")
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Facturat"
            ExcelFacturat Libro
            nom = "Contactes" 'Fins " & Now
            Descripcio = "Facturat Fins " & Now
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Materias Primas"
            ExcelMateriasPrimas Libro, dia
            nom = "Palets" 'Fins " & Now
            Descripcio = "Detall Magatzems " & dia
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "EQUIPS"
            ExcelProduccioEquips Libro, dia
            nom = "Palets" 'Fins " & Now
            Descripcio = "Produccio Per Equips Data " & dia
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Major"
            ExcelMajor Libro, dia
            nom = "Resum Major "  'Fins " & Now
            Descripcio = "Major Total " & dia
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Devolucions"
            If P2 = "" Then P2 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            If P3 = "" Then P3 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            Di = P2
            Df = P3
            ExcelDevolucions Libro, Di, Df
            nom = "Devolucions" & Di & " - " & Df
            Descripcio = "Devolucions" & Di & " - " & Df
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "Rentabilitat"
            If P2 = "" Then P2 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            If P3 = "" Then P3 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            Di = P2
            Df = P3
            ExcelRentabilitat Libro, Di, Df
            nom = "Rentabilitat" & Di & "  " & Df
            Descripcio = "Rentabilitat" & Di & "  " & Df
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "FacturesRebudes"
            If P2 = "" Then P2 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            If P3 = "" Then P3 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            Di = CVDate(Car(P2))
            Df = CVDate(Car(P3))
            ExcelFacturesRebudes Libro, Di, Df, P4
            nom = "Factures Rebudes " & Di & "  " & Df
            Descripcio = "Factures Rebudes : " & Di & "  " & Df
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "ModifGraella"
            D = dia
            ExcelModifGraella Libro, dia
            nom = "Modifiacions Graella" & dia
            Descripcio = "Modifiacions Graella" & dia
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "ServitTornat"
            rellenaHojaServitdiari Libro.Sheets(1), dia, "Servit"
            Set Rs = rec("Select Nom From Viatges Order By Nom")
            Kk = 2
            While Not Rs.EOF
                If Kk > Libro.Sheets.Count Then Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
                rellenaHojaServitdiari Libro.Sheets(Kk), dia, Rs("Nom"), Rs("Nom")
                Rs.MoveNext
                Kk = Kk + 1
            Wend
            nom = "SERVIT " & dateToParam(dia, "_")
            Descripcio = "Servit diari de " & P2
        Case "Hores"
            'If P2 = "" Then P2 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            'If P3 = "" Then P3 = "[" & Format(Now, "dd-mm-yyyy") & "]"
            'Di = CVDate(Car(P2))
            'Df = CVDate(Car(P3))
            Libro.Styles("Normal").Font.Size = 9
            'If P2 = "" Then P2 = DatePart("ww", Now, vbMonday, vbFirstFourDays)
            If P2 = "" Then P2 = DatePart("ww", Now)
            If P2 = "0" Then P2 = DatePart("ww", Now)
            'ExcelHores Libro.Sheets(1), Di, Df
            If IsNumeric(P2) Then ExcelHores2 Libro.Sheets(1), P2, P3, P4, P5
            nom = "Hores"
            Descripcio = "Hores " & Di & " - " & Df
            Libro.Sheets(1).Select
            Libro.Sheets(1).Cells.EntireColumn.AutoFit
            Libro.Sheets(1).Range("A1").Select
        Case "FacturacioAnual"
            nom = "FacturacionAnual_" & P2
            ExcelFacturacioAnual Libro.Sheets(1), Car(P2)
        Case "FacturacioAnualAcc"
            nom = "FacturacionAnualAcc_" & P2
            ExcelFacturacioAnualAcc Libro.Sheets(1), Car(P2)
        Case Else
             nom = "Cap"
    
    End Select
    
    If Not nom = "Cap" Then
        If Excel.Application.Version >= 12 Then
        
            Libro.SaveAs Excel.Application.DefaultFilePath & "\" & iD & ".xls", xlExcel8
        Else
            Libro.SaveAs Excel.Application.DefaultFilePath & "\" & iD & ".xls"
        End If
        
        Libro.Close
  
        Set Libro = Nothing
        Set MsExcel = Nothing
  
        a.Open
        a.LoadFromFile Excel.Application.DefaultFilePath & "\" & iD & ".xls"
        s = a.ReadText()

On Error GoTo norR
        
        Set Rs = rec("SET LANGUAGE Español ;   select * from " & tablaArchivo() & " where fecha=(select max(fecha) from " & tablaArchivo() & ")", True)
        
        ' ESTO LO HA PUESTO JORGE PARA EL ERROR DE GUARDADO.
        'rs.CursorLocation = adUseClient
        'NO GUARDA FICHERO !!!!!!!!!!!!!!!!!!!!
        
        Rs.AddNew
        Rs("id").Value = iD
        Rs("nombre").Value = nom
        Rs("descripcion").Value = Descripcio
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
        CalculaExcel2 = iD
        
        MyKill Excel.Application.DefaultFilePath & "\" & iD & ".xls"
    End If
    
norR:

    On Error Resume Next
        db2.Close
    On Error GoTo 0
    
End Function


Function CalculaExcelCentral(ByVal P2 As String, ByVal P3 As String, ByVal P4 As String, ByVal P5 As String, MsExcel As Object, Libro As Object)
    Dim An  As Double, mes As Double, data() As String
    On Error Resume Next
    Dim Rs As ADODB.Recordset, i As Integer, iD As String, a As New Stream, s() As Byte, dia As Date

    On Error Resume Next
        db2.Close
    On Error GoTo norR
   
    Set db2 = New ADODB.Connection
    db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"

    MsExcel.DisplayAlerts = False
    MsExcel.Visible = frmSplash.Debugant
    
    Set Rs = rec("select newid() i")
    iD = Rs("i")

    While Libro.Sheets.Count > 1
        Libro.Sheets(1).Delete
    Wend
    
    If P5 = "" Then CarregaLlistaBotiguesTotes P5
       
    CalculaExcelCentralCalcArticles data
    For An = 0 To 4
        Libro.Sheets.Add , Libro.Sheets(An + 1)
        CalculaExcelCentralAddArticles Libro.Sheets(An + 1), Year(Now) - An, data
        For mes = 1 To 12
            Informa " Expportant Excel any  " & Year(Now) - An & " Mes " & mes
            CalculaExcelCentralAddMes Libro.Sheets(An + 1), Year(Now) - An, mes, UBound(data, 1), P5
            DoEvents
        Next
    Next
    Dim Titol As String
    Titol = "Acumulat Anual " & Format(Now, "mmmm yyyy")
    If Len(P5) > 0 Then
        Dim Cli() As String
        Cli = Split(P5, ",")
        Libro.Sheets(Libro.Sheets.Count).Name = "Detall"
        Libro.Sheets(Libro.Sheets.Count).Cells(1, 1).Value = "Resum Anual Generat : " & Now
        Libro.Sheets(Libro.Sheets.Count).Cells(2, 1).Value = "Botigues : "
        For i = 0 To UBound(Cli)
            Libro.Sheets(Libro.Sheets.Count).Cells(2 + i, 2).Value = BotigaCodiNom(Cli(i))
            Titol = Titol & " " & BotigaCodiNom(Cli(i))
        Next
    End If
    
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
    Rs("descripcion").Value = Left(Titol, 250)
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
    CalculaExcelCentral = iD
norR:
    On Error Resume Next
        db2.Close
    On Error GoTo 0

End Function

Function CalculaExcelAnual(ByVal P2 As String, ByVal P3 As String, ByVal P4 As String, ByVal P5 As String, MsExcel As Object, Libro As Object)
    Dim An  As Double, mes As Double, data() As String, Col As Integer, Camp As String, Titol As String
    On Error Resume Next
    Dim Rs As ADODB.Recordset, i As Integer, iD As String, a As New Stream, s() As Byte, dia As Date
    
    Camp = "QuantitatDemanada"
    An = Car(P2)
    On Error Resume Next
        db2.Close
    On Error GoTo norR
   
    Set db2 = New ADODB.Connection
    db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"

    MsExcel.DisplayAlerts = False
    MsExcel.Visible = frmSplash.Debugant
    
    Set Rs = rec("select newid() i")
    iD = Rs("i")

    While Libro.Sheets.Count > 1
        Libro.Sheets(1).Delete
    Wend
    
    If P5 = "" Then CarregaLlistaBotiguesTotes P5
       On Error GoTo 0
       
    CalculaExcelCentralCalcArticles data
    dia = DateSerial(An, 1, 1)
    
    Libro.Sheets(1).Name = "Any " & Year(dia)
    Libro.Sheets(1).Range(Libro.Sheets(1).Cells(1, 1), Libro.Sheets(1).Cells(UBound(data, 1), 7)).Value = data
    
    Col = 0
    While An = Year(dia)
        Informa " Expportant Excel Servit " & dia
        CalculaExcelAnualAddMes Libro.Sheets(1), dia, UBound(data, 1), P5, Col, Camp
        DoEvents
        Col = Col + 1
        dia = DateAdd("m", 1, dia)
    Wend
    
    Libro.Sheets.Add , Libro.Sheets(1)
    Titol = "Acumulat Anual " & Camp & " " & An
    If Len(P5) > 0 Then
        Dim Cli() As String
        Cli = Split(P5, ",")
        Libro.Sheets(Libro.Sheets.Count).Name = "Detall"
        Libro.Sheets(Libro.Sheets.Count).Cells(1, 1).Value = "Resum Anual Generat : " & Now
        Libro.Sheets(Libro.Sheets.Count).Cells(2, 1).Value = "Botigues : "
        For i = 0 To UBound(Cli)
            Libro.Sheets(Libro.Sheets.Count).Cells(2 + i, 2).Value = BotigaCodiNom(Cli(i))
            Titol = Titol & " " & BotigaCodiNom(Cli(i))
        Next
    End If
    
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
  
    Rs("nombre").Value = Left(Titol, 20)
    Rs("descripcion").Value = Left(Titol, 250)
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
   CalculaExcelAnual = iD
norR:
    On Error Resume Next
        db2.Close
    On Error GoTo 0

End Function


Function CalculaExcelResultatBotiga(MsExcel As Object, Libro As Object, botiga, dia As Date)
    Dim An  As Double, mes As Double, data() As String, Col As Integer, Camp As String, Titol As String, sql As String, rsC
    On Error Resume Next
    Dim Rs As ADODB.Recordset, i As Integer, iD As String, a As New Stream, s() As Byte

    Camp = "QuantitatDemanada"
    An = 2014
    On Error Resume Next
        db2.Close
    On Error GoTo 0
   
    Set db2 = New ADODB.Connection
    db2.Open "WSID=" & db2MyId & ";UID=" & db2User & ";PWD=" & db2Psw & ";Database=" & db2NomDb & ";Server=" & db2Server & ";Driver={SQL Server};DSN='';"

    MsExcel.DisplayAlerts = False
    MsExcel.Visible = True 'frmSplash.Debugant
    
    Set Rs = rec("select newid() i")
    iD = Rs("i")

    While Libro.Sheets.Count > 1
        Libro.Sheets(1).Delete
    Wend
    
    dia = DateSerial(Year(Now), Month(Now), Day(Now))
    
    Libro.Sheets(1).Name = Format(dia, "yyyy mmmm")
    
    sql = "select data,import,motiu from [" & NomTaulaMovi(dia) & "] where botiga = " & botiga & " order by DATa"
    rellenaHojaSql "Gastos", sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    sql = "select SUM(import),DAY(data) from [" & NomTaulaVentas(dia) & "] where botiga = " & botiga & "  group by DAY(data) order by DAY(data)"
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    rellenaHojaSql "Gastos", sql, Libro.Sheets(Libro.Sheets.Count), 0
    
    
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
  
    Rs("nombre").Value = Left(Titol, 20)
    Rs("descripcion").Value = Left(Titol, 250)
    Rs("extension").Value = "XLS"
    Rs("mime").Value = "application/vnd.ms-excel"
    Rs("propietario").Value = ""
    Rs("archivo").Value = s
    Rs("fecha").Value = Now
    Rs("tmp").Value = 0
    Rs("down").Value = 1
    Rs.Update
    Rs.Close
    a.Close
   CalculaExcelResultatBotiga = iD
norR:
    On Error Resume Next
        db2.Close
    On Error GoTo 0

End Function




Private Sub rellenaHojaDia(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, ByVal nombre As String)
    Dim Rs As ADODB.Recordset, rsC As ADODB.Recordset, sql As String, i As Integer, c As Integer
        
    Hoja.Name = nombre
    Hoja.Cells(1, 1).Value = dateToParam(dia, "-")

    'Set rsC = rec("select count(distinct valor1)+2 c from paramshw")
    Set rsC = rec("select count(distinct codi)+2 c from clients")
    c = rsC("c")

    '    Set rs = rec("select nom from clients where codi in (select distinct valor1 from paramshw) order by codi")
    Set Rs = rec("select nom from clients order by codi")
On Error Resume Next
    Hoja.Range("B1", Hoja.Cells(1, c)).Value = Rs.GetRows
    Hoja.Range("B1", Hoja.Cells(1, c)).Font.Bold = True
    Hoja.Range("B1", Hoja.Cells(1, c)).HorizontalAlignment = xlCenter
    Hoja.Range("A1", Hoja.Cells(1, c)).Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Range("A1", Hoja.Cells(1, c)).Borders(xlEdgeTop).Weight = xlMedium

    Set rsC = rec("select count(codi)+2 c from articles")
    c = rsC("c")
    
    Set Rs = rec("select nom from articles order by codi")
    Hoja.Range("A2", Hoja.Cells(c, 1)).CopyFromRecordset Rs
    Hoja.Range("A1", Hoja.Cells(c, 1)).Font.Bold = True
    Hoja.Range("A1", Hoja.Cells(c, 1)).Borders(xlEdgeLeft).Weight = xlMedium
    Hoja.Range("A1", Hoja.Cells(c, 1)).Borders(xlEdgeRight).Weight = xlMedium

    i = 2    'Set rs = rec("select distinct valor1 from paramshw order by valor1")
    Set Rs = rec("select distinct codi from clients order by codi")
    While Not Rs.EOF
      sql = "select isnull(e.quantitat,0) q,a.codi from articles a inner join [" & nomTaulaEstadistic(dia) & "] " & _
            " e on a.codi=e.producte Where e.tipus=2 And e.client=" & Rs("codi") & " And e.dia=" & Day(dia) & _
            " union select 0 q,codi From articles where codi not in (select producte From " & _
            "[" & nomTaulaEstadistic(dia) & "] Where tipus=2 And client=" & Rs("codi") & " And dia=" & Day(dia) & _
            ") order by codi"
      Set rsC = rec(sql)
      Hoja.Range(Hoja.Cells(2, i), Hoja.Cells(c, i)).CopyFromRecordset rsC
      Hoja.Range(Hoja.Cells(2, i), Hoja.Cells(c, i)).NumberFormat = "0.00"
      i = i + 1
      Rs.MoveNext
    Wend

    Hoja.Columns(i).Clear
    Set Hoja = Nothing

End Sub

Private Sub rellenaHojaDem(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, ByVal nombre As String, Optional Viatge As String = "")
  Dim Rs As ADODB.Recordset
  Dim rsC As ADODB.Recordset
  Dim sql As String
  Dim i As Integer
  Dim c As Integer
  
On Error Resume Next
  With Hoja
    
    .Name = Left(Replace(nombre, ":", " "), 30)
    
    'Set rsC = rec("select count(distinct valor1)+3 c from paramshw")
    
    sql = "select * From [" & DonamNomTaulaServit(dia) & "] "
    If Len(Viatge) > 0 Then sql = sql & " Where Viatge = '" & Viatge & "' "
    Set Rs = rec(sql)
    If Not Rs.EOF Then
    
        .Cells(1, 1).Value = dateToParam(dia, "-")
    
        'Set rsC = rec("select count(distinct codi) + 3 c from (select distinct codi from clients union select distinct codi from clients_zombis) t")
        Set rsC = rec("select count(distinct codi) + 3 c from clients")
        c = rsC("c")
    
        'Set rs = rec("select nom from (select nom, codi from clients union select nom, codi from clients_zombis) t order by codi")
        Set Rs = rec("select nom from clients order by codi")
    
        With .Cells(1, 2)
            .Value = "TOTALS"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    
        With .Range("C1", .Cells(1, c))
            .Value = Rs.GetRows
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

        With .Range("A1", .Cells(1, c))
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
        End With

        'Set rsC = rec("select count(*) c from (select distinct codi from articles union select distinct codi from articles_zombis) t")
        Set rsC = rec("select count(*) c from articles")
        c = rsC("c")
    
        'Set rs = rec("select nom from (select nom, codi from articles union select nom, codi from articles_zombis) t order by codi")
        Set Rs = rec("select nom from articles order by codi")
        .Range("A2", .Cells(c, 1)).CopyFromRecordset Rs

        With .Range("A1", .Cells(c, 1))
            .Font.Bold = True
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With

        With .Range("B1", .Cells(c, 1))
            .Font.Bold = True
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With
    
        i = 3

        'Set rs = rec("select distinct codi from (select codi from clients union select codi from clients_zombis) t order by codi")
        Set Rs = rec("select distinct codi from clients order by codi")
        While Not Rs.EOF
            sql = "select sum(isnull(s.quantitatdemanada,0)) q, a.codi  Codi "
            sql = sql & "from articles a "
            sql = sql & "left join [" & DonamNomTaulaServit(dia) & "] s on a.codi=s.codiarticle and s.client=" & Rs("codi") & " "
            If Len(Viatge) > 0 Then sql = sql & "and s.viatge='" & Viatge & "' "
            sql = sql & "group by a.codi "
            sql = sql & "order by a.codi "

            Set rsC = rec(sql)
            With .Range(.Cells(2, i), .Cells(c, i))
                .CopyFromRecordset rsC
                .NumberFormat = "0.00"
            End With
            i = i + 1
            Rs.MoveNext
        Wend
        .Columns(i).Clear
        With .Range("B2", .Cells(c, 2))
            .Formula = "=SUM(RC[1]:RC[" & i - 3 & "])"
            .NumberFormat = "0.00"
        End With
  
    End If
  
  End With
  
  Set Hoja = Nothing

End Sub


Private Sub rellenaHojaServitdiari(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, ByVal nombre As String, Optional Viatge As String = "")
  Dim Rs As ADODB.Recordset
  Dim rsC As ADODB.Recordset
  Dim sql As String
  Dim i As Integer
  Dim c As Integer
  
On Error Resume Next
  With Hoja
    
    .Name = Left(Replace(nombre, ":", " "), 30)
    
    'Set rsC = rec("select count(distinct valor1)+3 c from paramshw")
    
    sql = "select * From " & DonamTaulaServit(dia) & " with (nolock) "
    If Len(Viatge) > 0 Then sql = sql & " Where Viatge = '" & Viatge & "' "
    Set Rs = rec(sql)
    If Not Rs.EOF Then
    
        .Cells(1, 1).Value = " " & Format(dia, "dd/mm/yyyy")
    
        'Set rsC = rec("select count(distinct codi) + 3 c from (select distinct codi from clients union select distinct codi from clients_zombis) t")
        'Set rsC = rec("select count(distinct codi) + 3 c from clients with (nolock)")
        Set rsC = rec("select count(distinct c.codi) + 3 c from ParamsHw p left join  clients c on c.Codi=p.valor1 Where c.nom Is Not Null")
        c = rsC("c")
    
        'Set rs = rec("select nom from (select nom, codi from clients union select nom, codi from clients_zombis) t order by codi")
        'Set rs = rec("select nom from clients with (nolock) order by codi")
        Set Rs = rec("select c.nom from ParamsHw p left join  clients c on c.Codi=p.valor1 Where c.nom Is Not Null order by c.codi")
    
        With .Cells(1, 2)
            .Value = "TOTALS"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    
        With .Range("C1", .Cells(1, c))
            .Value = Rs.GetRows
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

        With .Range("A1", .Cells(1, c))
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
        End With

        Set rsC = rec("select max(a.codi) c from Articles a")
        c = rsC("c")
    
        Set Rs = rec("select distinct isnull(a.nom, '')  nom, isnull(a.codi, 0) codi , n.num from Articles a right join hit.dbo.nums n on a.codi=n.num left join " & DonamTaulaServit(dia) & " v on v.codiarticle = a.codi Where n.Num <= " & c & " order by n.num")
        .Range("A2", .Cells(c, 1)).CopyFromRecordset Rs

        With .Range("A1", .Cells(c + 1, 1))
            .Font.Bold = True
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With

        With .Range("B1", .Cells(c + 1, 1))
            .Font.Bold = True
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With
    
        i = 3

        'Set rs = rec("select distinct codi from (select codi from clients union select codi from clients_zombis) t order by codi")
        'Set rs = rec("select distinct codi from clients with (nolock) order by codi")
        Set Rs = rec("select c.codi from ParamsHw p left join  clients c on c.Codi=p.valor1 Where c.nom Is Not Null order by c.codi")
        While Not Rs.EOF
        
            sql = "select Sum(case WHEN client=" & Rs(0) & " "
            If Len(Viatge) > 0 Then sql = sql & " and viatge='" & Viatge & "' "
            sql = sql & "THEN  e.quantitatServida else 0 end) q, n.num  Codi "
            sql = sql & "from articles a with (nolock) "
            sql = sql & "right join hit.dbo.nums n with (nolock) on a.codi=n.num "
            sql = sql & "left join " & DonamTaulaServit(dia) & " e with (nolock) on a.codi=e.codiarticle "
            sql = sql & "Where n.Num <= " & c & " "
            sql = sql & "Group By n.num "
            sql = sql & "order by n.num"
        
            'sql = "select sum(isnull(s.quantitatservida,0)) q, a.codi  Codi "
            'sql = sql & "from articles a with (nolock) "
            'sql = sql & "left join [" & DonamNomTaulaServit(dia) & "] s with (nolock) on a.codi=s.codiarticle and s.client=" & rs("codi") & " "
            'If Len(Viatge) > 0 Then sql = sql & "and s.viatge='" & Viatge & "' "
            'sql = sql & "group by a.codi "
            'sql = sql & "order by a.codi "

            Set rsC = rec(sql)
            With .Range(.Cells(2, i), .Cells(c + 1, i))
                .CopyFromRecordset rsC
                .NumberFormat = "0.00"
            End With
            i = i + 1
            Rs.MoveNext
        Wend
        .Columns(i).Clear
        With .Range("B2", .Cells(c + 1, 2))
            .Formula = "=SUM(RC[1]:RC[" & i - 3 & "])"
            .NumberFormat = "0.00"
        End With
  
    End If
 
  End With
  
  Set Hoja = Nothing

End Sub



Private Sub rellenaHojaCuadre(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, cliente As String)
  Dim Rs As ADODB.Recordset
  Dim rsC As ADODB.Recordset
  Dim sql As String
  Dim i As Integer
  Dim c As Integer

On Error GoTo 0
  With Hoja

    sql = "select aa.nom, sum(Inventari0) as Inventari0 ,Sum(Servit)  as servit,sum(venut)*-1 as venut,sum(Tornat)*-1 as Tornat , sum(Inventari1)*-1 as Inventari1  , (sum(Inventari0)*-1  + Sum(Servit)*-1   + sum(venut)  + sum(Tornat) + sum(Inventari1) )  as Cuadre "
    sql = sql & "from ( "
    sql = sql & "select plu as article , sum(quantitat) as Tornat ,0 as Venut ,0 as Servit ,0 as Inventari0 ,0 as Inventari1 from [" & NomTaulaDevol(dia) & "] where botiga = " & cliente & " and day(Data) = " & Day(dia) & " group by plu "
    sql = sql & "Union "
    sql = sql & "select plu as article ,0 as Tornat , sum(quantitat) as venut ,0 as Servit ,0 as Inventari0 ,0 as Inventari1 from [" & NomTaulaVentas(dia) & "] where botiga = " & cliente & " and day(Data) = " & Day(dia) & " group by plu "
    sql = sql & "Union "
    sql = sql & "select CodiArticle as article ,0 as Tornat ,0 as venut, sum(quantitatservida) as Servit,0 as Inventari0,0 as Inventari1   from [" & DonamNomTaulaServit(dia) & "] where client = " & cliente & " group by CodiArticle "
    sql = sql & "Union "
    sql = sql & "select plu as article ,0 as Tornat , 0 as venut,0 as Servit, sum(quantitat) as Inventari0,0 as Inventari1 from [" & NomTaulaInventari(dia) & "] where botiga = " & cliente & " and day(Data) = " & Day(dia) & " group by plu "
    sql = sql & "Union "
    sql = sql & "select plu as article ,0 as Tornat , 0 as venut,0 as Servit, 0 as Inventari0,sum(quantitat) as Inventari1 from [" & NomTaulaInventari(DateAdd("d", -1, dia)) & "] where botiga = " & cliente & " and day(Data) = " & Day(DateAdd("d", -1, dia)) & " group by plu "
    sql = sql & ") a "
    sql = sql & "left join EquivalenciaProductes e on e.prodvenut = a.article "
    sql = sql & "join articles aa on isnull(e.prodservit,a.article) = aa.codi "
    sql = sql & "join families f  on aa.familia = f.nom "
    sql = sql & "join families ff on f.pare = ff.nom "
    sql = sql & "group by ff.pare,f.pare,f.nom,aa.nom "
    sql = sql & "Having Sum(venut) > 0 And (Sum(Inventari0) * -1 + Sum(Servit) * -1 + Sum(venut) + Sum(Tornat) + Sum(Inventari1)) <> 0 "
    sql = sql & "order by ff.pare,f.pare,f.nom,aa.nom "
    Set Rs = rec(sql)
    If Not Rs.EOF Then .Cells.CopyFromRecordset Rs
  End With
  Set Hoja = Nothing

End Sub



Private Sub rellenaHojaSer(ElMes As Boolean, ByRef Hoja As Excel.Worksheet, ByVal dia As Date, ByVal nombre As String, Optional Viatge As String = "")
  Dim Rs As ADODB.Recordset
  Dim rsC As ADODB.Recordset
  Dim sql As String
  Dim i As Integer
  Dim c As Integer

  Informa " Expportant Excel Servit"

  ExecutaComandaSql "Drop Table ServitTmp "
  CreaTaulaServit "ServitTmp"
  ExecutaComandaSql "DROP TRIGGER [M_ServitTmp] "
  If ElMes Then
    Dim dd As Date
    dd = DateSerial(Year(dia), Month(dia), 1)
    While Month(dd) = Month(dia)
        Informa " Expportant Excel Servit " & dd
         sql = "insert into ServitTmp  select * From [" & DonamNomTaulaServit(dd) & "] "
        If Len(Viatge) > 0 Then sql = sql & " Where Viatge = '" & Viatge & "' "
        ExecutaComandaSql sql
        dd = DateAdd("d", 1, dd)
    Wend
  Else
     sql = "insert into ServitTmp  select * From [" & DonamNomTaulaServit(dia) & "] "
     If Len(Viatge) > 0 Then sql = sql & " Where Viatge = '" & Viatge & "' "
     ExecutaComandaSql sql
  End If
  
  With Hoja
    .Name = Left(Replace(nombre, ":", " "), 30)
    'Set rsC = rec("select count(distinct valor1)+3 c from paramshw")
    
    sql = "select * From ServitTmp "
    If Len(Viatge) > 0 Then sql = sql & " Where Viatge = '" & Viatge & "' "
    Set Rs = rec(sql)
    If Not Rs.EOF Then
    
        '.Cells(1, 1).Value = dateToParam(dia, "/")
        .Cells(1, 1).Value = dia
        .Cells(1, 1).Interior.ColorIndex = 19
    
        Set rsC = rec("select count(distinct client) c from ServitTmp ")
        c = rsC("c")
    
        'Set rs = rec("select nom from clients where codi in (select distinct valor1 from paramshw) order by codi")
        Set Rs = rec("select distinct c.nom,c.codi  nom from clients c join  ServitTmp v on v.client = c.codi  order by c.codi")
    
        With .Cells(1, 2)
            .Value = "TOTALS"
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
    
        If c > 254 Then c = 254
        
        '.Cells(1, 1).Interior.ColorIndex = 19
        With .Range("C1", .Cells(1, c + 2))
            .Value = Rs.GetRows
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With

        With .Range("A1", .Cells(1, c))
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
        End With

        Set rsC = rec("select max(a.codi) c from Articles a")
        'Set rsC = rec("select count(distinct v.codiarticle ) c  from Articles a join ServitTmp v on v.codiarticle = a.codi  ")
        c = rsC("c")
    
        Set Rs = rec("select distinct isnull(a.nom, '')  nom, isnull(a.codi, 0) codi , n.num from Articles a right join hit.dbo.nums n on a.codi=n.num left join ServitTmp v on v.codiarticle = a.codi Where n.Num <= " & c & " order by n.num")
        'Set rs = rec("select distinct a.nom  nom, a.codi codi from Articles a join ServitTmp v on v.codiarticle = a.codi order by a.codi ")
        .Range("A2", .Cells(c, 1)).CopyFromRecordset Rs

        With .Range("A1", .Cells(c + 1, 1))
            .Font.Bold = True
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With

        With .Range("B1", .Cells(c + 1, 1))
            .Font.Bold = True
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
        End With
    
        i = 3
        'Set rs = rec("select distinct valor1 from paramshw order by valor1")
        Set Rs = rec("select distinct c.codi codi from clients c join  ServitTmp v on v.client = c.codi  order by c.codi")
        While (Not Rs.EOF) And (i < 256)
        
            sql = "select Sum(case Client when " & Rs(0) & " then  e.quantitatServida else 0 end) q,n.num  Codi "
            sql = sql & "from articles a "
            sql = sql & "right join hit.dbo.nums n on a.codi=n.num "
            sql = sql & "left join ServitTmp e on a.codi=e.codiarticle "
            sql = sql & "Where n.Num <= " & c & " "
            If Len(Viatge) > 0 Then sql = sql & " and Viatge = '" & Viatge & "' "
            sql = sql & "Group By n.num "
            sql = sql & "order by n.num"
        
            'sql = "select Sum(case Client when " & rs(0) & " then  e.quantitatServida else 0 end) q,a.codi  Codi from articles a inner join ServitTmp  e on a.codi=e.codiarticle  "
            'If Len(Viatge) > 0 Then sql = sql & " Where Viatge = '" & Viatge & "' "
            'sql = sql & " Group By Codi order by codi "

            Set rsC = rec(sql)
            With .Range(.Cells(2, i), .Cells(c + 1, i))
                .CopyFromRecordset rsC
                .NumberFormat = "0.00"
            End With
            i = i + 1
            Rs.MoveNext
        Wend
        .Columns(i).Clear
        With .Range("B2", .Cells(c + 1, 2))
            .Formula = "=SUM(RC[1]:RC[" & i - 3 & "])"
            .NumberFormat = "0.00"
        End With
  
    End If

  End With
  
  Set Hoja = Nothing

End Sub



Private Sub rellenaHojaDemAnual(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, ByVal nombre As String, Optional Viatge As String = "")
  Dim Rs As ADODB.Recordset
  Dim rsC As ADODB.Recordset
  Dim sql As String
  Dim i As Integer
  Dim c As Integer

  With Hoja
    
    .Name = Left(Replace(nombre, ":", " "), 30)
    
    'Set rsC = rec("select count(distinct valor1)+3 c from paramshw")
    
sql = "select * From [" & DonamNomTaulaServit(dia) & "] "
If Len(Viatge) > 0 Then sql = sql & " Where Viatge = '" & Viatge & "' "
Set Rs = rec(sql)
If Not Rs.EOF Then
    
    .Cells(1, 1).Value = dateToParam(dia, "-")
    
    Set rsC = rec("select count(distinct codi)+3 c from clients")
    c = rsC("c")
    
    'Set rs = rec("select nom from clients where codi in (select distinct valor1 from paramshw) order by codi")
    Set Rs = rec("select nom from clients order by codi")
    
    With .Cells(1, 2)
      .Value = "TOTALS"
      .Font.Bold = True
      .HorizontalAlignment = xlCenter
    End With
    
    With .Range("C1", .Cells(1, c))
      .Value = Rs.GetRows
      .Font.Bold = True
      .HorizontalAlignment = xlCenter
    End With

    With .Range("A1", .Cells(1, c))
      .Borders(xlEdgeBottom).Weight = xlMedium
      .Borders(xlEdgeTop).Weight = xlMedium
    End With

    Set rsC = rec("select max(codi)+2 c from articles")
    c = rsC("c")
    
    Set Rs = rec("select nom from (select codi,nom from articles union select num as codi,'** Eliminado' as nom from hit.dbo.nums where num <= (select max(codi) from articles) and not num in (select codi from articles)) ff order by codi")
    .Range("A2", .Cells(c, 1)).CopyFromRecordset Rs

    With .Range("A1", .Cells(c, 1))
      .Font.Bold = True
      .Borders(xlEdgeLeft).Weight = xlMedium
      .Borders(xlEdgeRight).Weight = xlMedium
    End With

    With .Range("B1", .Cells(c, 1))
      .Font.Bold = True
      .Borders(xlEdgeLeft).Weight = xlMedium
      .Borders(xlEdgeRight).Weight = xlMedium
    End With
    
    i = 3
    'Set rs = rec("select distinct valor1 from paramshw order by valor1")
        Set Rs = rec("select distinct codi from clients order by codi")
        While Not Rs.EOF
            sql = "select sum(q) q,codi From (select isnull(e.quantitatdemanada,0) q,a.codi  Codi from articles a inner join [" & DonamNomTaulaServit(dia) & "] e on a.codi=e.codiarticle "
            If Len(Viatge) > 0 Then sql = sql & " And Viatge = '" & Viatge & "' "
            sql = sql & " Where e.client=" & Rs("codi") & " Union select 0 q,num as codi From hit.dbo.nums where num <= (select max(codi) from Articles ) ) ff  Group By Codi order by codi"

            Set rsC = rec(sql)
            With .Range(.Cells(2, i), .Cells(c, i))
                .CopyFromRecordset rsC
                .NumberFormat = "0.00"
            End With
            i = i + 1
            Rs.MoveNext
        Wend
      .Columns(i).Clear
      With .Range("B1", .Cells(c, 2))
        .Formula = "=SUM(RC[1]:RC[" & i - 3 & "])"
        .NumberFormat = "0.00"
      End With
  
End If
  
  End With
  
  Set Hoja = Nothing

End Sub


Private Sub rellenaHojaMes(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, nom)
  Dim i As Integer
  Dim DiaS As Integer
  Dim Rs As ADODB.Recordset
  Dim rsC As ADODB.Recordset
  Dim data() As Double
  Dim mes As Integer
  Dim sql As String

  With Hoja

    mes = Month(dia)
    nom = meses(mes) & " " & Year(dia)
    .Name = nom

    For i = 1 To 31
      .Cells(i + 2, 1).Value = " " & Day(dia) & " de " & Month(dia)
      dia = DateAdd("d", -1, dia)
    Next
    dia = DateAdd("d", 31, dia)
    
    i = 2
    Set Rs = rec("select codi,nom from clients where codi in (select distinct valor1 from paramshw) and not codi in (select distinct valor1 from TpvEquivalents) order by codi")
    'Set Rs = rec("select codi,nom from clients order by codi")
    While Not Rs.EOF
      
      sql = Rs("codi")
      
      sql = "select " & _
            "  datepart(y,v.data) as d, " & _
            "  sum(case when v.data > m.data  then 0 else import end) as M, " & _
            "  sum(case when v.data > m.data  then import else 0 end) as T, " & _
            "  count(distinct case when v.data > m.data  then 0 else v.num_tick end) as Cm, " & _
            "  count(distinct case when v.data > m.data  then v.num_tick else 0 end) as Ct " & _
            " from (Select * from [" & NomTaulaVentas(DateAdd("m", -1, dia)) & "] union select * from [" & NomTaulaVentas(dia) & "]) v " & _
            "  join (select  min(data) Data " & _
            "        from (Select * from [" & NomTaulaMovi(DateAdd("m", -1, dia)) & "] union select * from [" & NomTaulaMovi(dia) & "]) k " & _
            "        where botiga=" & sql & " and tipus_moviment='Z' " & _
            "        group by datepart(y,data)) m " & _
            "  on datepart(y,v.data) = datepart(y,m.data) " & _
            " where v.Botiga = " & sql & _
            " group by datepart(y,v.data) " & _
            " order by datepart(y,v.data) "

      
      'Sql = "Select import M,0 T,day(data) d From [" & NomTaulaMovi(Dia) & "] where botiga=" & Sql & " and " & _
      '      "tipus_moviment='Z' and datepart(hh,data)>17 Union select 0 M,import T,day(data) From " & _
      '      "[" & NomTaulaMovi(Dia) & "] where botiga=" & Sql & " and tipus_moviment='Z' and datepart(hh,data)<=17 " & _
      '      "order by day(data)"
      
On Error GoTo nooo
Dim DiaTope
      DiaTope = DatePart("y", dia)
      DiaS = 31
      Set rsC = rec(sql)
If Not Rs.EOF Then
      ReDim data(DiaS, 3)
         While Not rsC.EOF
            If rsC("d") <= DiaTope And rsC("d") >= (DiaTope - 28) Then
               data(DiaTope - rsC("d"), 0) = rsC("M")
               data(DiaTope - rsC("d"), 1) = rsC("CM") - 1
               data(DiaTope - rsC("d"), 2) = rsC("T")
               data(DiaTope - rsC("d"), 3) = rsC("CT") - 1
            End If
           rsC.MoveNext
         Wend
    
      .Range(.Cells(2, i), .Cells(3 + DiaS, i + 3)).Value = data
      .Cells(1, i).Value = Rs("nom")
      .Cells(2, i).Value = "MATÍ  €"
      .Cells(2, i + 1).Value = "c"
      .Cells(2, i + 2).Value = "TARDA  €"
      .Cells(2, i + 3).Value = "c"
      .Range(.Cells(1, i), .Cells(1, i + 3)).Merge

      With .Range(.Cells(1, i), .Cells(Day(dia) + 2, i))
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlThin
      End With

      Rs.MoveNext
      i = i + 4
End If
    
    Wend
    .Range(.Cells(1, i), .Cells(Day(dia) + 2, i)).Borders(xlEdgeLeft).Weight = xlMedium
    .Range("A1", .Cells(2, i - 1)).Borders(xlEdgeBottom).Weight = xlMedium
    .Range("A" & Day(dia) + 2, .Cells(Day(dia) + 2, i - 1)).Borders(xlEdgeBottom).Weight = xlMedium
    .Range("A1:A" & Day(dia) + 2).Borders(xlEdgeLeft).Weight = xlMedium
    .Range("A1", .Cells(1, i - 1)).Borders(xlEdgeTop).Weight = xlMedium

    With .Range("B1", .Cells(2, i))
      .Font.Bold = True
      .HorizontalAlignment = xlCenter
    End With

    With .Range("A3:A" & Day(dia) + 2)
      .Font.Bold = True
      .NumberFormat = "0"
    End With

    .Columns("A:A").ColumnWidth = 4

  End With
nooo:

  Set Hoja = Nothing

End Sub





 Private Sub RellenaHojaMesResumen(ByRef x As Object, dia As Date, Nom1, Nom2)
  Dim i As Integer, K As Integer, Kk As Integer, Acumu1 As Double, Acumu2 As Double, Acumu As Double, AcumuStr As String
  Dim DiaS As Integer, ii As Integer
  Dim Rs As ADODB.Recordset
  Dim rsC As ADODB.Recordset
  Dim data() As Double
  Dim mes As Integer
  Dim sql As String
  Dim Acu1 As Double, Acu2 As Double
  Dim AcuDiners As Double, AcuClients As Double, AcuDinersImp1 As Double, AcuClientsImp1 As Double, AcuDinersImp2 As Double, AcuClientsImp2 As Double, StrParcial As String, Parcial As Double
  Dim StC1 As Double, StE1 As Double, StC2 As Double, StE2 As Double
  
On Error GoTo nooo
  
  With x.Sheets(3)

    mes = Month(dia)
    x.Sheets(3).Name = "Parcials  " & meses(mes) & " a " & Format(Now, "dd mm yyyy")
    x.Sheets(4).Name = "Resum " & meses(mes) & " a " & Format(Now, "dd mm yyyy")

    dia = CDate("01/" & mes & "/" & Year(dia))
    DiaS = Day(DateAdd("d", -1, DateAdd("m", 1, dia)))
    While Month(dia) = mes
      .Cells(Day(dia) + 2, 1).Value = Day(dia)
      dia = DateAdd("d", 1, dia)
    Wend

    dia = DateAdd("d", -1, dia)

    i = 2
    ii = 4
    StC1 = 0
    StC2 = 0
    StE1 = 0
    StE2 = 0
    Set Rs = rec("select codi,nom from clients where codi in (select distinct valor1 from paramshw) and not codi in (select distinct valor1 from TpvEquivalents) order by codi ")
    While Not Rs.EOF
      sql = Rs("codi")

      AcuDiners = 0
      AcuClients = 0
      
      AcuDinersImp1 = 0
      AcuDinersImp2 = 0
      
      AcuClientsImp1 = 0
      AcuClientsImp2 = 0
      
      For Kk = i To i + 3  'Per cada concepte
        Acumu1 = 0
        Acumu2 = 0
        For K = 28 + 2 To 3 Step -1 ' Per cada dia
            Acu1 = 0
            Acu2 = 0
            If Not IsEmpty(x.Sheets(1).Cells(K, Kk).Value) Then If x.Sheets(1).Cells(K, Kk).Value > 0 Then Acu1 = x.Sheets(1).Cells(K, Kk).Value
            If Not IsEmpty(x.Sheets(2).Cells(K, Kk).Value) Then If x.Sheets(2).Cells(K, Kk).Value > 0 Then Acu2 = x.Sheets(2).Cells(K, Kk).Value
            Acumu1 = Acumu1 + Acu1
            Acumu2 = Acumu2 + Acu2
            Acumu = Int((Acumu1 / IIf(Acumu2 = 0, 1, Acumu2) - 1) * 100)
            Parcial = Int((Acu1 / IIf(Acu2 = 0, 1, Acu2) - 1) * 100)
'            x.Sheets(3).Cells(k, kk).Value = k
            StrParcial = Parcial
            If Not (Acu2 = 0 Or Acu1 = 0) Then StrParcial = ""
            x.Sheets(3).Cells(K, Kk).Value = Acumu & " (" & StrParcial & ") %"
            
            If Acumu < 0 Then x.Sheets(3).Cells(K, Kk).Font.ColorIndex = 3
            If Acumu > 10 Then x.Sheets(3).Cells(K, Kk).Font.ColorIndex = 10
        Next
        AcumuStr = ""
        
        If Not (Acumu2 = 0 And Acumu1 = 0) Then AcumuStr = Acumu & " %"
        If Kk = i + 0 Then .Cells(2, i).Value = "MATÍ  € " & AcumuStr
        If Kk = i + 1 Then .Cells(2, i + 1).Value = "c " & AcumuStr
        If Kk = i + 2 Then .Cells(2, i + 2).Value = "TARDA  € " & AcumuStr
        If Kk = i + 3 Then .Cells(2, i + 3).Value = "c " & AcumuStr
        
        If Kk = i + 0 Or Kk = i + 2 Then '' parlem de diners
           AcuDiners = AcuDiners + Acumu
           AcuDinersImp1 = AcuDinersImp1 + Acumu1
           AcuDinersImp2 = AcuDinersImp2 + Acumu2
        End If
        
        If Kk = i + 1 Or Kk = i + 3 Then  ' Parlem de clients
            AcuClients = AcuClients + Acumu
            AcuClientsImp1 = AcuClientsImp1 + Acumu1
            AcuClientsImp2 = AcuClientsImp2 + Acumu2
        End If
      Next
      
      With .Range(.Cells(1, i), .Cells(Day(dia) + 2, i))
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlThin
      End With
      .Cells(1, i).Value = Rs("nom") & " " & AcuDiners & " € " & AcuClients & " c"
      .Range(.Cells(1, i), .Cells(1, i + 3)).Merge
      
      x.Sheets(4).Cells(ii, 1) = Rs("nom")
      x.Sheets(4).Cells(ii, 2) = Int((AcuDinersImp1 / IIf(AcuDinersImp2 = 0, 1, AcuDinersImp2) - 1) * 100) & " %"
      If Int((AcuDinersImp1 / IIf(AcuDinersImp2 = 0, 1, AcuDinersImp2) - 1) * 100) < 0 Then x.Sheets(4).Cells(ii, 2).Font.ColorIndex = 3
      If Int((AcuDinersImp1 / IIf(AcuDinersImp2 = 0, 1, AcuDinersImp2) - 1) * 100) > 10 Then x.Sheets(4).Cells(ii, 2).Font.ColorIndex = 10
      
      x.Sheets(4).Cells(ii, 3) = Int((AcuClientsImp1 / IIf(AcuClientsImp2 = 0, 1, AcuClientsImp2) - 1) * 100) & " %"
      If Int((AcuClientsImp1 / IIf(AcuClientsImp2 = 0, 1, AcuClientsImp2) - 1) * 100) < 0 Then x.Sheets(4).Cells(ii, 3).Font.ColorIndex = 3
      If Int((AcuClientsImp1 / IIf(AcuClientsImp2 = 0, 1, AcuClientsImp2) - 1) * 100) > 10 Then x.Sheets(4).Cells(ii, 3).Font.ColorIndex = 10
      
      x.Sheets(4).Cells(ii, 4) = AcuDinersImp1 & " € "
      x.Sheets(4).Cells(ii, 5) = AcuDinersImp2 & " €"
      x.Sheets(4).Cells(ii, 6) = AcuClientsImp1 & " c"
      x.Sheets(4).Cells(ii, 7) = AcuClientsImp2 & " c"
      StC1 = StC1 + AcuClientsImp1
      StC2 = StC2 + AcuClientsImp2
      StE1 = StE1 + AcuDinersImp1
      StE2 = StE2 + AcuDinersImp2
     
      Rs.MoveNext
      ii = ii + 1
      i = i + 4
    Wend
    
    x.Sheets(4).Cells(2, 1) = "Totals : "
    x.Sheets(4).Cells(2, 2) = Int((StE1 / IIf(StE2 = 0, 1, StE2) - 1) * 100) & " %"
    If Int((StE1 / IIf(StE2 = 0, 1, StE2) - 1) * 100) < 0 Then x.Sheets(4).Cells(2, 2).Font.ColorIndex = 3
    If Int((StE1 / IIf(StE2 = 0, 1, StE2) - 1) * 100) > 10 Then x.Sheets(4).Cells(2, 2).Font.ColorIndex = 10
    x.Sheets(4).Cells(2, 3) = Int((StC1 / IIf(StC2 = 0, 1, StC2) - 1) * 100) & " %"
    If Int((StC1 / IIf(StC2 = 0, 1, StC2) - 1) * 100) < 0 Then x.Sheets(4).Cells(2, 3).Font.ColorIndex = 3
    If Int((StC1 / IIf(StC2 = 0, 1, StC2) - 1) * 100) > 10 Then x.Sheets(4).Cells(2, 3).Font.ColorIndex = 10
      
      
    x.Sheets(4).Cells(2, 4) = StE1 & " e "
    x.Sheets(4).Cells(2, 5) = StE2 & " e"
    x.Sheets(4).Cells(2, 6) = StC1 & " c"
    x.Sheets(4).Cells(2, 7) = StC2 & " c"
    
    x.Sheets(4).Cells(1, 2) = "Increment e "
    x.Sheets(4).Cells(1, 3) = "Increment c "
    x.Sheets(4).Cells(1, 4) = Year(dia) & " e "
    x.Sheets(4).Cells(1, 5) = Year(DateAdd("yyyy", -1, dia)) & " e "
    x.Sheets(4).Cells(1, 6) = Year(dia) & " c"
    x.Sheets(4).Cells(1, 7) = Year(DateAdd("yyyy", -1, dia)) & " c"
    
    x.Sheets(3).Range(x.Sheets(3).Cells(1, 1), x.Sheets(3).Cells(33, i)).HorizontalAlignment = xlRight
    With .Range("B1", .Cells(2, i))
      .Font.Bold = True
      .HorizontalAlignment = xlCenter
    End With

    With .Range("A3:A" & Day(dia) + 2)
      .Font.Bold = True
      .NumberFormat = "0"
    End With

    x.Sheets(4).Columns("A:A").EntireColumn.AutoFit
    x.Sheets(4).Columns("B:G").Select

    x.Sheets(4).Columns("B:G").HorizontalAlignment = xlRight
    x.Sheets(4).Columns("B:G").VerticalAlignment = xlBottom
    x.Sheets(4).Columns("B:G").WrapText = False
    x.Sheets(4).Columns("B:G").Orientation = 0
    x.Sheets(4).Columns("B:G").AddIndent = False
    x.Sheets(4).Columns("B:G").ShrinkToFit = False
    x.Sheets(4).Columns("B:G").MergeCells = False
    x.Sheets(4).Range("A1").Select

  End With
nooo:

  'Set HOJA = Nothing

End Sub






Private Sub rellenaHojaVenut(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, Df As Date, clients As String)
    Dim i As Integer, DiaS As Integer, Rs As ADODB.Recordset, rsC As ADODB.Recordset, data() As Double, mes As Integer, sql As String, D As Date, Boti As Double
    
    ExecutaComandaSql "Drop table TemporalExcel"
    ExecutaComandaSql "Create Table [TemporalExcel] (Botiga float,Data datetime,Dependenta float,Num_tick float,Estat [nvarchar] (25), Plu float, Quantitat float, Import float,Tipus_venta [nvarchar] (25),FormaMarcar [nvarchar] (255) Default (''),Otros [nvarchar] (255) Default ('')) "
    
    D = dia
    While D < Df
        ExecutaComandaSql "Insert into [TemporalExcel] Select * from [" & NomTaulaVentas(dia) & "] where botiga in(" & clients & ") and day(data)= " & Day(D) & "  "
        D = DateAdd("D", 1, D)
        DoEvents
    Wend
    
    mes = Month(dia)
    Hoja.Name = "Venut"

    dia = CDate("01/" & mes & "/" & Year(dia))
    DiaS = Day(DateAdd("d", -1, DateAdd("m", 1, dia)))

    For i = 0 To 24
        Hoja.Cells(i + 4, 1).Value = i & " h."
    Next

    i = 2
    ReDim dataaCc(24, 3)
    ReDim data(24, 3)
    Set Rs = rec("select codi,nom from clients where codi in (select distinct valor1 from paramshw) order by codi")
    While Not Rs.EOF
        Boti = Rs("codi")
        sql = "Select datepart(hh,data),sum(import) as i,count(distinct num_tick) as c ,count(distinct dependenta) as d from TemporalExcel "
        sql = sql & "where botiga = " & Boti & " group by datepart(hh,data) "
            
        On Error GoTo nooo
        Set rsC = rec(sql)
        If Not Rs.EOF Then
            While Not rsC.EOF
                data(rsC(0), 0) = rsC("i")
                data(rsC(0), 1) = rsC("c")
                data(rsC(0), 2) = rsC("d")
                dataaCc(rsC(0), 0) = dataaCc(rsC(0), 0) + rsC("i")
                dataaCc(rsC(0), 1) = dataaCc(rsC(0), 1) + rsC("c")
                dataaCc(rsC(0), 2) = dataaCc(rsC(0), 2) + rsC("d")
           
                rsC.MoveNext
            Wend
    
            Hoja.Range(Hoja.Cells(4, i), Hoja.Cells(4 + 24, i + 2)).Value = data
            Hoja.Cells(1, i).Value = Rs("nom")
            Hoja.Cells(2, i).Value = "Import"
            Hoja.Cells(2, i + 1).Value = "Clients"
            Hoja.Cells(2, i + 2).Value = "Dependentes"
            Hoja.Range(Hoja.Cells(1, i), Hoja.Cells(1, i + 2)).Merge
            Hoja.Range(Hoja.Cells(1, i), Hoja.Cells(Day(dia) + 2, i)).Borders(xlEdgeLeft).Weight = xlMedium
            Hoja.Range(Hoja.Cells(1, i), Hoja.Cells(Day(dia) + 2, i)).Borders(xlEdgeRight).Weight = xlThin
            Rs.MoveNext
            i = i + 3
        End If
    Wend
    Hoja.Range(Hoja.Cells(1, i), Hoja.Cells(Day(dia) + 2, i)).Borders(xlEdgeLeft).Weight = xlMedium
    Hoja.Range("A1", Hoja.Cells(2, i - 1)).Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Range("A" & Day(dia) + 2, Hoja.Cells(Day(dia) + 2, i - 1)).Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Range("A1:A" & Day(dia) + 2).Borders(xlEdgeLeft).Weight = xlMedium
    Hoja.Range("A1", Hoja.Cells(1, i - 1)).Borders(xlEdgeTop).Weight = xlMedium
    Hoja.Range("B1", Hoja.Cells(2, i)).Font.Bold = True
    Hoja.Range("B1", Hoja.Cells(2, i)).HorizontalAlignment = xlCenter
    Hoja.Range("A3:A" & Day(dia) + 2).Font.Bold = True
    Hoja.Range("A3:A" & Day(dia) + 2).NumberFormat = "0"
    Hoja.Columns("A:A").ColumnWidth = 4
    Hoja.Cells(2 + 24 + 5, 2).Value = "Des de : "
    Hoja.Cells(2 + 24 + 6, 2).Value = "Fins A : "
    Hoja.Cells(2 + 24 + 5, 3).Value = dia
    Hoja.Cells(2 + 24 + 6, 3).Value = Df
    
nooo:
  Set Hoja = Nothing
End Sub






Private Sub rellenaHojaVenutMes(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, clients As String)
    Dim i As Integer, DiaS As Integer, Rs As ADODB.Recordset, rsC As ADODB.Recordset, data() As Double, mes As Integer, sql As String, D As Date, Boti As Double
    
On Error GoTo nooo
    ExecutaComandaSql "Drop table TemporalExcel"
    ExecutaComandaSql "Drop table TemporalExcelCaixes"
   
    AjuntaTpvS Year(dia), Month(dia), Day(dia)
    
    
' [" & NomTaulaMovi(dia) & "]
    sql = "select max(isnull(dm.motiu,'')) + max(isnull(dt.motiu,'')) Explicacio,min(z.d) Z,z.botiga botiga ,isnull(sum(dM.import),0) D_Mati,isnull(sum(dT.import),0) D_Tarda "
    sql = sql & "Into TemporalExcelCaixes  From "
    sql = sql & "(Select min(data) d ,Botiga from [" & NomTaulaMovi(dia) & "] z where tipus_moviment = 'Z' group by day(data),Botiga) z left join "
    sql = sql & "(Select Data d ,Botiga,Import,motiu from [" & NomTaulaMovi(dia) & "] z where tipus_moviment = 'J') Dm on day(dm.d) = day(z.d) and dm.Botiga = z.Botiga and dm.d <= z.d left join "
    sql = sql & "(Select Data d ,Botiga,Import,motiu from [" & NomTaulaMovi(dia) & "] z where tipus_moviment = 'J') Dt on day(dt.d) = day(z.d) and dt.Botiga = z.Botiga and dt.d > z.d "
    sql = sql & "group by z.botiga,day(z.d) "

    ExecutaComandaSql sql

    sql = "select Explicacio,day(v.data) vData,v.botiga Botiga , "
    sql = sql & "Case when v.data<isnull(z.Z,v.data) then v.import else 0 end Z_Mati,d_Mati, "
    sql = sql & "Case when v.data<isnull(z.Z,v.data) then v.num_tick else 0 end C_Mati, "
    sql = sql & "Case when v.data>=isnull(z.Z,v.data) then v.import else 0 end Z_Tarda,d_Tarda, "
    sql = sql & "Case when v.data>=isnull(z.Z,v.data) then v.num_tick else 0 end C_Tarda "
    sql = sql & "Into TemporalExcel "
    sql = sql & "from [" & NomTaulaVentas(dia) & "] v left join TemporalExcelCaixes z "
    sql = sql & "on v.botiga = z.botiga and day(v.data)=day(z.Z) "
    ExecutaComandaSql sql
    
    sql = "select c.nom," & Month(dia) & "," & Year(dia) & ",cast(dia as nvarchar) + '/" & Month(dia) & "/" & Year(dia) & "' , dia,z_mati+z_tarda+D_Mati+D_Tarda,z_mati+z_tarda,z_mati,D_Mati,z_tarda,D_Tarda,c_mati,c_tarda,c_mati + c_tarda , explicacio   from "
    sql = sql & "( "
    sql = sql & "select Botiga,vdata dia ,sum(z_mati + z_tarda) Z,sum(Z_Mati) Z_Mati,max(D_Mati) D_Mati, "
    sql = sql & "count(distinct C_Mati) C_Mati, "
    sql = sql & "sum(Z_Tarda) Z_Tarda,Max(D_Tarda) D_Tarda, "
    sql = sql & "count(distinct C_Tarda) C_Tarda, "
    sql = sql & "count(distinct C_Tarda + C_Mati) C, "
    sql = sql & "max(explicacio) explicacio "
    sql = sql & "from TemporalExcel v "
    sql = sql & "group by Botiga,vdata "
    sql = sql & ") d join clients c on c.codi = d.botiga "
    sql = sql & "order by c.nom,dia "
    Set Rs = rec(sql)
    
    mes = Month(dia)
    Hoja.Name = Format(dia, "mmmm yyyy")
    Hoja.Cells(1, 1).Value = "Tienda"
    Hoja.Cells(1, 2).Value = "Mes"
    Hoja.Cells(1, 3).Value = "Año"
    Hoja.Cells(1, 4).Value = "Fecha"
    Hoja.Cells(1, 5).Value = "Dia"
    Hoja.Cells(1, 6).Value = "Recaudat"
    Hoja.Cells(1, 7).Value = "Total"
    Hoja.Cells(1, 8).Value = "Mati"
    Hoja.Cells(1, 9).Value = "Desc Mati"
    Hoja.Cells(1, 10).Value = "Tarda"
    Hoja.Cells(1, 11).Value = "Desc Tarda"
    Hoja.Cells(1, 12).Value = "Cl Mati"
    Hoja.Cells(1, 13).Value = "Cl Tarda"
    Hoja.Cells(1, 14).Value = "Total Clients"
    Hoja.Cells(1, 15).Value = "Explicacio"
    Hoja.Range("A2").CopyFromRecordset Rs
    Hoja.Range("A1..Z1").Font.Bold = True
    Hoja.Cells.EntireColumn.AutoFit
nooo:

End Sub
Private Sub rellenaHojaDiaDeLaSetmana2(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, client As String)
    Dim i As Integer, DiaS As Integer, Rs As ADODB.Recordset, rsC As ADODB.Recordset, data() As Double, mes As Integer, sql As String, D As Date, Boti As Double

    ExecutaComandaSql "Drop table TemporalExcel"
    ExecutaComandaSql "Drop table TemporalExcelCaixes"
    
' [" & NomTaulaMovi(dia) & "]

    ExecutaComandaSql " select min(data) z ,Botiga Into TemporalExcelCaixes from [" & NomTaulaMovi(dia) & "] z where tipus_moviment = 'Z' group by day(data),Botiga"
    ExecutaComandaSql " Insert Into TemporalExcelCaixes select min(data) z ,Botiga from [" & NomTaulaMovi(DateAdd("m", -1, dia)) & "] z where tipus_moviment = 'Z' group by day(data),Botiga"

    For i = 1 To 5
        D = DateAdd("d", -(i * 7), dia)
        sql = ""
        If i > 1 Then sql = sql & "Into TemporalExcel "
        sql = sql & "select day(v.data) vData,v.botiga Botiga , "
        sql = sql & "Case when v.data<isnull(z.Z,v.data) then v.import else 0 end Z_Mati, "
        sql = sql & "Case when v.data<isnull(z.Z,v.data) then v.num_tick else 0 end C_Mati, "
        sql = sql & "Case when v.data>=isnull(z.Z,v.data) then v.import else 0 end Z_Tarda, "
        sql = sql & "Case when v.data>=isnull(z.Z,v.data) then v.num_tick else 0 end C_Tarda "
        If i = 1 Then sql = sql & "Into TemporalExcel "
        sql = sql & "from [" & NomTaulaVentas(dia) & "] v left join TemporalExcelCaixes z "
        sql = sql & "on v.botiga = z.botiga and day(v.data)=day(z.Z) and day(v.data) = " & Day(D) & " "
        ExecutaComandaSql sql
    Next
    
    sql = "select c.nom," & Month(dia) & "," & Year(dia) & ",cast(dia as nvarchar) + '/" & Month(dia) & "/" & Year(dia) & "' , dia,z_mati+z_tarda,z_mati,z_tarda,c_mati,c_tarda,c_mati + c_tarda   from "
    sql = sql & "( "
    sql = sql & "select Botiga,vdata dia ,sum(z_mati + z_tarda) Z,sum(Z_Mati) Z_Mati, "
    sql = sql & "count(distinct C_Mati) C_Mati, "
    sql = sql & "sum(Z_Tarda) Z_Tarda, "
    sql = sql & "count(distinct C_Tarda) C_Tarda, "
    sql = sql & "count(distinct C_Tarda + C_Mati) C "
    sql = sql & "from TemporalExcel v "
    sql = sql & "group by Botiga,vdata "
    sql = sql & ") d join clients c on c.codi = d.botiga "
    sql = sql & "order by c.nom,dia "
    Set Rs = rec(sql)
    
    mes = Month(dia)
    Hoja.Name = Format(dia, "mmmm yyyy")
    Hoja.Cells(1, 1).Value = "Tienda"
    Hoja.Cells(1, 2).Value = "Mes"
    Hoja.Cells(1, 3).Value = "Año"
    Hoja.Cells(1, 4).Value = "Fecha"
    Hoja.Cells(1, 5).Value = "Dia"
    Hoja.Cells(1, 6).Value = "Total"
    Hoja.Cells(1, 7).Value = "Mati"
    Hoja.Cells(1, 8).Value = "Tarda"
    Hoja.Cells(1, 9).Value = "Cl Mati"
    Hoja.Cells(1, 10).Value = "Cl Tarda"
    Hoja.Cells(1, 11).Value = "Total Clients"
    Hoja.Range("A2").CopyFromRecordset Rs
    Hoja.Range("A1..Z1").Font.Bold = True
    
'    HOJA.Range(HOJA.Cells(1, i), HOJA.Cells(Day(dia) + 2, i)).Borders(xlEdgeLeft).Weight = xlMedium
'    HOJA.Range("A1", HOJA.Cells(2, i - 1)).Borders(xlEdgeBottom).Weight = xlMedium
'    HOJA.Range("A" & Day(dia) + 2, HOJA.Cells(Day(dia) + 2, i - 1)).Borders(xlEdgeBottom).Weight = xlMedium
'    HOJA.Range("A1:A" & Day(dia) + 2).Borders(xlEdgeLeft).Weight = xlMedium
'    HOJA.Range("A1", HOJA.Cells(1, i - 1)).Borders(xlEdgeTop).Weight = xlMedium
'    HOJA.Range("B1", HOJA.Cells(2, i)).Font.Bold = True
'    HOJA.Range("B1", HOJA.Cells(2, i)).HorizontalAlignment = xlCenter
'    HOJA.Range("A3:A" & Day(dia) + 2).Font.Bold = True
'    HOJA.Range("A3:A" & Day(dia) + 2).NumberFormat = "0"
'    HOJA.Columns("A:A").ColumnWidth = 4
'    HOJA.Cells(2 + 24 + 5, 2).Value = "Des de : "
'    HOJA.Cells(2 + 24 + 6, 2).Value = "Fins A : "
'    HOJA.Cells(2 + 24 + 5, 3).Value = dia
'    HOJA.Cells(2 + 24 + 6, 3).Value = Df
    
nooo:
  'Set HOJA = Nothing
End Sub

Private Sub rellenaHojaDiaDeLaSetmana(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, client, ByRef Resum As Excel.Worksheet, NumTi)
    Dim i As Integer, DiaS As Integer, Rs As rdoResultset, rsC As ADODB.Recordset, data() As Double, mes As Integer, sql As String, D As Date, Boti As Double
    Dim cM, cT, zM, zT, hM, hT, DepsM(), DepsT(), DescM, DescT, Families(), FamiliesPct(), FamiliesPctAcu(), GraficX(), GraficY()
    Dim cMm, cTm, zMm, zTm, hMm, hTm, Col, cha As Chart, Kk As Object, MaxFam, j, Se, Dv, SeM, DvM
    On Error Resume Next
    Hoja.Name = Join(Split(BotigaCodiNom(client), " "), "")
    Hoja.Cells(1, 3).Value = Join(Split(BotigaCodiNom(client), " "), "")
    Hoja.Cells(1, 1).Value = Format(dia, "dd/mm/yyyy dddd")
    Hoja.Cells(4, 1).Value = "Mañana"
    
    Hoja.Rows("4:4").Select
    Hoja.Rows("4:4").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Rows("4:4").Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Rows("4:4").Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    Hoja.Rows("4:4").RowHeight = 30
    
    Hoja.Cells(4, 1).Font.Bold = True
    Hoja.Cells(5, 2).Value = "Recaudacion"
    Hoja.Cells(6, 2).Value = "Clientes"
    Hoja.Cells(7, 2).Value = "Media tiket"
    Hoja.Cells(8, 2).Value = "Horas"
    Hoja.Cells(9, 2).Value = "Rec/Hora"
    Hoja.Cells(10, 2).Value = "Descuadre"
    Hoja.Cells(13, 1).Value = "Tarde"
    Hoja.Rows("13:13").RowHeight = 30
    Hoja.Rows("13:13").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Rows("13:13").Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Rows("13:13").Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    Hoja.Cells(13, 1).Font.Bold = True
    Hoja.Cells(14, 2).Value = "Recaudacion"
    Hoja.Cells(15, 2).Value = "Clientes"
    Hoja.Cells(16, 2).Value = "Media tiket"
    Hoja.Cells(17, 2).Value = "Horas"
    Hoja.Cells(18, 2).Value = "Rec/Hora"
    
    Hoja.Rows("19:19").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(20, 1).Value = "Euros"
    Hoja.Cells(20, 1).Font.Bold = True
    Hoja.Cells(20, 2).Value = "Devolucion"
    Hoja.Cells(21, 2).Value = "Servido"
    Hoja.Cells(22, 2).Value = "Ingreso"
    Hoja.Cells(23, 2).Value = "(Err) Retorno"
    Hoja.Rows("23:23").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Hoja.Cells(24, 1).Value = "Familia"
    Hoja.Cells(24, 1).Font.Bold = True
    
    Set Rs = Db.OpenResultset("Select * from families Where Pare = 'Article' Order by nom ")
    ReDim Families(0)
    ReDim FamiliesPct(0)
    ReDim FamiliesPctAcu(0)
    While Not Rs.EOF
        ReDim Preserve Families(UBound(Families) + 1)
        ReDim Preserve FamiliesPct(UBound(FamiliesPct) + 1)
        ReDim Preserve FamiliesPctAcu(UBound(FamiliesPctAcu) + 1)
        FamiliesPct(UBound(FamiliesPct)) = 0
        FamiliesPctAcu(UBound(FamiliesPctAcu)) = 0
        Families(UBound(Families)) = Rs("Nom")
        Rs.MoveNext
    Wend

    cMm = 0: cTm = 0: zMm = 0: zTm = 0: hMm = 0: hTm = 0: MaxFam = 0: SeM = 0: DvM = 0
    
    Col = 2
    
    For i = 4 To 1 Step -1
        Col = Col + 1
        rellenaHojaDiaDeLaSetmanaBuscaDades DateAdd("d", -(i * 7), dia), client, cM, cT, zM, zT, hM, hT, DepsM, DepsT, DescM, DescT, Families, FamiliesPct, Dv, Se
        If DatePart("w", DateAdd("d", -(i * 7), dia)) = 1 Then Resum.Cells(3, Col).Font.ColorIndex = 46
        Hoja.Cells(3, Col).Value = Format(DateAdd("d", -(i * 7), dia), "dd/mm/yyyy") & Chr(10) & Format(DateAdd("d", -(i * 7), dia), "dddd")
        
        If UBound(DepsM) > 0 Then Hoja.Cells(4, Col).Value = Right(Join(DepsM, Chr(10)), Len(Join(DepsM, Chr(10))) - 1)
        Hoja.Cells(5, Col).Value = zM
        Hoja.Cells(6, Col).Value = cM
        Hoja.Cells(7, Col).Value = Round((zM / cM), 2)
        Hoja.Cells(8, Col).Value = Int(hM / 60)
        Hoja.Cells(20, Col).Value = Dv
        Hoja.Cells(21, Col).Value = Se
        Hoja.Cells(22, Col).Value = zM + zT
        If (Se) > 0 Then Hoja.Cells(23, Col).Value = "(" & Round(1 - (zM + zT + Dv) / Se, 2) & ") " & Int((zM + zT) / Se * 100) - 100 & " %"
                
        If hM > 0 Then Hoja.Cells(9, Col).Value = Round((zM / (hM / 60)), 0)
        Hoja.Cells(10, Col).Value = DescM
            
        If UBound(DepsT) > 0 Then Hoja.Cells(13, Col).Value = Right(Join(DepsT, Chr(10)), Len(Join(DepsT, Chr(10))) - 1)
        Hoja.Cells(14, Col).Value = zT
        Hoja.Cells(15, Col).Value = cT
        Hoja.Cells(16, Col).Value = Round((zT / cT), 2)
        Hoja.Cells(17, Col).Value = Int(hT / 60)
        If hT > 0 Then Hoja.Cells(18, Col).Value = Round((zT / (hT / 60)), 0)
        Hoja.Cells(19, Col).Value = DescT
        
        For j = 1 To UBound(Families)
            FamiliesPctAcu(j) = FamiliesPctAcu(j) + FamiliesPct(j)
            Hoja.Cells(24 + j, 2).Value = Families(j)
            Hoja.Cells(24 + j, Col).Value = Round(FamiliesPct(j)) & " %"
            If MaxFam < j Then MaxFam = j
        Next
        cMm = cMm + cM
        cTm = cTm + cT
        zMm = zMm + zM
        zTm = zTm + zT
        hMm = hMm + hM
        hTm = hTm + hT
        SeM = SeM + Se
        DvM = DvM + Dv
    Next
    
    Col = Col + 1
    Hoja.Cells(3, Col).Value = "Medias"
    Hoja.Cells(3, Col).Font.Bold = True
    Hoja.Cells(5, Col).Value = Int(zMm / 4)
    Hoja.Cells(6, Col).Value = Int(cMm / 4)
    If cMm > 0 Then Hoja.Cells(7, Col).Value = Round((zMm / cMm), 2)
    Hoja.Cells(8, Col).Value = Int(hMm / 60 / 4)
    If hMm > 0 Then Hoja.Cells(9, Col).Value = Round((zMm / (hMm / 60)), 0)
    Hoja.Cells(14, Col).Value = Int(zTm / 4)
    Hoja.Cells(15, Col).Value = Int(cTm / 4)
    If cTm > 0 Then Hoja.Cells(16, Col).Value = Round((zTm / cTm), 2)
    Hoja.Cells(17, Col).Value = Int(hTm / 60 / 4)
    If hTm > 0 Then Hoja.Cells(18, Col).Value = Round((zTm / (hTm / 60)), 0)
    
    Hoja.Cells(20, Col).Value = Int(DvM / 4)
    Hoja.Cells(21, Col).Value = Int(SeM / 4)
    Hoja.Cells(22, Col).Value = Int((zMm + zTm) / 4)
        
    
    Resum.Cells(2, NumTi).Value = Join(Split(BotigaCodiNom(client), " "), "")
    Resum.Cells(2, NumTi).Font.Bold = True
    rellenaHojaDiaDeLaSetmanaBuscaDades dia, client, cM, cT, zM, zT, hM, hT, DepsM, DepsT, DescM, DescT, Families, FamiliesPct, Dv, Se
    Resum.Cells(3, NumTi).Value = dia
'    If UBound(DepsM) > 0 Then Resum.Cells(4, NumTi).Value = Right(Join(DepsM, Chr(10)), Len(Join(DepsM, Chr(10))) - 1)
    Resum.Cells(5, NumTi).Value = zM
    Resum.Cells(6, NumTi).Value = cM
    Resum.Cells(7, NumTi).Value = Round((zM / cM), 2)
    Resum.Cells(8, NumTi).Value = Int(hM / 60)
    Resum.Cells(20, NumTi).Value = Dv
    Resum.Cells(21, NumTi).Value = Se
    Resum.Cells(22, NumTi).Value = zM + zT
    If (Se) > 0 Then Resum.Cells(23, NumTi).Value = "(" & Round(1 - (zM + zT + Dv) / Se, 2) & ") " & Int((zM + zT) / Se * 100) - 100 & " %"
    
    If hM > 0 Then Resum.Cells(9, NumTi).Value = Round((zM / (hM / 60)), 0)
    Resum.Cells(10, NumTi).Value = DescM
            
'    If UBound(DepsT) > 0 Then Resum.Cells(13, NumTi).Value = Right(Join(DepsT, Chr(10)), Len(Join(DepsT, Chr(10))) - 1)
    Resum.Cells(14, NumTi).Value = zT
    Resum.Cells(15, NumTi).Value = cT
    Resum.Cells(16, NumTi).Value = Round((zT / cT), 2)
    Resum.Cells(17, NumTi).Value = Int(hT / 60)
    If hT > 0 Then Resum.Cells(18, NumTi).Value = Round((zT / (hT / 60)), 0)
    Resum.Cells(19, NumTi).Value = DescT
    For j = 1 To UBound(Families)
        Resum.Cells(24 + j, 2).Value = Families(j)
        Resum.Cells(24 + j, NumTi).Value = Int(FamiliesPct(j)) & " %"
        If FamiliesPctAcu(j) > 0 Then Resum.Cells(24 + j, NumTi + 1).Value = Int(FamiliesPct(j) / (FamiliesPctAcu(j) / 4) * 100) - 100 & " %"
        If FamiliesPctAcu(j) > 0 Then Resum.Cells(24 + j, NumTi - 1).Value = Int((FamiliesPctAcu(j) / 4)) & " %"
        If MaxFam < j Then MaxFam = j
    Next
    
'    Resum.Cells(3, NumTi + 1).Value = "Desvio"
    Resum.Cells(3, NumTi + 1).Font.Bold = True
    Resum.Cells(3, NumTi + 1).Font.Italic = True
    If zMm > 0 Then Resum.Cells(5, NumTi + 1).Value = Int((zM / Int(zMm / 4) - 1) * 100)
    If Resum.Cells(5, NumTi + 1).Value > 0 Then Resum.Cells(5, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(5, NumTi + 1).Font.ColorIndex = 3
    If cMm > 0 Then Resum.Cells(6, NumTi + 1).Value = Int((cM / Int(cMm / 4) - 1) * 100)
    If Resum.Cells(6, NumTi + 1).Value > 0 Then Resum.Cells(6, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(6, NumTi + 1).Font.ColorIndex = 3
    If cMm > 0 Then Resum.Cells(7, NumTi + 1).Value = Int((Round((zM / cM), 2) / Round((zMm / cMm), 2) - 1) * 100)
    If Resum.Cells(7, NumTi + 1).Value > 0 Then Resum.Cells(7, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(7, NumTi + 1).Font.ColorIndex = 3
    If hMm > 0 Then Resum.Cells(8, NumTi + 1).Value = Int(((hM) / (hMm / 4) - 1) * 100)
    If Resum.Cells(8, NumTi + 1).Value < 0 Then Resum.Cells(8, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(8, NumTi + 1).Font.ColorIndex = 3
    If hM > 0 Then If hMm > 0 Then Resum.Cells(9, NumTi + 1).Value = Int(((zM / hM) / (zMm / hMm) - 1) * 100)
    If Resum.Cells(9, NumTi + 1).Value > 0 Then Resum.Cells(9, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(9, NumTi + 1).Font.ColorIndex = 3
    
    If zTm > 0 Then Resum.Cells(14, NumTi + 1).Value = Int((zT / Int(zTm / 4) - 1) * 100)
    If Resum.Cells(14, NumTi + 1).Value > 0 Then Resum.Cells(14, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(14, NumTi + 1).Font.ColorIndex = 3
    If Int(cTm / 4) And cTm > 0 And cT > 0 Then Resum.Cells(15, NumTi + 1).Value = Int((cT / Int(cTm / 4) - 1) * 100)
    If Resum.Cells(15, NumTi + 1).Value > 0 Then Resum.Cells(15, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(15, NumTi + 1).Font.ColorIndex = 3
    If cTm > 0 Then Resum.Cells(16, NumTi + 1).Value = Int((Round((zT / cT), 2) / Round((zTm / cTm), 2) - 1) * 100)
    If Resum.Cells(16, NumTi + 1).Value > 0 Then Resum.Cells(16, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(16, NumTi + 1).Font.ColorIndex = 3
    If hTm > 0 Then Resum.Cells(17, NumTi + 1).Value = Int(((hT) / (hTm / 4) - 1) * 100)
    If Resum.Cells(17, NumTi + 1).Value > 0 Then Resum.Cells(17, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(17, NumTi + 1).Font.ColorIndex = 3
    If hT > 0 And hTm > 0 And zTm > 0 Then Resum.Cells(18, NumTi + 1).Value = Int(((zT / hT) / (zTm / hTm) - 1) * 100)
    If Resum.Cells(18, NumTi + 1).Value > 0 Then Resum.Cells(18, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(18, NumTi + 1).Font.ColorIndex = 3
    
    If DvM > 0 Then Resum.Cells(20, NumTi + 1).Value = Int((Dv / Int(DvM / 4) - 1) * 100)
    If Resum.Cells(20, NumTi + 1).Value > 0 Then Resum.Cells(20, NumTi + 1).Font.ColorIndex = 3 Else Resum.Cells(20, NumTi + 1).Font.ColorIndex = 4
    If SeM > 0 Then Resum.Cells(21, NumTi + 1).Value = Int((Se / Int(SeM / 4) - 1) * 100)
    If Resum.Cells(21, NumTi + 1).Value > 0 Then Resum.Cells(21, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(21, NumTi + 1).Font.ColorIndex = 3
    If (zMm + zTm) > 0 Then Resum.Cells(22, NumTi + 1).Value = Int(((zM + zT) / Int((zMm + zTm) / 4) - 1) * 100)
    If Resum.Cells(22, NumTi + 1).Value > 0 Then Resum.Cells(22, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(22, NumTi + 1).Font.ColorIndex = 3
    
    
    For i = 6 To 0 Step -1
        Col = Col + 1
        rellenaHojaDiaDeLaSetmanaBuscaDades DateAdd("d", -i, dia), client, cM, cT, zM, zT, hM, hT, DepsM, DepsT, DescM, DescT, Families, FamiliesPct, Dv, Se
        Hoja.Cells(3, Col).Value = Format(DateAdd("d", -i, dia), "dd/mm/yyyy") & Chr(10) & Format(DateAdd("d", -i, dia), "dddd")
        If DatePart("w", DateAdd("d", -i, dia)) = 1 Then Hoja.Cells(3, Col).Font.ColorIndex = 46
        
        If UBound(DepsM) > 0 Then Hoja.Cells(4, Col).Value = Right(Join(DepsM, Chr(10)), Len(Join(DepsM, Chr(10))) - 1)
        Hoja.Cells(5, Col).Value = zM
        Hoja.Cells(6, Col).Value = cM
        Hoja.Cells(7, Col).Value = Round((zM / cM), 2)
        Hoja.Cells(8, Col).Value = Int(hM / 60)
        Hoja.Cells(20, Col).Value = Dv
        Hoja.Cells(21, Col).Value = Se
        Hoja.Cells(22, Col).Value = zM + zT
        If (Se) > 0 Then Hoja.Cells(23, Col).Value = "(" & Round(1 - (zM + zT + Dv) / Se, 2) & ") " & Int((zM + zT) / Se * 100) - 100 & " %"
        
        If hM > 0 Then Hoja.Cells(9, Col).Value = Round((zM / (hM / 60)), 0)
        Hoja.Cells(10, Col).Value = DescM
            
        If UBound(DepsT) > 0 Then Hoja.Cells(13, Col).Value = Right(Join(DepsT, Chr(10)), Len(Join(DepsT, Chr(10))) - 1)
        Hoja.Cells(14, Col).Value = zT
        Hoja.Cells(15, Col).Value = cT
        Hoja.Cells(16, Col).Value = Round((zT / cT), 2)
        Hoja.Cells(17, Col).Value = Int(hT / 60)
        If hT > 0 Then Hoja.Cells(18, Col).Value = Round((zT / (hT / 60)), 0)
        Hoja.Cells(19, Col).Value = DescT
        
        For j = 1 To UBound(Families)
            FamiliesPctAcu(j) = FamiliesPctAcu(j) + FamiliesPct(j)
            Hoja.Cells(24 + j, 2).Value = Families(j)
            Hoja.Cells(24 + j, Col).Value = Round(FamiliesPct(j)) & " %"
            If MaxFam < j Then MaxFam = j
        Next
    Next
    
    
'    Hoja.Range("H:I").Copy
'    Resum.Range("C:D").Select
'    Resum.Paste , True
'    MsExcel.CutCopyMode = False
    
    'Hoja.Range("H:I").Copy
    'Resum.Range("C:D").Select
    'Resum.PasteSpecial "Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= False, Transpose:=False"
    'MsExcel.CutCopyMode = False
    'Application.CutCopyMode = False
'    Sheets("T--13OBRADOR").Select
'    Range("B5:H10").Select
'    Selection.Copy
'    Sheets("Resumen").Select
'    Range("B14").Select
'    ActiveSheet.Paste Link:=True
    
    
'    MaxFam = MaxFam + 1
'    rellenaHojaDiaDeLaSetmanaBuscaDadesGrafic Dia, Client, GraficX, GraficY
'    For i = 1 To UBound(GraficX)
'        Hoja.Cells(21 + MaxFam, i).Value = GraficX(i)
'        Hoja.Cells(22 + MaxFam, i).Value = GraficY(i)
'    Next
'    DoEvents
'    Hoja.Rows(21 + MaxFam & ":" & 22 + MaxFam).Select
'    DoEvents
'    Set Kk = Hoja.ChartObjects.Add(Left:=600, Width:=375, Top:=75, Height:=225)
'    DoEvents
'    HOJA.ChartObjects(kk.Name).Chart.ChartType = xlLineMarkers
'    HOJA.ChartObjects(kk.Name).Chart.SetSourceData Source:=HOJA.Range("A30:CK31"), PlotBy:=xlRows
'    HOJA.ChartObjects(kk.Name).Chart.Location Where:=xlLocationAsObject, Name:="T--01"
'    HOJA.ChartObjects(kk.Name).Chart.HasAxis(xlCategory, xlPrimary) = True
'    HOJA.ChartObjects(kk.Name).Chart.HasAxis(xlValue, xlPrimary) = True
'    HOJA.ChartObjects(kk.Name).Chart.Axes(xlCategory, xlPrimary).CategoryType = xlAutomatic
        
nooo:
End Sub


Private Sub rellenaHojaDiaDeLaSetmana3(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, client, ByRef Resum As Excel.Worksheet, NumTi, PrintTotales, ByRef Tz, ByRef TzM, ByRef TzT, ByRef TMCostH, ByRef TTCostH, ByRef TCostH, ByRef tServit, ByRef TDevol, ByRef TNeto, ByRef TServitIVA, ByRef TDevolIVA)
    Dim i As Integer, DiaS As Integer, Rs As rdoResultset, rsC As ADODB.Recordset, data() As Double, mes As Integer
    Dim sql As String, D As Date, Boti As Double
    Dim cM, cT, zM, zT, hM, hT, tM, tMI, tT, tTI, Deps(), DepsM(), DepsT(), DescM, DescT, Families(), FamiliesPct(), FamiliesPctAcu(), GraficX(), GraficY()
    Dim cMm, cTm, zMm, zTm, hMm, hTm, tMm, tMIm, tTm, tTIm, Col, cha As Chart, Kk As Object, MaxFam, j, H, Se, Dv
    Dim Sef, Dvf, DvfIVA, SefIVA, NetofIVA, NetoPor, SeM, DvM, SefM, DvfM, Hores As String, numTreb, arrayTreb, trebEncontrado
    'Filas
    Dim Fila, FilaTrebM, FilaVentesM, FilaTrebT, FilaVentesT, FilaEuros, FilaFabrica, FilaFamilies, FilaHoresR, FilaHoresRR, FilaIng, FilaIngR, CostH, PreuH, PerCostH, calculM, calculT
    Dim THoresM, THoresT, TIngM, TIngT, TIngMPer, TIngTPer, TCostHM, TCostHT, TNetoIva
    Dim ArrayDies(7)
    ArrayDies(0) = "Lunes"
    ArrayDies(1) = "Martes"
    ArrayDies(2) = "Miercoles"
    ArrayDies(3) = "Jueves"
    ArrayDies(4) = "Viernes"
    ArrayDies(5) = "Sabado"
    ArrayDies(6) = "Domingo"
    
    Hoja.Name = Left(Join(Split(Join(Split(BotigaCodiNom(client), " "), ""), "/"), ""), 31)
    Hoja.Cells(1, 1).Value = Format(dia, "dd/mm/yyyy dddd")
    Hoja.Cells(1, 1).Font.Bold = True
    Hoja.Cells(1, 3).Value = Join(Split(BotigaCodiNom(client), " "), "")
    Hoja.Cells(1, 3).Font.Bold = True
    'Mañana
    Fila = 4
    Hoja.Cells(Fila, 1).Value = "Mañana"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    Hoja.Columns("B:B").ColumnWidth = 25
    Hoja.Columns("C:AZ").ColumnWidth = 8
    Fila = Fila + 1
    FilaTrebM = Fila
    Hoja.Cells(Fila, 1).Value = "Trabajadores"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Col = 2
    numTreb = 0
    For i = 4 To 1 Step -1
        'Son tres columnes per dia, per poder separar valors de fitxatges de treballadors
        Col = Col + 3
        Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeLeft).LineStyle = xlNone
        Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeRight).LineStyle = xlContinuous
        Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeRight).Weight = xlMedium
        Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        rellenaHojaDiaDeLaSetmanaBuscaDades3 DateAdd("d", -(i * 7), dia), client, Deps, DepsM, DepsT, DescM, DescT
        'rellenaHojaDiaDeLaSetmanaBuscaDades2 DateAdd("d", -(i * 7), dia), Client, cM, cT, zM, zT, hM, hT, tM, tMI, tT, tTI, Deps, DepsM, DepsT, DescM, DescT, Families, FamiliesPct, Dv, Se, Dvf, Sef, DvfIVA, SefIVA, NetofIVA, NetoPor
        MergeCelda Hoja, 4, Col - 2, 4, Col, Format(DateAdd("d", -(i * 7), dia), "dd/mm/yyyy") & Chr(10) & Format(DateAdd("d", -(i * 7), dia), "dddd")
        If DatePart("w", DateAdd("d", -(i * 7), dia)) = 1 Then Resum.Cells(4, Col - 2).Font.ColorIndex = 46
        Hoja.Cells(4, Col - 2).Font.Bold = True
        'Mañana
        For j = 0 To UBound(DepsM)
            trebEncontrado = 0
            arrayTreb = Split(DepsM(j), "|")
            If UBound(arrayTreb) > 0 Then
                For H = 0 To numTreb
                    If arrayTreb(0) = Hoja.Cells(FilaTrebM + H, 2).Value Then
                        trebEncontrado = 1
                        If UBound(arrayTreb) >= 1 Then
                            If InStr(arrayTreb(1), "E") >= 1 Then Hoja.Cells(FilaTrebM + H, Col - 2).Value = Replace(arrayTreb(1), "E", "")
                            Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                        End If
                        If UBound(arrayTreb) >= 2 Then
                            If InStr(arrayTreb(2), "P") >= 1 Then Hoja.Cells(FilaTrebM + H, Col - 1).Value = Replace(arrayTreb(2), "P", "")
                            Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                        End If
                        If UBound(arrayTreb) >= 3 Then Hoja.Cells(FilaTrebM + H, Col).Value = arrayTreb(3)
                        Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                        Exit For
                    End If
                Next
                If trebEncontrado = 0 Then
                    Hoja.Cells(FilaTrebM + H, 2).Value = arrayTreb(0)
                    If UBound(arrayTreb) >= 1 Then
                        If InStr(arrayTreb(1), "E") >= 1 Then Hoja.Cells(FilaTrebM + H, Col - 2).Value = Replace(arrayTreb(1), "E", "")
                        Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                    End If
                    If UBound(arrayTreb) >= 2 Then
                        If InStr(arrayTreb(2), "P") >= 1 Then Hoja.Cells(FilaTrebM + H, Col - 1).Value = Replace(arrayTreb(2), "P", "")
                        Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                    End If
                    If UBound(arrayTreb) >= 3 Then Hoja.Cells(FilaTrebM + H, Col).Value = arrayTreb(3)
                    Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                    numTreb = numTreb + 1
                End If
            End If
        Next
    Next
    Col = Col + 1
    'Ultims 7 dies
    For i = 6 To 0 Step -1
        'Son tres columnes per dia, per poder separar valors de fitxatges de treballadors
        Col = Col + 3
        Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeLeft).LineStyle = xlNone
        Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeRight).LineStyle = xlContinuous
        Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeRight).Weight = xlMedium
        Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
        rellenaHojaDiaDeLaSetmanaBuscaDades3 DateAdd("d", -i, dia), client, Deps, DepsM, DepsT, DescM, DescT
        
        MergeCelda Hoja, 4, Col - 2, 4, Col, ArrayDies(Weekday(DateAdd("d", -(i + 1), dia)) - 1) & " " & Format(DateAdd("d", -i, dia), "dd/mm/yyyy") & Chr(10) & Format(DateAdd("d", -i, dia), "dddd")
        
        
        If DatePart("w", DateAdd("d", -i, dia)) = 1 Then Hoja.Cells(4, Col - 2).Font.ColorIndex = 46
        Hoja.Cells(4, Col - 2).Font.Bold = True
        'Mañana
        For j = 0 To UBound(DepsM)
            trebEncontrado = 0
            arrayTreb = Split(DepsM(j), "|")
            If UBound(arrayTreb) > 0 Then
                For H = 0 To numTreb
                    If arrayTreb(0) = Hoja.Cells(FilaTrebM + H, 2).Value Then
                        trebEncontrado = 1
                        If UBound(arrayTreb) >= 1 Then
                            If InStr(arrayTreb(1), "E") >= 1 Then Hoja.Cells(FilaTrebM + H, Col - 2).Value = Replace(arrayTreb(1), "E", "")
                            Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                        End If
                        If UBound(arrayTreb) >= 2 Then
                            If InStr(arrayTreb(2), "P") >= 1 Then Hoja.Cells(FilaTrebM + H, Col - 1).Value = Replace(arrayTreb(2), "P", "")
                            Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                        End If
                        If UBound(arrayTreb) >= 3 Then Hoja.Cells(FilaTrebM + H, Col).Value = arrayTreb(3)
                        Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                        Exit For
                    End If
                Next
                If trebEncontrado = 0 Then
                    Hoja.Cells(FilaTrebM + H, 2).Value = arrayTreb(0)
                    If UBound(arrayTreb) >= 1 Then
                        If InStr(arrayTreb(1), "E") >= 1 Then Hoja.Cells(FilaTrebM + H, Col - 2).Value = Replace(arrayTreb(1), "E", "")
                        Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                    End If
                    If UBound(arrayTreb) >= 2 Then
                        If InStr(arrayTreb(2), "P") >= 1 Then Hoja.Cells(FilaTrebM + H, Col - 1).Value = Replace(arrayTreb(2), "P", "")
                        Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                    End If
                    If UBound(arrayTreb) >= 3 Then Hoja.Cells(FilaTrebM + H, Col).Value = arrayTreb(3)
                    Hoja.Cells(FilaTrebM + H, Col).HorizontalAlignment = xlRight
                    numTreb = numTreb + 1
                End If
            End If
        Next
    Next
    FilaVentesM = Fila + numTreb + 1
    Fila = Fila + numTreb + 1
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeTop).LineStyle = xlContinuous
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeTop).Weight = xlMedium
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    Hoja.Cells(Fila, 1).Value = "Ventas"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Hoja.Cells(Fila, 2).Value = "Recaudacion"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Clientes"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Media ticket"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Tickets anulados"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Horas"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Rec/Hora"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Descuadre"
    Fila = Fila + 2
    'Tarde
    Hoja.Cells(Fila, 1).Value = "Tarde"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).Weight = xlMedium
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).ColorIndex = xlAutomatic
    Fila = Fila + 1
    FilaTrebT = Fila
    Hoja.Cells(Fila, 1).Value = "Trabajadores"
    Hoja.Cells(Fila, 1).Font.Bold = True
    'FilaTrebT = Fila
    Col = 2
    numTreb = 0
    For i = 4 To 1 Step -1
        'Son tres columnes per dia, per poder separar valors de fitxatges de treballadors
        Col = Col + 3
        rellenaHojaDiaDeLaSetmanaBuscaDades3 DateAdd("d", -(i * 7), dia), client, Deps, DepsM, DepsT, DescM, DescT
        'Tarde
        For j = 0 To UBound(DepsT)
            trebEncontrado = 0
            arrayTreb = Split(DepsT(j), "|")
            If UBound(arrayTreb) > 0 Then
                For H = 0 To numTreb
                    If arrayTreb(0) = Hoja.Cells(FilaTrebT + H, 2).Value Then
                        trebEncontrado = 1
                        If UBound(arrayTreb) >= 1 Then
                            If InStr(arrayTreb(1), "E") >= 1 Then Hoja.Cells(FilaTrebT + H, Col - 2).Value = Replace(arrayTreb(1), "E", "")
                            Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                        End If
                        If UBound(arrayTreb) >= 2 Then
                            If InStr(arrayTreb(2), "P") >= 1 Then Hoja.Cells(FilaTrebT + H, Col - 1).Value = Replace(arrayTreb(2), "P", "")
                            Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                        End If
                        If UBound(arrayTreb) >= 3 Then Hoja.Cells(FilaTrebT + H, Col).Value = arrayTreb(3)
                        Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                        Exit For
                    End If
                Next
                If trebEncontrado = 0 Then
                    Hoja.Cells(FilaTrebT + H, 2).Value = arrayTreb(0)
                    If UBound(arrayTreb) >= 1 Then
                        If InStr(arrayTreb(1), "E") >= 1 Then Hoja.Cells(FilaTrebT + H, Col - 2).Value = Replace(arrayTreb(1), "E", "")
                        Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                    End If
                    If UBound(arrayTreb) >= 2 Then
                        If InStr(arrayTreb(2), "P") >= 1 Then Hoja.Cells(FilaTrebT + H, Col - 1).Value = Replace(arrayTreb(2), "P", "")
                        Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                    End If
                    If UBound(arrayTreb) >= 3 Then Hoja.Cells(FilaTrebT + H, Col).Value = arrayTreb(3)
                    Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                    numTreb = numTreb + 1
                End If
            End If
        Next
    Next
    Col = Col + 1
     'Ultims 7 dies
    For i = 6 To 0 Step -1
        'Son tres columnes per dia, per poder separar valors de fitxatges de treballadors
        Col = Col + 3
        rellenaHojaDiaDeLaSetmanaBuscaDades3 DateAdd("d", -i, dia), client, Deps, DepsM, DepsT, DescM, DescT
        'rellenaHojaDiaDeLaSetmanaBuscaDades2 DateAdd("d", -i, dia), Client, cM, cT, zM, zT, hM, hT, tM, tMI, tT, tTI, Deps, DepsM, DepsT, DescM, DescT, Families, FamiliesPct, Dv, Se, Dvf, Sef, DvfIVA, SefIVA, NetofIVA, NetoPor
        'Tarde
        For j = 0 To UBound(DepsT)
            trebEncontrado = 0
            arrayTreb = Split(DepsT(j), "|")
            If UBound(arrayTreb) > 0 Then
                For H = 0 To numTreb
                    Hoja.Cells(FilaTrebT + H, 2).Select
                    If arrayTreb(0) = Hoja.Cells(FilaTrebT + H, 2).Value Then
                        trebEncontrado = 1
                        If UBound(arrayTreb) >= 1 Then
                            If InStr(arrayTreb(1), "E") >= 1 Then Hoja.Cells(FilaTrebT + H, Col - 2).Value = Replace(arrayTreb(1), "E", "")
                            Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                        End If
                        If UBound(arrayTreb) >= 2 Then
                            If InStr(arrayTreb(2), "P") >= 1 Then Hoja.Cells(FilaTrebT + H, Col - 1).Value = Replace(arrayTreb(2), "P", "")
                            Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                        End If
                        If UBound(arrayTreb) >= 3 Then Hoja.Cells(FilaTrebT + H, Col).Value = arrayTreb(3)
                        Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                        Exit For
                    End If
                Next
                If trebEncontrado = 0 Then
                    Hoja.Cells(FilaTrebT + H, 2).Value = arrayTreb(0)
                    If UBound(arrayTreb) >= 1 Then
                        If InStr(arrayTreb(1), "E") >= 1 Then Hoja.Cells(FilaTrebT + H, Col - 2).Value = Replace(arrayTreb(1), "E", "")
                        Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                    End If
                    If UBound(arrayTreb) >= 2 Then
                        If InStr(arrayTreb(2), "P") >= 1 Then Hoja.Cells(FilaTrebT + H, Col - 1).Value = Replace(arrayTreb(2), "P", "")
                        Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                    End If
                    If UBound(arrayTreb) >= 3 Then Hoja.Cells(FilaTrebT + H, Col).Value = arrayTreb(3)
                    Hoja.Cells(FilaTrebT + H, Col).HorizontalAlignment = xlRight
                    numTreb = numTreb + 1
                End If
            End If
        Next
    Next
    'Ventes
    FilaVentesT = Fila + numTreb + 1
    Fila = Fila + numTreb + 1
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeTop).LineStyle = xlContinuous
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeTop).Weight = xlMedium
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeTop).ColorIndex = xlAutomatic
    Hoja.Cells(Fila, 1).Value = "Ventas"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Hoja.Cells(Fila, 2).Value = "Recaudacion"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Clientes"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Media ticket"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Tickets anulados"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Horas"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Rec/Hora"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Descuadre"
    Fila = Fila + 1
    'Euros
    FilaEuros = Fila
    Hoja.Cells(Fila, 1).Value = "Euros"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Hoja.Rows(Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Devolucion"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Servido"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Ingreso"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "(Err) Retorno"
    Fila = Fila + 1
    'Fabrica
    FilaFabrica = Fila
    Hoja.Rows(Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(Fila, 1).Value = "Fabrica"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Servido fabrica + IVA"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Devolucion fabrica + IVA"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Neto + % sobre ventas"
    Fila = Fila + 1
    'Familia
    FilaFamilies = Fila
    Hoja.Rows(Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(Fila, 1).Value = "Familia"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Fila = Fila + 1
    Set Rs = Db.OpenResultset("Select * from families Where Pare = 'Article' Order by nom ")
    ReDim Families(0)
    ReDim FamiliesPct(0)
    ReDim FamiliesPctAcu(0)
    While Not Rs.EOF
        ReDim Preserve Families(UBound(Families) + 1)
        ReDim Preserve FamiliesPct(UBound(FamiliesPct) + 1)
        ReDim Preserve FamiliesPctAcu(UBound(FamiliesPctAcu) + 1)
        FamiliesPct(UBound(FamiliesPct)) = 0
        FamiliesPctAcu(UBound(FamiliesPctAcu)) = 0
        Families(UBound(Families)) = Rs("Nom")
        Fila = Fila + 1
        Rs.MoveNext
    Wend
    'Horas resumen
    PreuH = 11 'Cost hora treballador
    Fila = Fila + 1
    FilaHoresR = Fila
    FilaHoresRR = 31 + UBound(Families) + 1
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(Fila, 1).Value = "Resumen Horas"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Horas mañana"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Horas tarde"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Horas totales"
    Fila = Fila + 1
    'Ingreso
    FilaIng = Fila
    FilaIngR = FilaHoresRR + 5
    Hoja.Rows(Fila & ":" & Fila).Borders(xlEdgeBottom).LineStyle = xlContinuous
    Hoja.Cells(Fila, 1).Value = "Ingreso"
    Hoja.Cells(Fila, 1).Font.Bold = True
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Ingreso mañana"
    Fila = Fila + 1
    Hoja.Cells(Fila, 2).Value = "Ingreso tarde"
        
    cMm = 0: cTm = 0: zMm = 0: zTm = 0: hMm = 0: hTm = 0: MaxFam = 0: SeM = 0: DvM = 0
    
    Col = 2
    '4 Dias anteriores de las 4 ultimas semanas
    For i = 4 To 1 Step -1
        'Son tres columnes per dia, per poder separar valors de fitxatges de treballadors
        Col = Col + 3
        rellenaHojaDiaDeLaSetmanaBuscaDades2 DateAdd("d", -(i * 7), dia), client, cM, cT, zM, zT, hM, hT, tM, tMI, tT, tTI, Deps, DepsM, DepsT, DescM, DescT, Families, FamiliesPct, Dv, Se, Dvf, Sef, DvfIVA, SefIVA, NetofIVA, NetoPor
        'Mañana
        MergeCelda Hoja, FilaVentesM, Col - 2, FilaVentesM, Col, zM
        MergeCelda Hoja, FilaVentesM + 1, Col - 2, FilaVentesM + 1, Col, cM
        MergeCelda Hoja, FilaVentesM + 2, Col - 2, FilaVentesM + 2, Col, Round((zM / cM), 2)
        MergeCelda Hoja, FilaVentesM + 3, Col - 2, FilaVentesM + 3, Col, tM & "(" & tMI & ")"
        MergeCelda Hoja, FilaVentesM + 4, Col - 2, FilaVentesM + 4, Col, Int(hM / 60)
        MergeCelda Hoja, FilaVentesM + 5, Col - 2, FilaVentesM + 5, Col
        If hM > 0 Then Hoja.Cells(FilaVentesM + 5, Col - 2).Value = Round((zM / (hM / 60)), 0)
        MergeCelda Hoja, FilaVentesM + 6, Col - 2, FilaVentesM + 6, Col, DescM
        MergeCelda Hoja, FilaEuros + 1, Col - 2, FilaEuros + 1, Col, Dv
        MergeCelda Hoja, FilaEuros + 2, Col - 2, FilaEuros + 2, Col, Se
        MergeCelda Hoja, FilaEuros + 3, Col - 2, FilaEuros + 3, Col, zM + zT
        MergeCelda Hoja, FilaEuros + 4, Col - 2, FilaEuros + 4, Col
        If (Se) > 0 Then Hoja.Cells(FilaEuros + 4, Col - 2).Value = "(" & Round(1 - (zM + zT + Dv) / Se, 2) & ") " & Int((zM + zT) / Se * 100) - 100 & " %"
        MergeCelda Hoja, FilaFabrica + 1, Col - 2, FilaFabrica + 1, Col, SefIVA
        MergeCelda Hoja, FilaFabrica + 2, Col - 2, FilaFabrica + 2, Col, DvfIVA
        MergeCelda Hoja, FilaFabrica + 3, Col - 2, FilaFabrica + 3, Col, NetofIVA & " (" & NetoPor & ")"
        'Tarde
        MergeCelda Hoja, FilaVentesT, Col - 2, FilaVentesT, Col, zT
        MergeCelda Hoja, FilaVentesT + 1, Col - 2, FilaVentesT + 1, Col, cT
        MergeCelda Hoja, FilaVentesT + 2, Col - 2, FilaVentesT + 2, Col, Round((zT / cT), 2)
        MergeCelda Hoja, FilaVentesT + 3, Col - 2, FilaVentesT + 3, Col, tT & "(" & tTI & ")"
        MergeCelda Hoja, FilaVentesT + 4, Col - 2, FilaVentesT + 4, Col, Int(hT / 60)
        MergeCelda Hoja, FilaVentesT + 5, Col - 2, FilaVentesT + 5, Col
        If hT > 0 Then Hoja.Cells(FilaVentesT + 5, Col - 2).Value = Round((zT / (hT / 60)), 0)
        MergeCelda Hoja, FilaVentesT + 6, Col - 2, FilaVentesT + 6, Col, DescT
        CostH = ((Int(hM / 60)) * PreuH)
        If zM <> 0 Then
            PerCostH = Round((CostH / zM) * 100, 1)
        End If
        MergeCelda Hoja, FilaHoresR + 1, Col - 2, FilaHoresR + 1, Col, CostH & " (" & PerCostH & ")"
        CostH = ((Int(hT / 60)) * PreuH)
        If zT <> 0 Then PerCostH = Round((CostH / zT) * 100, 1)
        
        MergeCelda Hoja, FilaHoresR + 2, Col - 2, FilaHoresR + 2, Col, CostH & " (" & PerCostH & ")"
        CostH = ((Int(hM / 60) + Int(hT / 60)) * PreuH)
        PerCostH = Round((CostH / (zM + zT)) * 100, 1)
        MergeCelda Hoja, FilaHoresR + 3, Col - 2, FilaHoresR + 3, Col, CostH & " (" & PerCostH & ")"
    
    
        ReDim Preserve FamiliesPctAcu(UBound(FamiliesPct))
        For j = 1 To UBound(Families)
            If FamiliesPctAcu(j) = "" Then FamiliesPctAcu(j) = "0"
            FamiliesPctAcu(j) = FamiliesPctAcu(j) + FamiliesPct(j)
            Hoja.Cells(FilaFamilies + j, 2).Value = Families(j)
            If FamiliesPct(j) = "" Then FamiliesPct(j) = "0"
            MergeCelda Hoja, FilaFamilies + j, Col - 2, FilaFamilies + j, Col, Round(FamiliesPct(j)) & " %"
            If MaxFam < j Then MaxFam = j
        Next
    
        cMm = cMm + cM
        cTm = cTm + cT
        zMm = zMm + zM
        zTm = zTm + zT
        hMm = hMm + hM
        hTm = hTm + hT
        SeM = SeM + Se
        DvM = DvM + Dv
        SefM = SefM + Sef
        DvfM = DvfM + Dvf
        tMm = tMm + tM
        tMIm = tMIm + tMI
        tTm = tTm + tT
        tTIm = tTIm + tTI
        
        calculM = 0
        calculT = 0
        If Int(zMm / 4) > 0 Then calculM = Int((zM / Int(zMm / 4) - 1) * 100)
        If Int(zTm / 4) > 0 Then calculT = Int((zT / Int(zTm / 4) - 1) * 100)
        MergeCelda Hoja, FilaIng + 1, Col - 2, FilaIng + 1, Col, Round(zM, 1) & "(" & Round(zM * (calculM / 100), 1) & ")"
        MergeCelda Hoja, FilaIng + 2, Col - 2, FilaIng + 2, Col, Round(zT, 1) & "(" & Round(zT * (calculT / 100), 1) & ")"
        'THoresM = THoresM + hM
        'THoresT = THoresT + hT
        'TIngM = TIngM + zM
        'TIngT = TIngT + zT
    Next
    'Medias
    Col = Col + 1
    Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeRight).Weight = xlMedium
    Hoja.Range(Hoja.Columns(Col), Hoja.Columns(Col)).Borders(xlEdgeRight).ColorIndex = xlAutomatic
    Hoja.Cells(4, Col).Value = "Medias"
    Hoja.Cells(4, Col).Font.Bold = True
    'Medias Mañana
    Hoja.Cells(FilaVentesM, Col).Value = Int(zMm / 4)
    Hoja.Cells(FilaVentesM + 1, Col).Value = Int(cMm / 4)
    If cMm > 0 Then Hoja.Cells(FilaVentesM + 2, Col).Value = Round((zMm / cMm), 2)
    Hoja.Cells(FilaVentesM + 3, Col).Value = tMm & "(" & tMIm & ")"
    Hoja.Cells(FilaVentesM + 4, Col).Value = Int(hMm / 60 / 4)
    If hMm > 0 Then Hoja.Cells(FilaVentesM + 5, Col).Value = Round((zMm / (hMm / 60)), 0)
    'Medias Tarde
    Hoja.Cells(FilaVentesT, Col).Value = Int(zTm / 4)
    Hoja.Cells(FilaVentesT + 1, Col).Value = Int(cTm / 4)
    If cTm > 0 Then Hoja.Cells(FilaVentesT + 2, Col).Value = Round((zTm / cTm), 2)
    Hoja.Cells(FilaVentesT + 3, Col).Value = tTm & "(" & tTIm & ")"
    Hoja.Cells(FilaVentesT + 4, Col).Value = Int(hTm / 60 / 4)
    If hTm > 0 Then Hoja.Cells(FilaVentesT + 5, Col).Value = Round((zTm / (hTm / 60)), 0)
    
    Hoja.Cells(FilaEuros + 1, Col).Value = Int(DvM / 4)
    Hoja.Cells(FilaEuros + 2, Col).Value = Int(SeM / 4)
    Hoja.Cells(FilaEuros + 3, Col).Value = Int((zMm + zTm) / 4)
    Hoja.Cells(FilaFabrica + 1, Col).Value = Int(SefM / 4)
    Hoja.Cells(FilaFabrica + 2, Col).Value = Int(DvfM / 4)
    
    
    'Resum. Ordre de fileres diferent del ordre de pagines botiga
    Resum.Cells(2, NumTi).Value = Join(Split(BotigaCodiNom(client), " "), "")
    Resum.Cells(2, NumTi).Font.Bold = True
    rellenaHojaDiaDeLaSetmanaBuscaDades2 dia, client, cM, cT, zM, zT, hM, hT, tM, tMI, tT, tTI, Deps, DepsM, DepsT, DescM, DescT, Families, FamiliesPct, Dv, Se, Dvf, Sef, DvfIVA, SefIVA, NetofIVA, NetoPor
    Tz = Tz + (IIf(IsNull(zM), 0, zM) + IIf(IsNull(zT), 0, zT)) 'Total Ingreso
    TzM = TzM + IIf(IsNull(zM), 0, zM) 'Total Ingreso mati
    TzT = TzT + IIf(IsNull(zT), 0, zT) 'Total Ingreso tarda
    TServitIVA = TServitIVA + IIf(IsNull(SefIVA), 0, SefIVA) 'Total Servit
    TDevolIVA = TDevolIVA + IIf(IsNull(DvfIVA), 0, DvfIVA) 'Total Devolucions
    TNeto = TNeto + IIf(IsNull(NetofIVA), 0, NetofIVA) 'Total Neto
    TDevol = TDevolIVA + IIf(IsNull(Dv), 0, Dv)
    tServit = tServit + IIf(IsNull(Se), 0, Se)
    Resum.Cells(3, NumTi).Value = dia
'    If UBound(DepsM) > 0 Then Resum.Cells(4, NumTi).Value = Right(Join(DepsM, Chr(10)), Len(Join(DepsM, Chr(10))) - 1)
    Resum.Cells(5, NumTi).Value = zM
    Resum.Cells(6, NumTi).Value = cM
    Resum.Cells(7, NumTi).Value = Round((zM / cM), 2)
    Resum.Cells(8, NumTi).Value = tM
    Resum.Cells(8, NumTi + 1).Value = tMI
    Resum.Cells(9, NumTi).Value = Int(hM / 60)
    If hM > 0 Then Resum.Cells(10, NumTi).Value = Round((zM / (hM / 60)), 0)
    Resum.Cells(11, NumTi).Value = DescM
    Resum.Cells(22, NumTi).Value = Dv
    Resum.Cells(23, NumTi).Value = Se
    Resum.Cells(24, NumTi).Value = zM + zT
    If (Se) > 0 Then Resum.Cells(25, NumTi).Value = "(" & Round(1 - (zM + zT + Dv) / Se, 2) & ") " & Int((zM + zT) / Se * 100) - 100 & " %"
    Resum.Cells(27, NumTi).Value = SefIVA
    Resum.Cells(28, NumTi).Value = DvfIVA
    Resum.Cells(29, NumTi).Value = NetofIVA
    Resum.Cells(29, NumTi + 1).Value = NetoPor
     
            
'    If UBound(DepsT) > 0 Then Resum.Cells(13, NumTi).Value = Right(Join(DepsT, Chr(10)), Len(Join(DepsT, Chr(10))) - 1)
    Resum.Cells(14, NumTi).Value = zT
    Resum.Cells(15, NumTi).Value = cT
    Resum.Cells(16, NumTi).Value = Round((zT / cT), 2)
    Resum.Cells(17, NumTi).Value = tT
    Resum.Cells(17, NumTi + 1).Value = tTI
    Resum.Cells(18, NumTi).Value = Int(hT / 60)
    If hT > 0 Then Resum.Cells(19, NumTi).Value = Round((zT / (hT / 60)), 0)
    Resum.Cells(20, NumTi).Value = DescT
    CostH = ((Int(hM / 60)) * PreuH)
    If zM <> 0 Then
        PerCostH = Round((CostH / zM) * 100, 1)
    End If
    TMCostH = TMCostH + IIf(IsNull(CostH), 0, CostH)
    Resum.Cells(FilaHoresRR + 1, NumTi).Value = CostH & " (" & PerCostH & ")"
    CostH = ((Int(hT / 60)) * PreuH)
    If zT <> 0 Then
        PerCostH = Round((CostH / zT) * 100, 1)
    End If
    TTCostH = TTCostH + IIf(IsNull(CostH), 0, CostH)
    Resum.Cells(FilaHoresRR + 2, NumTi).Value = CostH & " (" & PerCostH & ")"
    CostH = ((Int(hM / 60) + Int(hT / 60)) * PreuH)
    PerCostH = Round((CostH / (zM + zT)) * 100, 1)
    TCostH = TCostH + IIf(IsNull(CostH), 0, CostH)
    Resum.Cells(FilaHoresRR + 3, NumTi).Value = CostH & " (" & PerCostH & ")"
    Resum.Cells(FilaIngR, NumTi).Value = Round(zM, 2)
    calculM = Int((zM / Int(zMm / 4) - 1) * 100)
    calculT = Int((zT / Int(zTm / 4) - 1) * 100)
    Resum.Cells(FilaIngR, NumTi + 1).Value = Round(zM * (calculM / 100), 1)
    'Round(zM * ((zM - (zMm / 4)) / 100), 1)
    Resum.Cells(FilaIngR + 1, NumTi).Value = Round(zT, 2)
    Resum.Cells(FilaIngR + 1, NumTi + 1).Value = Round(zT * (calculT / 100), 1)
    'Round(zT * ((zT - (zTm / 4)) / 100), 1)
    
    For j = 1 To UBound(Families)
        Resum.Cells(31 + j, 2).Value = Families(j)
        Resum.Cells(31 + j, NumTi).Value = Int(FamiliesPct(j)) & " %"
        If FamiliesPctAcu(j) > 0 Then Resum.Cells(31 + j, NumTi + 1).Value = Int(FamiliesPct(j) / (FamiliesPctAcu(j) / 4) * 100) - 100 & " %"
        If FamiliesPctAcu(j) > 0 Then Resum.Cells(31 + j, NumTi - 1).Value = Int((FamiliesPctAcu(j) / 4)) & " %"
        If MaxFam < j Then MaxFam = j
    Next
    
'    Resum.Cells(3, NumTi + 1).Value = "Desvio"
    Resum.Cells(3, NumTi + 1).Font.Bold = True
    Resum.Cells(3, NumTi + 1).Font.Italic = True
    If zMm > 0 Then Resum.Cells(5, NumTi + 1).Value = Int((zM / Int(zMm / 4) - 1) * 100)
    'If zMm > 0 Then Resum.Cells(5, NumTi + 1).Value = Int((zM - Int(zMm / 4)))
    If Resum.Cells(5, NumTi + 1).Value > 0 Then Resum.Cells(5, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(5, NumTi + 1).Font.ColorIndex = 3
    
    If cM > 0 Then If cMm > 0 Then If Int(cMm / 4) > 0 Then Resum.Cells(6, NumTi + 1).Value = Int((cM / Int(cMm / 4) - 1) * 100)
    'If cMm > 0 Then Resum.Cells(6, NumTi + 1).Value = Int((cM - Int(cMm / 4)))
    If Resum.Cells(6, NumTi + 1).Value > 0 Then Resum.Cells(6, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(6, NumTi + 1).Font.ColorIndex = 3
    If zM > 0 Then If cMm > 0 Then Resum.Cells(7, NumTi + 1).Value = Int((Round((zM / cM), 2) / Round((zMm / cMm), 2) - 1) * 100)
    If Resum.Cells(7, NumTi + 1).Value > 0 Then Resum.Cells(7, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(7, NumTi + 1).Font.ColorIndex = 3
    If hM = 0 Then hM = 1
    If hMm = 0 Then hMm = 1
    If hMm > 0 Then Resum.Cells(9, NumTi + 1).Value = Int(((hM) / (hMm / 4) - 1) * 100)
    'If hMm > 0 Then Resum.Cells(9, NumTi + 1).Value = Int(((hM) - (hMm / 4)))
    If Resum.Cells(9, NumTi + 1).Value < 0 Then Resum.Cells(9, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(9, NumTi + 1).Font.ColorIndex = 3
    If hM > 0 Then If hMm > 0 Then Resum.Cells(10, NumTi + 1).Value = Int(((zM / hM) / (zMm / hMm) - 1) * 100)
    If Resum.Cells(10, NumTi + 1).Value > 0 Then Resum.Cells(10, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(10, NumTi + 1).Font.ColorIndex = 3
    
    If zTm > 0 Then Resum.Cells(14, NumTi + 1).Value = Int((zT / Int(zTm / 4) - 1) * 100)
    'If zTm > 0 Then Resum.Cells(14, NumTi + 1).Value = Int((zT - Int(zTm / 4)))
    If Resum.Cells(14, NumTi + 1).Value > 0 Then Resum.Cells(14, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(14, NumTi + 1).Font.ColorIndex = 3
    If Int(cTm / 4) And cTm > 0 And cT > 0 Then Resum.Cells(15, NumTi + 1).Value = Int((cT / Int(cTm / 4) - 1) * 100)
    'If Int(cTm / 4) And cTm > 0 And cT > 0 Then Resum.Cells(15, NumTi + 1).Value = Int((cT - Int(cTm / 4)))
    If Resum.Cells(15, NumTi + 1).Value > 0 Then Resum.Cells(15, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(15, NumTi + 1).Font.ColorIndex = 3
    If cTm > 0 Then Resum.Cells(16, NumTi + 1).Value = Int((Round((zT / cT), 2) / Round((zTm / cTm), 2) - 1) * 100)
    If Resum.Cells(16, NumTi + 1).Value > 0 Then Resum.Cells(16, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(16, NumTi + 1).Font.ColorIndex = 3
    If hTm > 0 Then Resum.Cells(18, NumTi + 1).Value = Int(((hT) / (hTm / 4) - 1) * 100)
    'If hTm > 0 Then Resum.Cells(17, NumTi + 1).Value = Int(((hT) - (hTm / 4)))
    If Resum.Cells(18, NumTi + 1).Value > 0 Then Resum.Cells(18, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(18, NumTi + 1).Font.ColorIndex = 3
    If hT > 0 And hTm > 0 And zTm > 0 Then Resum.Cells(19, NumTi + 1).Value = Int(((zT / hT) / (zTm / hTm) - 1) * 100)
    If Resum.Cells(19, NumTi + 1).Value > 0 Then Resum.Cells(19, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(19, NumTi + 1).Font.ColorIndex = 3
    If Int(DvM / 4) > 0 Then Resum.Cells(22, NumTi + 1).Value = Int((Dv / Int(DvM / 4) - 1) * 100)
    If Resum.Cells(22, NumTi + 1).Value > 0 Then Resum.Cells(22, NumTi + 1).Font.ColorIndex = 3 Else Resum.Cells(22, NumTi + 1).Font.ColorIndex = 4
    If SeM > 0 Then Resum.Cells(23, NumTi + 1).Value = Int((Se / Int(SeM / 4) - 1) * 100)
    If Resum.Cells(23, NumTi + 1).Value > 0 Then Resum.Cells(23, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(23, NumTi + 1).Font.ColorIndex = 3
    If (zMm + zTm) > 0 Then Resum.Cells(24, NumTi + 1).Value = Int(((zM + zT) / Int((zMm + zTm) / 4) - 1) * 100)
    If Resum.Cells(24, NumTi + 1).Value > 0 Then Resum.Cells(24, NumTi + 1).Font.ColorIndex = 4 Else Resum.Cells(24, NumTi + 1).Font.ColorIndex = 3
    
    'Ultims 7 dies
    For i = 6 To 0 Step -1
        'Son tres columnes per dia, per poder separar valors de fitxatges de treballadors
        Col = Col + 3
        rellenaHojaDiaDeLaSetmanaBuscaDades2 DateAdd("d", -i, dia), client, cM, cT, zM, zT, hM, hT, tM, tMI, tT, tTI, Deps, DepsM, DepsT, DescM, DescT, Families, FamiliesPct, Dv, Se, Dvf, Sef, DvfIVA, SefIVA, NetofIVA, NetoPor
        'Mañana
        Hoja.Cells(FilaTrebM, Col).Font.Size = 8
        MergeCelda Hoja, FilaVentesM, Col - 2, FilaVentesM, Col, zM
        MergeCelda Hoja, FilaVentesM + 1, Col - 2, FilaVentesM + 1, Col, cM
        MergeCelda Hoja, FilaVentesM + 2, Col - 2, FilaVentesM + 2, Col, Round((zM / cM), 2)
        MergeCelda Hoja, FilaVentesM + 3, Col - 2, FilaVentesM + 3, Col, tM & "(" & tMI & ")"
        MergeCelda Hoja, FilaVentesM + 4, Col - 2, FilaVentesM + 4, Col, Int(hM / 60)
        MergeCelda Hoja, FilaVentesM + 5, Col - 2, FilaVentesM + 5, Col
        If hM > 0 Then Hoja.Cells(FilaVentesM + 5, Col - 2).Value = Round((zM / (hM / 60)), 0)
        MergeCelda Hoja, FilaVentesM + 6, Col - 2, FilaVentesM + 6, Col, DescM
        MergeCelda Hoja, FilaEuros + 1, Col - 2, FilaEuros + 1, Col, Dv
        MergeCelda Hoja, FilaEuros + 2, Col - 2, FilaEuros + 2, Col, Se
        MergeCelda Hoja, FilaEuros + 3, Col - 2, FilaEuros + 3, Col, zM + zT
        MergeCelda Hoja, FilaEuros + 4, Col - 2, FilaEuros + 4, Col
        If (Se) > 0 Then Hoja.Cells(FilaEuros + 4, Col - 2).Value = "(" & Round(1 - (zM + zT + Dv) / Se, 2) & ") " & Int((zM + zT) / Se * 100) - 100 & " %"
        MergeCelda Hoja, FilaFabrica + 1, Col - 2, FilaFabrica + 1, Col, SefIVA
        MergeCelda Hoja, FilaFabrica + 2, Col - 2, FilaFabrica + 2, Col, DvfIVA
        MergeCelda Hoja, FilaFabrica + 3, Col - 2, FilaFabrica + 3, Col, NetofIVA & " (" & NetoPor & ")"
        'Tarde
        Hoja.Cells(FilaTrebT, Col).Font.Size = 8
        MergeCelda Hoja, FilaVentesT, Col - 2, FilaVentesT, Col, zT
        MergeCelda Hoja, FilaVentesT + 1, Col - 2, FilaVentesT + 1, Col, cT
        MergeCelda Hoja, FilaVentesT + 2, Col - 2, FilaVentesT + 2, Col, Round((zT / cT), 2)
        MergeCelda Hoja, FilaVentesT + 3, Col - 2, FilaVentesT + 3, Col, tT & "(" & tTI & ")"
        MergeCelda Hoja, FilaVentesT + 4, Col - 2, FilaVentesT + 4, Col, Int(hT / 60)
        MergeCelda Hoja, FilaVentesT + 5, Col - 2, FilaVentesT + 5, Col
        If hT > 0 Then Hoja.Cells(FilaVentesT + 5, Col - 2).Value = Round((zT / (hT / 60)), 0)
        MergeCelda Hoja, FilaVentesT + 6, Col - 2, FilaVentesT + 6, Col, DescT
        CostH = ((Int(hM / 60)) * PreuH)
        THoresM = THoresM + CostH
        If zM > 0 Then
            PerCostH = Round((CostH / zM) * 100, 1)
        Else
            PerCostH = 0
        End If
        MergeCelda Hoja, FilaHoresR + 1, Col - 2, FilaHoresR + 1, Col, CostH & " (" & PerCostH & ")"
        CostH = ((Int(hT / 60)) * PreuH)
        THoresT = THoresT + CostH
        If zT <> 0 Then PerCostH = Round((CostH / zT) * 100, 1)
        MergeCelda Hoja, FilaHoresR + 2, Col - 2, FilaHoresR + 2, Col, CostH & " (" & PerCostH & ")"
        CostH = ((Int(hM / 60) + Int(hT / 60)) * PreuH)
        PerCostH = Round((CostH / (zM + zT)) * 100, 1)
        MergeCelda Hoja, FilaHoresR + 3, Col - 2, FilaHoresR + 3, Col, CostH & " (" & PerCostH & ")"
        
                
        ReDim Preserve FamiliesPctAcu(UBound(FamiliesPct))
        For j = 1 To UBound(Families)
            If FamiliesPctAcu(j) = "" Then FamiliesPctAcu(j) = "0"
            FamiliesPctAcu(j) = FamiliesPctAcu(j) + FamiliesPct(j)
            Hoja.Cells(FilaFamilies + j, 2).Value = Families(j)
            If FamiliesPct(j) = "" Then FamiliesPct(j) = "0"
            MergeCelda Hoja, FilaFamilies + j, Col - 2, FilaFamilies + j, Col, Round(FamiliesPct(j)) & " %"
            If MaxFam < j Then MaxFam = j
        Next
        calculM = Int((zM / Int(zMm / 4) - 1) * 100)
        calculT = Int((zT / Int(zTm / 4) - 1) * 100)
        MergeCelda Hoja, FilaIng + 1, Col - 2, FilaIng + 1, Col, Round(zM, 1) & "(" & Round(zM * (calculM / 100), 1) & ")"
        MergeCelda Hoja, FilaIng + 2, Col - 2, FilaIng + 2, Col, Round(zT, 1) & "(" & Round(zT * (calculT / 100), 1) & ")"
        TIngM = TIngM + zM
        TIngMPer = TIngMPer + Round(zM * (calculM / 100), 1)
        TIngT = TIngT + zT
        TIngTPer = TIngTPer + Round(zT * (calculT / 100), 1)
        TNetoIva = TNetoIva + NetofIVA
    Next
    'Totales per full actual
    Hoja.Cells(4, Col + 1).Value = "TOTALES" & " " & Join(Split(BotigaCodiNom(client), " "), "")
    Hoja.Cells(4, Col + 1).Font.Bold = True
    Hoja.Cells(FilaEuros + 1, Col + 1).FormulaR1C1 = "=TRUNC(SUM(RC[-" & Col - 15 & "]:RC[-1]),1)"
    Hoja.Cells(FilaEuros + 1, Col + 1).Font.Bold = True
    Hoja.Cells(FilaEuros + 1, Col + 2).FormulaR1C1 = "=TRUNC((R" & FilaEuros + 1 & "C" & Col + 1 & "*100)/R" & FilaEuros + 2 & "C" & Col + 1 & ",1)"
    Hoja.Cells(FilaEuros + 1, Col + 2).Font.Bold = True
    Hoja.Cells(FilaEuros + 2, Col + 1).FormulaR1C1 = "=TRUNC(SUM(RC[-" & Col - 15 & "]:RC[-1]),1)"
    Hoja.Cells(FilaEuros + 2, Col + 1).Font.Bold = True
    Hoja.Cells(FilaEuros + 3, Col + 1).FormulaR1C1 = "=TRUNC(SUM(RC[-" & Col - 15 & "]:RC[-1]),1)"
    Hoja.Cells(FilaEuros + 3, Col + 1).Font.Bold = True
    Hoja.Cells(FilaFabrica + 1, Col + 1).FormulaR1C1 = "=TRUNC(SUM(RC[-" & Col - 15 & "]:RC[-1]),1)"
    Hoja.Cells(FilaFabrica + 1, Col + 1).Font.Bold = True
    Hoja.Cells(FilaFabrica + 2, Col + 1).FormulaR1C1 = "=TRUNC(SUM(RC[-" & Col - 15 & "]:RC[-1]),1)"
    Hoja.Cells(FilaFabrica + 2, Col + 1).Font.Bold = True
    Hoja.Cells(FilaFabrica + 2, Col + 2).FormulaR1C1 = "=TRUNC((R" & FilaFabrica + 2 & "C" & Col + 1 & "*100)/R" & FilaFabrica + 1 & "C" & Col + 1 & ",1)"
    Hoja.Cells(FilaFabrica + 2, Col + 2).Font.Bold = True
    Hoja.Cells(FilaFabrica + 3, Col + 1).Value = TNetoIva
    Hoja.Cells(FilaFabrica + 3, Col + 1).Font.Bold = True
    Hoja.Cells(FilaFabrica + 3, Col + 2).Value = "=TRUNC((R" & FilaFabrica + 3 & "C" & Col + 1 & "*100)/R" & FilaEuros + 3 & "C" & Col + 1 & ",1)"
    Hoja.Cells(FilaFabrica + 3, Col + 2).Font.Bold = True
    Hoja.Cells(FilaHoresR + 1, Col + 1).FormulaR1C1 = THoresM
    Hoja.Cells(FilaHoresR + 1, Col + 1).Font.Bold = True
    PerCostH = Round((THoresM / TIngM) * 100, 1)
    Hoja.Cells(FilaHoresR + 1, Col + 2).FormulaR1C1 = PerCostH
    Hoja.Cells(FilaHoresR + 1, Col + 2).Font.Bold = True
    Hoja.Cells(FilaHoresR + 2, Col + 1).FormulaR1C1 = THoresT
    Hoja.Cells(FilaHoresR + 2, Col + 1).Font.Bold = True
    PerCostH = Round((THoresT / TIngT) * 100, 1)
    Hoja.Cells(FilaHoresR + 2, Col + 2).FormulaR1C1 = PerCostH
    Hoja.Cells(FilaHoresR + 2, Col + 2).Font.Bold = True
    Hoja.Cells(FilaHoresR + 3, Col + 1).FormulaR1C1 = "=TRUNC(SUM(R" & FilaHoresR + 1 & "C" & Col + 1 & ":R" & FilaHoresR + 2 & "C" & Col + 1 & "),1)"
    Hoja.Cells(FilaHoresR + 3, Col + 1).Font.Bold = True
    PerCostH = Round(((THoresM + THoresT) / (TIngT + TIngM)) * 100, 1)
    Hoja.Cells(FilaHoresR + 3, Col + 2).FormulaR1C1 = PerCostH
    Hoja.Cells(FilaHoresR + 3, Col + 2).Font.Bold = True
    Hoja.Cells(FilaIng + 1, Col + 1).Value = TIngM
    Hoja.Cells(FilaIng + 1, Col + 1).Font.Bold = True
    Hoja.Cells(FilaIng + 2, Col + 1).Value = TIngT
    Hoja.Cells(FilaIng + 2, Col + 1).Font.Bold = True
    
    'Totals per full resum
    If PrintTotales = True Then
        Resum.Cells(2, NumTi + 2).Value = "TOTALES"
        Resum.Cells(2, NumTi + 2).Font.Bold = True
        'FilaEuros
        Resum.Cells(22, NumTi + 2).Value = Round(TDevol, 2)
        Resum.Cells(22, NumTi + 2).Font.Bold = True
        Resum.Cells(22, NumTi + 3).Value = "=TRUNC((R22C" & NumTi + 2 & "*100)/R23C" & NumTi + 2 & ",1)"
        Resum.Cells(23, NumTi + 2).Value = Round(tServit, 2)
        Resum.Cells(23, NumTi + 2).Font.Bold = True
        Resum.Cells(24, NumTi + 2).Value = Round(Tz, 2)
        Resum.Cells(24, NumTi + 2).Font.Bold = True
        'Resum.Cells(24, NumTi + 3).Value = 0
        'Fila Fabrica
        Resum.Cells(27, NumTi + 2).Value = Round(TServitIVA, 2)
        'Round(((TServit * 100) / Tz), 2)
        Resum.Cells(27, NumTi + 2).Font.Bold = True
        Resum.Cells(28, NumTi + 2).Value = Round(TDevolIVA, 2)
        Resum.Cells(28, NumTi + 2).Font.Bold = True
        Resum.Cells(28, NumTi + 3).Value = "=TRUNC((RC" & NumTi + 2 & "*100)/R27C" & NumTi + 2 & ",1)"
        'Round(((TDevol * 100) / TServit), 2)
        Resum.Cells(28, NumTi + 3).Font.Bold = True
        Resum.Cells(29, NumTi + 2).Value = Round(TNeto, 2)
        Resum.Cells(29, NumTi + 2).Font.Bold = True
        Resum.Cells(29, NumTi + 3).Value = "=TRUNC((RC" & NumTi + 2 & "*100)/R24C" & NumTi + 2 & ",1)"
        'Round(((TNeto * 100) / Tz), 2)
        Resum.Cells(29, NumTi + 3).Font.Bold = True
        
        Resum.Cells(FilaIngR, NumTi + 2).Value = Round(TzM, 2)
        Resum.Cells(FilaIngR, NumTi + 2).Font.Bold = True
        Resum.Cells(FilaIngR + 1, NumTi + 2).Value = Round(TzT, 2)
        Resum.Cells(FilaIngR + 1, NumTi + 2).Font.Bold = True
        
        Resum.Cells(FilaHoresRR + 1, NumTi + 2).Value = Round(TMCostH, 2)
        Resum.Cells(FilaHoresRR + 1, NumTi + 2).Font.Bold = True
        Resum.Cells(FilaHoresRR + 1, NumTi + 3).Value = 0
        PerCostH = Round((THoresM / TIngM) * 100, 1)
             
        Resum.Cells(FilaHoresRR + 1, NumTi + 3).Value = "=TRUNC((RC" & NumTi + 2 & "/R" & FilaIngR & "C" & NumTi + 2 & ")*100,1)"
        'Round(((TMCostH * 100) / Tz), 2)
        Resum.Cells(FilaHoresRR + 1, NumTi + 3).Font.Bold = True
        
        Resum.Cells(FilaHoresRR + 2, NumTi + 2).Value = Round(TTCostH, 2)
        Resum.Cells(FilaHoresRR + 2, NumTi + 2).Font.Bold = True
        Resum.Cells(FilaHoresRR + 2, NumTi + 3).Value = 0
        Resum.Cells(FilaHoresRR + 2, NumTi + 3).Value = "=TRUNC((RC" & NumTi + 2 & "/R" & FilaIngR + 1 & "C" & NumTi + 2 & ")*100,1)"
        'Round(((TTCostH * 100) / Tz), 2)
        Resum.Cells(FilaHoresRR + 2, NumTi + 3).Font.Bold = True
        
        
        Resum.Cells(FilaHoresRR + 3, NumTi + 2).Value = Round(TCostH, 2)
        Resum.Cells(FilaHoresRR + 3, NumTi + 2).Font.Bold = True
        Resum.Cells(FilaHoresRR + 3, NumTi + 3).Value = 0
        Resum.Cells(FilaHoresRR + 3, NumTi + 3).Value = "=TRUNC((RC" & NumTi + 2 & "/R24C" & NumTi + 2 & ")*100,1)"
        'Round(((TCostH * 100) / Tz), 2)
        Resum.Cells(FilaHoresRR + 3, NumTi + 3).Font.Bold = True
    End If
    
'    Hoja.Range("H:I").Copy
'    Resum.Range("C:D").Select
'    Resum.Paste , True
'    MsExcel.CutCopyMode = False
    
    'Hoja.Range("H:I").Copy
    'Resum.Range("C:D").Select
    'Resum.PasteSpecial "Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= False, Transpose:=False"
    'MsExcel.CutCopyMode = False
    'Application.CutCopyMode = False
'    Sheets("T--13OBRADOR").Select
'    Range("B5:H10").Select
'    Selection.Copy
'    Sheets("Resumen").Select
'    Range("B14").Select
'    ActiveSheet.Paste Link:=True
    
    
'    MaxFam = MaxFam + 1
'    rellenaHojaDiaDeLaSetmanaBuscaDadesGrafic Dia, Client, GraficX, GraficY
'    For i = 1 To UBound(GraficX)
'        Hoja.Cells(21 + MaxFam, i).Value = GraficX(i)
'        Hoja.Cells(22 + MaxFam, i).Value = GraficY(i)
'    Next
'    DoEvents
'    Hoja.Rows(21 + MaxFam & ":" & 22 + MaxFam).Select
'    DoEvents
'    Set Kk = Hoja.ChartObjects.Add(Left:=600, Width:=375, Top:=75, Height:=225)
'    DoEvents
'    HOJA.ChartObjects(kk.Name).Chart.ChartType = xlLineMarkers
'    HOJA.ChartObjects(kk.Name).Chart.SetSourceData Source:=HOJA.Range("A30:CK31"), PlotBy:=xlRows
'    HOJA.ChartObjects(kk.Name).Chart.Location Where:=xlLocationAsObject, Name:="T--01"
'    HOJA.ChartObjects(kk.Name).Chart.HasAxis(xlCategory, xlPrimary) = True
'    HOJA.ChartObjects(kk.Name).Chart.HasAxis(xlValue, xlPrimary) = True
'    HOJA.ChartObjects(kk.Name).Chart.Axes(xlCategory, xlPrimary).CategoryType = xlAutomatic
    
    
nooo:


End Sub

Private Sub rellenaHojaHoresPerBotiga(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, client As String)
    Dim i As Integer, DiaS As Integer, Rs As rdoResultset, rsC As ADODB.Recordset, data() As Double, mes As Integer, sql As String, D As Date, Boti As Double, LaSql As String, Hi, Acc, Ac, dep, Ndep, depSeg, Di, Mi, Ai, Df, Mf, Af
    Dim cM, cT, zM, zT, hM, hT, DepsM(), DepsT(), DescM, DescT, Dependentes(), DependentesPct(), DependentesPctAcu(), GraficX(), GraficY()
    Dim cMm, cTm, zMm, zTm, hMm, hTm, Col, cha As Chart, Kk As Object, MaxFam, j
    Dim color As Integer
    
    Hoja.Name = BotigaCodiNom(client)
    For i = 0 To 6
        D = DateAdd("d", i, dia)
        Hoja.Cells(1, 2 + i + 1).Value = Format(D, "dd mmmm yy")
        Hoja.Cells(2, 2 + i + 1).Value = Format(D, "dddd")
    Next

    Di = Day(dia)
    Mi = Month(dia)
    Ai = Right(Year(dia), 2)
    
    Df = Day(DateAdd("d", 6, dia))
    Mf = Month(DateAdd("d", 6, dia))
    Af = Right(Year(DateAdd("d", 6, dia)), 2)
       
    Set Rs = Db.OpenResultset("Select * from (select * from [" & NomTaulaHoraris(dia) & "] union select * from [" & NomTaulaHoraris(DateAdd("m", 1, dia)) & "] ) a Where botiga = " & client & " and data between convert(datetime,'" & Di & "/" & Mi & "/" & Ai & "',3)  and convert(datetime,'" & Df & "/" & Mf & "/" & Af & "',3)+convert(datetime,'23:59:59',8) Order by dependenta,data ")
    Hi = ""
    Acc = 0
    dep = 0
    Ndep = 0
    color = 0
    While Not Rs.EOF
        If Rs("Operacio") = "E" Then Hi = Rs("Data")
        If Rs("Operacio") = "P" Then
            If Hi = "" Then
                color = 1
            Else
                If (Day(Hi) = Day(Rs("Data"))) Then
                    Ac = Round(DateDiff("n", Hi, Rs("Data")) / 60, 2)
                    i = DateDiff("d", dia, Hi)
                    Acc = Acc + Ac
                    If Hoja.Cells(4 + Ndep, 3 + i).Value = "" Then
                        Hoja.Cells(4 + Ndep, 3 + i).Value = Ac
                    Else
                        color = 2
'                        HOJA.Cells(4 + Ndep, 3 + i).Interior.ColorIndex = 36
                        Hoja.Cells(4 + Ndep, 3 + i).Value = Hoja.Cells(4 + Ndep, 3 + i).Value + Ac
                    End If
                End If
            End If
            Hi = ""
        End If
        Rs.MoveNext
        depSeg = 0
        If Not Rs.EOF Then depSeg = Rs("Dependenta")
          
        If Not dep = depSeg Or dep = 0 Then
            If dep = 0 Then dep = depSeg
            Hoja.Cells(4 + Ndep, 1).Value = DependentaCodiNom(CDbl(dep))
            If Not dep = depSeg Then Ndep = Ndep + 1
            dep = depSeg
            color = 0
        End If
    Wend
    For i = 1 To Ndep
        Hoja.Cells(3 + i, 10).FormulaR1C1 = "=SUM(RC[-7]:RC[-1])"
        Hoja.Cells(3 + i, 10).Font.Bold = True
    Next
    For i = 1 To 7
        Hoja.Cells(5 + Ndep, 2 + i).FormulaR1C1 = "=SUM(R[-" & (Ndep + 2) & "]C:R[-1]C)"
        Hoja.Cells(5 + Ndep, 2 + i).Font.Bold = True
    Next
nooo:
End Sub





Private Sub rellenaHojaIngredients(ByRef Hoja As Excel.Worksheet, ByVal MatPrima As String, ByVal codiBotiga As String, ByVal nomBotiga As String, ByVal codiArticle As String, ByRef ArrDatesComanda() As Date)

    Dim Rs As ADODB.Recordset, rsServit As ADODB.Recordset, rsVenut As ADODB.Recordset, rsProds As ADODB.Recordset
    Dim MatPrimaNom As String, f As Integer, D As Integer, sql As String
    Dim PFin As Date, PIni As Date
    
    Hoja.Name = nomBotiga
    
    Set Rs = rec("select nombre from ccMateriasPrimas  where id = '" & MatPrima & "'")
    If Not Rs.EOF Then MatPrimaNom = Rs("nombre")
    
    Hoja.Cells(1, 1).Value = nomBotiga
    Hoja.Cells(1, 1).Font.Bold = True
    Hoja.Cells(2, 1).Value = MatPrimaNom
    Hoja.Cells(2, 1).Font.Bold = True
    
    'Dates
    For D = 1 To UBound(ArrDatesComanda)
        Hoja.Cells(4, D * 2).Value = Format(ArrDatesComanda(D), "dd/mm/yyyy")
        Hoja.Cells(4, D * 2).Font.Bold = True
        If D > 1 Then
            Set rsServit = rec("select sum(quantitatServida) Qs from " & DonamTaulaServit(ArrDatesComanda(D)) & " where CodiArticle='" & codiArticle & "' and client='" & codiBotiga & "'")
            If Not rsServit.EOF Then Hoja.Cells(4, (D * 2) - 1).Value = rsServit("Qs")
        End If
    Next
    
    'Productos que tienen como ingrediente la materia en estudio
    Set rsProds = rec("select a.NOM, a.codi, i.quantitat from ingredients i left join articles a on i.article=a.codi where i.materia='" & MatPrima & "'")
    f = 6
    While Not rsProds.EOF
        Hoja.Cells(f, 1).Value = rsProds("Nom")
        'Hoja.Cells(f, 1).AutoFit
        For D = 1 To UBound(ArrDatesComanda)
            PFin = ArrDatesComanda(D)
            If D + 1 <= UBound(ArrDatesComanda) Then
                PIni = ArrDatesComanda(D + 1)
                sql = "select isnull(SUM(quantitat), 0) venut "
                If Month(PIni) = Month(PFin) Then
                    sql = sql & "from [" & NomTaulaVentas(PIni) & "] v "
                Else
                    sql = sql & "from (select * from [" & NomTaulaVentas(PIni) & "] union all select * from [" & NomTaulaVentas(PFin) & "]) v "
                End If
                sql = sql & "where botiga='" & codiBotiga & "' and plu='" & rsProds("codi") & "' and "
                sql = sql & "Data between CONVERT(datetime,'" & Format(PIni, "dd/mm/yyyy") & "', 103) and CONVERT(datetime, '" & Format(PFin, "dd/mm/yyyy") & "', 103) + convert(datetime,'23:59:59',8)"
                Set rsVenut = rec(sql)
                If Not rsVenut.EOF Then
                    Hoja.Cells(f, D * 2).Value = rsVenut("venut")
                    Hoja.Cells(f, (D * 2) + 1).Value = rsVenut("venut") * rsProds("quantitat")
                End If
            End If
        Next
        
        f = f + 1
        rsProds.MoveNext
    Wend
    
    For D = 1 To UBound(ArrDatesComanda) - 1
      Hoja.Cells(5, (D * 2) + 1).FormulaR1C1 = "=SUM(R[1]C:R[" & (f - 6) & "]C)"
      Hoja.Cells(5, D * 2).FormulaR1C1 = "=RC[1]/R[-1]C[1]"
      Hoja.Cells(5, D * 2).NumberFormat = "0.00%"
    Next
    
    Hoja.Cells.EntireColumn.AutoFit
End Sub

Sub rellenaCobraments(Libro, dia As Date)
    Dim Hoja As Excel.Worksheet, D As Date, Rs, i, j, Rs2, K, Posat As Boolean, sql As String, Acc As Double, Pagades As String, Rs4
    Dim HojaResum As Excel.Worksheet
    
On Error Resume Next

    Set HojaResum = Libro.Sheets(Libro.Sheets.Count)
    HojaResum.Name = "Resum "
    
    ExecutaComandaSql "Drop Table CobramentsTmp"
    HojaResum.Cells(1, 1).Value = "Nombre"
    HojaResum.Cells(1, 1).Font.Bold = True
    HojaResum.Cells(1, 2).Value = "Fecha vencimiento"
    HojaResum.Cells(1, 2).Font.Bold = True
    HojaResum.Cells(1, 3).Value = "Debe"
    HojaResum.Cells(1, 3).Font.Bold = True
    HojaResum.Cells(1, 4).Value = "Haber"
    HojaResum.Cells(1, 4).Font.Bold = True
    HojaResum.Cells(1, 5).Value = "Saldo"
    HojaResum.Cells(1, 5).Font.Bold = True
    HojaResum.Cells(1, 6).Value = "Factura Num"
    HojaResum.Cells(1, 6).Font.Bold = True
        
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Hoja.Name = "Detall Clients "
    
    sql = "Select c.nom nom , c.codi Codi , c.[Nom Llarg] [Nom Llarg] , isnull(Cc.valor ,c.codi ) CodiContable "
    sql = sql & "from Clients c left join constantsclient cc on cc.variable ='CodiContable' and cc.codi = c.codi "
    sql = sql & "where c.codi not in (select codi from constantsclient where variable='EsCalaix' and valor='EsCalaix') "
    sql = sql & "order by [Nom Llarg]"
    
    Set Rs = Db.OpenResultset(sql)
    
    i = 2
    j = 2
        
    Hoja.Cells(1, 1).Value = "Nombre"
    Hoja.Cells(1, 1).Font.Bold = True
    Hoja.Cells(1, 2).Value = "Fecha vencimiento"
    Hoja.Cells(1, 2).Font.Bold = True
    Hoja.Cells(1, 3).Value = "Debe"
    Hoja.Cells(1, 3).Font.Bold = True
    Hoja.Cells(1, 4).Value = "Haber"
    Hoja.Cells(1, 4).Font.Bold = True
    Hoja.Cells(1, 5).Value = "Saldo"
    Hoja.Cells(1, 5).Font.Bold = True
    Hoja.Cells(1, 6).Value = "Factura Num"
    Hoja.Cells(1, 6).Font.Bold = True
    
'    Hoja.Columns("B:B").NumberFormat = "dd-mm-yy"
    Hoja.Columns("C:E").NumberFormat = "#,##0.00"
    
    ExecutaComandaSql "Drop TABLE TmpCobraments"
    ExecutaComandaSql "CREATE TABLE TmpCobraments ([Tipo]     [nvarchar] (255) NULL ,[Deve]     [nvarchar] (255) NULL ,[Haver]     [nvarchar] (255) NULL ,[Data]     [datetime] NULL ,[Concepto]     [nvarchar] (255) NULL)"
    
    
    While Not Rs.EOF
        Informa "Cobraments per : " & BotigaCodiNom(Rs("Codi"))
        Posat = False
        Acc = 0

        ExecutaComandaSql "Delete TmpCobraments"
        
        'FACTURAS
        For K = 1 To 12
            D = DateSerial(Year(dia), K, 1)
            If ExisteixTaula(NomTaulaFacturaIva(D)) Then
                ExecutaComandaSql "insert into TmpCobraments Select 'F' ,total,0,dataVenciment,numfactura  from [" & NomTaulaFacturaIva(D) & "] where clientcodi = " & Rs("Codi") & " "
            End If
        Next
      
        'Sql = "Insert Into TmpCobraments  Select Distinct 'P' Tipo,c4.valor deve,0.0001 Haver,c5.valor Data,c2.valor Descrip from [" & DonamNomTaulaNorma43Conta & "]  c1 "
        'Sql = Sql & "join  [norma43conta]  c2 on c1.idnorma43 = c2.idnorma43 and c2.concepto = 'FacturaNum' and c2.orden=2 "
        'Sql = Sql & "join  [norma43conta]  c3 on c1.idnorma43 = c3.idnorma43 and c3.concepto = 'DEBE' and c3.orden=1 "
        'Sql = Sql & "join  [norma43conta]  c4 on c1.idnorma43 = c4.idnorma43 and c4.concepto = 'DEBE' and c4.orden=1 "
        'Sql = Sql & "join  [norma43conta]  c5 on c1.idnorma43 = c5.idnorma43 and c5.concepto = 'FECHA' and c5.orden=1 "
        'Sql = Sql & "join  [norma43conta]  c6 on c1.idnorma43 = c6.idnorma43 and c6.concepto = 'SUBCUENTA' and c6.orden=2 and c6.valor = '" & (43000000 + Rs("CodiContable")) & "' "
        'Sql = Sql & "Where c1.concepto = 'FacturaId' "

        'PAGOS
        sql = "Insert Into TmpCobraments Select Distinct 'P' Tipo, DEBE deve, HABER haver, FECHA Data, FacturaNum Descrip "
        sql = sql & "From " & DonamNomTaulaAsientosContables(dia) & " "
        sql = sql & "where SUBCUENTA='" & (43000000 + Rs("CodiContable")) & "' "
        ExecutaComandaSql sql
        
        Set Rs2 = Db.OpenResultset("Select Concepto from TmpCobraments Where  Tipo = 'P' ")
        While Not Rs2.EOF
            Pagades = Pagades & "," & Rs2("Concepto") & ","
            Rs2.MoveNext
        Wend
        
        Set Rs2 = Db.OpenResultset("Select * from TmpCobraments order by data ")
        While Not Rs2.EOF
            If Not Posat Then
                Hoja.Cells(i, 1).Value = Rs("Nom Llarg") & " (" & Rs("Nom") & ")"
                HojaResum.Cells(j, 1).Value = Rs("Nom Llarg") & " (" & Rs("Nom") & ")"
                Posat = True
            End If
            
            If Rs2("Tipo") = "P" Then
                Hoja.Cells(i, 2).Value = Format(Rs2("data"), "mm-dd-yyyy")
                Hoja.Cells(i, 4).Value = Format(Round(Rs2("haver"), 2), "#.00")
                Hoja.Cells(i, 6).Value = Rs2("Concepto")
                Acc = Acc + Round(Rs2("haver"), 2)
            Else
                Hoja.Cells(i, 2).Value = Format(Rs2("data"), "mm-dd-yyyy")
                Hoja.Cells(i, 3).Value = Format(Round(Rs2("Deve"), 2), "#.00")
                Hoja.Cells(i, 6).Value = Rs2("Concepto")
                Acc = Acc - Round(Rs2("Deve"), 2)
                If InStr(Pagades, "," & Rs2("Concepto") & ",") > 0 Then
                    Hoja.Cells(i, 3).Font.ColorIndex = 5
                Else
                    If CDate(Rs2("data")) < Now() Then 'Si ha vencido el pago
                        HojaResum.Cells(j, 2).Value = Format(Rs2("data"), "mm-dd-yyyy")
                        HojaResum.Cells(j, 3).Value = Format(Round(Rs2("Deve"), 2), "#.00")
                        HojaResum.Cells(j, 6).Value = Rs2("Concepto")
                        j = j + 1
                    End If
                End If
            End If

            Hoja.Cells(i, 5).Value = Format(Round(Acc, 2), "#.00")
            If Acc < 0 Then Hoja.Cells(i, 5).Font.ColorIndex = 3 Else Hoja.Cells(i, 5).Font.ColorIndex = 4
            Rs2.MoveNext
            i = i + 1
        Wend
        Rs2.Close
        Rs.MoveNext
    Wend
    Rs.Close
    Hoja.Columns("A:A").EntireColumn.AutoFit
    Hoja.Columns("B:B").EntireColumn.AutoFit
        
    HojaResum.Columns("A:A").EntireColumn.AutoFit
    
'************************* Control Rutas
'    Set Rs4 = Db.OpenResultset("select * from rutas Order by nom ")
'    While Not Rs4.EOF
'        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
'        Set Hoja = Libro.Sheets(Libro.Sheets.Count)
'        Hoja.Name = Left("Ruta " & Rs4("Nom"), 30)
'        Set rs = Db.OpenResultset("Select c.nom nom , c.codi Codi , c.[Nom Llarg] [Nom Llarg] , isnull(Cc.valor ,c.codi ) CodiContable  from Clients c left join constantsclient cc on cc.variable ='CodiContable' and cc.codi = c.codi where c.codi in (select Cli  from rutascli where ruta = '" & Rs4("Id") & "') order by [Nom Llarg]")
'        i = 4
'        Hoja.Cells(1, 1).Value = "Nombre"
'        Hoja.Cells(1, 1).Font.Bold = True
'        Hoja.Cells(1, 2).Value = "Fecha"
'        Hoja.Cells(1, 2).Font.Bold = True
'        Hoja.Cells(1, 3).Value = "Debe"
'        Hoja.Cells(1, 3).Font.Bold = True
'        Hoja.Cells(1, 4).Value = "Haber"
'        Hoja.Cells(1, 4).Font.Bold = True
'        Hoja.Cells(1, 5).Value = "Saldo"
'        Hoja.Cells(1, 5).Font.Bold = True
'        Hoja.Cells(1, 5).Value = "Saldo"
'        Hoja.Cells(1, 5).Font.Bold = True
'        Hoja.Cells(1, 7).Value = "Factura Num"
'        Hoja.Cells(1, 7).Font.Bold = True
'        Hoja.Columns("C:E").NumberFormat = "#,##0.00"
'        ExecutaComandaSql "Drop TABLE TmpCobraments"
'        ExecutaComandaSql "CREATE TABLE TmpCobraments ([Tipo]     [nvarchar] (255) NULL ,[Deve]     [nvarchar] (255) NULL ,[Haver]     [nvarchar] (255) NULL ,[Data]     [datetime] NULL ,[Concepto]     [nvarchar] (255) NULL)"
'        While Not rs.EOF
'            Informa "Cobraments per : " & BotigaCodiNom(rs("Codi"))
'            If InStr(BotigaCodiNom(rs("Codi")), "arolina") > 0 Then
'                Acc = Acc
'            End If
'            Posat = False
'            Acc = 0
'            ExecutaComandaSql "Delete TmpCobraments"
'            For K = 1 To 12
'            D = DateSerial(Year(dia), K, 1)
'                If ExisteixTaula(NomTaulaFacturaIva(D)) Then
'                    ExecutaComandaSql "insert into TmpCobraments Select 'F' ,total,0,dataFactura,numfactura  from [" & NomTaulaFacturaIva(D) & "] where clientcodi = " & rs("Codi") & " "
'                End If
'            Next
'            'Sql = "Insert Into TmpCobraments  Select Distinct 'P' Tipo,c4.valor deve,0.0001 Haver,c5.valor Data,c2.valor Descrip from [" & DonamNomTaulaNorma43Conta & "]  c1 "
'            'Sql = Sql & "join  [norma43conta]  c2 on c1.idnorma43 = c2.idnorma43 and c2.concepto = 'FacturaNum' and c2.orden=2 "
'            'Sql = Sql & "join  [norma43conta]  c3 on c1.idnorma43 = c3.idnorma43 and c3.concepto = 'DEBE' and c3.orden=1 "
'            'Sql = Sql & "join  [norma43conta]  c4 on c1.idnorma43 = c4.idnorma43 and c4.concepto = 'DEBE' and c4.orden=1 "
'            'Sql = Sql & "join  [norma43conta]  c5 on c1.idnorma43 = c5.idnorma43 and c5.concepto = 'FECHA' and c5.orden=1 "
'            'Sql = Sql & "join  [norma43conta]  c6 on c1.idnorma43 = c6.idnorma43 and c6.concepto = 'SUBCUENTA' and c6.orden=2 and c6.valor = '" & (43000000 + Rs("CodiContable")) & "' "
'            'Sql = Sql & "Where c1.concepto = 'FacturaId' "
'
'            sql = "Insert Into TmpCobraments Select Distinct 'P' Tipo, DEBE deve, 0.0001 Haver, FECHA Data, FacturaNum Descrip "
'            sql = sql & "From " & DonamNomTaulaAsientosContables(dia) & " "
'            sql = sql & "where SUBCUENTA='" & (43000000 + rs("CodiContable")) & "' "
'            ExecutaComandaSql sql
'
'            Set Rs2 = Db.OpenResultset("Select Concepto from TmpCobraments Where  Tipo = 'P' ")
'            While Not Rs2.EOF
'                Pagades = Pagades & "," & Rs2("Concepto") & ","
'                Rs2.MoveNext
'            Wend
'
'            Set Rs2 = Db.OpenResultset("Select * from TmpCobraments order by data ")
'            While Not Rs2.EOF
'                If Not Posat Then Hoja.Cells(i, 1).Value = rs("Nom Llarg") & " (" & rs("Nom") & ")"
'                Posat = True
'                If Rs2("Tipo") = "P" Then
'                    Hoja.Cells(i, 2).Value = Format(Rs2("data"), "mm-dd-yyyy")
'                    Hoja.Cells(i, 4).Value = Format(Round(Rs2("Deve"), 2), "#.00")
'                    Hoja.Cells(i, 7).Value = Rs2("Concepto")
'                    Acc = Acc + Round(Rs2("Deve"), 2)
'                Else
'                    Hoja.Cells(i, 2).Value = Format(Rs2("data"), "mm-dd-yyyy")
'                    Hoja.Cells(i, 3).Value = Format(Round(Rs2("Deve"), 2), "#.00")
'                    Hoja.Cells(i, 7).Value = Rs2("Concepto")
'                    Acc = Acc - Round(Rs2("Deve"), 2)
'                    If InStr(Pagades, "," & Rs2("Concepto") & ",") > 0 Then
'                        Hoja.Cells(i, 3).Font.ColorIndex = 5
'                    End If
'                End If
'                Hoja.Cells(i, 5).Value = Format(Round(Acc, 2), "#.00")
'                If Acc < 0 Then Hoja.Cells(i, 5).Font.ColorIndex = 3 Else Hoja.Cells(i, 5).Font.ColorIndex = 4
'                Rs2.MoveNext
'                i = i + 1
'            Wend
'            Rs2.Close
'            rs.MoveNext
'        Wend
'        rs.Close
'        Hoja.Columns("A:A").EntireColumn.AutoFit
'        Hoja.Columns("B:B").EntireColumn.AutoFit
'        Rs4.MoveNext
'    Wend
'    Rs4.Close
'hasta aquí**************************
'    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
'    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
'    Hoja.Name = "Detall Proveedors "
'    Set Rs = Db.OpenResultset("select * from CcProveedores order by nombre")
'    i = 4
'    Hoja.Cells(1, 1).Value = "Nom"
'    Hoja.Cells(1, 1).Font.Bold = True
'    Hoja.Cells(1, 2).Value = "Saldo"
'    Hoja.Cells(1, 2).Font.Bold = True
'    Hoja.Cells(1, 3).Value = "Acc Debe"
'    Hoja.Cells(1, 3).Font.Bold = True
'    Hoja.Cells(1, 4).Value = "Acc Haber"
'    Hoja.Cells(1, 4).Font.Bold = True
'
'    While Not Rs.EOF
'        Hoja.Cells(i, 1).Value = Rs("nombre")
'        i = i + 1
'        Rs.MoveNext
'    Wend
'    Rs.Close
'    Hoja.Columns("A:A").EntireColumn.AutoFit
'    Hoja.Columns("B:B").EntireColumn.AutoFit
'    Hoja.Cells.Select
'    Hoja.Cells.EntireColumn.AutoFit
'
'
nooo:
End Sub



Sub rellenaPreus(Libro)
    Dim Hoja As Excel.Worksheet, D As Date, Rs As ADODB.Recordset, i, Rs2, K, Posat As Boolean, sql As String, Acc As Double
On Error Resume Next
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Hoja.Name = "Preus "
    Set Rs = rec("select p.valor CodiExtern,a.codi CodiIntern,a.nom Nom ,ff.pare Familia1,f.pare Familia2,a.familia Familia3 ,t.Iva ,a.preu PreuVp,a.preumajor PreuVm from articles a left join tipusiva2012 t on t.tipus = a.tipoiva left join articlespropietats p on p.codiarticle=a.codi and p.variable = 'CODI_PROD' left join families f on f.nom = a.familia left join families ff on f.pare = ff.nom  order by Familia1,Familia2,Familia3,a.Nom ,a.codi ")
    Hoja.Range(Hoja.Cells(3, 1), Hoja.Cells(Rs.Fields.Count, Rs.Fields.Count)).CopyFromRecordset Rs
    Hoja.Cells(2, 1).Value = "Codi Extern"
    Hoja.Cells(2, 2).Value = "Codi Intern"
    Hoja.Cells(2, 3).Value = "Nom"
    Hoja.Cells(2, 4).Value = "Familia 1"
    Hoja.Cells(2, 5).Value = "Familia 2"
    Hoja.Cells(2, 6).Value = "Familia 3"
    Hoja.Cells(2, 7).Value = "Iva"
    Hoja.Cells(2, 8).Value = "Preu Vd"
    Hoja.Cells(2, 9).Value = "Preu Vm"
    
    i = 0
    Set Rs2 = Db.OpenResultset("Select distinct TarifaCodi,TarifaNom from tarifesespecials order by tarifanom")
    While Not Rs2.EOF
        Set Rs = rec("select isnull(cast(e.preu as nvarchar),'') PreuVp,isnull(cast(e.preumajor as nvarchar),'') PreuVm from articles a left join families f on f.nom = a.familia left join families ff on f.pare = ff.nom left join tarifesespecials e on e.codi = a.codi and tarifacodi='" & Rs2("TarifaCodi") & "' order by ff.pare,ff.nom,a.familia,a.nom,a.codi")
        Hoja.Range(Hoja.Cells(3, 10 + (2 * i)), Hoja.Cells(Rs.Fields.Count + 3, 12 + (2 * i))).CopyFromRecordset Rs
        Hoja.Cells(1, 10 + (2 * i)).Value = "Tarifa"
        Hoja.Cells(1, 10 + (2 * i) + 1).Value = Rs2("TarifaNom")
        
'        Hoja.Cells(1, 10 + (2 * i) + 1).AddComment
'        Hoja.Cells(1, 10 + (2 * i) + 1).Comment.Text Text:="" & Rs2("Tarifanom")
'        Hoja.Cells(1, 10 + (2 * i) + 1).Comment.Visible = True
        
        Rs2.MoveNext
        i = i + 1
    Wend
    
    Set Rs2 = Db.OpenResultset("Select distinct Client,nom From tarifesespecialsclients  join Clients C on c.codi = tarifesespecialsclients .Client order by nom")
    While Not Rs2.EOF
        Set Rs = rec("select isnull(cast(e.preu as nvarchar),'') PreuVp,isnull(cast(e.preumajor as nvarchar),'') PreuVm from articles a left join families f on f.nom = a.familia left join families ff on f.pare = ff.nom left join tarifesespecialsclients e on e.codi = a.codi and Client='" & Rs2("Client") & "' order by ff.pare,ff.nom,a.familia,a.nom,a.codi")
        Hoja.Range(Hoja.Cells(3, 10 + (2 * i)), Hoja.Cells(Rs.Fields.Count + 3, 12 + (2 * i))).CopyFromRecordset Rs
        Hoja.Cells(1, 10 + (2 * i)).Value = "Client"
        Hoja.Cells(2, 10 + (2 * i)).Value = Rs2("nom")
        Rs2.MoveNext
        i = i + 1
    Wend
    
    
    Hoja.Columns("H:I").Select
    Hoja.Columns("H:I").Replace ".", ",", xlPart, xlByRows, False
    Hoja.Columns("H:I").HorizontalAlignment = xlRight

'    Selection.NumberFormat = "0.00"
        
    Hoja.Rows("2:2").Interior.ColorIndex = 36
 '       .Pattern = xlSolid
    Hoja.Cells.HorizontalAlignment = xlRight
    Hoja.Columns("C:F").HorizontalAlignment = xlLeft
    Hoja.Rows("1:2").HorizontalAlignment = xlLeft
    Hoja.Rows("1:1").Font.Bold = True
    Hoja.Rows("2:2").Font.Bold = True
    Hoja.Cells.Select
    Hoja.Cells.EntireColumn.AutoFit
    Hoja.Cells(1, 1).Value = Format(Now, "dddd dd-mm-yy hh:mm")
    Hoja.Range("D3").Select
    Hoja.Application.ActiveWindow.FreezePanes = True
    Columns("A:A").EntireColumn.Hidden = True
    Columns("B:B").EntireColumn.Hidden = True
    
    
nooo:
End Sub



Sub rellenaPreusPerClients(Libro)
    Dim Hoja As Excel.Worksheet, D As Date, Rs As ADODB.Recordset, i, Rs2, K, Posat As Boolean, sql As String, Acc As Double
On Error Resume Next

    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Hoja.Name = "Preus Especials"
    Set Rs = rec("Select isnull(codi,0) ,isnull(nom,'') Nom From articles order by Nom ,codi ")
    
'    Hoja.Columns("H:I").Select
'    Hoja.Columns("H:I").Replace ".", ",", xlPart, xlByRows, False
'    Hoja.Columns("H:I").HorizontalAlignment = xlRight
    
    Hoja.Range(Hoja.Cells(3, 1), Hoja.Cells(Rs.Fields.Count, Rs.Fields.Count)).CopyFromRecordset Rs
    Hoja.Cells(2, 1).Value = "Codi"
    Hoja.Cells(2, 2).Value = "Nom"
    i = 0
    Set Rs2 = Db.OpenResultset("Select distinct Client,nom From tarifesespecialsclients  join Clients C on c.codi = tarifesespecialsclients .Client order by nom")
    While Not Rs2.EOF
        Set Rs = rec("select isnull(cast(e.preu as nvarchar),'') PreuVp,isnull(cast(e.preumajor as nvarchar),'') PreuVm from articles a left join tarifesespecialsclients e on e.codi = a.codi and Client='" & Rs2("Client") & "' order by a.nom,a.codi")
        Hoja.Range(Hoja.Cells(3, 3 + (2 * i)), Hoja.Cells(Rs.Fields.Count + 3, 5 + (2 * i))).CopyFromRecordset Rs
        
        
        Hoja.Cells(1, 3 + (2 * i)).Value = "Client"
        Hoja.Cells(2, 3 + (2 * i)).Value = Rs2("nom")
        Rs2.MoveNext
        i = i + 1
    Wend
    
    Hoja.Rows("2:2").Interior.ColorIndex = 36
    Hoja.Cells.HorizontalAlignment = xlRight
    Hoja.Columns("C:F").HorizontalAlignment = xlLeft
    Hoja.Rows("1:2").HorizontalAlignment = xlLeft
    Hoja.Rows("1:1").Font.Bold = True
    Hoja.Rows("2:2").Font.Bold = True
    Hoja.Cells.Select
    Hoja.Cells.EntireColumn.AutoFit
    Hoja.Cells(1, 1).Value = Format(Now, "dddd dd-mm-yy hh:mm")
    Hoja.Range("D3").Select
    Hoja.Application.ActiveWindow.FreezePanes = True
    Columns("A:A").EntireColumn.Hidden = True
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Hoja.Name = "Descomptes Per Producte"
    Set Rs = rec("select codi ,nom Nom From articles order by Nom ,codi ")
    
    Hoja.Range(Hoja.Cells(3, 1), Hoja.Cells(Rs.Fields.Count, Rs.Fields.Count)).CopyFromRecordset Rs
    Hoja.Cells(2, 1).Value = "Codi"
    Hoja.Cells(2, 2).Value = "Nom"
    
    i = 0
    Set Rs2 = Db.OpenResultset("Select distinct cc.Codi,Nom From constantsclient cc join Clients C on c.codi = cc.Codi where variable = 'DtoProducte' ")
    While Not Rs2.EOF
        Set Rs = rec("Select isnull(cast(right(Valor,len(Valor) - charindex('|', Valor)) as nvarchar),'') PreuVp From articles a Left Join constantsclient cc on cc.codi ='" & Rs2("Codi") & "' And a.codi = left(Valor, charindex('|', Valor) -1)  and variable ='DtoProducte' Order by Nom ,a.codi")
        
        Hoja.Range(Hoja.Cells(3, 3 + (2 * i)), Hoja.Cells(Rs.Fields.Count + 3, 4 + (2 * i))).CopyFromRecordset Rs
        Hoja.Cells(1, 3 + (2 * i)).Value = "Client"
        Hoja.Cells(2, 3 + (2 * i)).Value = Rs2("nom")
        Rs2.MoveNext
        i = i + 1
    Wend
    
    
    Hoja.Rows("2:2").Interior.ColorIndex = 36
 '       .Pattern = xlSolid
    Hoja.Cells.HorizontalAlignment = xlRight
    Hoja.Columns("C:F").HorizontalAlignment = xlLeft
    Hoja.Rows("1:2").HorizontalAlignment = xlLeft
    Hoja.Rows("1:1").Font.Bold = True
    Hoja.Rows("2:2").Font.Bold = True
    Hoja.Cells.Select
    Hoja.Cells.EntireColumn.AutoFit
    Hoja.Cells(1, 1).Value = Format(Now, "dddd dd-mm-yy hh:mm")
    Hoja.Range("D3").Select
    Hoja.Application.ActiveWindow.FreezePanes = True
'    Columns("A:A").EntireColumn.Hidden = True
'    Columns("B:B").EntireColumn.Hidden = True
    
    Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Hoja.Name = "Descomptes Per Familia"
    Set Rs = rec("select  Nom From Families order by Nom ")
    
    Hoja.Range(Hoja.Cells(3, 1), Hoja.Cells(Rs.Fields.Count, Rs.Fields.Count)).CopyFromRecordset Rs
    Hoja.Cells(2, 1).Value = "Nom"
    
    i = 0
    Set Rs2 = Db.OpenResultset("Select distinct cc.Codi,Nom From constantsclient cc join Clients C on c.codi = cc.Codi where variable = 'DtoFamilia' ")
    While Not Rs2.EOF
        Set Rs = rec("Select Isnull(cast(right(Valor,len(Valor) - charindex('|', Valor)) as nvarchar),'') PreuVp From Families f Left Join constantsclient cc on cc.codi ='" & Rs2("Codi") & "' And f.nom COLLATE SQL_Latin1_General_CP1_CI_AS  = left(Valor, charindex('|', Valor) -1)  and variable ='DtoFamilia' order by f.nom")

        Hoja.Range(Hoja.Cells(3, 3 + (2 * i)), Hoja.Cells(Rs.Fields.Count + 3, 4 + (2 * i))).CopyFromRecordset Rs
        Hoja.Cells(1, 3 + (2 * i)).Value = "Client"
        Hoja.Cells(2, 3 + (2 * i)).Value = Rs2("nom")
        Rs2.MoveNext
        i = i + 1
    Wend
    
    
    Hoja.Rows("2:2").Interior.ColorIndex = 36
    Hoja.Cells.HorizontalAlignment = xlRight
    Hoja.Columns("C:F").HorizontalAlignment = xlLeft
    Hoja.Rows("1:2").HorizontalAlignment = xlLeft
    Hoja.Rows("1:1").Font.Bold = True
    Hoja.Rows("2:2").Font.Bold = True
    Hoja.Cells.Select
    Hoja.Cells.EntireColumn.AutoFit
    Hoja.Cells(1, 1).Value = Format(Now, "dddd dd-mm-yy hh:mm")
    Hoja.Range("D3").Select
    Hoja.Application.ActiveWindow.FreezePanes = True
'    Columns("A:A").EntireColumn.Hidden = True
'    Columns("B:B").EntireColumn.Hidden = True
    
    
nooo:
End Sub




Sub RendimentEquipsUn(Hoja, dia As Date, equip As String)
    Dim Rs As ADODB.Recordset, sql As String, i
    
    Hoja.Cells(1, 2).Value = "Producto"
    Hoja.Cells(1, 3).Value = "Cantidad"
    Hoja.Cells(1, 4).Value = "Importe"
    
    sql = "Select a.nom,sum(quantitatservida),sum(a.preu * quantitatservida) "
    sql = sql & "From [" & DonamNomTaulaServit(dia) & "] s "
    sql = sql & "left join articles a on a.codi = s.codiarticle "
    sql = sql & "left join families f on a.familia = f.nom "
    sql = sql & "left join families ff on ff.nom=f.pare "
    sql = sql & "where equip = '" & equip & "' "
    sql = sql & "group by a.nom,ff.pare,f.pare,a.familia,a.nom "
    sql = sql & "order by ff.pare,f.pare,a.familia,a.nom "
    Set Rs = rec(sql)
    
    Hoja.Range(Hoja.Cells(1, 2), Hoja.Cells(Rs.Fields.Count + 1, 4)).CopyFromRecordset Rs
    
    CarregaHoresEquip dia, equip
'    Hoja.Range(Hoja.Cells(1, 2), Hoja.Cells(Rs.Fields.Count + 1, 4)).CopyFromRecordset Rs
    
    

    
'    Hoja.Cells.HorizontalAlignment = xlRight
'    Hoja.Columns("C:F").HorizontalAlignment = xlLeft
'    Hoja.Rows("1:2").HorizontalAlignment = xlLeft
'    Hoja.Rows("1:1").Font.Bold = True
'    Hoja.Rows("2:2").Font.Bold = True
'    Hoja.Cells.Select
'    Hoja.Cells.EntireColumn.AutoFit
'    Hoja.Cells(1, 1).Value = Format(Now, "dddd dd-mm-yy hh:mm")
'    Hoja.Range("D3").Select
'    Hoja.Application.ActiveWindow.FreezePanes = True
'    Columns("A:A").EntireColumn.Hidden = True
'    Columns("B:B").EntireColumn.Hidden = True


End Sub



Sub CondicionsEspecials(Libro)
    Dim Hoja As Excel.Worksheet, D As Date, Rs As ADODB.Recordset, i, Rs2, K, Posat As Boolean, sql As String, Acc As Double
On Error Resume Next
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Hoja.Name = "Preus Especials"
    
    Set Rs = rec("select p.valor CodiExtern,a.codi CodiIntern,a.nom Nom ,ff.pare Familia1,f.pare Familia2,a.familia Familia3 ,t.Iva ,a.preu PreuVp,a.preumajor PreuVm from articles a left join tipusiva t on t.tipus = a.tipoiva left join articlespropietats p on p.codiarticle=a.codi and p.variable = 'CODI_PROD' left join families f on f.nom = a.familia left join families ff on f.pare = ff.nom  order by Familia1,Familia2,Familia3,a.Nom ,a.codi ")
    Hoja.Range(Hoja.Cells(3, 1), Hoja.Cells(Rs.Fields.Count, Rs.Fields.Count)).CopyFromRecordset Rs
    Hoja.Cells(2, 1).Value = "Codi Extern"
    Hoja.Cells(2, 2).Value = "Codi Intern"
    Hoja.Cells(2, 3).Value = "Nom"
    Hoja.Cells(2, 4).Value = "Familia 1"
    Hoja.Cells(2, 5).Value = "Familia 2"
    Hoja.Cells(2, 6).Value = "Familia 3"
    Hoja.Cells(2, 7).Value = "Iva"
    Hoja.Cells(2, 8).Value = "Preu Vd"
    Hoja.Cells(2, 9).Value = "Preu Vm"
    
    i = 0
    Set Rs2 = Db.OpenResultset("Select distinct TarifaCodi,TarifaNom from tarifesespecials order by tarifanom")
    While Not Rs2.EOF
        Set Rs = rec("select isnull(cast(e.preu as nvarchar),'') PreuVp,isnull(cast(e.preumajor as nvarchar),'') PreuVm from articles a left join families f on f.nom = a.familia left join families ff on f.pare = ff.nom left join tarifesespecials e on e.codi = a.codi and tarifacodi='" & Rs2("TarifaCodi") & "' order by ff.pare,ff.nom,a.familia,a.nom,a.codi")
        Hoja.Range(Hoja.Cells(3, 10 + (2 * i)), Hoja.Cells(Rs.Fields.Count + 3, 12 + (2 * i))).CopyFromRecordset Rs
        Hoja.Cells(1, 10 + (2 * i)).Value = "Tarifa"
        Hoja.Cells(1, 10 + (2 * i) + 1).Value = Rs2("TarifaNom")
        
'        Hoja.Cells(1, 10 + (2 * i) + 1).AddComment
'        Hoja.Cells(1, 10 + (2 * i) + 1).Comment.Text Text:="" & Rs2("Tarifanom")
'        Hoja.Cells(1, 10 + (2 * i) + 1).Comment.Visible = True
        
        Rs2.MoveNext
        i = i + 1
    Wend
    
    Set Rs2 = Db.OpenResultset("Select distinct Client,nom From tarifesespecialsclients  join Clients C on c.codi = tarifesespecialsclients .Client order by nom")
    While Not Rs2.EOF
        Set Rs = rec("select isnull(cast(e.preu as nvarchar),'') PreuVp,isnull(cast(e.preumajor as nvarchar),'') PreuVm from articles a left join families f on f.nom = a.familia left join families ff on f.pare = ff.nom left join tarifesespecialsclients e on e.codi = a.codi and Client='" & Rs2("Client") & "' order by ff.pare,ff.nom,a.familia,a.nom,a.codi")
        Hoja.Range(Hoja.Cells(3, 10 + (2 * i)), Hoja.Cells(Rs.Fields.Count + 3, 12 + (2 * i))).CopyFromRecordset Rs
        Hoja.Cells(1, 10 + (2 * i)).Value = "Client"
        Hoja.Cells(2, 10 + (2 * i)).Value = Rs2("nom")
        Rs2.MoveNext
        i = i + 1
    Wend
    
    Hoja.Rows("2:2").Interior.ColorIndex = 36
 '       .Pattern = xlSolid
    Hoja.Cells.HorizontalAlignment = xlRight
    Hoja.Columns("C:F").HorizontalAlignment = xlLeft
    Hoja.Rows("1:2").HorizontalAlignment = xlLeft
    Hoja.Rows("1:1").Font.Bold = True
    Hoja.Rows("2:2").Font.Bold = True
    Hoja.Cells.Select
    Hoja.Cells.EntireColumn.AutoFit
    Hoja.Cells(1, 1).Value = Format(Now, "dddd dd-mm-yy hh:mm")
    Hoja.Range("D3").Select
    Hoja.Application.ActiveWindow.FreezePanes = True
    Columns("A:A").EntireColumn.Hidden = True
    Columns("B:B").EntireColumn.Hidden = True
    
    
nooo:
End Sub




Sub RendimentEquips(Libro, dia As Date)
    Dim Hoja As Excel.Worksheet, D As Date, Rs As rdoResultset, i, Rs2, K, Posat As Boolean, sql As String, Acc As Double
    Dim equip As String
    
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Set Rs = Db.OpenResultset("Select Distinct Equip From [" & DonamNomTaulaServit(dia) & "]  ")
    While Not Rs.EOF
        equip = Rs("Equip")
        If equip = "" Then equip = "Otros"
        Hoja.Name = equip
        RendimentEquipsUn Hoja, dia, equip
        Rs.MoveNext
        If Not Rs.EOF Then
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            Set Hoja = Libro.Sheets(Libro.Sheets.Count)
        End If
    Wend
    Rs.Close
    
End Sub



Sub FeinaSetmanal(Libro, dia As Date)
    Dim Hoja As Excel.Worksheet, D As Date, Rs As rdoResultset, i, Rs2, K, Posat As Boolean, sql As String, Acc As Double
    Dim equip As String
    
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Set Rs = Db.OpenResultset("Select Distinct Equip From [" & DonamNomTaulaServit(dia) & "]  ")
    While Not Rs.EOF
        equip = Rs("Equip")
        If equip = "" Then equip = "Otros"
        Hoja.Name = equip
        RendimentEquipsUn Hoja, dia, equip
        Rs.MoveNext
        If Not Rs.EOF Then
            Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
            Set Hoja = Libro.Sheets(Libro.Sheets.Count)
        End If
    Wend
    Rs.Close
    
End Sub




Sub ExcelInventari(Libro, dia As Date)
    Dim Hoja As Excel.Worksheet, D As Date, Rs As rdoResultset, i, Rs2, K, Posat As Boolean, sql As String, Acc As Double
    Dim equip As String

    
    D = Now
    ExecutaComandaSql "drop table tmp"
    ExecutaComandaSql "Select * into Tmp  from [" & NomTaulaVentas(D) & "] "
    ExecutaComandaSql "Insert into Tmp select * from [" & NomTaulaVentas(DateAdd("m", -1, D)) & "] "
    ExecutaComandaSql "Insert into Tmp select * from [" & NomTaulaVentas(DateAdd("m", -2, D)) & "] "
    
    ExecutaComandaSql "Insert into Tmp Select Botiga,DATA,dependenta,num_tick,Estat,plu,Quantitat,Import,Tipus_venta,'999' FormaMarcar,'999' Otros from  [" & NomTaulaInventari(D) & "] "
    ExecutaComandaSql "Insert into Tmp Select Botiga,DATA,dependenta,num_tick,Estat,plu,Quantitat,Import,Tipus_venta,'999' FormaMarcar,'999' Otros from  [" & NomTaulaInventari(DateAdd("m", -1, D)) & "] "
    ExecutaComandaSql "Insert into Tmp Select Botiga,DATA,dependenta,num_tick,Estat,plu,Quantitat,Import,Tipus_venta,'999' FormaMarcar,'999' Otros from  [" & NomTaulaInventari(DateAdd("m", -2, D)) & "] "
    
    'ExecutaComandaSql "Delete a from tmp a join (select MAX(data) data,botiga,Plu from Tmp   where otros = '999' group by botiga ,Plu) b on a.botiga = b.botiga and a.Plu = b.Plu and a.data < b.data"
    
    
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Set Rs = Db.OpenResultset("select c.Nom Nom ,c.Codi Codi from paramshw w join clients c on c.Codi = w.valor1 order by c.nom")
    
    While Not Rs.EOF
        equip = Rs("Nom")
        
        sql = ""
        'Sql = Sql & "Select ff.pare F1,f.pare F2 ,f.Nom F3 ,a.nom Article, ISNULL(v.plu,ISNULL(i.plu,isnull(i.plu,0))) plu ,isnull(i.i,0) Inventari,ISNULL( v.v,0) Venut,isnull(r.r,0) recepcionat, a.preu, a.preu * (isnull(i.i,0)  -  ISNULL( v.v,0) + isnull(r.r,0)) ValorStock  from "
        sql = sql & "Select ff.pare F1,f.pare F2 ,f.Nom F3 ,a.nom Article, a.codi plu ,isnull(i.i,0) Inventari,ISNULL( v.v,0) Venut,isnull(r.r,0) recepcionat, a.preu, a.preu * (isnull(i.i,0)  -  ISNULL( v.v,0) + isnull(r.r,0)) ValorStock  from "
        sql = sql & "(select SUM(quantitat) v ,0 i ,0 r,plu from Tmp  where botiga = " & Rs("Codi") & " and otros='0' group by plu ) v "
        sql = sql & "Left Join "
        sql = sql & "(select 0 v ,SUM(quantitat) I,0 r ,plu from Tmp  where botiga = " & Rs("Codi") & " and otros='999' and tipus_venta = 1 group by plu ) i "
        sql = sql & "on v.plu=i.plu "
        sql = sql & "Left Join "
        sql = sql & "(select 0 v ,0 i ,SUM(quantitat) r,plu from Tmp  where botiga = " & Rs("Codi") & " and otros='999' and tipus_venta = 2 group by plu ) r "
        sql = sql & "on v.plu=r.plu "
        sql = sql & "right join articles a on a.codi = ISNULL(v.plu,ISNULL(i.plu,isnull(r.plu,0))) "
        sql = sql & "left join families f on f.Nom = a.Familia left join families ff on ff.Nom = f.pare "
        sql = sql & "order by ff.pare,f.pare,f.Nom,a.Nom "

        Libro.Sheets.Add , Libro.Sheets(Libro.Sheets.Count)
        Set Hoja = Libro.Sheets(Libro.Sheets.Count)
        rellenaHojaSql equip, sql, Libro.Sheets(Libro.Sheets.Count), 0
        Rs.MoveNext
    Wend
    Rs.Close
    
End Sub





Sub rellenaClients(Libro)
    Dim Hoja As Excel.Worksheet, D As Date, Rs As ADODB.Recordset, i, Rs2, K, Posat As Boolean, sql As String, Acc As Double
On Error Resume Next
    ExecutaComandaSql "CREATE INDEX [ConstantsClient_Ind_CodiVariable] ON [ConstantsClient] (Codi,Variable) "
    Set Hoja = Libro.Sheets(Libro.Sheets.Count)
    Hoja.Name = "Clients"
    
    
    ExecutaComandaSql "Select * into [ConstantsClient" & Format(Now, "yyyymmddhhnnss") & "]  from ConstantsClient "
    ExecutaComandaSql "Drop Table Tmp1 "
    ExecutaComandaSql "select distinct codi,variable,valor into Tmp1 from ConstantsClient"
    ExecutaComandaSql "Delete ConstantsClient"
    ExecutaComandaSql "insert into ConstantsClient  select * from Tmp1 "
    ExecutaComandaSql "delete cc From constantsclient cc left join Clients c on c.codi = cc.codi  where c.codi is null "
    
    sql = ""
    sql = sql & "select top 1000 "
    sql = sql & "c.Codi CodiClient,"
    sql = sql & "Nom NomCurt,"
    sql = sql & "Nif,"
    sql = sql & "Adresa,"
    sql = sql & "ciutat,"
    sql = sql & "cp,"
    sql = sql & "lliure,"
    sql = sql & "[Nom Llarg],"
    sql = sql & "[Tipus Iva],"
    sql = sql & "[Preu Base],"
    sql = sql & "[Desconte ProntoPago],"
    sql = sql & "isnull(T.TarifaNom,'') Tarifa,"
    sql = sql & "AlbaraValorat,"
    sql = sql & "isnull(C1.Valor,'') Grup_client,"
    sql = sql & "isnull(C2.Valor,'') EsCalaix,"
    sql = sql & "isnull(C3.Valor,'') Fax,"
    sql = sql & "isnull(C4.Valor,'') Drebuts,"
    sql = sql & "isnull(C5.Valor,'') COPIES_ALB,"
    sql = sql & "isnull(C6.Valor,'') Nrebuts,"
    sql = sql & "isnull(C7.Valor,'') Adr_Entrega,"
    sql = sql & "isnull(C8.Valor,'') Acreedor,"
    sql = sql & "isnull(C9.Valor,'') DescTe,"
    sql = sql & "isnull(C10.Valor,'') EsClient,"
    sql = sql & "isnull(C11.Valor,'') CFINAL,"
    sql = sql & "isnull(C12.Valor,'') Email,"
    sql = sql & "isnull(C13.Valor,'') DiaPagament,"
    sql = sql & "isnull(C14.Valor,'') CodiContable,"
    sql = sql & "isnull(C15.Valor,'') P_Contacte,"
    sql = sql & "isnull(C16.Valor,'') Venciment,"
    sql = sql & "isnull(C17.Valor,'') Per_Facturacio,"
    sql = sql & "isnull(C18.Valor,'') Idioma,"
    sql = sql & "isnull(C19.Valor,'') NomClientFactura,"
    sql = sql & "isnull(C20.Valor,'') DtoFamilia,"
    sql = sql & "isnull(C21.Valor,'') Tel,"
    sql = sql & "isnull(C22.Valor,'') CompteCorrent,"
    sql = sql & "isnull(C23.Valor,'') FormaPago,"
    sql = sql & "isnull(C24.Valor,'') FormaPagoLlista,"
    sql = sql & "isnull(C25.Valor,'') NoDevolucions,"
    sql = sql & "isnull(C26.Valor,'') NoPagaEnTienda,"
    sql = sql & "isnull(C27.Valor,'') AlbaransValorats "
    
    sql = sql & "from clients C with (nolock) "
    
    sql = sql & "Left Join ConstantsClient C1 with (nolock) on C1.Codi = C.Codi And c1.Variable = 'Grup_client' "
    sql = sql & "Left Join ConstantsClient C2 with (nolock) on C2.Codi = C.Codi And c2.Variable = 'EsCalaix' "
    sql = sql & "Left Join ConstantsClient C3 with (nolock) on C3.Codi = C.Codi And c3.Variable = 'Fax' "
    sql = sql & "Left Join ConstantsClient C4 with (nolock) on C4.Codi = C.Codi And c4.Variable = 'Drebuts' "
    sql = sql & "Left Join ConstantsClient C5 with (nolock) on C5.Codi = C.Codi And c5.Variable = 'COPIES_ALB' "
    sql = sql & "Left Join ConstantsClient C6 with (nolock) on C6.Codi = C.Codi And c6.Variable = 'Nrebuts' "
    sql = sql & "Left Join ConstantsClient C7 with (nolock) on C7.Codi = C.Codi And c7.Variable = 'Adr_Entrega' "
    sql = sql & "Left Join ConstantsClient C8 with (nolock) on C8.Codi = C.Codi And c8.Variable = 'Acreedor' "
    sql = sql & "Left Join ConstantsClient C9 with (nolock) on C9.Codi = C.Codi And c9.Variable = 'DescTe' "
    sql = sql & "Left Join ConstantsClient C10 with (nolock) on C10.Codi = C.Codi And c10.Variable = 'EsClient' "
    sql = sql & "Left Join ConstantsClient C11 with (nolock) on C11.Codi = C.Codi And c11.Variable = 'CFINAL' "
    sql = sql & "Left Join ConstantsClient C12 with (nolock) on C12.Codi = C.Codi And c12.Variable = 'Email' "
    sql = sql & "Left Join ConstantsClient C13 with (nolock) on C13.Codi = C.Codi And c13.Variable = 'DiaPagament' "
    sql = sql & "Left Join ConstantsClient C14 with (nolock) on C14.Codi = C.Codi And c14.Variable = 'codiContable' "
    sql = sql & "Left Join ConstantsClient C15 with (nolock) on C15.Codi = C.Codi And c15.Variable = 'P_Contacte' "
    sql = sql & "Left Join ConstantsClient C16 with (nolock) on C16.Codi = C.Codi And c16.Variable = 'Venciment' "
    sql = sql & "Left Join ConstantsClient C17 with (nolock) on C17.Codi = C.Codi And c17.Variable = 'Per_Facturacio' "
    sql = sql & "Left Join ConstantsClient C18 with (nolock) on C18.Codi = C.Codi And c18.Variable = 'Idioma' "
    sql = sql & "Left Join ConstantsClient C19 with (nolock) on C19.Codi = C.Codi And c19.Variable = 'NomClientFactura' "
    sql = sql & "Left Join ConstantsClient C20 with (nolock) on C20.Codi = C.Codi And c20.Variable = 'DtoFamiliaNO OK' "
    sql = sql & "Left Join ConstantsClient C21 with (nolock) on C21.Codi = C.Codi And c21.Variable = 'Tel' "
    sql = sql & "Left Join ConstantsClient C22 with (nolock) on C22.Codi = C.Codi And c22.Variable = 'CompteCorrent' "
    sql = sql & "Left Join ConstantsClient C23 with (nolock) on C23.Codi = C.Codi And c23.Variable = 'FormaPago' "
    sql = sql & "Left Join ConstantsClient C24 with (nolock) on C24.Codi = C.Codi And c24.Variable = 'FormaPagoLlista' "
    sql = sql & "Left Join ConstantsClient C25 with (nolock) on C25.Codi = C.Codi And c25.Variable = 'NoDevolucions' "
    sql = sql & "Left Join ConstantsClient C26 with (nolock) on C26.Codi = C.Codi And c26.Variable = 'NoPagaEnTienda' "
    sql = sql & "Left Join ConstantsClient C27 with (nolock) on C27.Codi = C.Codi And c27.Variable = 'AlbaransValorats' "
    sql = sql & "Left Join (select distinct Tarifacodi,Tarifanom From TarifesEspecials with (nolock) ) T on t.TarifaCodi = C.[Desconte 5]  "
      
    sql = sql & " Order By NomCurt"
   
    Set Rs = rec(sql)
    
    Hoja.Range(Hoja.Cells(3, 1), Hoja.Cells(Rs.Fields.Count, Rs.Fields.Count)).CopyFromRecordset Rs
    
    For i = 0 To Rs.Fields.Count
        Hoja.Cells(1, i + 1).Value = Rs.Fields(i).Name
        Hoja.Cells(2, i + 1).Value = Rs.Fields(i).Name
    Next
    
    Hoja.Rows("2:2").Interior.ColorIndex = 36
 '       .Pattern = xlSolid
    Hoja.Cells.HorizontalAlignment = xlLeft
    Hoja.Columns("C:F").HorizontalAlignment = xlLeft
    Hoja.Rows("1:2").HorizontalAlignment = xlLeft
    Hoja.Rows("1:1").Font.Bold = True
    Hoja.Rows("2:2").Font.Bold = True
    Hoja.Cells.Select
    Hoja.Cells.EntireColumn.AutoFit
    Hoja.Cells(1, 1).Value = Format(Now, "dddd dd-mm-yy hh:mm")
    Hoja.Range("C3").Select
    Hoja.Application.ActiveWindow.FreezePanes = True
    Columns("A:A").EntireColumn.Hidden = True
    Hoja.Rows("1:1").EntireRow.Hidden = True

nooo:
End Sub




Public Sub rellenaHojaSql(Titol As String, sql, ByRef Hoja As Excel.Worksheet, Optional Off As Integer = 0)
    Dim Rs As ADODB.Recordset
    Dim i
    Dim html As String
    
On Error GoTo nor44:

    Set Rs = rec(sql)
    
    If Off = 0 Then
        For i = 0 To Rs.Fields.Count - 1
            Hoja.Cells(1, i + 1).Value = Rs.Fields(i).Name
            Hoja.Cells(1, i + 1).Font.Bold = True
            'Hoja.Columns("A:Z"). _
                NumberFormat = "@"
        Next
        ExcelNombreHoja Hoja, Titol
    End If

    Hoja.Range(Hoja.Cells(2, 1 + Off), Hoja.Cells(2, Off + Rs.Fields.Count)).CopyFromRecordset Rs
    
    If Not Off = 0 Then
        Hoja.Cells(1, Off + 1).Value = Titol
        Hoja.Cells(1, Off + 1).Font.Bold = True
    End If
    
    Hoja.Cells.Select
    Hoja.Cells.EntireColumn.AutoFit
    Hoja.Cells.EntireRow.AutoFit
    
nor44:
 If err.Number <> 0 Then
        html = "<p><h3>Error excel </h3></p>"
        html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
        html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
        html = html & "<p><b>ERROR:</b>" & err.Source & "</p>"
        html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
            
        sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR! ExcelFR ha fallat", html, "", ""
            
        Informa "Error : " & err.Description
    End If
End Sub


Sub ExcelNombreHoja(Hoja, Titol)

On Error GoTo nor:
    Hoja.Select
    Hoja.Name = Left(Titol, 30)

nor:
    

End Sub

Private Sub rellenaHojaRentabilidad(ByRef Hoja As Excel.Worksheet, ByVal dia As Date, EmpresaNom, EmpresaCodi)
    Dim i As Integer, DiaS As Integer, Rs As ADODB.Recordset, rsC As ADODB.Recordset, data() As Double, mes As Integer, sql As String, D As Date, Boti As Double
    Dim cM, cT, zM, zT, hM, hT, DepsM(), DepsT(), DescM, DescT, Families(), GraficX(), GraficY(), RecMa, RecTa
    Dim cMm, cTm, zMm, zTm, hMm, hTm, Col, cha As Chart, Kk As Object, MaxFam, j, ProveeNom() As String, ProveeImport() As String, codiEmp As String
    Dim Rss
    
On Error GoTo nooo

    codiEmp = "0"
    If InStr(EmpresaCodi, "_") > 0 Then codiEmp = Left(EmpresaCodi, InStr(EmpresaCodi, "_") - 1)
    ReDim ProveeNom(0)
    ReDim ProveeImport(0)
    If EmpresaNom = "" Then EmpresaNom = "nova" & Now
    Hoja.Name = Left(EmpresaNom, 30)
    Hoja.Cells(1, 1).Value = EmpresaNom & " Hoja de rentabilidad"
    Hoja.Cells(1, 1).Font.Bold = True
    Hoja.Cells(3, 1).Value = "Semana Del " & Format(dia, "dd") & " Al " & Format(DateAdd("d", 7, dia), "dd") & " De " & Format(DateAdd("d", 7, dia), "mmmm")
    Hoja.Cells(3, 17).Value = "Total Semanal"
    Hoja.Cells(3, 17).Font.Bold = True
    
    Hoja.Cells(5, 1).Value = "Recaudacion "
    Hoja.Cells(5, 1).Font.Bold = True
    Hoja.Cells(6, 1).Value = "Produccion "
    Hoja.Cells(6, 1).Font.Bold = True

    Hoja.Cells(8, 1).Value = "TOTAL Recaudacion"
    Hoja.Rows("8:8").Font.Bold = True


    cMm = 0: cTm = 0: zMm = 0: zTm = 0: hMm = 0: hTm = 0: MaxFam = 0
    Col = 1
    For i = 0 To 6
       CreaTaulesDadesTpv DateAdd("d", i, dia)
       CarregaProveedors codiEmp, EmpresaNom, DateAdd("d", i, dia), ProveeNom, ProveeImport, RecMa, RecTa
    Next
    
    Hoja.Cells(10 + UBound(ProveeNom) + 2, 1).Value = "TOTAL GASTOS"
    Hoja.Rows(10 + UBound(ProveeNom) + 2 & ":" & 10 + UBound(ProveeNom) + 2).Font.Bold = True
    Hoja.Cells(10 + UBound(ProveeNom) + 6, 1).Value = "DIFERENCIA"
    Hoja.Rows(10 + UBound(ProveeNom) + 6 & ":" & 10 + UBound(ProveeNom) + 6).Font.Bold = True
    
    For i = 0 To 6
        zM = 0
        Col = Col + 2
        Hoja.Cells(3, Col).Value = Format(DateAdd("d", i, dia), "dddd") & " " & Format(DateAdd("d", i, dia), "dd")
        CarregaProveedors codiEmp, EmpresaNom, DateAdd("d", i, dia), ProveeNom, ProveeImport, RecMa, RecTa
        Hoja.Cells(5, Col).Value = RecMa
        Hoja.Cells(6, Col).Value = RecTa
        Hoja.Cells(8, Col).Value = RecMa + RecTa

        For j = 1 To UBound(ProveeNom)
            Hoja.Cells(10 + j, 1).Value = ProveeNom(j)
            If ProveeImport(j) = "" Then ProveeImport(j) = 0
            Hoja.Cells(10 + j, Col).Value = Round(ProveeImport(j), 2)
            zM = zM + ProveeImport(j)
        Next
        Hoja.Cells(10 + j + 1, Col).Value = Round(zM, 2)
        Hoja.Cells(10 + j + 5, Col).Value = (RecMa + RecTa) - Round(zM, 2)
        If RecMa + RecTa > 0 Then Hoja.Cells(10 + j + 5, Col + 1).Value = 100 - Round(Round(zM, 2) / (RecMa + RecTa), 2) * 100
        If RecMa + RecTa > 0 Then Hoja.Cells(10 + j + 1, Col + 1).Value = Int(Round(zM, 2) / (RecMa + RecTa) * 100)
        
        If codiEmp = 3 Then
            Set Rs = rec("select a.nom,sum(case a.desconte when 1 then QuantitatServida*preumajor*(1-cast(c.[Desconte 1] as real) / 100 ) when 2 then quantitatservida*preumajor*(1-cast(c.[Desconte 2] as real) / 100 ) when 3 then quantitatservida*preumajor*(1-cast(c.[Desconte 3] as real) / 100 ) else quantitatservida*preumajor end) from [" & DonamNomTaulaServit(DateAdd("d", i, dia)) & "] s join articles a on a.codi = s.codiarticle join Clients  c on c.codi = s.client  where codiarticle in (select codiarticle from articlespropietats   where variable = 'EMP_FACTURA' and valor = 3 ) Group by  a.nom")
            Hoja.Range(Hoja.Cells(22 + UBound(ProveeNom), Col), Hoja.Cells(22 + UBound(ProveeNom), Col)).CopyFromRecordset Rs
        Else
            Set Rs = rec("select  a.nom,sum(case a.desconte when 1 then QuantitatServida*preumajor*(1-cast(c.[Desconte 1] as real) / 100 ) when 2 then quantitatservida*preumajor*(1-cast(c.[Desconte 2] as real) / 100 ) when 3 then quantitatservida*preumajor*(1-cast(c.[Desconte 3] as real) / 100 ) else quantitatservida*preumajor end) from [" & DonamNomTaulaServit(DateAdd("d", i, dia)) & "] s join articles a on a.codi = s.codiarticle join Clients  c on c.codi = s.client  where codiarticle not in (select codiarticle from articlespropietats   where variable = 'EMP_FACTURA' and valor = 3 ) Group By  a.nom")
            Hoja.Range(Hoja.Cells(22 + UBound(ProveeNom), Col), Hoja.Cells(22 + UBound(ProveeNom), Col)).CopyFromRecordset Rs
        End If
        
    Next
    
    
    Hoja.Columns("A:A").EntireColumn.AutoFit
    Hoja.Columns("B:B").EntireColumn.AutoFit
    
    Hoja.Columns("A:Z").NumberFormat = "#,##0.00"
    
    Hoja.Columns("D:D").NumberFormat = "#,##0"
    Hoja.Columns("D:D").ColumnWidth = 2.43
    Hoja.Columns("F:F").NumberFormat = "#,##0"
    Hoja.Columns("F:F").ColumnWidth = 2.43
    Hoja.Columns("H:H").NumberFormat = "#,##0"
    Hoja.Columns("H:H").ColumnWidth = 2.43
    Hoja.Columns("J:J").NumberFormat = "#,##0"
    Hoja.Columns("J:J").ColumnWidth = 2.43
    Hoja.Columns("L:L").NumberFormat = "#,##0"
    Hoja.Columns("L:L").ColumnWidth = 2.43
    Hoja.Columns("N:N").NumberFormat = "#,##0"
    Hoja.Columns("N:N").ColumnWidth = 2.43
    Hoja.Columns("P:P").NumberFormat = "#,##0"
    Hoja.Columns("P:P").ColumnWidth = 2.43
    
    Hoja.Cells(5, 17).FormulaR1C1 = "=RC[-2]+RC[-4]+RC[-6]+RC[-8]+RC[-10]+RC[-12]+RC[-14]"
    Hoja.Cells(6, 17).FormulaR1C1 = "=RC[-2]+RC[-4]+RC[-6]+RC[-8]+RC[-10]+RC[-12]+RC[-14]"
    Hoja.Cells(8, 17).FormulaR1C1 = "=RC[-2]+RC[-4]+RC[-6]+RC[-8]+RC[-10]+RC[-12]+RC[-14]"
    For j = 1 To UBound(ProveeNom)
        Hoja.Cells(10 + j, 17).Value = "=RC[-2]+RC[-4]+RC[-6]+RC[-8]+RC[-10]+RC[-12]+RC[-14]"
    Next
    Hoja.Cells(10 + j + 1, 17).Value = "=RC[-2]+RC[-4]+RC[-6]+RC[-8]+RC[-10]+RC[-12]+RC[-14]"
    Hoja.Cells(10 + j + 5, 17).Value = "=RC[-2]+RC[-4]+RC[-6]+RC[-8]+RC[-10]+RC[-12]+RC[-14]"
        
nooo:
End Sub




Private Sub CalculaExcelCentralCalcArticles(data() As String)
    Dim i As Integer
    Dim rsC As ADODB.Recordset
    Dim F2Pare As String, F2Nom As String, FNom As String
    Dim sql As String, MaxArt As Double
    
    sql = "Select Count(*) from Articles a join families f on a.familia = f.nom  left join families f2 on f.pare = f2.nom  "
    Set rsC = rec(sql)
    
    If rsC.EOF Then Exit Sub
    Dim Tope
    Tope = rsC(0) + 3
    ReDim data(Tope, 6)
    
    sql = "Select f2.pare F2Pare,f2.nom F2Nom,f.nom FNom,a.nom ANom from Articles a join families f on a.familia = f.nom  left join families f2 on f.pare = f2.nom order by f2.pare,f2.nom,f.nom,a.nom "
    
    sql = "Select sum(import) as Euros,sum(quantitat) as quantitat,avi,pare,familia,nom "
    sql = sql & "from ( "
    sql = sql & "Select 0 import,0 quantitat,isnull(f2.pare,'ZZ Eliminat') Avi ,isnull(f2.nom,'ZZ Eliminat') Pare ,isnull(f.nom,'ZZ Eliminat') Familia ,isnull(a.nom,'ZZ Eliminat') nom  From Articles a left join families f on a.familia = f.nom  left join families f2 on f.pare = f2.nom "
'    Sql = Sql & "union "
'    Sql = Sql & "Select Sum(Quantitat) Quantitat,Sum(import) Import ,isnull(f2.pare,'ZZ Eliminat') Avi ,isnull(f2.nom,'ZZ Eliminat') Pare ,isnull(f.nom,'ZZ Eliminat') Familia ,isnull(a.nom,'ZZ Eliminat') nom  From Articles a left join families f on a.familia = f.nom  left join families f2 on f.pare = f2.nom "
'    Sql = Sql & "right join [" & NomTaulaVentas(DateSerial(An, Mes, 1)) & "] v on v.plu = a.codi group by isnull(f2.pare,'ZZ Eliminat'),isnull(f2.nom,'ZZ Eliminat'),isnull(f.nom,'ZZ Eliminat'),isnull(a.nom,'ZZ Eliminat') "
    sql = sql & ") f "
    sql = sql & "group by avi,pare,familia,nom "
    sql = sql & "order by avi,pare,familia,nom "
    
    Set rsC = rec(sql)
    MaxArt = 0
    F2Pare = ""
    F2Nom = ""
    FNom = ""
    data(1, 5) = "TOTAL €"
    data(1, 6) = "TOTAL Q"
    MaxArt = 2
    While Not rsC.EOF
        If MaxArt < Tope Then
            If Not F2Pare = rsC("avi") Then data(MaxArt, 1) = rsC("avi")
            If Not F2Nom = rsC("pare") Then data(MaxArt, 2) = rsC("pare")
            If Not FNom = rsC("familia") Then data(MaxArt, 3) = rsC("familia")
            data(MaxArt, 4) = rsC("nom")
        End If
        
        F2Pare = rsC("avi")
        F2Nom = rsC("pare")
        FNom = rsC("familia")
        
        rsC.MoveNext
        MaxArt = MaxArt + 1
    Wend
    
End Sub


Private Sub CalculaExcelCentralAddArticles(ByRef Hoja As Excel.Worksheet, Titol As String, data() As String)
    Dim i As Integer, Formula As String
    
    Hoja.Name = Titol
    
    Hoja.Range(Hoja.Cells(1, 1), Hoja.Cells(UBound(data, 1), 7)).Value = data
    
    Formula = "=SUM("
    For i = 1 To 12
        If Not Formula = "=SUM(" Then Formula = Formula & ","
        Formula = Formula & "RC[" & i * 3 & "]"
    Next
    Hoja.Range(Hoja.Cells(3, 6), Hoja.Cells(UBound(data, 1), 6)).Formula = Formula  ' "=SUM(RC[2],RC[5],RC[8],RC[11],RC[14],RC[17],RC[14],RC[16],RC[18],RC[20],RC[22],RC[24])"
    Hoja.Range(Hoja.Cells(3, 7), Hoja.Cells(UBound(data, 1), 7)).Formula = Formula '"=SUM(RC[3],RC[5],RC[7],RC[9],RC[11],RC[13],RC[15],RC[17],RC[19],RC[21],RC[23],RC[25])"
    Hoja.Range(Hoja.Cells(3, 8), Hoja.Cells(UBound(data, 1), 8)).Formula = Formula ' "=SUM(RC[3],RC[5],RC[7],RC[9],RC[11],RC[13],RC[15],RC[17],RC[19],RC[21],RC[23],RC[25])"
   
    Hoja.Range(Hoja.Cells(UBound(data, 1) + 3, 6), Hoja.Cells(UBound(data, 1) + 3, 8 + (12 * 3))).Formula = "=SUM(R[-" & UBound(data, 1) & "]C:R[-1]C)"
    Hoja.Range(Hoja.Cells(3, 4), Hoja.Cells(UBound(data, 1) + 3, 8 + (3 * 12))).NumberFormat = "#,##0.00"
   
    Hoja.Range(Hoja.Cells(UBound(data, 1) + 4, 6), Hoja.Cells(UBound(data, 1) + 4, 8 + (12 * 3))).Formula = "=SUM(R[" & 1 & "]C:R[" & 1 + 24 & "]C)"
    Hoja.Range(Hoja.Cells(UBound(data, 1) + 4, 1), Hoja.Cells(UBound(data, 1) + 4, 1)).Value = "Detall (Import X Hora) , (Clients X Hora)"
    
    
    Set Hoja = Nothing

End Sub



Private Sub CalculaExcelCentralAddMes(ByRef Hoja As Excel.Worksheet, An As Double, mes As Double, Maxim As Double, clients As String)
    Dim i As Integer
    Dim rsC As ADODB.Recordset
    Dim data() As Double, F2Pare As String, F2Nom As String, FNom As String
    Dim sql As String, MaxArt As Double
On Error GoTo nor
    ReDim data(Maxim, 2)
    
    sql = "Select sum(NumClients) as NumClients,sum(import) as Euros,sum(quantitat) as quantitat,avi,pare,familia,nom "
    sql = sql & "from ( "
    sql = sql & "Select 0 as NumClients,0 import,0 quantitat,isnull(f2.pare,'ZZ Eliminat') Avi ,isnull(f2.nom,'ZZ Eliminat') Pare ,isnull(f.nom,'ZZ Eliminat') Familia ,isnull(a.nom,'ZZ Eliminat') nom  From Articles a left join families f on a.familia = f.nom  left join families f2 on f.pare = f2.nom "
    sql = sql & "union "
    sql = sql & "Select count(distinct ([Num_tick] + botiga + data + dependenta) ) as NumClients,Sum(import) Import ,Sum(Quantitat) Quantitat,isnull(f2.pare,'ZZ Eliminat') Avi ,isnull(f2.nom,'ZZ Eliminat') Pare ,isnull(f.nom,'ZZ Eliminat') Familia ,isnull(a.nom,'ZZ Eliminat') nom  From Articles a left join families f on a.familia = f.nom  left join families f2 on f.pare = f2.nom "
    sql = sql & "right join [" & NomTaulaVentas(DateSerial(An, mes, 1)) & "] v on v.plu = a.codi Where botiga in (" & clients & ") group by isnull(f2.pare,'ZZ Eliminat'),isnull(f2.nom,'ZZ Eliminat'),isnull(f.nom,'ZZ Eliminat'),isnull(a.nom,'ZZ Eliminat') "
    sql = sql & ") f "
    sql = sql & "group by avi,pare,familia,nom "
    sql = sql & "order by avi,pare,familia,nom "
    
    Set rsC = rec(sql)
    If rsC.EOF Then Exit Sub
    i = 2
    While Not rsC.EOF
        data(i, 0) = data(i, 0) + rsC("Euros")  ' Format(rsC("Euros"), "#,#.00")
        data(i, 1) = data(i, 1) + rsC("Quantitat") ' Format(rsC("Quantitat"), "#,#.000")
        data(i, 2) = data(i, 2) + rsC("NumClients") ' Format(rsC("NumClients"), "#,#.000")
        
        rsC.MoveNext
        i = i + 1
        If i > Maxim Then i = Maxim
    Wend
    
    Hoja.Range(Hoja.Cells(1, 3 * mes + 6), Hoja.Cells(i, 3 * mes + 8)).Value = data
    
    Hoja.Range(Hoja.Cells(1, 3 * mes + 6), Hoja.Cells(1, 3 * mes + 8)).Value = UCase(Format(DateSerial(An, mes, 1), "mmmm"))
    Hoja.Range(Hoja.Cells(2, 3 * mes + 6), Hoja.Cells(2, 3 * mes + 6)).Value = "Import"
    Hoja.Range(Hoja.Cells(2, 3 * mes + 7), Hoja.Cells(2, 3 * mes + 7)).Value = "Quantitat"
    Hoja.Range(Hoja.Cells(2, 3 * mes + 8), Hoja.Cells(2, 3 * mes + 8)).Value = "Clients"
    
    sql = "Select datepart(hh,Data) as Hora, Sum(Import) as Venut,count(distinct ([Num_tick] + botiga + data + dependenta) ) as NumClients "
    sql = sql & "From [" & NomTaulaVentas(DateSerial(An, mes, 1)) & "] "
    sql = sql & "Where botiga in (" & clients & ") Group by datepart(hh,Data)"
    
    Set rsC = rec(sql)
    If rsC.EOF Then Exit Sub
    DoEvents
    ReDim data(24, 2)
    For i = 0 To 24
       data(i, 0) = i
    Next
    
    i = 0
    While Not rsC.EOF
        data(rsC("Hora"), 1) = data(rsC("Hora"), 1) + rsC("NumClients")
        data(rsC("Hora"), 2) = data(rsC("Hora"), 2) + rsC("Venut")
        rsC.MoveNext
    Wend
    
    Hoja.Range(Hoja.Cells(Maxim + 5, 3 * mes + 6), Hoja.Cells(Maxim + (5 + 24), 3 * mes + 8)).Value = data
    DoEvents
    
    Set Hoja = Nothing

nor:


End Sub



Private Sub CalculaExcelAnualAddMes(ByRef Hoja As Excel.Worksheet, dia As Date, Maxim As Double, clients As String, Col As Integer, Camp As String)
    Dim i As Integer, dd As Date, mes As Integer
    Dim rsC As ADODB.Recordset
    Dim data() As Double, F2Pare As String, F2Nom As String, FNom As String
    Dim sql As String, MaxArt As Double
'On Error GoTo nor
    
    ReDim data(Maxim, 1)
    dd = dia
    mes = Month(dd)
    sql = "Select sum(Q) ,a.codi,isnull(f2.pare,'ZZ Eliminat') Avi ,isnull(f2.nom,'ZZ Eliminat') Pare ,isnull(f.nom,'ZZ Eliminat') Familia ,isnull(a.nom,'ZZ Eliminat') nom "
    sql = sql & "from ( "
    sql = sql & "Select 0 as Q,Codi From Articles  "
    sql = sql & ""
    While Month(dd) = mes
        sql = sql & "union  Select sum(" & Camp & ") as Q,CodiArticle as codi From [" & DonamNomTaulaServit(dd) & "] group by codiarticle "
        dd = DateAdd("d", 1, dd)
    Wend
    sql = sql & ") s "
    sql = sql & "join articles a on a.codi = s.codi Left join families f on a.familia = f.nom left join families f2 on f.pare = f2.nom "
    sql = sql & "group by  a.codi,isnull(f2.pare,'ZZ Eliminat') ,isnull(f2.nom,'ZZ Eliminat') ,isnull(f.nom,'ZZ Eliminat') ,isnull(a.nom,'ZZ Eliminat') "
    sql = sql & "order by avi,pare,familia,nom "
    Set rsC = rec(sql)
    
    If rsC.EOF Then Exit Sub
    i = 0
    While Not rsC.EOF
        If i < Maxim Then data(i, 0) = CDbl(rsC(0)) ' Replace(rsC(0), ".", ",")
        i = i + 1
        rsC.MoveNext
    Wend
    
    Hoja.Range(Hoja.Cells(3, 7 + Col), Hoja.Cells(Maxim + 3, 7 + Col + 1)).Value = data
   
    Set Hoja = Nothing
nor:

End Sub




Sub TancaExcel(MsExcel As Object, Lobro As Object)
On Error Resume Next
  Set Lobro = Nothing
'MsExcel.Visible = True
DoEvents
MsExcel.Quit
DoEvents
  Set MsExcel = Nothing

End Sub


Sub MergeCelda(ByRef Hoja As Excel.Worksheet, F1, C1, F2, C2, Optional Str)
   Dim rng
   Hoja.Activate
   Hoja.Range(Hoja.Cells(F1, C1), Hoja.Cells(F2, C2)).MergeCells = True
   Hoja.Cells(F1, C1).Value = Str
End Sub

