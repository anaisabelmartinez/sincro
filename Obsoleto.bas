Attribute VB_Name = "Obsoleto"

Function XX_SincroDbVendesAmetller(p1, P2, P3, P4, idTasca) As Boolean
Dim botiguesCad As String, mes As String, anyo As String, desde As String, debugSincro As Boolean
Dim codiBotiga As String, codiBotigaextern As String, tabla As String, fecha, fecha_caracter
Dim sql As String, sql2 As String, sql3 As String, sqlSP As String, numCab As Integer, maxIdTicket As String
Dim parametros As String, tablaTmp As String, Rs As rdoResultset, rsClients As rdoResultset, pos As Integer
Dim rsCab As rdoResultset, rsLin As rdoResultset, html As String, hasta As String, desde0
Dim connMysql As ADODB.Connection

'FUNCIO ANTIGA!!! 09/08/2012

On Error GoTo norVendes
'Parametros
'Si pasamos parametros vacios se recorreran todas las tiendas de paramsHw
',mirando ventas (fechas actuales), se eliminaran datos en servidor remoto
'si se le indica una fecha desde y se ejecutaran inserciones (cabeceras y lineas)

'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
'connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=mysql.casaametller.net;Port=3307;Database=sys_datos;User=hituser; Password=aM3fP6x8;Option=3;"
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
botiguesCad = p1 'Cadena de tiendas separadas por ,
If botiguesCad <> "" Then
    botiguesCad = Replace(botiguesCad, "[", "")
    botiguesCad = Replace(botiguesCad, "]", "")
End If
'desde = P4 'Fecha desde la que eliminan datos en tablas ventas de serv. remoto
'mes = P2
'anyo = P3
desde = P2 'Fecha desde la que eliminan datos en tablas ventas de serv. remoto
If desde <> "" Then
    desde = Replace(desde, "[", "")
    desde = Replace(desde, "]", "")
End If
desde0 = desde
mes = Month(desde) 'Mes para tabla ventas
If mes <> "" Then If Len(mes) = 1 Then mes = "0" & mes
anyo = Year(desde) 'Año para tabla ventas
Desti = P3 'Email
'desde = CDate(desde)
hasta = P4 'Fecha desde la que eliminan datos en tablas ventas de serv. remoto
If hasta <> "" Then
    hasta = Replace(hasta, "[", "")
    hasta = Replace(hasta, "]", "")
End If
debugSincro = False 'Solo impresion o ejecucion completa
'Crear txt log
If debugSincro = True Then
    Desti = EmailGuardia
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\Vendes_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
End If
'--------------------------------------------------------------
numCab = 0
If mes = "" Then
    mes = Month(Now)
    If Len(mes) = 1 Then mes = "0" & mes
End If
If anyo = "" Then anyo = Year(Now)
tabla = "[Fac_laforneria].[dbo].[V_Venut_" & anyo & "-" & mes & "] " 'Tabla ventas
InformaMiss "INICIO SINCRO_VENDES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "INICIO SINCRO_VENDES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'--TABLAS TEMPORALES
'--------------------------------------------------------------------------------
'Creamos tablas temporales de las quales podemos obtener datos de familia, secciones, etc
InformaMiss "CREANDO TABLAS TEMPORALES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CREANDO TABLAS TEMPORALES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
nTmp = Now
tablaTmp = "[sincro_vendesTmpVb_" & botiguesCad & "_" & nTmp & "]"
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmp la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp)
sql = " SELECT * INTO " & tablaTmp & " FROM OPENQUERY(AMETLLER,'"
sql = sql & "SELECT a.idArticulo,a.Descripcion,a.Descripcion1,a.idFamilia,f.NombreFamilia,a.idSubFamilia,"
sql = sql & "sf.NombreSubFamilia,a.idDepartamento,d.NombreDepartamento,a.idSeccion,s.NombreSeccion,"
sql = sql & "a.PrecioSinIVA , a.PrecioConIva, i.IdIVA, i.PorcentajeIva "
sql = sql & "FROM dat_articulo a "
sql = sql & "LEFT JOIN dat_familia f on (a.idFamilia=f.idFamilia) "
sql = sql & "LEFT JOIN dat_subfamilia sf on a.idFamilia=sf.idFamilia and a.idsubfamilia = sf.idsubfamilia "
sql = sql & "LEFT JOIN dat_departamento d on (a.idDepartamento=d.idDepartamento) "
sql = sql & "LEFT JOIN dat_seccion s on (a.idSeccion=s.idSeccion) "
sql = sql & "LEFT JOIN dat_iva i on (a.idIva=i.idIva) "
sql = sql & "Where a.idEmpresa = 1 And IdArticulo < 90000') "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
html = "<p><h3>Resum Vendes Ametller</h3></p>"
'--------------------------------------------------------------------------------
'--CURSOR TIENDAS
'--------------------------------------------------------------------------------
'--Creamos cursor para tiendas
InformaMiss "CREANDO CURSOR TIENDAS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CREANDO CURSOR TIENDAS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
sql = "SELECT p.valor1 FROM Clients c LEFT JOIN ParamsHw p ON (c.codi=p.valor1) WHERE p.valor1 is not null "
If botiguesCad <> "" Then sql = sql & "and p.valor1 IN ( " & botiguesCad & ")  "
sql = sql & " order by p.valor1 "
Set rsClients = Db.OpenResultset(sql)
Do While Not rsClients.EOF
    codiBotiga = rsClients("valor1")
    If codiBotiga = "518" Then
        codiBotigaextern = 1061
    Else
        codiBotigaextern = codiBotiga
    End If
    pos = Len(codiBotigaextern)
    If pos > 3 Then
        pos = 3
    Else
        pos = 2
    End If
    InformaMiss "TRATANDO TIENDA " & codiBotiga & "(" & codiBotigaextern & "),M:" & mes & ",A:" & anyo & "-->" & Now
    Informa2 "T:" & codiBotiga & "(" & codiBotigaextern & "),DESDE:" & desde
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "TRATANDO TIENDA " & codiBotiga & "(" & codiBotigaextern & "),M:" & mes & ",A:" & anyo & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'--------------------------------------------------------------------------------
'--ACTUALIZACION DE DUPLICADOS. TIQUETS CON MAS DE UNA FECHA DIFERENTE QUE AL AGRUPARSE
'--POR NUMERO DE TIQUET Y FECHA, GENERA DOS LINEAS PARA ESE MISMO NUMERO DE TIQUET
'--Y A LA HORA DE IMPORTAR A MYSQL DA ERRORES.
'--------------------------------------------------------------------------------
'--Si existe sincro_duplicadosTmp la borramos y volvemos a generar
    'Sql = "SELECT object_id FROM sys.objects with (nolock) WHERE name='sincro_duplicadosTmp' AND type='U' "
    'Set Rs = Db.OpenResultset(Sql)
    'If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE sincro_duplicadosTmp")
'--Si existe sincro_duplicadosTmp2 la borramos y volvemos a generar
    'Sql = "SELECT object_id FROM sys.objects with (nolock) WHERE name='sincro_duplicadosTmp2' AND type='U' "
    'Set Rs = Db.OpenResultset(Sql)
    'If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE sincro_duplicadosTmp2")
    'Sql = "SELECT data,Num_tick,botiga INTO sincro_duplicadosTmp FROM " & tabla & " GROUP BY data,Num_tick,Botiga "
    'Set Rs = Db.OpenResultset(Sql)
    'Sql = "SELECT Num_tick,Botiga INTO sincro_duplicadosTmp2 FROM sincro_duplicadosTmp GROUP BY Num_tick,Botiga "
    'Sql = Sql & "Having Count(Num_tick) >= 2 "
    'Set Rs = Db.OpenResultset(Sql)
    'Sql = "UPDATE " & tabla & " SET Data=v.data FROM ( "
    'Sql = Sql & "SELECT MIN(data) data,Num_tick,botiga FROM " & tabla & " "
    'Sql = Sql & "WHERE Num_tick IN(SELECT Num_tick FROM sincro_duplicadosTmp2 s WHERE " & tabla & ".Botiga=s.botiga) "
    'Sql = Sql & "AND botiga IN (SELECT botiga FROM sincro_duplicadosTmp2 s WHERE " & tabla & ".Botiga=s.botiga) "
    'Sql = Sql & "GROUP BY Num_tick,botiga "
    'Sql = Sql & ") v "
    'Sql = Sql & "Where " & tabla & ".Num_tick=v.Num_tick "
    'Sql = Sql & "AND " & tabla & ".Botiga=v.Botiga "
    'Set Rs = Db.OpenResultset(Sql)
    'If debugSincro = True Then
    '    Txt.WriteLine "--------------------------------------------------------------------------------"
    '    Txt.WriteLine "SQL:" & Sql & "-->" & Now
    '    Txt.WriteLine "--------------------------------------------------------------------------------"
    'End If
'--------------------------------------------------------------------------------
'--BORRADO DE DATOS EN SERV.REMOTO
'--CUIDADO! NO ELIMINA LAS LINEAS QUE SE HAYAN PODIDO QUEDAR SIN CABECERA!!!
'--------------------------------------------------------------------------------
    InformaMiss "BORRADO DE DATOS EN SERV.REMOTO DESDE " & desde & "-->" & Now
    If desde <> "" Then
        desde = Year(desde) & "-" & Month(desde) & "-" & Day(desde) & " " & Hour(desde) & ":" & Minute(desde) & ":" & Second(desde) & ".000"
'----------------------------------------------------------------------------------------------------
'---- LINEAS DE TICKET
'----------------------------------------------------------------------------------------------------
        sqlSP = "delete FROM dat_ticket_linea "
        sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
        sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
        sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' and idTicket "
        sqlSP = sqlSP & "in (select idTicket from dat_ticket_cabecera where idEmpresa=1 "
        sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' and idBalanzaEsclava=-1 "
        sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
        sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and timestamp>='" & desde & "' "
        'If hasta <> "" Then sqlSP = sqlSP & " and timestamp<='" & hasta & "' "
        sqlSP = sqlSP & ")"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sqlSP
        End If
'----------------------------------------------------------------------------------------------------
'---- CABECERAS DE TICKET
'----------------------------------------------------------------------------------------------------
        sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
        sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
        sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
        sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and timestamp>'" & desde & "' "
        'If hasta <> "" Then sqlSP = sqlSP & " and timestamp<='" & hasta & "' "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sqlSP
        End If
    End If
'--------------------------------------------------------------------------------
'--MAX IDTICKET
'--------------------------------------------------------------------------------
    InformaMiss "OBTENIENDO ULTIMO NUMERO TICKET-->" & Now
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "OBTENIENDO ULTIMO NUMERO TICKET-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'----------------------------------------------------------------------------------------------------
'---- SACAMOS EL IDTICKET MAXIMO Y LA FECHA MAXIMA DE LA TIENDA
'----------------------------------------------------------------------------------------------------
    sql = "select isNull(IdTicket,0) idTicket,isNull(fecha,'" & anyo & "-" & mes & "-01 00:00:01.000') fecha FROM openquery(AMETLLER, "
    sql = sql & "'SELECT MAX(idTicket)as IdTicket,MAX(TimeStamp)as fecha From dat_ticket_cabecera "
    sql = sql & " WHERE IdEmpresa=1 AND IdTienda=''" & Left(codiBotigaextern, pos) & "'' "
    sql = sql & "AND IdBalanzaMaestra=''" & Right(codiBotigaextern, 1) & "'' "
    sql = sql & "GROUP BY IdEmpresa,IdTienda,IdBalanzaMaestra') "
    Set Rs = Db.OpenResultset(sql)
    maxIdTicket = 0
    fecha = desde0
    fecha_caracter = desde0
    If Not Rs.EOF Then
        maxIdTicket = Rs("idTicket")
        fecha = Rs("fecha")
        'LA SQL NO FUNCIONA CON ESTE TIPO FECHA! fecha_caracter = Left(fecha, 4) & "-" & Mid(fecha, 6, 2) & "-" & Mid(fecha, 9, 2) & " " & Mid(fecha, 12, 2) & ":" & Mid(fecha, 15, 2) & ":" & Mid(fecha, 18, 2) & ".000"
        fecha_caracter = Mid(fecha, 9, 2) & "-" & Mid(fecha, 6, 2) & "-" & Left(fecha, 4) & " " & Mid(fecha, 12, 2) & ":" & Mid(fecha, 15, 2) & ":" & Mid(fecha, 18, 2) & ".000"
    End If
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'--------------------------------------------------------------------------------
'--LINEAS Y CABECERAS
'--------------------------------------------------------------------------------
    InformaMiss "INSERTANDO LINEAS Y CABECERAS-->" & Now
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "INSERTANDO LINEAS Y CABECERAS-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
    ImporteTotal = 0
    ImporteTotal2 = 0
    numLineas = 0
    'SQL ventas SQL SERVER
    sql = "SELECT ROW_NUMBER() OVER (PARTITION BY ven.num_tick ORDER BY ven.num_tick, ven.Plu,ven.Quantitat,ven.Import,ven.Data) AS IdLineaTicket, "
    sql = sql & "1 AS IdEmpresa,LEFT('" & codiBotigaextern & "'," & pos & ") AS IdTienda,RIGHT('" & codiBotigaextern & "',1) AS IdBalanzaMaestra, "
    sql = sql & "-1 AS IdBalanzaEsclava,DENSE_RANK() OVER (ORDER BY ven.num_tick) + '" & maxIdTicket & "'  AS IdTicket,2 AS TipoVenta,0 AS EstadoLinea, "
    sql = sql & "isNull(tmp.idArticulo,'89992') AS IdArticulo,isNull(tmp.Descripcion,'Preu Directe Hit') AS Descripcion, "
    sql = sql & "isNull(tmp.Descripcion1,'Preu Directe Hit') AS Descripcion1,isNull(art.EsSumable,1) AS Comportamiento,0 AS Tara, "
    sql = sql & "ven.Quantitat AS Peso,NULL AS PesoRegalado,round(ven.import/ven.quantitat,2) AS Precio,0 AS PrecioSinIVA, "
    sql = sql & "round(ven.import/ven.quantitat,2) AS PrecioConIVASinDtoL,isNull(tmp.idIva,0) AS IdIVA, "
    sql = sql & "isNull(tmp.PorcentajeIVA,0) AS PorcentajeIVA,NULL AS Descuento, "
    sql = sql & "NULL AS ImporteSinIVASinDtoL,NULL AS ImporteConIVASinDtoL,NULL AS ImporteDelDescuento, "
    sql = sql & "isNull(ROUND(ven.Import,2),0) AS Importe, "
    sql = sql & "isNull(ROUND(CAST(ven.Import AS FLOAT)-((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100),2),0) AS ImporteSinIVAConDtoL, "
    sql = sql & "isNull(ROUND((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100,2),0) AS ImporteDelIVAConDtoL, "
    sql = sql & "isNull(ROUND(CAST(ven.Import AS FLOAT)-((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100),2),0) AS ImporteSinIVAConDtoLConDtoTotal, "
    sql = sql & "isNull(ROUND((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100,2),0) AS ImporteDelIVAConDtoLConDtoTotal, "
    sql = sql & "isNull(ROUND(ven.Import,2),0) AS ImporteConDtoTotal,0 AS TaraFija,0 AS TaraPorcentual,isNull(tmp.idArticulo,'89992') AS CodInterno, "
    sql = sql & "'' as EANScannerArticulo,0 AS IdClase,NULL AS NombreClase,0 AS IdElemAsociado, "
    sql = sql & "isNull(tmp.IdFamilia,0) AS IdFamilia,isNull(tmp.NombreFamilia,'Familia no asignada') AS NombreFamilia, "
    sql = sql & "isNull(tmp.IdSeccion,0) AS IdSeccion, "
    sql = sql & "isNull(tmp.NombreSeccion,'Seccio no asignada') AS NombreSeccion,isNull(tmp.IdSubFamilia,0) AS IdSubFamilia, "
    sql = sql & "isNull(tmp.NombreSubFamilia,'Subfamilia no asignada') AS NombreSubFamilia,isNull(tmp.IdDepartamento,0) AS IdDepartamento, "
    sql = sql & "isNull(tmp.NombreDepartamento,'Departament no asignat') AS NombreDepartamento,1 AS Modificado,'A' AS Operacion, "
    sql = sql & "'Comunicaciones' AS Usuario,ven.Data AS TimeStamp, "
    sql = sql & "'Casa Ametller S.L.' AS NombreEmpresa,CASE WHEN CHARINDEX('_',cli.nom)>0 THEN SUBSTRING(cli.nom,1,CAST(CHARINDEX('_',cli.nom)AS INTEGER)-1)  "
    sql = sql & "ELSE cli.Nom END AS NombreTienda,'T' AS Tipo,isNull(dE.valor,'0000') AS IdVendedor,isNull(dep.nom,'Dependenta sense asignar') AS NombreVendedor, "
    sql = sql & " 0 AS IdCliente,NULL AS NombreCliente, "
    sql = sql & "NULL AS DNICliente,NULL AS DireccionCliente,NULL AS PoblacionCliente, "
    sql = sql & "NULL AS ProvinciaCliente,NULL AS CPCliente,0 AS TelefonoCliente,NULL AS ImporteLineas,NULL AS PorcDescuento,NULL AS ImporteDescuento, "
    sql = sql & "NULL AS ImporteLineas2,NULL AS PorcDescuento2,NULL AS ImporteDescuento2, "
    sql = sql & "NULL AS ImporteLineas3,NULL AS PorcDescuento3,NULL AS ImporteDescuento3, "
    sql = sql & "NULL AS ImporteTotal3,NULL AS ImporteSinRedondeo,NULL AS ImporteDelRedondeo, "
    sql = sql & "NULL AS SerieLIdFinDeDia,0 AS SerieLTicketErroneo,0 AS ImporteDtoTotalSinIVA, "
    sql = sql & "ven.Data AS Fecha,'C' AS EstadoTicket,0 AS ImporteDevuelto,0 AS PuntosFidelidad,ven.num_tick AS NumTicket "
    sql = sql & "FROM " & tabla & " ven "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[clients] cli ON (ven.Botiga=cli.codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[dependentes] dep ON (ven.dependenta=dep.codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[dependentesExtes] dE on (dep.CODI=dE.id and dE.nom='CODI_DEP') "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[Articles] art ON (ven.Plu=art.Codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[ArticlesPropietats] artP1 ON (art.Codi=artP1.CodiArticle and artP1.Variable='CODI_PROD') "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[TipusIva] iva ON (art.TipoIva=iva.Tipus) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo]." & tablaTmp & " tmp ON (artP1.valor=tmp.IdArticulo) "
    sql = sql & "WHERE ven.Botiga='" & codiBotiga & "' and ven.data>'" & fecha_caracter & "' "
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
    Set rsLin = Db.OpenResultset(sql)
    If Not rsLin.EOF Then
        idTicketAnt = rsLin("idTicket")
    End If
    Do While Not rsLin.EOF
'--------------------------------------------------------------------------------
'--INSERTANDO CABECERA
'--------------------------------------------------------------------------------
        If rsLin("IdTicket") <> idTicketAnt Then
            'Importe de ivas de los 3 tipos
            sql = "SELECT ("
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=1 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=1))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=2 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=2))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=3 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=3))/100,0),2) "
            sql = sql & ") as resultado "
            Set Rs = Db.OpenResultset(sql)
            If Not Rs.EOF Then
                ImporteTotalSinIVAConDtoL = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalSinIVAConDtoLConDtoTotal = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalDelIVAConDtoLConDtoTotal = Rs("resultado")
            End If
            ImporteEntregado = ImporteTotal
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR CABECERA!!! BORRADO DE CABECERA ACTUAL POR SI YA EXISTIERA
'----------------------------------------------------------------------------------------------------
            sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
            sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
            sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and idTicket='" & idTicketAnt & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
            'SQL INSERT Cabecera MYSQL
            'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_cabecera limit 1'') "
            sql = "INSERT INTO dat_ticket_cabecera "
            sql = sql & "(IdTicket, NumTicket, IdEmpresa, NombreEmpresa,IdTienda,NombreTienda,IdBalanzaMaestra,"
            sql = sql & "NombreBalanzaMaestra, IdBalanzaEsclava, Tipo,"
            sql = sql & "IdVendedor,NombreVendedor,IdCliente,NombreCliente,"
            sql = sql & "DNICliente,DireccionCliente,PoblacionCliente,"
            sql = sql & "ProvinciaCliente,CPCliente,TelefonoCliente,"
            sql = sql & "TipoVenta,ImporteLineas,PorcDescuento,ImporteDescuento,"
            sql = sql & "ImporteTotal,ImporteLineas2,PorcDescuento2,ImporteDescuento2,"
            sql = sql & "ImporteTotal2,ImporteLineas3,PorcDescuento3,ImporteDescuento3,"
            sql = sql & "ImporteTotal3,ImporteTotalSinIVAConDtoL,ImporteTotalSinIVAConDtoLConDtoTotal,"
            sql = sql & "ImporteDtoTotalSinIVA,ImporteTotalDelIVAConDtoLConDtoTotal,ImporteSinRedondeo,"
            sql = sql & "ImporteDelRedondeo,Fecha,NumLineas,SerieLIdFinDeDia,SerieLTicketErroneo,Modificado,"
            sql = sql & "Operacion,Usuario,TimeStamp,EstadoTicket,ImporteEntregado,ImporteDevuelto,PuntosFidelidad) values ("
            sql = sql & "'" & idTicketAnt & "','" & NumTicket & "','" & idEmpresa & "','" & NombreEmpresa & "',"
            sql = sql & "'" & IdTienda & "','" & NombreTienda & "','" & idBalanzaMaestra & "','-Balança " & idBalanzaMaestra & "',"
            sql = sql & "'" & idBalanzaEsclava & "','" & tipo & "','" & IdVendedor & "','" & NombreVendedor & "',"
            sql = sql & "'" & idCliente & "','" & NombreCliente & "','" & DNICliente & "','" & DireccionCliente & "',"
            sql = sql & "'" & PoblacionCliente & "','" & ProvinciaCliente & "','" & CPCliente & "','" & TelefonoCliente & "',"
            sql = sql & "'" & tipoVenta & "','" & ImporteLineas & "','" & PorcDescuento & "','" & ImporteDescuento & "',"
            sql = sql & "'" & ImporteTotal & "','" & ImporteLineas2 & "','" & PorcDescuento2 & "','" & ImporteDescuento2 & "',"
            sql = sql & "'" & ImporteTotal2 & "','" & ImporteLineas3 & "','" & PorcDescuento3 & "','" & ImporteDescuento3 & "',"
            sql = sql & "'" & ImporteTotal3 & "','" & ImporteTotalSinIVAConDtoL & "','" & ImporteTotalSinIVAConDtoLConDtoTotal & "','" & ImporteDtoTotalSinIVA & "',"
            sql = sql & "'" & ImporteTotalDelIVAConDtoLConDtoTotal & "','" & ImporteSinRedondeo & "','" & ImporteDelRedondeo & "','" & Format(fecha, "yyyy-mm-dd hh:nn:ss") & ".000',"
            sql = sql & "'" & numLineas & "','" & SerieLIdFinDeDia & "','" & SerieLTicketErroneo & "','" & Modificado & "',"
            sql = sql & "'" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000','" & EstadoTicket & "',"
            sql = sql & "'" & ImporteEntregado & "','" & ImporteDevuelto & "','" & PuntosFidelidad & "')"
            InformaMiss "CABECERA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "CABECERA!"
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "INSERTANDO CABECERA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sql
            End If
            ImporteTotal = 0
            ImporteTotal2 = 0
            numLineas = 0
            idTicketAnt = rsLin("idTicket")
        End If
        'Campos lineas
        import = rsLin("Importe")
        IdLineaTicket = rsLin("IdLineaTicket")
        idEmpresa = rsLin("idEmpresa")
        IdTienda = rsLin("IdTienda")
        idBalanzaMaestra = rsLin("IdBalanzaMaestra")
        idBalanzaEsclava = rsLin("IdBalanzaEsclava")
        idTicket = rsLin("IdTicket")
        tipoVenta = rsLin("TipoVenta")
        EstadoLinea = rsLin("EstadoLinea")
        IdArticulo = rsLin("IdArticulo")
        DESCRIPCIoN = rsLin("Descripcion")
        Descripcion1 = rsLin("Descripcion1")
        Comportamiento = rsLin("Comportamiento")
        Tara = rsLin("Tara")
        Peso = rsLin("Peso")
        PesoRegalado = rsLin("PesoRegalado")
        precio = rsLin("Precio")
        PrecioSinIva = rsLin("PrecioSinIva")
        PrecioConIVASinDtoL = rsLin("PrecioConIVASinDtoL")
        IdIVA = rsLin("IdIVA")
        PorcentajeIva = rsLin("PorcentajeIva")
        descuento = rsLin("descuento")
        ImporteSinIVASinDtoL = rsLin("ImporteSinIVASinDtoL")
        ImporteConIVASinDtoL = rsLin("ImporteConIVASinDtoL")
        ImporteDelDescuento = rsLin("ImporteDelDescuento")
        importe = rsLin("Importe")
        ImporteSinIVAConDtoL = rsLin("ImporteSinIVAConDtoL")
        ImporteDelIVAConDtoL = rsLin("ImporteDelIVAConDtoL")
        ImporteSinIVAConDtoLConDtoTotal = rsLin("ImporteSinIVAConDtoLConDtoTotal")
        ImporteDelIVAConDtoLConDtoTotal = rsLin("ImporteDelIVAConDtoLConDtoTotal")
        ImporteConDtoTotal = rsLin("ImporteConDtoTotal")
        TaraFija = rsLin("TaraFija")
        TaraPorcentual = rsLin("TaraPorcentual")
        CodInterno = rsLin("CodInterno")
        EANScannerArticulo = rsLin("EANScannerArticulo")
        IdClase = rsLin("IdClase")
        NombreClase = rsLin("NombreClase")
        IdElemAsociado = rsLin("IdElemAsociado")
        IdFamilia = rsLin("IdFamilia")
        NombreFamilia = rsLin("NombreFamilia")
        IdSeccion = rsLin("IdSeccion")
        nombreSeccion = rsLin("NombreSeccion")
        IdSubFamilia = rsLin("IdSubFamilia")
        NombreSubFamilia = rsLin("NombreSubFamilia")
        IdDepartamento = rsLin("IdDepartamento")
        NombreDepartamento = rsLin("NombreDepartamento")
        Modificado = rsLin("Modificado")
        Operacion = rsLin("Operacion")
        usuario = rsLin("Usuario")
        TimeStamp = rsLin("TimeStamp")
        'Campos cabecera
        NombreEmpresa = rsLin("NombreEmpresa")
        NombreTienda = rsLin("NombreTienda")
        tipo = rsLin("tipo")
        IdVendedor = rsLin("IdVendedor")
        NombreVendedor = rsLin("NombreVendedor")
        idCliente = rsLin("IdCliente")
        NombreCliente = rsLin("NombreCliente")
        DNICliente = rsLin("DNICliente")
        DireccionCliente = rsLin("DireccionCliente")
        PoblacionCliente = rsLin("PoblacionCliente")
        ProvinciaCliente = rsLin("ProvinciaCliente")
        CPCliente = rsLin("CPCliente")
        TelefonoCliente = rsLin("TelefonoCliente")
        ImporteLineas = rsLin("ImporteLineas")
        PorcDescuento = rsLin("PorcDescuento")
        ImporteDescuento = rsLin("ImporteDescuento")
        ImporteLineas2 = rsLin("ImporteLineas2")
        PorcDescuento2 = rsLin("PorcDescuento2")
        ImporteDescuento2 = rsLin("ImporteDescuento2")
        ImporteLineas3 = rsLin("ImporteLineas3")
        PorcDescuento3 = rsLin("PorcDescuento3")
        ImporteDescuento3 = rsLin("ImporteDescuento3")
        ImporteTotal3 = rsLin("ImporteTotal3")
        ImporteSinRedondeo = rsLin("ImporteSinRedondeo")
        ImporteDelRedondeo = rsLin("ImporteDelRedondeo")
        SerieLIdFinDeDia = rsLin("SerieLIdFinDeDia")
        SerieLTicketErroneo = rsLin("SerieLTicketErroneo")
        ImporteDtoTotalSinIVA = rsLin("ImporteDtoTotalSinIVA")
        fecha = rsLin("fecha")
        EstadoTicket = rsLin("EstadoTicket")
        ImporteDevuelto = rsLin("ImporteDevuelto")
        PuntosFidelidad = rsLin("PuntosFidelidad")
        NumTicket = rsLin("numTicket")
        'ImporteTotal en ImporteTotal y ImporteTotal2
        ImporteTotal = ImporteTotal + Round(importe, 2)
        ImporteTotal2 = ImporteTotal2 + Round(importe, 2)
        numLineas = numLineas + 1
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR LINEA!!! BORRADO DE LINEAS DE TICKET PARA EL NUMERO DE TICKET ACTUAL
'----  ES POSIBLE QUE ALGUNA VEZ SE HAYA QUEDADO LA LINEA SIN CABECERA
'----------------------------------------------------------------------------------------------------
        If IdLineaTicket = 1 Then
            sqlSP = "delete FROM dat_ticket_linea "
            sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
            sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
            sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and idTicket='" & idTicket & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
        End If
'--------------------------------------------------------------------------------
'--INSERTANDO LINEA
'--------------------------------------------------------------------------------
        'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_linea limit 1'') ( "
        sql = "INSERT INTO dat_ticket_linea ( "
        sql = sql & "IdLineaTicket,IdEmpresa,IdTienda,IdBalanzaMaestra,IdBalanzaEsclava,IdTicket,"
        sql = sql & "TipoVenta,EstadoLinea,IdArticulo,Descripcion,Descripcion1,Comportamiento,Tara,"
        sql = sql & "Peso,PesoRegalado,Precio,PrecioSinIVA,PrecioConIVASinDtoL,IdIVA,PorcentajeIVA,"
        sql = sql & "Descuento,ImporteSinIVASinDtoL,ImporteConIVASinDtoL,ImporteDelDescuento,Importe,"
        sql = sql & "ImporteSinIVAConDtoL,ImporteDelIVAConDtoL,ImporteSinIVAConDtoLConDtoTotal,"
        sql = sql & "ImporteDelIVAConDtoLConDtoTotal,ImporteConDtoTotal,TaraFija,TaraPorcentual,"
        sql = sql & "CodInterno,EANScannerArticulo,IdClase,NombreClase,IdElemAsociado,IdFamilia,"
        sql = sql & "NombreFamilia,IdSeccion,NombreSeccion,IdSubFamilia,NombreSubFamilia,IdDepartamento,"
        sql = sql & "NombreDepartamento,Modificado,Operacion,Usuario,TimeStamp) "
        sql = sql & " values ('" & IdLineaTicket & "','" & idEmpresa & "','" & IdTienda & "','"
        sql = sql & idBalanzaMaestra & "','" & idBalanzaEsclava & "','" & idTicket & "','"
        sql = sql & tipoVenta & "','" & EstadoLinea & "','" & IdArticulo & "','" & Replace(DESCRIPCIoN, "'", "''") & "','" & Replace(Descripcion1, "'", "''") & "','" & CInt(Comportamiento) & "','" & Tara & "','"
        sql = sql & Peso & "','" & PesoRegalado & "','" & precio & "','" & PrecioSinIva & "','"
        sql = sql & PrecioConIVASinDtoL & "','" & IdIVA & "','" & PorcentajeIva & "','" & descuento & "','" & ImporteSinIVASinDtoL & "','" & ImporteConIVASinDtoL & "','"
        sql = sql & ImporteDelDescuento & "',' " & importe & "',' " & ImporteSinIVAConDtoL & "','"
        sql = sql & ImporteDelIVAConDtoL & "','" & ImporteSinIVAConDtoLConDtoTotal & "','"
        sql = sql & ImporteDelIVAConDtoLConDtoTotal & "','" & ImporteConDtoTotal & "','" & TaraFija & "','" & TaraPorcentual & "','"
        sql = sql & CodInterno & "','" & EANScannerArticulo & "','" & IdClase & "','" & NombreClase & "','" & IdElemAsociado & "','" & IdFamilia & "','"
        sql = sql & NombreFamilia & "','" & IdSeccion & "','" & nombreSeccion & "','" & IdSubFamilia & "','" & NombreSubFamilia & "','" & IdDepartamento & "','"
        sql = sql & NombreDepartamento & "','" & Modificado & "','" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000') "
        InformaMiss "LINEA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "INSERTANDO LINEA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sql
        End If
        rsLin.MoveNext
'----------------------------------------------------------------------------------------------------
'----  ULTIMA CABECERA
'----------------------------------------------------------------------------------------------------
        If rsLin.EOF Then
            'Importe de ivas de los 3 tipos
            sql = "SELECT ("
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=1 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=1))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=2 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=2))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=3 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=3))/100,0),2) "
            sql = sql & ") as resultado "
            Set Rs = Db.OpenResultset(sql)
            If Not Rs.EOF Then
                ImporteTotalSinIVAConDtoL = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalSinIVAConDtoLConDtoTotal = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalDelIVAConDtoLConDtoTotal = Rs("resultado")
            End If
            ImporteEntregado = ImporteTotal
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR CABECERA!!! BORRADO DE CABECERA ACTUAL POR SI YA EXISTIERA
'----------------------------------------------------------------------------------------------------
            sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
            sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
            sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and idTicket='" & idTicket & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
            'SQL INSERT Cabecera MYSQL
            'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_cabecera limit 1'') "
            sql = "INSERT INTO dat_ticket_cabecera "
            sql = sql & "(IdTicket, NumTicket, IdEmpresa, NombreEmpresa,IdTienda,NombreTienda,IdBalanzaMaestra,"
            sql = sql & "NombreBalanzaMaestra, IdBalanzaEsclava, Tipo,"
            sql = sql & "IdVendedor,NombreVendedor,IdCliente,NombreCliente,"
            sql = sql & "DNICliente,DireccionCliente,PoblacionCliente,"
            sql = sql & "ProvinciaCliente,CPCliente,TelefonoCliente,"
            sql = sql & "TipoVenta,ImporteLineas,PorcDescuento,ImporteDescuento,"
            sql = sql & "ImporteTotal,ImporteLineas2,PorcDescuento2,ImporteDescuento2,"
            sql = sql & "ImporteTotal2,ImporteLineas3,PorcDescuento3,ImporteDescuento3,"
            sql = sql & "ImporteTotal3,ImporteTotalSinIVAConDtoL,ImporteTotalSinIVAConDtoLConDtoTotal,"
            sql = sql & "ImporteDtoTotalSinIVA,ImporteTotalDelIVAConDtoLConDtoTotal,ImporteSinRedondeo,"
            sql = sql & "ImporteDelRedondeo,Fecha,NumLineas,SerieLIdFinDeDia,SerieLTicketErroneo,Modificado,"
            sql = sql & "Operacion,Usuario,TimeStamp,EstadoTicket,ImporteEntregado,ImporteDevuelto,PuntosFidelidad) values ("
            sql = sql & "'" & idTicket & "','" & NumTicket & "','" & idEmpresa & "','" & NombreEmpresa & "',"
            sql = sql & "'" & IdTienda & "','" & NombreTienda & "','" & idBalanzaMaestra & "','-Balança " & idBalanzaMaestra & "',"
            sql = sql & "'" & idBalanzaEsclava & "','" & tipo & "','" & IdVendedor & "','" & NombreVendedor & "',"
            sql = sql & "'" & idCliente & "','" & NombreCliente & "','" & DNICliente & "','" & DireccionCliente & "',"
            sql = sql & "'" & PoblacionCliente & "','" & ProvinciaCliente & "','" & CPCliente & "','" & TelefonoCliente & "',"
            sql = sql & "'" & tipoVenta & "','" & ImporteLineas & "','" & PorcDescuento & "','" & ImporteDescuento & "',"
            sql = sql & "'" & ImporteTotal & "','" & ImporteLineas2 & "','" & PorcDescuento2 & "','" & ImporteDescuento2 & "',"
            sql = sql & "'" & ImporteTotal2 & "','" & ImporteLineas3 & "','" & PorcDescuento3 & "','" & ImporteDescuento3 & "',"
            sql = sql & "'" & ImporteTotal3 & "','" & ImporteTotalSinIVAConDtoL & "','" & ImporteTotalSinIVAConDtoLConDtoTotal & "','" & ImporteDtoTotalSinIVA & "',"
            sql = sql & "'" & ImporteTotalDelIVAConDtoLConDtoTotal & "','" & ImporteSinRedondeo & "','" & ImporteDelRedondeo & "','" & Format(fecha, "yyyy-mm-dd hh:nn:ss") & ".000',"
            sql = sql & "'" & numLineas & "','" & SerieLIdFinDeDia & "','" & SerieLTicketErroneo & "','" & Modificado & "',"
            sql = sql & "'" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000','" & EstadoTicket & "',"
            sql = sql & "'" & ImporteEntregado & "','" & ImporteDevuelto & "','" & PuntosFidelidad & "')"
            InformaMiss "CABECERA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "CABECERA!"
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "INSERTANDO CABECERA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sql
            End If
        End If
    Loop
'--------------------------------------------------------------------------------
'--SIGUIENTE TIENDA
'--------------------------------------------------------------------------------
    rsClients.MoveNext
    'Borramos tablas temp
    Db.OpenResultset ("DROP TABLE " & tablaTmp)
    html = html & "<p><b>Botiga: </b>" & codiBotiga & "</p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & desde & "</p>"
    html = html & "<p><b>Finalitzat: </b>" & Now() & "</p>"
Loop
If P3 <> "" Then sf_enviarMail "secrehit@hit.cat", Desti, " Vendes sincronitzades " & codiBotiga, html, "", ""

connMysql.Close
Set connMysql = Nothing

InformaMiss "FIN SINCRO_VENDES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "FIN SINCRO_VENDES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.Close
End If
Exit Function

norVendes:
'----------------------------------------------------------------------------------------------------
'----  SI HAY ERROR: BORRADO ULTIMAS LINEAS DE TICKET HUERFANAS SIN CABECERA
'----------------------------------------------------------------------------------------------------
    Db.OpenResultset ("DROP TABLE " & tablaTmp)
    html = "<p><h3>Resum Vendes Ametller </h3></p>"
    html = html & "<p><b>Botiga: </b>" & botiguesCad & "</p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & desde & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    If P3 <> "" Then sf_enviarMail "secrehit@hit.cat", Desti, "ERROR! Sincronitzacio de vendes de " & codiBotiga & " ha fallat", html, "", ""
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    html = html & "<p><b>Ultima sqlSP:</b>" & sqlSP & "</p>"
    
    sf_enviarMail "secrehit@hit.cat", EmailGuardia, "ERROR! Sincronitzacio de vendes de " & codiBotiga & " ha fallat", html, "", ""
        
    Set connMysql = New ADODB.Connection
    connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
    connMysql.ConnectionTimeout = 1000 '16 min
    connMysql.Open

    sqlSP = "delete FROM dat_ticket_linea "
    sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
    sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
    sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
    sqlSP = sqlSP & "and idTicket='" & idTicket & "' "
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "Error:" & err.Description & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        connMysql.Execute sqlSP
    End If
    
    connMysql.Close
    Set connMysql = Nothing
    
    ExecutaComandaSql "Delete from  " & TaulaCalculsEspecials & " Id = '" & idTasca & "' "
    
    Informa "Error : " & err.Description
    
End Function


Function XX_SincroDbVendesAmetller1(p1, P2, P3, P4, idTasca) As Boolean
Dim botiguesCad As String, mes As String, anyo As String, desde As String, debugSincro As Boolean
Dim codiBotiga As String, codiBotigaextern As String, tabla As String, fecha, fecha_caracter
Dim sql As String, sql2 As String, sql3 As String, sqlSP As String, numCab As Integer, maxIdTicket As String
Dim parametros As String, tablaTmp As String, Rs As rdoResultset, rsClients As rdoResultset, pos As Integer
Dim rsCab As rdoResultset, rsLin As rdoResultset, html As String, hasta As String, desde0
Dim connMysql As ADODB.Connection

'JUNTA ALBARANS AMB VENDES!


On Error GoTo norVendes
'Parametros
'Si pasamos parametros vacios se recorreran todas las tiendas de paramsHw
',mirando ventas (fechas actuales), se eliminaran datos en servidor remoto
'si se le indica una fecha desde y se ejecutaran inserciones (cabeceras y lineas)

'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
botiguesCad = p1 'Cadena de tiendas separadas por ,
If botiguesCad <> "" Then
    botiguesCad = Replace(botiguesCad, "[", "")
    botiguesCad = Replace(botiguesCad, "]", "")
End If
desde = P2 'Fecha desde la que eliminan datos en tablas ventas de serv. remoto
If desde <> "" Then
    desde = Replace(desde, "[", "")
    desde = Replace(desde, "]", "")
End If
desde0 = desde
mes = Month(desde) 'Mes para tabla ventas
If mes <> "" Then If Len(mes) = 1 Then mes = "0" & mes
anyo = Year(desde) 'Año para tabla ventas
Desti = P3 'Email
'desde = CDate(desde)
hasta = P4 'Fecha desde la que eliminan datos en tablas ventas de serv. remoto
If hasta <> "" Then
    hasta = Replace(hasta, "[", "")
    hasta = Replace(hasta, "]", "")
End If
debugSincro = False 'Solo impresion o ejecucion completa
'Crear txt log
If debugSincro = True Then
    Desti = EmailGuardia
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\Vendes_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
End If
'--------------------------------------------------------------
numCab = 0
If mes = "" Then
    mes = Month(Now)
    If Len(mes) = 1 Then mes = "0" & mes
End If
If anyo = "" Then anyo = Year(Now)
InformaMiss "INICIO SINCRO_VENDES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "INICIO SINCRO_VENDES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'--TABLAS TEMPORALES
'--------------------------------------------------------------------------------
'Creamos tablas temporales de las quales podemos obtener datos de familia, secciones, etc
InformaMiss "CREANDO TABLAS TEMPORALES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CREANDO TABLAS TEMPORALES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
nTmp = Now
tablaTmp = "[sincro_vendesTmpArticles_" & botiguesCad & "_" & nTmp & "]"
tablaTmp2 = "[sincro_vendesTmpClients_" & botiguesCad & "_" & nTmp & "]"
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmpArticles la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp)
sql = " SELECT * INTO " & tablaTmp & " FROM OPENQUERY(AMETLLER,'"
sql = sql & "SELECT a.idArticulo,a.Descripcion,a.Descripcion1,a.idFamilia,f.NombreFamilia,a.idSubFamilia,"
sql = sql & "sf.NombreSubFamilia,a.idDepartamento,d.NombreDepartamento,a.idSeccion,s.NombreSeccion,"
sql = sql & "a.PrecioSinIVA , a.PrecioConIva, i.IdIVA, i.PorcentajeIva "
sql = sql & "FROM dat_articulo a "
sql = sql & "LEFT JOIN dat_familia f on (a.idFamilia=f.idFamilia) "
sql = sql & "LEFT JOIN dat_subfamilia sf on a.idFamilia=sf.idFamilia and a.idsubfamilia = sf.idsubfamilia "
sql = sql & "LEFT JOIN dat_departamento d on (a.idDepartamento=d.idDepartamento) "
sql = sql & "LEFT JOIN dat_seccion s on (a.idSeccion=s.idSeccion) "
sql = sql & "LEFT JOIN dat_iva i on (a.idIva=i.idIva) "
sql = sql & "Where a.idEmpresa = 1 And IdArticulo < 90000') "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmpClients la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp2 & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp2)
sql = " SELECT * INTO " & tablaTmp2 & " FROM OPENQUERY(AMETLLER,'"
sql = sql & "SELECT idCliente, Nombre,TRIM(DNI) AS Cif, TRIM(Direccion) AS Dir,"
sql = sql & "TRIM(CodPostal) AS CP, TRIM(Poblacion) AS Ciutat,TimeStamp AS fecMod "
sql = sql & "FROM dat_cliente WHERE IdEmpresa=1') "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
html = "<p><h3>Resum Vendes Ametller</h3></p>"
'--------------------------------------------------------------------------------
'--CURSOR TIENDAS
'--------------------------------------------------------------------------------
'--Creamos cursor para tiendas
InformaMiss "CREANDO CURSOR TIENDAS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CREANDO CURSOR TIENDAS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
sql = "SELECT p.valor1 FROM Clients c LEFT JOIN ParamsHw p ON (c.codi=p.valor1) WHERE p.valor1 is not null "
If botiguesCad <> "" Then sql = sql & "and p.valor1 IN ( " & botiguesCad & ")  "
sql = sql & " order by p.valor1 "
Set rsClients = Db.OpenResultset(sql)
Do While Not rsClients.EOF
    codiBotiga = rsClients("valor1")
    If codiBotiga = "518" Then
        codiBotigaextern = 1061
    Else
        codiBotigaextern = codiBotiga
    End If
    pos = Len(codiBotigaextern)
    If pos > 3 Then
        pos = 3
    Else
        pos = 2
    End If
    InformaMiss "TRATANDO TIENDA " & codiBotiga & "(" & codiBotigaextern & "),M:" & mes & ",A:" & anyo & "-->" & Now
    Informa2 "T:" & codiBotiga & "(" & codiBotigaextern & "),DESDE:" & desde
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "TRATANDO TIENDA " & codiBotiga & "(" & codiBotigaextern & "),M:" & mes & ",A:" & anyo & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'--------------------------------------------------------------------------------
'--ACTUALIZACION DE DUPLICADOS. TIQUETS CON MAS DE UNA FECHA DIFERENTE QUE AL AGRUPARSE
'--POR NUMERO DE TIQUET Y FECHA, GENERA DOS LINEAS PARA ESE MISMO NUMERO DE TIQUET
'--Y A LA HORA DE IMPORTAR A MYSQL DA ERRORES.
'--------------------------------------------------------------------------------
'--Si existe sincro_duplicadosTmp la borramos y volvemos a generar
    'Sql = "SELECT object_id FROM sys.objects with (nolock) WHERE name='sincro_duplicadosTmp' AND type='U' "
    'Set Rs = Db.OpenResultset(Sql)
    'If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE sincro_duplicadosTmp")
'--Si existe sincro_duplicadosTmp2 la borramos y volvemos a generar
    'Sql = "SELECT object_id FROM sys.objects with (nolock) WHERE name='sincro_duplicadosTmp2' AND type='U' "
    'Set Rs = Db.OpenResultset(Sql)
    'If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE sincro_duplicadosTmp2")
    'Sql = "SELECT data,Num_tick,botiga INTO sincro_duplicadosTmp FROM " & tabla & " GROUP BY data,Num_tick,Botiga "
    'Set Rs = Db.OpenResultset(Sql)
    'Sql = "SELECT Num_tick,Botiga INTO sincro_duplicadosTmp2 FROM sincro_duplicadosTmp GROUP BY Num_tick,Botiga "
    'Sql = Sql & "Having Count(Num_tick) >= 2 "
    'Set Rs = Db.OpenResultset(Sql)
    'Sql = "UPDATE " & tabla & " SET Data=v.data FROM ( "
    'Sql = Sql & "SELECT MIN(data) data,Num_tick,botiga FROM " & tabla & " "
    'Sql = Sql & "WHERE Num_tick IN(SELECT Num_tick FROM sincro_duplicadosTmp2 s WHERE " & tabla & ".Botiga=s.botiga) "
    'Sql = Sql & "AND botiga IN (SELECT botiga FROM sincro_duplicadosTmp2 s WHERE " & tabla & ".Botiga=s.botiga) "
    'Sql = Sql & "GROUP BY Num_tick,botiga "
    'Sql = Sql & ") v "
    'Sql = Sql & "Where " & tabla & ".Num_tick=v.Num_tick "
    'Sql = Sql & "AND " & tabla & ".Botiga=v.Botiga "
    'Set Rs = Db.OpenResultset(Sql)
    'If debugSincro = True Then
    '    Txt.WriteLine "--------------------------------------------------------------------------------"
    '    Txt.WriteLine "SQL:" & Sql & "-->" & Now
    '    Txt.WriteLine "--------------------------------------------------------------------------------"
    'End If
'--------------------------------------------------------------------------------
'--BORRADO DE DATOS EN SERV.REMOTO
'--CUIDADO! NO ELIMINA LAS LINEAS QUE SE HAYAN PODIDO QUEDAR SIN CABECERA!!!
'--------------------------------------------------------------------------------
    InformaMiss "BORRADO DE DATOS EN SERV.REMOTO DESDE " & desde & "-->" & Now
    If desde <> "" Then
        desde = Year(desde) & "-" & Month(desde) & "-" & Day(desde) & " " & Hour(desde) & ":" & Minute(desde) & ":" & Second(desde) & ".000"
'----------------------------------------------------------------------------------------------------
'---- LINEAS DE TICKET
'----------------------------------------------------------------------------------------------------
        sqlSP = "delete FROM dat_ticket_linea "
        sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
        sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
        sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' and idTicket "
        sqlSP = sqlSP & "in (select idTicket from dat_ticket_cabecera where idEmpresa=1 "
        sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' and idBalanzaEsclava=-1 "
        sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
        sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and timestamp>='" & desde & "' "
        'If hasta <> "" Then sqlSP = sqlSP & " and timestamp<='" & hasta & "' "
        sqlSP = sqlSP & ")"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sqlSP
        End If
'----------------------------------------------------------------------------------------------------
'---- CABECERAS DE TICKET
'----------------------------------------------------------------------------------------------------
        sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
        sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
        sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
        sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and timestamp>'" & desde & "' "
        'If hasta <> "" Then sqlSP = sqlSP & " and timestamp<='" & hasta & "' "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sqlSP
        End If
    End If
'--------------------------------------------------------------------------------
'--MAX IDTICKET
'--------------------------------------------------------------------------------
    InformaMiss "OBTENIENDO ULTIMO NUMERO TICKET-->" & Now
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "OBTENIENDO ULTIMO NUMERO TICKET-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'----------------------------------------------------------------------------------------------------
'---- SACAMOS EL IDTICKET MAXIMO Y LA FECHA MAXIMA DE LA TIENDA
'----------------------------------------------------------------------------------------------------
    sql = "select isNull(IdTicket,0) idTicket,isNull(fecha,'" & anyo & "-" & mes & "-01 00:00:01.000') fecha FROM openquery(AMETLLER, "
    sql = sql & "'SELECT MAX(idTicket)as IdTicket,MAX(TimeStamp)as fecha From dat_ticket_cabecera "
    sql = sql & " WHERE IdEmpresa=1 AND IdTienda=''" & Left(codiBotigaextern, pos) & "'' "
    sql = sql & "AND IdBalanzaMaestra=''" & Right(codiBotigaextern, 1) & "'' "
    sql = sql & "GROUP BY IdEmpresa,IdTienda,IdBalanzaMaestra') "
    Set Rs = Db.OpenResultset(sql)
    maxIdTicket = 0
    fecha = desde0
    fecha_caracter = desde0
    If Not Rs.EOF Then
        maxIdTicket = Rs("idTicket")
        fecha = Rs("fecha")
        'LA SQL NO FUNCIONA CON ESTE TIPO FECHA! fecha_caracter = Left(fecha, 4) & "-" & Mid(fecha, 6, 2) & "-" & Mid(fecha, 9, 2) & " " & Mid(fecha, 12, 2) & ":" & Mid(fecha, 15, 2) & ":" & Mid(fecha, 18, 2) & ".000"
        fecha_caracter = Mid(fecha, 9, 2) & "-" & Mid(fecha, 6, 2) & "-" & Left(fecha, 4) & " " & Mid(fecha, 12, 2) & ":" & Mid(fecha, 15, 2) & ":" & Mid(fecha, 18, 2) & ".000"
    End If
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'--------------------------------------------------------------------------------
'--TABLA TEMPORAL DE VENTAS Y ALBARANES
'--------------------------------------------------------------------------------
    tabla = "[sincro_vendesTmpVendes_" & botiguesCad & "_" & nTmp & "]"
    'Si existe sincro_vendesTmp2 la borramos y volvemos a generar
    sql = "SELECT object_id FROM sys.objects with (nolock) "
    sql = sql & "WHERE name=REPLACE(REPLACE('" & tabla & "','[',''),']','') AND type='U' "
    Set Rs = Db.OpenResultset(sql)
    If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tabla)
    sql = "SELECT * INTO " & tabla & " FROM ("
    sql = sql & "SELECT * FROM [Fac_laforneria].[dbo].[V_Venut_" & anyo & "-" & mes & "] "
    sql = sql & "WHERE data>'" & fecha_caracter & "' "
    sql = sql & "UNION ALL "
    sql = sql & "SELECT * FROM [Fac_laforneria].[dbo].[V_Albarans_" & anyo & "-" & mes & "] "
    sql = sql & "WHERE data>'" & fecha_caracter & "' "
    sql = sql & ") t "
    sql = sql & "WHERE botiga='" & botiguesCad & "'"
    Db.QueryTimeout = 0
    Set Rs = Db.OpenResultset(sql)
    Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
'--LINEAS Y CABECERAS
'--------------------------------------------------------------------------------
    InformaMiss "INSERTANDO LINEAS Y CABECERAS-->" & Now
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "INSERTANDO LINEAS Y CABECERAS-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
    ImporteTotal = 0
    ImporteTotal2 = 0
    numLineas = 0
    'SQL ventas SQL SERVER
    sql = "SELECT ROW_NUMBER() OVER (PARTITION BY ven.num_tick ORDER BY ven.num_tick, ven.Plu,ven.Quantitat,ven.Import,ven.Data) AS IdLineaTicket, "
    sql = sql & "1 AS IdEmpresa,LEFT('" & codiBotigaextern & "'," & pos & ") AS IdTienda,RIGHT('" & codiBotigaextern & "',1) AS IdBalanzaMaestra, "
    sql = sql & "-1 AS IdBalanzaEsclava,DENSE_RANK() OVER (ORDER BY ven.num_tick) + '" & maxIdTicket & "'  AS IdTicket,2 AS TipoVenta,0 AS EstadoLinea, "
    sql = sql & "isNull(tmp.idArticulo,'89992') AS IdArticulo,isNull(tmp.Descripcion,'Preu Directe Hit') AS Descripcion, "
    sql = sql & "isNull(tmp.Descripcion1,'Preu Directe Hit') AS Descripcion1,isNull(art.EsSumable,1) AS Comportamiento,0 AS Tara, "
    sql = sql & "ven.Quantitat AS Peso,NULL AS PesoRegalado,round(ven.import/ven.quantitat,2) AS Precio,0 AS PrecioSinIVA, "
    sql = sql & "round(ven.import/ven.quantitat,2) AS PrecioConIVASinDtoL,isNull(tmp.idIva,0) AS IdIVA, "
    sql = sql & "isNull(tmp.PorcentajeIVA,0) AS PorcentajeIVA,NULL AS Descuento, "
    sql = sql & "NULL AS ImporteSinIVASinDtoL,NULL AS ImporteConIVASinDtoL,NULL AS ImporteDelDescuento, "
    sql = sql & "isNull(ROUND(ven.Import,2),0) AS Importe, "
    sql = sql & "isNull(ROUND(CAST(ven.Import AS FLOAT)-((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100),2),0) AS ImporteSinIVAConDtoL, "
    sql = sql & "isNull(ROUND((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100,2),0) AS ImporteDelIVAConDtoL, "
    sql = sql & "isNull(ROUND(CAST(ven.Import AS FLOAT)-((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100),2),0) AS ImporteSinIVAConDtoLConDtoTotal, "
    sql = sql & "isNull(ROUND((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100,2),0) AS ImporteDelIVAConDtoLConDtoTotal, "
    sql = sql & "isNull(ROUND(ven.Import,2),0) AS ImporteConDtoTotal,0 AS TaraFija,0 AS TaraPorcentual,isNull(tmp.idArticulo,'89992') AS CodInterno, "
    sql = sql & "'' as EANScannerArticulo,0 AS IdClase,NULL AS NombreClase,0 AS IdElemAsociado, "
    sql = sql & "isNull(tmp.IdFamilia,0) AS IdFamilia,isNull(tmp.NombreFamilia,'Familia no asignada') AS NombreFamilia, "
    sql = sql & "isNull(tmp.IdSeccion,0) AS IdSeccion, "
    sql = sql & "isNull(tmp.NombreSeccion,'Seccio no asignada') AS NombreSeccion,isNull(tmp.IdSubFamilia,0) AS IdSubFamilia, "
    sql = sql & "isNull(tmp.NombreSubFamilia,'Subfamilia no asignada') AS NombreSubFamilia,isNull(tmp.IdDepartamento,0) AS IdDepartamento, "
    sql = sql & "isNull(tmp.NombreDepartamento,'Departament no asignat') AS NombreDepartamento,1 AS Modificado,'A' AS Operacion, "
    sql = sql & "'Comunicaciones' AS Usuario,ven.Data AS TimeStamp, "
    sql = sql & "'Casa Ametller S.L.' AS NombreEmpresa,CASE WHEN CHARINDEX('_',cli.nom)>0 THEN SUBSTRING(cli.nom,1,CAST(CHARINDEX('_',cli.nom)AS INTEGER)-1)  "
    sql = sql & "ELSE cli.Nom END AS NombreTienda,'T' AS Tipo,isNull(dE.valor,'0000') AS IdVendedor,isNull(dep.nom,'Dependenta sense asignar') AS NombreVendedor, "
    sql = sql & "isNull(tmp2.idCliente,'0') AS IdCliente,isNull(tmp2.Nombre,NULL) AS NombreCliente, "
    sql = sql & "isNull(tmp2.Cif,'0') AS DNICliente,isNull(tmp2.Dir,'0')AS DireccionCliente,isNull(tmp2.Ciutat,'0') AS PoblacionCliente, "
    sql = sql & "NULL AS ProvinciaCliente,isNull(tmp2.CP,'0')AS CPCliente,0 AS TelefonoCliente,NULL AS ImporteLineas,NULL AS PorcDescuento,NULL AS ImporteDescuento, "
    sql = sql & "NULL AS ImporteLineas2,NULL AS PorcDescuento2,NULL AS ImporteDescuento2, "
    sql = sql & "NULL AS ImporteLineas3,NULL AS PorcDescuento3,NULL AS ImporteDescuento3, "
    sql = sql & "NULL AS ImporteTotal3,NULL AS ImporteSinRedondeo,NULL AS ImporteDelRedondeo, "
    sql = sql & "NULL AS SerieLIdFinDeDia,0 AS SerieLTicketErroneo,0 AS ImporteDtoTotalSinIVA, "
    sql = sql & "ven.Data AS Fecha,'C' AS EstadoTicket,0 AS ImporteDevuelto,0 AS PuntosFidelidad,ven.num_tick AS NumTicket "
    sql = sql & "FROM " & tabla & " ven "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[clients] cli ON (ven.Botiga=cli.codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[dependentes] dep ON (ven.dependenta=dep.codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[dependentesExtes] dE on (dep.CODI=dE.id and dE.nom='CODI_DEP') "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[Articles] art ON (ven.Plu=art.Codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[ArticlesPropietats] artP1 ON (art.Codi=artP1.CodiArticle and artP1.Variable='CODI_PROD') "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[TipusIva] iva ON (art.TipoIva=iva.Tipus) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[ConstantsClient] cc ON (CAST(ven.otros as nvarchar(10))=CAST(cc.codi as nvarchar(10)) and cc.Variable='CodiClientOrigen' ) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo]." & tablaTmp & " tmp ON (artP1.valor=tmp.IdArticulo) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo]." & tablaTmp2 & " tmp2 ON (cc.valor=tmp2.IdCliente) "
    sql = sql & "WHERE ven.Botiga='" & codiBotiga & "' and ven.data>'" & fecha_caracter & "' "
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
    Set rsLin = Db.OpenResultset(sql)
    If Not rsLin.EOF Then
        idTicketAnt = rsLin("idTicket")
    End If
    Do While Not rsLin.EOF
'--------------------------------------------------------------------------------
'--INSERTANDO CABECERA
'--------------------------------------------------------------------------------
        If rsLin("IdTicket") <> idTicketAnt Then
            'Importe de ivas de los 3 tipos
            sql = "SELECT ("
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=1 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=1))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=2 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=2))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=3 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=3))/100,0),2) "
            sql = sql & ") as resultado "
            Set Rs = Db.OpenResultset(sql)
            If Not Rs.EOF Then
                ImporteTotalSinIVAConDtoL = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalSinIVAConDtoLConDtoTotal = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalDelIVAConDtoLConDtoTotal = Rs("resultado")
            End If
            ImporteEntregado = ImporteTotal
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR CABECERA!!! BORRADO DE CABECERA ACTUAL POR SI YA EXISTIERA
'----------------------------------------------------------------------------------------------------
            sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
            sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
            sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and idTicket='" & idTicketAnt & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
            'SQL INSERT Cabecera MYSQL
            'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_cabecera limit 1'') "
            sql = "INSERT INTO dat_ticket_cabecera "
            sql = sql & "(IdTicket, NumTicket, IdEmpresa, NombreEmpresa,IdTienda,NombreTienda,IdBalanzaMaestra,"
            sql = sql & "NombreBalanzaMaestra, IdBalanzaEsclava, Tipo,"
            sql = sql & "IdVendedor,NombreVendedor,IdCliente,NombreCliente,"
            sql = sql & "DNICliente,DireccionCliente,PoblacionCliente,"
            sql = sql & "ProvinciaCliente,CPCliente,TelefonoCliente,"
            sql = sql & "TipoVenta,ImporteLineas,PorcDescuento,ImporteDescuento,"
            sql = sql & "ImporteTotal,ImporteLineas2,PorcDescuento2,ImporteDescuento2,"
            sql = sql & "ImporteTotal2,ImporteLineas3,PorcDescuento3,ImporteDescuento3,"
            sql = sql & "ImporteTotal3,ImporteTotalSinIVAConDtoL,ImporteTotalSinIVAConDtoLConDtoTotal,"
            sql = sql & "ImporteDtoTotalSinIVA,ImporteTotalDelIVAConDtoLConDtoTotal,ImporteSinRedondeo,"
            sql = sql & "ImporteDelRedondeo,Fecha,NumLineas,SerieLIdFinDeDia,SerieLTicketErroneo,Modificado,"
            sql = sql & "Operacion,Usuario,TimeStamp,EstadoTicket,ImporteEntregado,ImporteDevuelto,PuntosFidelidad) values ("
            sql = sql & "'" & idTicketAnt & "','" & NumTicket & "','" & idEmpresa & "','" & NombreEmpresa & "',"
            sql = sql & "'" & IdTienda & "','" & NombreTienda & "','" & idBalanzaMaestra & "','-Balança " & idBalanzaMaestra & "',"
            sql = sql & "'" & idBalanzaEsclava & "','" & tipo & "','" & IdVendedor & "','" & NombreVendedor & "',"
            sql = sql & "'" & idCliente & "','" & NombreCliente & "','" & DNICliente & "','" & DireccionCliente & "',"
            sql = sql & "'" & PoblacionCliente & "','" & ProvinciaCliente & "','" & CPCliente & "','" & TelefonoCliente & "',"
            sql = sql & "'" & tipoVenta & "','" & ImporteLineas & "','" & PorcDescuento & "','" & ImporteDescuento & "',"
            sql = sql & "'" & ImporteTotal & "','" & ImporteLineas2 & "','" & PorcDescuento2 & "','" & ImporteDescuento2 & "',"
            sql = sql & "'" & ImporteTotal2 & "','" & ImporteLineas3 & "','" & PorcDescuento3 & "','" & ImporteDescuento3 & "',"
            sql = sql & "'" & ImporteTotal3 & "','" & ImporteTotalSinIVAConDtoL & "','" & ImporteTotalSinIVAConDtoLConDtoTotal & "','" & ImporteDtoTotalSinIVA & "',"
            sql = sql & "'" & ImporteTotalDelIVAConDtoLConDtoTotal & "','" & ImporteSinRedondeo & "','" & ImporteDelRedondeo & "','" & Format(fecha, "yyyy-mm-dd hh:nn:ss") & ".000',"
            sql = sql & "'" & numLineas & "','" & SerieLIdFinDeDia & "','" & SerieLTicketErroneo & "','" & Modificado & "',"
            sql = sql & "'" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000','" & EstadoTicket & "',"
            sql = sql & "'" & ImporteEntregado & "','" & ImporteDevuelto & "','" & PuntosFidelidad & "')"
            InformaMiss "CABECERA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "CABECERA!"
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "INSERTANDO CABECERA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sql
            End If
            ImporteTotal = 0
            ImporteTotal2 = 0
            numLineas = 0
            idTicketAnt = rsLin("idTicket")
        End If
        'Campos lineas
        import = rsLin("Importe")
        IdLineaTicket = rsLin("IdLineaTicket")
        idEmpresa = rsLin("idEmpresa")
        IdTienda = rsLin("IdTienda")
        idBalanzaMaestra = rsLin("IdBalanzaMaestra")
        idBalanzaEsclava = rsLin("IdBalanzaEsclava")
        idTicket = rsLin("IdTicket")
        tipoVenta = rsLin("TipoVenta")
        EstadoLinea = rsLin("EstadoLinea")
        IdArticulo = rsLin("IdArticulo")
        DESCRIPCIoN = rsLin("Descripcion")
        Descripcion1 = rsLin("Descripcion1")
        Comportamiento = rsLin("Comportamiento")
        Tara = rsLin("Tara")
        Peso = rsLin("Peso")
        PesoRegalado = rsLin("PesoRegalado")
        precio = rsLin("Precio")
        PrecioSinIva = rsLin("PrecioSinIva")
        PrecioConIVASinDtoL = rsLin("PrecioConIVASinDtoL")
        IdIVA = rsLin("IdIVA")
        PorcentajeIva = rsLin("PorcentajeIva")
        descuento = rsLin("descuento")
        ImporteSinIVASinDtoL = rsLin("ImporteSinIVASinDtoL")
        ImporteConIVASinDtoL = rsLin("ImporteConIVASinDtoL")
        ImporteDelDescuento = rsLin("ImporteDelDescuento")
        importe = rsLin("Importe")
        ImporteSinIVAConDtoL = rsLin("ImporteSinIVAConDtoL")
        ImporteDelIVAConDtoL = rsLin("ImporteDelIVAConDtoL")
        ImporteSinIVAConDtoLConDtoTotal = rsLin("ImporteSinIVAConDtoLConDtoTotal")
        ImporteDelIVAConDtoLConDtoTotal = rsLin("ImporteDelIVAConDtoLConDtoTotal")
        ImporteConDtoTotal = rsLin("ImporteConDtoTotal")
        TaraFija = rsLin("TaraFija")
        TaraPorcentual = rsLin("TaraPorcentual")
        CodInterno = rsLin("CodInterno")
        EANScannerArticulo = rsLin("EANScannerArticulo")
        IdClase = rsLin("IdClase")
        NombreClase = rsLin("NombreClase")
        IdElemAsociado = rsLin("IdElemAsociado")
        IdFamilia = rsLin("IdFamilia")
        NombreFamilia = rsLin("NombreFamilia")
        IdSeccion = rsLin("IdSeccion")
        nombreSeccion = rsLin("NombreSeccion")
        IdSubFamilia = rsLin("IdSubFamilia")
        NombreSubFamilia = rsLin("NombreSubFamilia")
        IdDepartamento = rsLin("IdDepartamento")
        NombreDepartamento = rsLin("NombreDepartamento")
        Modificado = rsLin("Modificado")
        Operacion = rsLin("Operacion")
        usuario = rsLin("Usuario")
        TimeStamp = rsLin("TimeStamp")
        'Campos cabecera
        NombreEmpresa = rsLin("NombreEmpresa")
        NombreTienda = rsLin("NombreTienda")
        tipo = rsLin("tipo")
        IdVendedor = rsLin("IdVendedor")
        NombreVendedor = rsLin("NombreVendedor")
        idCliente = rsLin("IdCliente")
        NombreCliente = rsLin("NombreCliente")
        DNICliente = rsLin("DNICliente")
        DireccionCliente = rsLin("DireccionCliente")
        PoblacionCliente = rsLin("PoblacionCliente")
        ProvinciaCliente = rsLin("ProvinciaCliente")
        CPCliente = rsLin("CPCliente")
        TelefonoCliente = rsLin("TelefonoCliente")
        ImporteLineas = rsLin("ImporteLineas")
        PorcDescuento = rsLin("PorcDescuento")
        ImporteDescuento = rsLin("ImporteDescuento")
        ImporteLineas2 = rsLin("ImporteLineas2")
        PorcDescuento2 = rsLin("PorcDescuento2")
        ImporteDescuento2 = rsLin("ImporteDescuento2")
        ImporteLineas3 = rsLin("ImporteLineas3")
        PorcDescuento3 = rsLin("PorcDescuento3")
        ImporteDescuento3 = rsLin("ImporteDescuento3")
        ImporteTotal3 = rsLin("ImporteTotal3")
        ImporteSinRedondeo = rsLin("ImporteSinRedondeo")
        ImporteDelRedondeo = rsLin("ImporteDelRedondeo")
        SerieLIdFinDeDia = rsLin("SerieLIdFinDeDia")
        SerieLTicketErroneo = rsLin("SerieLTicketErroneo")
        ImporteDtoTotalSinIVA = rsLin("ImporteDtoTotalSinIVA")
        fecha = rsLin("fecha")
        EstadoTicket = rsLin("EstadoTicket")
        ImporteDevuelto = rsLin("ImporteDevuelto")
        PuntosFidelidad = rsLin("PuntosFidelidad")
        NumTicket = rsLin("numTicket")
        'ImporteTotal en ImporteTotal y ImporteTotal2
        ImporteTotal = ImporteTotal + Round(importe, 2)
        ImporteTotal2 = ImporteTotal2 + Round(importe, 2)
        numLineas = numLineas + 1
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR LINEA!!! BORRADO DE LINEAS DE TICKET PARA EL NUMERO DE TICKET ACTUAL
'----  ES POSIBLE QUE ALGUNA VEZ SE HAYA QUEDADO LA LINEA SIN CABECERA
'----------------------------------------------------------------------------------------------------
        If IdLineaTicket = 1 Then
            sqlSP = "delete FROM dat_ticket_linea "
            sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
            sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
            sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and idTicket='" & idTicket & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
        End If
'--------------------------------------------------------------------------------
'--INSERTANDO LINEA
'--------------------------------------------------------------------------------
        'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_linea limit 1'') ( "
        sql = "INSERT INTO dat_ticket_linea ( "
        sql = sql & "IdLineaTicket,IdEmpresa,IdTienda,IdBalanzaMaestra,IdBalanzaEsclava,IdTicket,"
        sql = sql & "TipoVenta,EstadoLinea,IdArticulo,Descripcion,Descripcion1,Comportamiento,Tara,"
        sql = sql & "Peso,PesoRegalado,Precio,PrecioSinIVA,PrecioConIVASinDtoL,IdIVA,PorcentajeIVA,"
        sql = sql & "Descuento,ImporteSinIVASinDtoL,ImporteConIVASinDtoL,ImporteDelDescuento,Importe,"
        sql = sql & "ImporteSinIVAConDtoL,ImporteDelIVAConDtoL,ImporteSinIVAConDtoLConDtoTotal,"
        sql = sql & "ImporteDelIVAConDtoLConDtoTotal,ImporteConDtoTotal,TaraFija,TaraPorcentual,"
        sql = sql & "CodInterno,EANScannerArticulo,IdClase,NombreClase,IdElemAsociado,IdFamilia,"
        sql = sql & "NombreFamilia,IdSeccion,NombreSeccion,IdSubFamilia,NombreSubFamilia,IdDepartamento,"
        sql = sql & "NombreDepartamento,Modificado,Operacion,Usuario,TimeStamp) "
        sql = sql & " values ('" & IdLineaTicket & "','" & idEmpresa & "','" & IdTienda & "','"
        sql = sql & idBalanzaMaestra & "','" & idBalanzaEsclava & "','" & idTicket & "','"
        sql = sql & tipoVenta & "','" & EstadoLinea & "','" & IdArticulo & "','" & Replace(DESCRIPCIoN, "'", "''") & "','" & Replace(Descripcion1, "'", "''") & "','" & CInt(Comportamiento) & "','" & Tara & "','"
        sql = sql & Peso & "','" & PesoRegalado & "','" & precio & "','" & PrecioSinIva & "','"
        sql = sql & PrecioConIVASinDtoL & "','" & IdIVA & "','" & PorcentajeIva & "','" & descuento & "','" & ImporteSinIVASinDtoL & "','" & ImporteConIVASinDtoL & "','"
        sql = sql & ImporteDelDescuento & "',' " & importe & "',' " & ImporteSinIVAConDtoL & "','"
        sql = sql & ImporteDelIVAConDtoL & "','" & ImporteSinIVAConDtoLConDtoTotal & "','"
        sql = sql & ImporteDelIVAConDtoLConDtoTotal & "','" & ImporteConDtoTotal & "','" & TaraFija & "','" & TaraPorcentual & "','"
        sql = sql & CodInterno & "','" & EANScannerArticulo & "','" & IdClase & "','" & NombreClase & "','" & IdElemAsociado & "','" & IdFamilia & "','"
        sql = sql & NombreFamilia & "','" & IdSeccion & "','" & nombreSeccion & "','" & IdSubFamilia & "','" & NombreSubFamilia & "','" & IdDepartamento & "','"
        sql = sql & NombreDepartamento & "','" & Modificado & "','" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000') "
        InformaMiss "LINEA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "INSERTANDO LINEA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sql
        End If
        rsLin.MoveNext
'----------------------------------------------------------------------------------------------------
'----  ULTIMA CABECERA
'----------------------------------------------------------------------------------------------------
        If rsLin.EOF Then
            'Importe de ivas de los 3 tipos
            sql = "SELECT ("
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=1 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=1))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=2 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=2))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=3 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=3))/100,0),2) "
            sql = sql & ") as resultado "
            Set Rs = Db.OpenResultset(sql)
            If Not Rs.EOF Then
                ImporteTotalSinIVAConDtoL = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalSinIVAConDtoLConDtoTotal = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalDelIVAConDtoLConDtoTotal = Rs("resultado")
            End If
            ImporteEntregado = ImporteTotal
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR CABECERA!!! BORRADO DE CABECERA ACTUAL POR SI YA EXISTIERA
'----------------------------------------------------------------------------------------------------
            sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
            sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
            sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and idTicket='" & idTicket & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
            'SQL INSERT Cabecera MYSQL
            'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_cabecera limit 1'') "
            sql = "INSERT INTO dat_ticket_cabecera "
            sql = sql & "(IdTicket, NumTicket, IdEmpresa, NombreEmpresa,IdTienda,NombreTienda,IdBalanzaMaestra,"
            sql = sql & "NombreBalanzaMaestra, IdBalanzaEsclava, Tipo,"
            sql = sql & "IdVendedor,NombreVendedor,IdCliente,NombreCliente,"
            sql = sql & "DNICliente,DireccionCliente,PoblacionCliente,"
            sql = sql & "ProvinciaCliente,CPCliente,TelefonoCliente,"
            sql = sql & "TipoVenta,ImporteLineas,PorcDescuento,ImporteDescuento,"
            sql = sql & "ImporteTotal,ImporteLineas2,PorcDescuento2,ImporteDescuento2,"
            sql = sql & "ImporteTotal2,ImporteLineas3,PorcDescuento3,ImporteDescuento3,"
            sql = sql & "ImporteTotal3,ImporteTotalSinIVAConDtoL,ImporteTotalSinIVAConDtoLConDtoTotal,"
            sql = sql & "ImporteDtoTotalSinIVA,ImporteTotalDelIVAConDtoLConDtoTotal,ImporteSinRedondeo,"
            sql = sql & "ImporteDelRedondeo,Fecha,NumLineas,SerieLIdFinDeDia,SerieLTicketErroneo,Modificado,"
            sql = sql & "Operacion,Usuario,TimeStamp,EstadoTicket,ImporteEntregado,ImporteDevuelto,PuntosFidelidad) values ("
            sql = sql & "'" & idTicket & "','" & NumTicket & "','" & idEmpresa & "','" & NombreEmpresa & "',"
            sql = sql & "'" & IdTienda & "','" & NombreTienda & "','" & idBalanzaMaestra & "','-Balança " & idBalanzaMaestra & "',"
            sql = sql & "'" & idBalanzaEsclava & "','" & tipo & "','" & IdVendedor & "','" & NombreVendedor & "',"
            sql = sql & "'" & idCliente & "','" & NombreCliente & "','" & DNICliente & "','" & DireccionCliente & "',"
            sql = sql & "'" & PoblacionCliente & "','" & ProvinciaCliente & "','" & CPCliente & "','" & TelefonoCliente & "',"
            sql = sql & "'" & tipoVenta & "','" & ImporteLineas & "','" & PorcDescuento & "','" & ImporteDescuento & "',"
            sql = sql & "'" & ImporteTotal & "','" & ImporteLineas2 & "','" & PorcDescuento2 & "','" & ImporteDescuento2 & "',"
            sql = sql & "'" & ImporteTotal2 & "','" & ImporteLineas3 & "','" & PorcDescuento3 & "','" & ImporteDescuento3 & "',"
            sql = sql & "'" & ImporteTotal3 & "','" & ImporteTotalSinIVAConDtoL & "','" & ImporteTotalSinIVAConDtoLConDtoTotal & "','" & ImporteDtoTotalSinIVA & "',"
            sql = sql & "'" & ImporteTotalDelIVAConDtoLConDtoTotal & "','" & ImporteSinRedondeo & "','" & ImporteDelRedondeo & "','" & Format(fecha, "yyyy-mm-dd hh:nn:ss") & ".000',"
            sql = sql & "'" & numLineas & "','" & SerieLIdFinDeDia & "','" & SerieLTicketErroneo & "','" & Modificado & "',"
            sql = sql & "'" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000','" & EstadoTicket & "',"
            sql = sql & "'" & ImporteEntregado & "','" & ImporteDevuelto & "','" & PuntosFidelidad & "')"
            InformaMiss "CABECERA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "CABECERA!"
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "INSERTANDO CABECERA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sql
            End If
        End If
    Loop
'--------------------------------------------------------------------------------
'--SIGUIENTE TIENDA
'--------------------------------------------------------------------------------
    rsClients.MoveNext
    'Borramos tablas temp
    Db.OpenResultset ("DROP TABLE " & tablaTmp)
    Db.OpenResultset ("DROP TABLE " & tablaTmp2)
    Db.OpenResultset ("DROP TABLE " & tabla)
    html = html & "<p><b>Botiga: </b>" & codiBotiga & "</p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & desde & "</p>"
    html = html & "<p><b>Finalitzat: </b>" & Now() & "</p>"
Loop
If P3 <> "" Then sf_enviarMail "secrehit@hit.cat", Desti, " Vendes sincronitzades " & codiBotiga, html, "", ""

sf_enviarMail "secrehit@hit.cat", EmailGuardia, " Vendes sincronitzades " & codiBotiga, html, "", ""

connMysql.Close
Set connMysql = Nothing

InformaMiss "FIN SINCRO_VENDES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "FIN SINCRO_VENDES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.Close
End If
Exit Function

norVendes:
'----------------------------------------------------------------------------------------------------
'----  SI HAY ERROR: BORRADO ULTIMAS LINEAS DE TICKET HUERFANAS SIN CABECERA
'----------------------------------------------------------------------------------------------------
    html = "<p><h3>Resum Vendes Ametller </h3></p>"
    html = html & "<p><b>Botiga: </b>" & botiguesCad & "</p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & desde & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    If P3 <> "" Then sf_enviarMail "secrehit@hit.cat", Desti, "ERROR! Sincronitzacio de vendes de " & codiBotiga & " ha fallat", html, "", ""
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    html = html & "<p><b>Ultima sqlSP:</b>" & sqlSP & "</p>"
    
    sf_enviarMail "secrehit@hit.cat", EmailGuardia, "ERROR! Sincronitzacio de vendes de " & codiBotiga & " ha fallat", html, "", ""
        
    Set connMysql = New ADODB.Connection
    connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
    connMysql.ConnectionTimeout = 1000 '16 min
    connMysql.Open

    sqlSP = "delete FROM dat_ticket_linea "
    sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
    sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
    sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
    sqlSP = sqlSP & "and idTicket='" & idTicket & "' "
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "Error:" & err.Description & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        connMysql.Execute sqlSP
    End If
    
    connMysql.Close
    Set connMysql = Nothing
    
    'Borramos tablas temporales
    Db.OpenResultset ("DROP TABLE " & tablaTmp2)
    Db.OpenResultset ("DROP TABLE " & tablaTmp)
    Db.OpenResultset ("DROP TABLE " & tabla)
    
    ExecutaComandaSql "Delete from  " & TaulaCalculsEspecials & " Id = '" & idTasca & "' "
    
    Informa "Error : " & err.Description
    
End Function

Function XX_SincroDbVendesAmetller2(p1, P2, P3, P4, idTasca) As Boolean
Dim botiguesCad As String, mes As String, anyo As String, desde As String, debugSincro As Boolean
Dim codiBotiga As String, codiBotigaextern As String, tabla As String, fecha, fecha_caracter
Dim sql As String, sql2 As String, sql3 As String, sqlSP As String, numCab As Integer, maxIdTicket As String
Dim parametros As String, tablaTmp As String, Rs As rdoResultset, rsClients As rdoResultset, pos As Integer
Dim rsCab As rdoResultset, rsLin As rdoResultset, html As String, hasta As String, desde0, tablaTmp2 As String, tablaTmp3 As String
Dim connMysql As ADODB.Connection

'JUNTA ALBARANS AMB VENDES!
'MERMES 09/08/2012

'If P1 <> "9011" Then Exit Function

On Error GoTo norVendes
'Parametros
'Si pasamos parametros vacios se recorreran todas las tiendas de paramsHw
',mirando ventas (fechas actuales), se eliminaran datos en servidor remoto
'si se le indica una fecha desde y se ejecutaran inserciones (cabeceras y lineas)

'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
botiguesCad = p1 'Cadena de tiendas separadas por ,
If botiguesCad <> "" Then
    botiguesCad = Replace(botiguesCad, "[", "")
    botiguesCad = Replace(botiguesCad, "]", "")
End If
desde = P2 'Fecha desde la que eliminan datos en tablas ventas de serv. remoto
If desde <> "" Then
    desde = Replace(desde, "[", "")
    desde = Replace(desde, "]", "")
End If
desde0 = desde
mes = Month(desde) 'Mes para tabla ventas
If mes <> "" Then If Len(mes) = 1 Then mes = "0" & mes
anyo = Year(desde) 'Año para tabla ventas
Desti = P3 'Email
'desde = CDate(desde)
'hasta = P4 'Fecha desde la que eliminan datos en tablas ventas de serv. remoto
'If hasta <> "" Then
'    hasta = Replace(hasta, "[", "")
'    hasta = Replace(hasta, "]", "")
'End If
debugSincro = False 'Solo impresion o ejecucion completa
'Crear txt log
If debugSincro = True Then
    Desti = EmailGuardia
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\Vendes_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
End If
'--------------------------------------------------------------
numCab = 0
If mes = "" Then
    mes = Month(Now)
    If Len(mes) = 1 Then mes = "0" & mes
End If
If anyo = "" Then anyo = Year(Now)
InformaMiss "INICIO SINCRO_VENDES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "INICIO SINCRO_VENDES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'--TABLAS TEMPORALES
'--------------------------------------------------------------------------------
'Creamos tablas temporales de las quales podemos obtener datos de familia, secciones, etc
InformaMiss "CREANDO TABLAS TEMPORALES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CREANDO TABLAS TEMPORALES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
nTmp = Now
tablaTmp = "[Fac_laforneria].[dbo].[sincro_vendesTmpArticles_" & botiguesCad & "_" & nTmp & "]"
tablaTmp2 = "[Fac_laforneria].[dbo].[sincro_vendesTmpClients_" & botiguesCad & "_" & nTmp & "]"
tablaTmp3 = "[Fac_laforneria].[dbo].[sincro_vendesTmpMermes_" & botiguesCad & "_" & nTmp & "]"
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmpArticles la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp)
sql = " SELECT * INTO " & tablaTmp & " FROM OPENQUERY(AMETLLER,'"
sql = sql & "SELECT a.idArticulo,a.Descripcion,a.Descripcion1,a.idFamilia,f.NombreFamilia,a.idSubFamilia,"
sql = sql & "sf.NombreSubFamilia,a.idDepartamento,d.NombreDepartamento,a.idSeccion,s.NombreSeccion,"
sql = sql & "a.PrecioSinIVA , a.PrecioConIva, i.IdIVA, i.PorcentajeIva "
sql = sql & "FROM dat_articulo a "
sql = sql & "LEFT JOIN dat_familia f on (a.idFamilia=f.idFamilia) "
sql = sql & "LEFT JOIN dat_subfamilia sf on a.idFamilia=sf.idFamilia and a.idsubfamilia = sf.idsubfamilia "
sql = sql & "LEFT JOIN dat_departamento d on (a.idDepartamento=d.idDepartamento) "
sql = sql & "LEFT JOIN dat_seccion s on (a.idSeccion=s.idSeccion) "
sql = sql & "LEFT JOIN dat_iva i on (a.idIva=i.idIva) "
sql = sql & "Where a.idEmpresa = 1 And IdArticulo < 90000') "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmpClients la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp2 & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp2)
sql = " SELECT * INTO " & tablaTmp2 & " FROM OPENQUERY(AMETLLER,'"
sql = sql & "SELECT idCliente, Nombre,TRIM(DNI) AS Cif, TRIM(Direccion) AS Dir,"
sql = sql & "TRIM(CodPostal) AS CP, TRIM(Poblacion) AS Ciutat,TimeStamp AS fecMod "
sql = sql & "FROM dat_cliente WHERE IdEmpresa=1') "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmpMermes la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp3 & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp3)
sql = "SELECT Client as Botiga,[TimeStamp] as Data,'9999' as Dependenta,"
sql = sql & "DENSE_RANK() OVER (ORDER BY s.[timestamp],s.client) as Num_tick,"
sql = sql & "Client as Estat,PluUtilitzat as Plu,QuantitatTornada as Quantitat,"
sql = sql & "0 as Import,'V' as Tipus_venta,'0' as FormaMarcar,'9999' as Otros "
'Sql = Sql & "(QuantitatTornada*isNull(a.Preu,0)) as Import,'V' as Tipus_venta,0 as FormaMarcar,'9999' as Otros "
sql = sql & "INTO " & tablaTmp3 & " FROM [Fac_laforneria].[dbo].[Servit-" & Format(CDate(desde), "yy-mm-dd") & "] s "
sql = sql & "LEFT JOIN articles a on s.pluUtilitzat=a.Codi where quantitatTornada<>'0' "
sql = sql & "ORDER BY s.[timestamp],s.client "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
html = "<p><h3>Resum Vendes Ametller</h3></p>"
'--------------------------------------------------------------------------------
'--CURSOR TIENDAS
'--------------------------------------------------------------------------------
'--Creamos cursor para tiendas
InformaMiss "CREANDO CURSOR TIENDAS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CREANDO CURSOR TIENDAS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
sql = "SELECT p.valor1 FROM [Fac_laforneria].[dbo].Clients c LEFT JOIN [Fac_laforneria].[dbo].ParamsHw p "
sql = sql & "ON (c.codi=p.valor1) WHERE p.valor1 is not null "
If botiguesCad <> "" Then sql = sql & "and p.valor1 IN ( " & botiguesCad & ")  "
sql = sql & " order by p.valor1 "
Set rsClients = Db.OpenResultset(sql)
Do While Not rsClients.EOF
    codiBotiga = rsClients("valor1")
    If codiBotiga = "518" Then
        codiBotigaextern = 1061
    Else
        codiBotigaextern = codiBotiga
    End If
    pos = Len(codiBotigaextern)
    If pos > 3 Then
        pos = 3
    Else
        pos = 2
    End If
    InformaMiss "TRATANDO TIENDA " & codiBotiga & "(" & codiBotigaextern & "),M:" & mes & ",A:" & anyo & "-->" & Now
    Informa2 "T:" & codiBotiga & "(" & codiBotigaextern & "),DESDE:" & desde
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "TRATANDO TIENDA " & codiBotiga & "(" & codiBotigaextern & "),M:" & mes & ",A:" & anyo & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'--------------------------------------------------------------------------------
'--ACTUALIZACION DE DUPLICADOS. TIQUETS CON MAS DE UNA FECHA DIFERENTE QUE AL AGRUPARSE
'--POR NUMERO DE TIQUET Y FECHA, GENERA DOS LINEAS PARA ESE MISMO NUMERO DE TIQUET
'--Y A LA HORA DE IMPORTAR A MYSQL DA ERRORES.
'--------------------------------------------------------------------------------
'--Si existe sincro_duplicadosTmp la borramos y volvemos a generar
    'Sql = "SELECT object_id FROM sys.objects with (nolock) WHERE name='sincro_duplicadosTmp' AND type='U' "
    'Set Rs = Db.OpenResultset(Sql)
    'If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE sincro_duplicadosTmp")
'--Si existe sincro_duplicadosTmp2 la borramos y volvemos a generar
    'Sql = "SELECT object_id FROM sys.objects with (nolock) WHERE name='sincro_duplicadosTmp2' AND type='U' "
    'Set Rs = Db.OpenResultset(Sql)
    'If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE sincro_duplicadosTmp2")
    'Sql = "SELECT data,Num_tick,botiga INTO sincro_duplicadosTmp FROM " & tabla & " GROUP BY data,Num_tick,Botiga "
    'Set Rs = Db.OpenResultset(Sql)
    'Sql = "SELECT Num_tick,Botiga INTO sincro_duplicadosTmp2 FROM sincro_duplicadosTmp GROUP BY Num_tick,Botiga "
    'Sql = Sql & "Having Count(Num_tick) >= 2 "
    'Set Rs = Db.OpenResultset(Sql)
    'Sql = "UPDATE " & tabla & " SET Data=v.data FROM ( "
    'Sql = Sql & "SELECT MIN(data) data,Num_tick,botiga FROM " & tabla & " "
    'Sql = Sql & "WHERE Num_tick IN(SELECT Num_tick FROM sincro_duplicadosTmp2 s WHERE " & tabla & ".Botiga=s.botiga) "
    'Sql = Sql & "AND botiga IN (SELECT botiga FROM sincro_duplicadosTmp2 s WHERE " & tabla & ".Botiga=s.botiga) "
    'Sql = Sql & "GROUP BY Num_tick,botiga "
    'Sql = Sql & ") v "
    'Sql = Sql & "Where " & tabla & ".Num_tick=v.Num_tick "
    'Sql = Sql & "AND " & tabla & ".Botiga=v.Botiga "
    'Set Rs = Db.OpenResultset(Sql)
    'If debugSincro = True Then
    '    Txt.WriteLine "--------------------------------------------------------------------------------"
    '    Txt.WriteLine "SQL:" & Sql & "-->" & Now
    '    Txt.WriteLine "--------------------------------------------------------------------------------"
    'End If
'--------------------------------------------------------------------------------
'--BORRADO DE DATOS EN SERV.REMOTO
'--CUIDADO! NO ELIMINA LAS LINEAS QUE SE HAYAN PODIDO QUEDAR SIN CABECERA!!!
'--------------------------------------------------------------------------------
    InformaMiss "BORRADO DE DATOS EN SERV.REMOTO DESDE " & desde & "-->" & Now
    If desde <> "" Then
        desde = Year(desde) & "-" & Month(desde) & "-" & Day(desde) & " " & Hour(desde) & ":" & Minute(desde) & ":" & Second(desde) & ".000"
'----------------------------------------------------------------------------------------------------
'---- LINEAS DE TICKET
'----------------------------------------------------------------------------------------------------
        sqlSP = "delete FROM dat_ticket_linea "
        sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
        sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
        sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' and idTicket "
        sqlSP = sqlSP & "in (select idTicket from dat_ticket_cabecera where idEmpresa=1 "
        sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' and idBalanzaEsclava=-1 "
        sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
        sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and timestamp>='" & desde & "' "
        'If hasta <> "" Then sqlSP = sqlSP & " and timestamp<='" & hasta & "' "
        sqlSP = sqlSP & ")"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sqlSP
        End If
'----------------------------------------------------------------------------------------------------
'---- CABECERAS DE TICKET
'----------------------------------------------------------------------------------------------------
        sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
        sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
        sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
        sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and timestamp>'" & desde & "' "
        'If hasta <> "" Then sqlSP = sqlSP & " and timestamp<='" & hasta & "' "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sqlSP
        End If
    End If
'--------------------------------------------------------------------------------
'--MAX IDTICKET
'--------------------------------------------------------------------------------
    InformaMiss "OBTENIENDO ULTIMO NUMERO TICKET-->" & Now
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "OBTENIENDO ULTIMO NUMERO TICKET-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'----------------------------------------------------------------------------------------------------
'---- SACAMOS EL IDTICKET MAXIMO Y LA FECHA MAXIMA DE LA TIENDA
'----------------------------------------------------------------------------------------------------
    sql = "select isNull(IdTicket,0) idTicket,isNull(fecha,'" & anyo & "-" & mes & "-01 00:00:01.000') fecha FROM openquery(AMETLLER, "
    sql = sql & "'SELECT MAX(idTicket)as IdTicket,MAX(TimeStamp)as fecha From dat_ticket_cabecera "
    sql = sql & " WHERE IdEmpresa=1 AND IdTienda=''" & Left(codiBotigaextern, pos) & "'' "
    sql = sql & "AND IdBalanzaMaestra=''" & Right(codiBotigaextern, 1) & "'' "
    sql = sql & "GROUP BY IdEmpresa,IdTienda,IdBalanzaMaestra') "
    Set Rs = Db.OpenResultset(sql)
    maxIdTicket = 0
    fecha = desde0
    fecha_caracter = desde0
    If Not Rs.EOF Then
        maxIdTicket = Rs("idTicket")
        fecha = Rs("fecha")
        'LA SQL NO FUNCIONA CON ESTE TIPO FECHA! fecha_caracter = Left(fecha, 4) & "-" & Mid(fecha, 6, 2) & "-" & Mid(fecha, 9, 2) & " " & Mid(fecha, 12, 2) & ":" & Mid(fecha, 15, 2) & ":" & Mid(fecha, 18, 2) & ".000"
        fecha_caracter = Mid(fecha, 9, 2) & "-" & Mid(fecha, 6, 2) & "-" & Left(fecha, 4) & " " & Mid(fecha, 12, 2) & ":" & Mid(fecha, 15, 2) & ":" & Mid(fecha, 18, 2) & ".000"
    End If
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'--------------------------------------------------------------------------------
'--TABLA TEMPORAL DE VENTAS Y ALBARANES
'--------------------------------------------------------------------------------
    tabla = "[Fac_laforneria].[dbo].[sincro_vendesTmpVendes_" & botiguesCad & "_" & nTmp & "]"
    'Si existe sincro_vendesTmp2 la borramos y volvemos a generar
    sql = "SELECT object_id FROM sys.objects with (nolock) "
    sql = sql & "WHERE name=REPLACE(REPLACE('" & tabla & "','[',''),']','') AND type='U' "
    Set Rs = Db.OpenResultset(sql)
    If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tabla)
    sql = "SELECT * INTO " & tabla & " FROM ("
    sql = sql & "SELECT * FROM [Fac_laforneria].[dbo].[V_Venut_" & anyo & "-" & mes & "] "
    sql = sql & "WHERE data>'" & fecha_caracter & "' and botiga='" & botiguesCad & "' "
    sql = sql & "UNION ALL "
    sql = sql & "SELECT * FROM [Fac_laforneria].[dbo].[V_Albarans_" & anyo & "-" & mes & "] "
    sql = sql & "WHERE data>'" & fecha_caracter & "' and botiga='" & botiguesCad & "' "
    sql = sql & "UNION ALL "
    sql = sql & "SELECT * FROM " & tablaTmp3
    sql = sql & "WHERE data>'" & fecha_caracter & "' and botiga='" & botiguesCad & "' "
    sql = sql & ") t "
    sql = sql & "WHERE botiga='" & botiguesCad & "'"
    Db.QueryTimeout = 0
    Set Rs = Db.OpenResultset(sql)
    Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
'--LINEAS Y CABECERAS
'--------------------------------------------------------------------------------
    InformaMiss "INSERTANDO LINEAS Y CABECERAS-->" & Now
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "INSERTANDO LINEAS Y CABECERAS-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
    ImporteTotal = 0
    ImporteTotal2 = 0
    numLineas = 0
    'SQL ventas SQL SERVER
    sql = "SELECT ROW_NUMBER() OVER (PARTITION BY ven.num_tick ORDER BY ven.num_tick, ven.Plu,ven.Quantitat,ven.Import,ven.Data) AS IdLineaTicket, "
    sql = sql & "1 AS IdEmpresa,LEFT('" & codiBotigaextern & "'," & pos & ") AS IdTienda,RIGHT('" & codiBotigaextern & "',1) AS IdBalanzaMaestra, "
    sql = sql & "-1 AS IdBalanzaEsclava,DENSE_RANK() OVER (ORDER BY ven.num_tick) + '" & maxIdTicket & "'  AS IdTicket,2 AS TipoVenta,0 AS EstadoLinea, "
    sql = sql & "isNull(tmp.idArticulo,'89992') AS IdArticulo,isNull(tmp.Descripcion,'Preu Directe Hit') AS Descripcion, "
    sql = sql & "isNull(tmp.Descripcion1,'Preu Directe Hit') AS Descripcion1,isNull(art.EsSumable,1) AS Comportamiento,0 AS Tara, "
    sql = sql & "ven.Quantitat AS Peso,NULL AS PesoRegalado,round(ven.import/ven.quantitat,2) AS Precio,0 AS PrecioSinIVA, "
    sql = sql & "round(ven.import/ven.quantitat,2) AS PrecioConIVASinDtoL,isNull(tmp.idIva,0) AS IdIVA, "
    sql = sql & "isNull(tmp.PorcentajeIVA,0) AS PorcentajeIVA,NULL AS Descuento, "
    sql = sql & "NULL AS ImporteSinIVASinDtoL,NULL AS ImporteConIVASinDtoL,NULL AS ImporteDelDescuento, "
    sql = sql & "isNull(ROUND(ven.Import,2),0) AS Importe, "
    sql = sql & "isNull(ROUND(CAST(ven.Import AS FLOAT)-((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100),2),0) AS ImporteSinIVAConDtoL, "
    sql = sql & "isNull(ROUND((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100,2),0) AS ImporteDelIVAConDtoL, "
    sql = sql & "isNull(ROUND(CAST(ven.Import AS FLOAT)-((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100),2),0) AS ImporteSinIVAConDtoLConDtoTotal, "
    sql = sql & "isNull(ROUND((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100,2),0) AS ImporteDelIVAConDtoLConDtoTotal, "
    sql = sql & "isNull(ROUND(ven.Import,2),0) AS ImporteConDtoTotal,0 AS TaraFija,0 AS TaraPorcentual,isNull(tmp.idArticulo,'89992') AS CodInterno, "
    sql = sql & "'' as EANScannerArticulo,0 AS IdClase,NULL AS NombreClase,0 AS IdElemAsociado, "
    sql = sql & "isNull(tmp.IdFamilia,0) AS IdFamilia,isNull(tmp.NombreFamilia,'Familia no asignada') AS NombreFamilia, "
    sql = sql & "isNull(tmp.IdSeccion,0) AS IdSeccion, "
    sql = sql & "isNull(tmp.NombreSeccion,'Seccio no asignada') AS NombreSeccion,isNull(tmp.IdSubFamilia,0) AS IdSubFamilia, "
    sql = sql & "isNull(tmp.NombreSubFamilia,'Subfamilia no asignada') AS NombreSubFamilia,isNull(tmp.IdDepartamento,0) AS IdDepartamento, "
    sql = sql & "isNull(tmp.NombreDepartamento,'Departament no asignat') AS NombreDepartamento,1 AS Modificado,'A' AS Operacion, "
    sql = sql & "'Comunicaciones' AS Usuario,ven.Data AS TimeStamp, "
    sql = sql & "'Casa Ametller S.L.' AS NombreEmpresa,CASE WHEN CHARINDEX('_',cli.nom)>0 THEN SUBSTRING(cli.nom,1,CAST(CHARINDEX('_',cli.nom)AS INTEGER)-1)  "
    sql = sql & "ELSE cli.Nom END AS NombreTienda,'T' AS Tipo,isNull(dE.valor,'0000') AS IdVendedor,isNull(dep.nom,'Dependenta sense asignar') AS NombreVendedor, "
    sql = sql & "isNull(tmp2.idCliente,'0') AS IdCliente,isNull(tmp2.Nombre,NULL) AS NombreCliente, "
    sql = sql & "isNull(tmp2.Cif,'0') AS DNICliente,isNull(tmp2.Dir,'0')AS DireccionCliente,isNull(tmp2.Ciutat,'0') AS PoblacionCliente, "
    sql = sql & "NULL AS ProvinciaCliente,isNull(tmp2.CP,'0')AS CPCliente,0 AS TelefonoCliente,NULL AS ImporteLineas,NULL AS PorcDescuento,NULL AS ImporteDescuento, "
    sql = sql & "NULL AS ImporteLineas2,NULL AS PorcDescuento2,NULL AS ImporteDescuento2, "
    sql = sql & "NULL AS ImporteLineas3,NULL AS PorcDescuento3,NULL AS ImporteDescuento3, "
    sql = sql & "NULL AS ImporteTotal3,NULL AS ImporteSinRedondeo,NULL AS ImporteDelRedondeo, "
    sql = sql & "NULL AS SerieLIdFinDeDia,0 AS SerieLTicketErroneo,0 AS ImporteDtoTotalSinIVA, "
    sql = sql & "ven.Data AS Fecha,'C' AS EstadoTicket,0 AS ImporteDevuelto,0 AS PuntosFidelidad,ven.num_tick AS NumTicket "
    sql = sql & "FROM " & tabla & " ven "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[clients] cli ON (ven.Botiga=cli.codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[dependentes] dep ON (ven.dependenta=dep.codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[dependentesExtes] dE on (dep.CODI=dE.id and dE.nom='CODI_DEP') "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[Articles] art ON (ven.Plu=art.Codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[ArticlesPropietats] artP1 ON (art.Codi=artP1.CodiArticle and artP1.Variable='CODI_PROD') "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[TipusIva] iva ON (art.TipoIva=iva.Tipus) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[ConstantsClient] cc ON (CAST(ven.otros as nvarchar(10))=CAST(cc.codi as nvarchar(10)) and cc.Variable='CodiClientOrigen' ) "
    sql = sql & "LEFT JOIN " & tablaTmp & " tmp ON (artP1.valor=tmp.IdArticulo) "
    sql = sql & "LEFT JOIN " & tablaTmp2 & " tmp2 ON (cc.valor=tmp2.IdCliente) "
    sql = sql & "WHERE ven.Botiga='" & codiBotiga & "' and ven.data>'" & fecha_caracter & "' "
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
    Set rsLin = Db.OpenResultset(sql)
    If Not rsLin.EOF Then
        idTicketAnt = rsLin("idTicket")
    End If
    Do While Not rsLin.EOF
'--------------------------------------------------------------------------------
'--INSERTANDO CABECERA
'--------------------------------------------------------------------------------
        If rsLin("IdTicket") <> idTicketAnt Then
            'Importe de ivas de los 3 tipos
            sql = "SELECT ("
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=1 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=1))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=2 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=2))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=3 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=3))/100,0),2) "
            sql = sql & ") as resultado "
            Set Rs = Db.OpenResultset(sql)
            If Not Rs.EOF Then
                ImporteTotalSinIVAConDtoL = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalSinIVAConDtoLConDtoTotal = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalDelIVAConDtoLConDtoTotal = Rs("resultado")
            End If
            ImporteEntregado = ImporteTotal
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR CABECERA!!! BORRADO DE CABECERA ACTUAL POR SI YA EXISTIERA
'----------------------------------------------------------------------------------------------------
            sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
            sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
            sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and idTicket='" & idTicketAnt & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
            'SQL INSERT Cabecera MYSQL
            'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_cabecera limit 1'') "
            sql = "INSERT INTO dat_ticket_cabecera "
            sql = sql & "(IdTicket, NumTicket, IdEmpresa, NombreEmpresa,IdTienda,NombreTienda,IdBalanzaMaestra,"
            sql = sql & "NombreBalanzaMaestra, IdBalanzaEsclava, Tipo,"
            sql = sql & "IdVendedor,NombreVendedor,IdCliente,NombreCliente,"
            sql = sql & "DNICliente,DireccionCliente,PoblacionCliente,"
            sql = sql & "ProvinciaCliente,CPCliente,TelefonoCliente,"
            sql = sql & "TipoVenta,ImporteLineas,PorcDescuento,ImporteDescuento,"
            sql = sql & "ImporteTotal,ImporteLineas2,PorcDescuento2,ImporteDescuento2,"
            sql = sql & "ImporteTotal2,ImporteLineas3,PorcDescuento3,ImporteDescuento3,"
            sql = sql & "ImporteTotal3,ImporteTotalSinIVAConDtoL,ImporteTotalSinIVAConDtoLConDtoTotal,"
            sql = sql & "ImporteDtoTotalSinIVA,ImporteTotalDelIVAConDtoLConDtoTotal,ImporteSinRedondeo,"
            sql = sql & "ImporteDelRedondeo,Fecha,NumLineas,SerieLIdFinDeDia,SerieLTicketErroneo,Modificado,"
            sql = sql & "Operacion,Usuario,TimeStamp,EstadoTicket,ImporteEntregado,ImporteDevuelto,PuntosFidelidad) values ("
            sql = sql & "'" & idTicketAnt & "','" & NumTicket & "','" & idEmpresa & "','" & NombreEmpresa & "',"
            sql = sql & "'" & IdTienda & "','" & NombreTienda & "','" & idBalanzaMaestra & "','-Balança " & idBalanzaMaestra & "',"
            sql = sql & "'" & idBalanzaEsclava & "','" & tipo & "','" & IdVendedor & "','" & NombreVendedor & "',"
            sql = sql & "'" & idCliente & "','" & NombreCliente & "','" & DNICliente & "','" & DireccionCliente & "',"
            sql = sql & "'" & PoblacionCliente & "','" & ProvinciaCliente & "','" & CPCliente & "','" & TelefonoCliente & "',"
            sql = sql & "'" & tipoVenta & "','" & ImporteLineas & "','" & PorcDescuento & "','" & ImporteDescuento & "',"
            sql = sql & "'" & ImporteTotal & "','" & ImporteLineas2 & "','" & PorcDescuento2 & "','" & ImporteDescuento2 & "',"
            sql = sql & "'" & ImporteTotal2 & "','" & ImporteLineas3 & "','" & PorcDescuento3 & "','" & ImporteDescuento3 & "',"
            sql = sql & "'" & ImporteTotal3 & "','" & ImporteTotalSinIVAConDtoL & "','" & ImporteTotalSinIVAConDtoLConDtoTotal & "','" & ImporteDtoTotalSinIVA & "',"
            sql = sql & "'" & ImporteTotalDelIVAConDtoLConDtoTotal & "','" & ImporteSinRedondeo & "','" & ImporteDelRedondeo & "','" & Format(fecha, "yyyy-mm-dd hh:nn:ss") & ".000',"
            sql = sql & "'" & numLineas & "','" & SerieLIdFinDeDia & "','" & SerieLTicketErroneo & "','" & Modificado & "',"
            sql = sql & "'" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000','" & EstadoTicket & "',"
            sql = sql & "'" & ImporteEntregado & "','" & ImporteDevuelto & "','" & PuntosFidelidad & "')"
            InformaMiss "CABECERA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "CABECERA!"
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "INSERTANDO CABECERA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sql
            End If
            ImporteTotal = 0
            ImporteTotal2 = 0
            numLineas = 0
            idTicketAnt = rsLin("idTicket")
        End If
        'Campos lineas
        import = rsLin("Importe")
        IdLineaTicket = rsLin("IdLineaTicket")
        idEmpresa = rsLin("idEmpresa")
        IdTienda = rsLin("IdTienda")
        idBalanzaMaestra = rsLin("IdBalanzaMaestra")
        idBalanzaEsclava = rsLin("IdBalanzaEsclava")
        idTicket = rsLin("IdTicket")
        tipoVenta = rsLin("TipoVenta")
        EstadoLinea = rsLin("EstadoLinea")
        IdArticulo = rsLin("IdArticulo")
        DESCRIPCIoN = rsLin("Descripcion")
        Descripcion1 = rsLin("Descripcion1")
        Comportamiento = rsLin("Comportamiento")
        Tara = rsLin("Tara")
        Peso = rsLin("Peso")
        PesoRegalado = rsLin("PesoRegalado")
        precio = rsLin("Precio")
        PrecioSinIva = rsLin("PrecioSinIva")
        PrecioConIVASinDtoL = rsLin("PrecioConIVASinDtoL")
        IdIVA = rsLin("IdIVA")
        PorcentajeIva = rsLin("PorcentajeIva")
        descuento = rsLin("descuento")
        ImporteSinIVASinDtoL = rsLin("ImporteSinIVASinDtoL")
        ImporteConIVASinDtoL = rsLin("ImporteConIVASinDtoL")
        ImporteDelDescuento = rsLin("ImporteDelDescuento")
        importe = rsLin("Importe")
        ImporteSinIVAConDtoL = rsLin("ImporteSinIVAConDtoL")
        ImporteDelIVAConDtoL = rsLin("ImporteDelIVAConDtoL")
        ImporteSinIVAConDtoLConDtoTotal = rsLin("ImporteSinIVAConDtoLConDtoTotal")
        ImporteDelIVAConDtoLConDtoTotal = rsLin("ImporteDelIVAConDtoLConDtoTotal")
        ImporteConDtoTotal = rsLin("ImporteConDtoTotal")
        TaraFija = rsLin("TaraFija")
        TaraPorcentual = rsLin("TaraPorcentual")
        CodInterno = rsLin("CodInterno")
        EANScannerArticulo = rsLin("EANScannerArticulo")
        IdClase = rsLin("IdClase")
        NombreClase = rsLin("NombreClase")
        IdElemAsociado = rsLin("IdElemAsociado")
        IdFamilia = rsLin("IdFamilia")
        NombreFamilia = rsLin("NombreFamilia")
        IdSeccion = rsLin("IdSeccion")
        nombreSeccion = rsLin("NombreSeccion")
        IdSubFamilia = rsLin("IdSubFamilia")
        NombreSubFamilia = rsLin("NombreSubFamilia")
        IdDepartamento = rsLin("IdDepartamento")
        NombreDepartamento = rsLin("NombreDepartamento")
        Modificado = rsLin("Modificado")
        Operacion = rsLin("Operacion")
        usuario = rsLin("Usuario")
        TimeStamp = rsLin("TimeStamp")
        'Campos cabecera
        NombreEmpresa = rsLin("NombreEmpresa")
        NombreTienda = rsLin("NombreTienda")
        tipo = rsLin("tipo")
        IdVendedor = rsLin("IdVendedor")
        NombreVendedor = rsLin("NombreVendedor")
        idCliente = rsLin("IdCliente")
        NombreCliente = rsLin("NombreCliente")
        DNICliente = rsLin("DNICliente")
        DireccionCliente = rsLin("DireccionCliente")
        PoblacionCliente = rsLin("PoblacionCliente")
        ProvinciaCliente = rsLin("ProvinciaCliente")
        CPCliente = rsLin("CPCliente")
        TelefonoCliente = rsLin("TelefonoCliente")
        ImporteLineas = rsLin("ImporteLineas")
        PorcDescuento = rsLin("PorcDescuento")
        ImporteDescuento = rsLin("ImporteDescuento")
        ImporteLineas2 = rsLin("ImporteLineas2")
        PorcDescuento2 = rsLin("PorcDescuento2")
        ImporteDescuento2 = rsLin("ImporteDescuento2")
        ImporteLineas3 = rsLin("ImporteLineas3")
        PorcDescuento3 = rsLin("PorcDescuento3")
        ImporteDescuento3 = rsLin("ImporteDescuento3")
        ImporteTotal3 = rsLin("ImporteTotal3")
        ImporteSinRedondeo = rsLin("ImporteSinRedondeo")
        ImporteDelRedondeo = rsLin("ImporteDelRedondeo")
        SerieLIdFinDeDia = rsLin("SerieLIdFinDeDia")
        SerieLTicketErroneo = rsLin("SerieLTicketErroneo")
        ImporteDtoTotalSinIVA = rsLin("ImporteDtoTotalSinIVA")
        fecha = rsLin("fecha")
        EstadoTicket = rsLin("EstadoTicket")
        ImporteDevuelto = rsLin("ImporteDevuelto")
        PuntosFidelidad = rsLin("PuntosFidelidad")
        NumTicket = rsLin("numTicket")
        'ImporteTotal en ImporteTotal y ImporteTotal2
        ImporteTotal = ImporteTotal + Round(importe, 2)
        ImporteTotal2 = ImporteTotal2 + Round(importe, 2)
        numLineas = numLineas + 1
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR LINEA!!! BORRADO DE LINEAS DE TICKET PARA EL NUMERO DE TICKET ACTUAL
'----  ES POSIBLE QUE ALGUNA VEZ SE HAYA QUEDADO LA LINEA SIN CABECERA
'----------------------------------------------------------------------------------------------------
        If IdLineaTicket = 1 Then
            sqlSP = "delete FROM dat_ticket_linea "
            sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
            sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
            sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and idTicket='" & idTicket & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
        End If
'--------------------------------------------------------------------------------
'--INSERTANDO LINEA
'--------------------------------------------------------------------------------
        'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_linea limit 1'') ( "
        sql = "INSERT INTO dat_ticket_linea ( "
        sql = sql & "IdLineaTicket,IdEmpresa,IdTienda,IdBalanzaMaestra,IdBalanzaEsclava,IdTicket,"
        sql = sql & "TipoVenta,EstadoLinea,IdArticulo,Descripcion,Descripcion1,Comportamiento,Tara,"
        sql = sql & "Peso,PesoRegalado,Precio,PrecioSinIVA,PrecioConIVASinDtoL,IdIVA,PorcentajeIVA,"
        sql = sql & "Descuento,ImporteSinIVASinDtoL,ImporteConIVASinDtoL,ImporteDelDescuento,Importe,"
        sql = sql & "ImporteSinIVAConDtoL,ImporteDelIVAConDtoL,ImporteSinIVAConDtoLConDtoTotal,"
        sql = sql & "ImporteDelIVAConDtoLConDtoTotal,ImporteConDtoTotal,TaraFija,TaraPorcentual,"
        sql = sql & "CodInterno,EANScannerArticulo,IdClase,NombreClase,IdElemAsociado,IdFamilia,"
        sql = sql & "NombreFamilia,IdSeccion,NombreSeccion,IdSubFamilia,NombreSubFamilia,IdDepartamento,"
        sql = sql & "NombreDepartamento,Modificado,Operacion,Usuario,TimeStamp) "
        sql = sql & " values ('" & IdLineaTicket & "','" & idEmpresa & "','" & IdTienda & "','"
        sql = sql & idBalanzaMaestra & "','" & idBalanzaEsclava & "','" & idTicket & "','"
        sql = sql & tipoVenta & "','" & EstadoLinea & "','" & IdArticulo & "','" & Replace(DESCRIPCIoN, "'", "''") & "','" & Replace(Descripcion1, "'", "''") & "','" & CInt(Comportamiento) & "','" & Tara & "','"
        sql = sql & Peso & "','" & PesoRegalado & "','" & precio & "','" & PrecioSinIva & "','"
        sql = sql & PrecioConIVASinDtoL & "','" & IdIVA & "','" & PorcentajeIva & "','" & descuento & "','" & ImporteSinIVASinDtoL & "','" & ImporteConIVASinDtoL & "','"
        sql = sql & ImporteDelDescuento & "',' " & importe & "',' " & ImporteSinIVAConDtoL & "','"
        sql = sql & ImporteDelIVAConDtoL & "','" & ImporteSinIVAConDtoLConDtoTotal & "','"
        sql = sql & ImporteDelIVAConDtoLConDtoTotal & "','" & ImporteConDtoTotal & "','" & TaraFija & "','" & TaraPorcentual & "','"
        sql = sql & CodInterno & "','" & EANScannerArticulo & "','" & IdClase & "','" & NombreClase & "','" & IdElemAsociado & "','" & IdFamilia & "','"
        sql = sql & NombreFamilia & "','" & IdSeccion & "','" & nombreSeccion & "','" & IdSubFamilia & "','" & NombreSubFamilia & "','" & IdDepartamento & "','"
        sql = sql & NombreDepartamento & "','" & Modificado & "','" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000') "
        InformaMiss "LINEA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "INSERTANDO LINEA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sql
        End If
        rsLin.MoveNext
'----------------------------------------------------------------------------------------------------
'----  ULTIMA CABECERA
'----------------------------------------------------------------------------------------------------
        If rsLin.EOF Then
            'Importe de ivas de los 3 tipos
            sql = "SELECT ("
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=1 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=1))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=2 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=2))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=3 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=3))/100,0),2) "
            sql = sql & ") as resultado "
            Set Rs = Db.OpenResultset(sql)
            If Not Rs.EOF Then
                ImporteTotalSinIVAConDtoL = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalSinIVAConDtoLConDtoTotal = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalDelIVAConDtoLConDtoTotal = Rs("resultado")
            End If
            ImporteEntregado = ImporteTotal
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR CABECERA!!! BORRADO DE CABECERA ACTUAL POR SI YA EXISTIERA
'----------------------------------------------------------------------------------------------------
            sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
            sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
            sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and idTicket='" & idTicket & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
            'SQL INSERT Cabecera MYSQL
            'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_cabecera limit 1'') "
            sql = "INSERT INTO dat_ticket_cabecera "
            sql = sql & "(IdTicket, NumTicket, IdEmpresa, NombreEmpresa,IdTienda,NombreTienda,IdBalanzaMaestra,"
            sql = sql & "NombreBalanzaMaestra, IdBalanzaEsclava, Tipo,"
            sql = sql & "IdVendedor,NombreVendedor,IdCliente,NombreCliente,"
            sql = sql & "DNICliente,DireccionCliente,PoblacionCliente,"
            sql = sql & "ProvinciaCliente,CPCliente,TelefonoCliente,"
            sql = sql & "TipoVenta,ImporteLineas,PorcDescuento,ImporteDescuento,"
            sql = sql & "ImporteTotal,ImporteLineas2,PorcDescuento2,ImporteDescuento2,"
            sql = sql & "ImporteTotal2,ImporteLineas3,PorcDescuento3,ImporteDescuento3,"
            sql = sql & "ImporteTotal3,ImporteTotalSinIVAConDtoL,ImporteTotalSinIVAConDtoLConDtoTotal,"
            sql = sql & "ImporteDtoTotalSinIVA,ImporteTotalDelIVAConDtoLConDtoTotal,ImporteSinRedondeo,"
            sql = sql & "ImporteDelRedondeo,Fecha,NumLineas,SerieLIdFinDeDia,SerieLTicketErroneo,Modificado,"
            sql = sql & "Operacion,Usuario,TimeStamp,EstadoTicket,ImporteEntregado,ImporteDevuelto,PuntosFidelidad) values ("
            sql = sql & "'" & idTicket & "','" & NumTicket & "','" & idEmpresa & "','" & NombreEmpresa & "',"
            sql = sql & "'" & IdTienda & "','" & NombreTienda & "','" & idBalanzaMaestra & "','-Balança " & idBalanzaMaestra & "',"
            sql = sql & "'" & idBalanzaEsclava & "','" & tipo & "','" & IdVendedor & "','" & NombreVendedor & "',"
            sql = sql & "'" & idCliente & "','" & NombreCliente & "','" & DNICliente & "','" & DireccionCliente & "',"
            sql = sql & "'" & PoblacionCliente & "','" & ProvinciaCliente & "','" & CPCliente & "','" & TelefonoCliente & "',"
            sql = sql & "'" & tipoVenta & "','" & ImporteLineas & "','" & PorcDescuento & "','" & ImporteDescuento & "',"
            sql = sql & "'" & ImporteTotal & "','" & ImporteLineas2 & "','" & PorcDescuento2 & "','" & ImporteDescuento2 & "',"
            sql = sql & "'" & ImporteTotal2 & "','" & ImporteLineas3 & "','" & PorcDescuento3 & "','" & ImporteDescuento3 & "',"
            sql = sql & "'" & ImporteTotal3 & "','" & ImporteTotalSinIVAConDtoL & "','" & ImporteTotalSinIVAConDtoLConDtoTotal & "','" & ImporteDtoTotalSinIVA & "',"
            sql = sql & "'" & ImporteTotalDelIVAConDtoLConDtoTotal & "','" & ImporteSinRedondeo & "','" & ImporteDelRedondeo & "','" & Format(fecha, "yyyy-mm-dd hh:nn:ss") & ".000',"
            sql = sql & "'" & numLineas & "','" & SerieLIdFinDeDia & "','" & SerieLTicketErroneo & "','" & Modificado & "',"
            sql = sql & "'" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000','" & EstadoTicket & "',"
            sql = sql & "'" & ImporteEntregado & "','" & ImporteDevuelto & "','" & PuntosFidelidad & "')"
            InformaMiss "CABECERA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "CABECERA!"
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "INSERTANDO CABECERA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sql
            End If
        End If
    Loop
'--------------------------------------------------------------------------------
'--SIGUIENTE TIENDA
'--------------------------------------------------------------------------------
    rsClients.MoveNext
    'Borramos tablas temp
    Db.OpenResultset ("DROP TABLE " & tablaTmp)
    Db.OpenResultset ("DROP TABLE " & tablaTmp2)
    Db.OpenResultset ("DROP TABLE " & tablaTmp3)
    Db.OpenResultset ("DROP TABLE " & tabla)
    html = html & "<p><b>Botiga: </b>" & codiBotiga & "</p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & desde & "</p>"
    html = html & "<p><b>Finalitzat: </b>" & Now() & "</p>"
Loop
If P3 <> "" Then sf_enviarMail "secrehit@hit.cat", Desti, " Vendes sincronitzades " & codiBotiga, html, "", ""

connMysql.Close
Set connMysql = Nothing

'Update de feinesafer nit para marcar cuando acaba proceso
If P4 = "Nit" Then
    sql = "Select count(id) num from feinesafer where tipus='SincroDbVendesAmetller2' "
    Set Rs = Db.OpenResultset(sql)
    If Not Rs.EOF Then
        If Rs("num") <= 1 Then
            sql = "Update feinesafer set Param3='Fi " & Now() & "' where tipus='SincroDbVendesAmetllerNit' and param2 like '%Si%' "
            Set Rs = Db.OpenResultset(sql)
        End If
    End If
End If

InformaMiss "FIN SINCRO_VENDES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "FIN SINCRO_VENDES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.Close
End If
Exit Function

norVendes:
'----------------------------------------------------------------------------------------------------
'----  SI HAY ERROR: BORRADO ULTIMAS LINEAS DE TICKET HUERFANAS SIN CABECERA
'----------------------------------------------------------------------------------------------------
    html = "<p><h3>Resum Vendes Ametller </h3></p>"
    html = html & "<p><b>Botiga: </b>" & botiguesCad & "</p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & desde & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    If P3 <> "" Then sf_enviarMail "secrehit@hit.cat", Desti, "ERROR! Sincronitzacio de vendes de " & codiBotiga & " ha fallat", html, "", ""
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    html = html & "<p><b>Ultima sqlSP:</b>" & sqlSP & "</p>"
    
    sf_enviarMail "secrehit@hit.cat", EmailGuardia, "ERROR! Sincronitzacio de vendes de " & codiBotiga & " ha fallat", html, "", ""
        
    Set connMysql = New ADODB.Connection
    connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
    connMysql.ConnectionTimeout = 1000 '16 min
    connMysql.Open

    sqlSP = "delete FROM dat_ticket_linea "
    sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
    sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
    sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
    sqlSP = sqlSP & "and idTicket='" & idTicket & "' "
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "Error:" & err.Description & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        connMysql.Execute sqlSP
    End If
    
    connMysql.Close
    Set connMysql = Nothing
    
    'Borramos tablas temporales
    Db.OpenResultset ("DROP TABLE " & tablaTmp3)
    Db.OpenResultset ("DROP TABLE " & tablaTmp2)
    Db.OpenResultset ("DROP TABLE " & tablaTmp)
    Db.OpenResultset ("DROP TABLE " & tabla)
    
    ExecutaComandaSql "Delete from  " & TaulaCalculsEspecials & " Id = '" & idTasca & "' "
    
    Informa "Error : " & err.Description
    
End Function



Function SincroDbVendesAmetller(p1, P2, P3, P4, idTasca) As Boolean
    Dim botiguesCad As String, mes As String, anyo As String, desde As String, debugSincro As Boolean
    Dim codiBotiga As String, codiBotigaextern As String, tabla As String, fecha, fecha_caracter
    Dim sql As String, sql2 As String, sql3 As String, sqlSP As String, numCab As Integer, maxIdTicket As String
    Dim parametros As String, tablaTmp As String, Rs As rdoResultset, rsClients As rdoResultset, pos As Integer
    Dim rsCab As rdoResultset, rsLin As rdoResultset, html As String, hasta As String, desde0, tablaTmp2 As String, tablaTmp3 As String
    Dim connMysql As ADODB.Connection

'INSERTA ALBARANS I VENDES
'INSERTA MERMES COM A VENTES AMB IDVENEDOR ESPECIAL 09/08/2012
'IDENTIFICA CLIENTS TANT DE TAULA ALBARANS COM DE TAULA VENUTS 28/09/2012

'If P1 <> "9011" Then Exit Function

On Error GoTo norVendes
'Parametros
'Si pasamos parametros vacios se recorreran todas las tiendas de paramsHw
',mirando ventas (fechas actuales), se eliminaran datos en servidor remoto
'si se le indica una fecha desde y se ejecutaran inserciones (cabeceras y lineas)

'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
botiguesCad = p1 'Cadena de tiendas separadas por ,
If botiguesCad <> "" Then
    botiguesCad = Replace(botiguesCad, "[", "")
    botiguesCad = Replace(botiguesCad, "]", "")
End If
desde = P2 'Fecha desde la que eliminan datos en tablas ventas de serv. remoto
If desde <> "" Then
    desde = Replace(desde, "[", "")
    desde = Replace(desde, "]", "")
End If
desde0 = desde
mes = Month(desde) 'Mes para tabla ventas
If mes <> "" Then If Len(mes) = 1 Then mes = "0" & mes
anyo = Year(desde) 'Año para tabla ventas
Desti = P3 'Email
'desde = CDate(desde)
'hasta = P4 'Fecha desde la que eliminan datos en tablas ventas de serv. remoto
'If hasta <> "" Then
'    hasta = Replace(hasta, "[", "")
'    hasta = Replace(hasta, "]", "")
'End If
debugSincro = False 'Solo impresion o ejecucion completa
'Crear txt log
If debugSincro = True Then
    Desti = EmailGuardia
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\Vendes_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
End If
'--------------------------------------------------------------
numCab = 0
If mes = "" Then
    mes = Month(Now)
    If Len(mes) = 1 Then mes = "0" & mes
End If
If anyo = "" Then anyo = Year(Now)
InformaMiss "INICIO SINCRO_VENDES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "INICIO SINCRO_VENDES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'--UPDATE DEPENDENTES ANTIGUES
'--------------------------------------------------------------------------------
'   ANTIGUO!
'    Sql = "select dbk.CODI codiantic,dbk.NOM nomantic,d.CODI codinou,d.NOM nomnou,"
'    Sql = Sql & "'update [Fac_laforneria].[dbo].[V_Venut_" & anyo & "-" & mes & "] set dependenta='''+cast(d.codi as nvarchar(25))+''' where dependenta='''+cast(dbk.CODI as nvarchar(25))+''' ' upd "
'    Sql = Sql & "FROM [Fac_laforneria].[dbo].[V_Venut_" & anyo & "-" & mes & "] v left join dependentesbk dbk on v.Dependenta=dbk.CODI "
'    Sql = Sql & "left join Dependentes d on dbk.NOM=d.NOM  "
'    Sql = Sql & "where dbk.CODI is not null and d.codi is not null group by dbk.CODI,dbk.nom,d.CODI,d.NOM order by dbk.CODI,d.CODI"
'    Set Rs = Db.OpenResultset(Sql)
'    Do While Not Rs.EOF
'        Sql = Rs("upd")
'        ExecutaComandaSql Sql
'        Rs.MoveNext
'    Loop
'    Sql = "select dbk.CODI codiantic,dbk.NOM nomantic,d.CODI codinou,d.NOM nomnou,"
'    Sql = Sql & "'update [Fac_laforneria].[dbo].[V_Albarans_" & anyo & "-" & mes & "] set dependenta='''+cast(d.codi as nvarchar(25))+''' where dependenta='''+cast(dbk.CODI as nvarchar(25))+''' ' upd "
'    Sql = Sql & "FROM [Fac_laforneria].[dbo].[V_Albarans_" & anyo & "-" & mes & "] v left join dependentesbk dbk on v.Dependenta=dbk.CODI "
'    Sql = Sql & "left join Dependentes d on dbk.NOM=d.NOM  "
'    Sql = Sql & "where dbk.CODI is not null and d.codi is not null group by dbk.CODI,dbk.nom,d.CODI,d.NOM order by dbk.CODI,d.CODI"
'    Set Rs = Db.OpenResultset(Sql)
'    Do While Not Rs.EOF
'        Sql = Rs("upd")
'        ExecutaComandaSql Sql
'        Rs.MoveNext
'    Loop

    'Compara codigos de dep. sin variable CODI_DEP y busca si existen codigos de dep. con esa variable. Compara a traves de nombre
    'Tabla Venut
    sql = "select t.codiAnt,t.nom1,t.codiNou,de.valor,"
    sql = sql & "'update [Fac_laforneria].[dbo].[V_Venut_" & anyo & "-" & mes & "] set dependenta='''+cast(t.codiNou as nvarchar(25))+''' where dependenta='''+cast(t.codiAnt as nvarchar(25))+''' ' upd "
    sql = sql & "From (select v.dependenta codiAnt,d.NOM nom1,d2.codi codiNou,d2.NOM nom2 "
    sql = sql & "from [Fac_laforneria].[dbo].[V_Venut_" & anyo & "-" & mes & "] v "
    sql = sql & "left join Dependentes d on v.Dependenta=d.CODI "
    sql = sql & "left join dependentes d2 on d.NOM=d2.NOM "
    sql = sql & "where d.codi not in (select id from dependentesExtes where nom='CODI_DEP') "
    sql = sql & "and d2.CODI in (select id from dependentesExtes where nom='CODI_DEP') "
    sql = sql & "group by v.Dependenta,d.NOM,d2.CODI,d2.NOM) t "
    sql = sql & "left join dependentesextes de on t.codiNou=de.id where de.nom='CODI_DEP' "
    Set Rs = Db.OpenResultset(sql)
    Do While Not Rs.EOF
       sql = Rs("upd")
       ExecutaComandaSql sql
       Rs.MoveNext
    Loop
    
    'Tabla Albarans
    sql = "select t.codiAnt,t.nom1,t.codiNou,de.valor,"
    sql = sql & "'update [Fac_laforneria].[dbo].[V_Albarans_" & anyo & "-" & mes & "] set dependenta='''+cast(t.codiNou as nvarchar(25))+''' where dependenta='''+cast(t.codiAnt as nvarchar(25))+''' ' upd "
    sql = sql & "From (select v.dependenta codiAnt,d.NOM nom1,d2.codi codiNou,d2.NOM nom2 "
    sql = sql & "from [Fac_laforneria].[dbo].[V_Albarans_" & anyo & "-" & mes & "] v "
    sql = sql & "left join Dependentes d on v.Dependenta=d.CODI "
    sql = sql & "left join dependentes d2 on d.NOM=d2.NOM "
    sql = sql & "where d.codi not in (select id from dependentesExtes where nom='CODI_DEP') "
    sql = sql & "and d2.CODI in (select id from dependentesExtes where nom='CODI_DEP') "
    sql = sql & "group by v.Dependenta,d.NOM,d2.CODI,d2.NOM) t "
    sql = sql & "left join dependentesextes de on t.codiNou=de.id where de.nom='CODI_DEP' "
    Set Rs = Db.OpenResultset(sql)
    Do While Not Rs.EOF
        sql = Rs("upd")
        ExecutaComandaSql sql
        Rs.MoveNext
    Loop
'--------------------------------------------------------------------------------
'--TABLAS TEMPORALES
'--------------------------------------------------------------------------------
'Creamos tablas temporales de las quales podemos obtener datos de familia, secciones, etc
InformaMiss "CREANDO TABLAS TEMPORALES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CREANDO TABLAS TEMPORALES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
nTmp = Now
tablaTmp = "[Fac_laforneria].[dbo].[sincro_vendesTmpArticles_" & botiguesCad & "_" & nTmp & "]"
tablaTmp2 = "[Fac_laforneria].[dbo].[sincro_vendesTmpClients_" & botiguesCad & "_" & nTmp & "]"
tablaTmp3 = "[Fac_laforneria].[dbo].[sincro_vendesTmpMermes_" & botiguesCad & "_" & nTmp & "]"
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmpArticles la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp)
sql = " SELECT * INTO " & tablaTmp & " FROM OPENQUERY(AMETLLER,'"
sql = sql & "SELECT a.idArticulo,a.Descripcion,a.Descripcion1,a.idFamilia,f.NombreFamilia,a.idSubFamilia,"
sql = sql & "sf.NombreSubFamilia,a.idDepartamento,d.NombreDepartamento,a.idSeccion,s.NombreSeccion,"
sql = sql & "a.PrecioSinIVA , a.PrecioConIva, i.IdIVA, i.PorcentajeIva "
sql = sql & "FROM dat_articulo a "
sql = sql & "LEFT JOIN dat_familia f on (a.idFamilia=f.idFamilia) "
sql = sql & "LEFT JOIN dat_subfamilia sf on a.idFamilia=sf.idFamilia and a.idsubfamilia = sf.idsubfamilia "
sql = sql & "LEFT JOIN dat_departamento d on (a.idDepartamento=d.idDepartamento) "
sql = sql & "LEFT JOIN dat_seccion s on (a.idSeccion=s.idSeccion) "
sql = sql & "LEFT JOIN dat_iva i on (a.idIva=i.idIva) "
sql = sql & "Where a.idEmpresa = 1 And IdArticulo < 90000') "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmpClients la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp2 & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp2)
sql = " SELECT * INTO " & tablaTmp2 & " FROM OPENQUERY(AMETLLER,'"
sql = sql & "SELECT idCliente, Nombre,TRIM(DNI) AS Cif, TRIM(Direccion) AS Dir,"
sql = sql & "TRIM(CodPostal) AS CP, TRIM(Poblacion) AS Ciutat,TimeStamp AS fecMod "
sql = sql & "FROM dat_cliente WHERE IdEmpresa=1') "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmpMermes la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp3 & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp3)
diff = DateDiff("D", Now, desde)
If diff < 0 Then diff = Abs(diff)
sql = "SELECT Botiga,Data,Dependenta,Num_tick,Estat,Plu,Quantitat,Import,Tipus_venta,FormaMarcar,Otros INTO  " & tablaTmp3 & " From ( "
For Z = 0 To diff
    sql = sql & "SELECT Client as Botiga,[TimeStamp] as Data,'9999' as Dependenta,"
    sql = sql & "DENSE_RANK() OVER (ORDER BY s.[timestamp],s.client) as Num_tick,"
    sql = sql & "Client as Estat,PluUtilitzat as Plu,QuantitatTornada as Quantitat,"
    sql = sql & "0 as Import,'V' as Tipus_venta,'0' as FormaMarcar,'9999' as Otros "
    sql = sql & "FROM [Fac_laforneria].[dbo].[" & DonamNomTaulaServit(CDate(DateAdd("D", Z, desde))) & "] s "
    sql = sql & "LEFT JOIN articles a on s.pluUtilitzat=a.Codi where quantitatTornada<>'0'"
    If Z < diff Then sql = sql & " UNION ALL "
Next
sql = sql & ") t ORDER BY Data,Botiga "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
html = "<p><h3>Resum Vendes Ametller</h3></p>"
'--------------------------------------------------------------------------------
'--CURSOR TIENDAS
'--------------------------------------------------------------------------------
'--Creamos cursor para tiendas
InformaMiss "CREANDO CURSOR TIENDAS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CREANDO CURSOR TIENDAS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
sql = "SELECT p.valor1 FROM [Fac_laforneria].[dbo].Clients c LEFT JOIN [Fac_laforneria].[dbo].ParamsHw p "
sql = sql & "ON (c.codi=p.valor1) WHERE p.valor1 is not null "
If botiguesCad <> "" Then sql = sql & "and p.valor1 IN ( " & botiguesCad & ")  "
sql = sql & " order by p.valor1 "
Set rsClients = Db.OpenResultset(sql)
Do While Not rsClients.EOF
    codiBotiga = rsClients("valor1")
    If codiBotiga = "518" Then
        codiBotigaextern = 1061
    Else
        codiBotigaextern = codiBotiga
    End If
    pos = Len(codiBotigaextern)
    If pos > 3 Then
        pos = 3
    Else
        pos = 2
    End If
    InformaMiss "TRATANDO TIENDA " & codiBotiga & "(" & codiBotigaextern & "),M:" & mes & ",A:" & anyo & "-->" & Now
    Informa2 "T:" & codiBotiga & "(" & codiBotigaextern & "),DESDE:" & desde
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "TRATANDO TIENDA " & codiBotiga & "(" & codiBotigaextern & "),M:" & mes & ",A:" & anyo & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'--------------------------------------------------------------------------------
'--ACTUALIZACION DE DUPLICADOS. TIQUETS CON MAS DE UNA FECHA DIFERENTE QUE AL AGRUPARSE
'--POR NUMERO DE TIQUET Y FECHA, GENERA DOS LINEAS PARA ESE MISMO NUMERO DE TIQUET
'--Y A LA HORA DE IMPORTAR A MYSQL DA ERRORES.
'--------------------------------------------------------------------------------
'--Si existe sincro_duplicadosTmp la borramos y volvemos a generar
    'Sql = "SELECT object_id FROM sys.objects with (nolock) WHERE name='sincro_duplicadosTmp' AND type='U' "
    'Set Rs = Db.OpenResultset(Sql)
    'If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE sincro_duplicadosTmp")
'--Si existe sincro_duplicadosTmp2 la borramos y volvemos a generar
    'Sql = "SELECT object_id FROM sys.objects with (nolock) WHERE name='sincro_duplicadosTmp2' AND type='U' "
    'Set Rs = Db.OpenResultset(Sql)
    'If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE sincro_duplicadosTmp2")
    'Sql = "SELECT data,Num_tick,botiga INTO sincro_duplicadosTmp FROM " & tabla & " GROUP BY data,Num_tick,Botiga "
    'Set Rs = Db.OpenResultset(Sql)
    'Sql = "SELECT Num_tick,Botiga INTO sincro_duplicadosTmp2 FROM sincro_duplicadosTmp GROUP BY Num_tick,Botiga "
    'Sql = Sql & "Having Count(Num_tick) >= 2 "
    'Set Rs = Db.OpenResultset(Sql)
    'Sql = "UPDATE " & tabla & " SET Data=v.data FROM ( "
    'Sql = Sql & "SELECT MIN(data) data,Num_tick,botiga FROM " & tabla & " "
    'Sql = Sql & "WHERE Num_tick IN(SELECT Num_tick FROM sincro_duplicadosTmp2 s WHERE " & tabla & ".Botiga=s.botiga) "
    'Sql = Sql & "AND botiga IN (SELECT botiga FROM sincro_duplicadosTmp2 s WHERE " & tabla & ".Botiga=s.botiga) "
    'Sql = Sql & "GROUP BY Num_tick,botiga "
    'Sql = Sql & ") v "
    'Sql = Sql & "Where " & tabla & ".Num_tick=v.Num_tick "
    'Sql = Sql & "AND " & tabla & ".Botiga=v.Botiga "
    'Set Rs = Db.OpenResultset(Sql)
    'If debugSincro = True Then
    '    Txt.WriteLine "--------------------------------------------------------------------------------"
    '    Txt.WriteLine "SQL:" & Sql & "-->" & Now
    '    Txt.WriteLine "--------------------------------------------------------------------------------"
    'End If
'--------------------------------------------------------------------------------
'--BORRADO DE DATOS EN SERV.REMOTO
'--CUIDADO! NO ELIMINA LAS LINEAS QUE SE HAYAN PODIDO QUEDAR SIN CABECERA!!!
'--------------------------------------------------------------------------------
    InformaMiss "BORRADO DE DATOS EN SERV.REMOTO DESDE " & desde & "-->" & Now
    If desde <> "" Then
        desde = Year(desde) & "-" & Month(desde) & "-" & Day(desde) & " " & Hour(desde) & ":" & Minute(desde) & ":" & Second(desde) & ".000"
'----------------------------------------------------------------------------------------------------
'---- LINEAS DE TICKET
'----------------------------------------------------------------------------------------------------
        sqlSP = "delete FROM dat_ticket_linea "
        sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
        sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
        sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' and idTicket "
        sqlSP = sqlSP & "in (select idTicket from dat_ticket_cabecera where idEmpresa=1 "
        sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' and idBalanzaEsclava=-1 "
        sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
        sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and timestamp>='" & desde & "' "
        'If hasta <> "" Then sqlSP = sqlSP & " and timestamp<='" & hasta & "' "
        sqlSP = sqlSP & ")"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sqlSP
        End If
'----------------------------------------------------------------------------------------------------
'---- CABECERAS DE TICKET
'----------------------------------------------------------------------------------------------------
        sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
        sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
        sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
        sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and timestamp>'" & desde & "' "
        'If hasta <> "" Then sqlSP = sqlSP & " and timestamp<='" & hasta & "' "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sqlSP
        End If
    End If
'--------------------------------------------------------------------------------
'--MAX IDTICKET
'--------------------------------------------------------------------------------
    InformaMiss "OBTENIENDO ULTIMO NUMERO TICKET-->" & Now
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "OBTENIENDO ULTIMO NUMERO TICKET-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'----------------------------------------------------------------------------------------------------
'---- SACAMOS EL IDTICKET MAXIMO Y LA FECHA MAXIMA DE LA TIENDA
'----------------------------------------------------------------------------------------------------
    sql = "select isNull(IdTicket,0) idTicket,isNull(fecha,'" & anyo & "-" & mes & "-01 00:00:01.000') fecha FROM openquery(AMETLLER, "
    sql = sql & "'SELECT MAX(idTicket)as IdTicket,MAX(TimeStamp)as fecha From dat_ticket_cabecera "
    sql = sql & " WHERE IdEmpresa=1 AND IdTienda=''" & Left(codiBotigaextern, pos) & "'' "
    sql = sql & "AND IdBalanzaMaestra=''" & Right(codiBotigaextern, 1) & "'' "
    sql = sql & "GROUP BY IdEmpresa,IdTienda,IdBalanzaMaestra') "
    Set Rs = Db.OpenResultset(sql)
    maxIdTicket = 0
    fecha = desde0
    fecha_caracter = desde0
    If Not Rs.EOF Then
        maxIdTicket = Rs("idTicket")
        fecha = Rs("fecha")
        'LA SQL NO FUNCIONA CON ESTE TIPO FECHA! fecha_caracter = Left(fecha, 4) & "-" & Mid(fecha, 6, 2) & "-" & Mid(fecha, 9, 2) & " " & Mid(fecha, 12, 2) & ":" & Mid(fecha, 15, 2) & ":" & Mid(fecha, 18, 2) & ".000"
        fecha_caracter = Mid(fecha, 9, 2) & "-" & Mid(fecha, 6, 2) & "-" & Left(fecha, 4) & " " & Mid(fecha, 12, 2) & ":" & Mid(fecha, 15, 2) & ":" & Mid(fecha, 18, 2) & ".000"
    End If
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
'--------------------------------------------------------------------------------
'--TABLA TEMPORAL DE VENTAS Y ALBARANES
'--------------------------------------------------------------------------------
    tabla = "[Fac_laforneria].[dbo].[sincro_vendesTmpVendes_" & botiguesCad & "_" & nTmp & "]"
    'Si existe sincro_vendesTmp2 la borramos y volvemos a generar
    sql = "SELECT object_id FROM sys.objects with (nolock) "
    sql = sql & "WHERE name=REPLACE(REPLACE('" & tabla & "','[',''),']','') AND type='U' "
    Set Rs = Db.OpenResultset(sql)
    If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tabla)
    sql = "SELECT * INTO " & tabla & " FROM ("
    sql = sql & "SELECT * FROM [Fac_laforneria].[dbo].[V_Venut_" & anyo & "-" & mes & "] "
    sql = sql & "WHERE data>'" & fecha_caracter & "' and botiga='" & botiguesCad & "' "
    sql = sql & "UNION ALL "
    sql = sql & "SELECT * FROM [Fac_laforneria].[dbo].[V_Albarans_" & anyo & "-" & mes & "] "
    sql = sql & "WHERE data>'" & fecha_caracter & "' and botiga='" & botiguesCad & "' "
    sql = sql & "UNION ALL "
    sql = sql & "SELECT * FROM " & tablaTmp3
    sql = sql & "WHERE data>'" & fecha_caracter & "' and botiga='" & botiguesCad & "' "
    sql = sql & ") t "
    sql = sql & "WHERE botiga='" & botiguesCad & "'"
    Db.QueryTimeout = 0
    Set Rs = Db.OpenResultset(sql)
    Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
'--LINEAS Y CABECERAS
'--------------------------------------------------------------------------------
    InformaMiss "INSERTANDO LINEAS Y CABECERAS-->" & Now
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "INSERTANDO LINEAS Y CABECERAS-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
    ImporteTotal = 0
    ImporteTotal2 = 0
    numLineas = 0
    'SQL ventas SQL SERVER
    sql = "SELECT ROW_NUMBER() OVER (PARTITION BY ven.num_tick,ven.data ORDER BY ven.data,ven.num_tick,ven.Plu,ven.Quantitat,ven.Import) AS IdLineaTicket, "
    sql = sql & "1 AS IdEmpresa,LEFT('" & codiBotigaextern & "'," & pos & ") AS IdTienda,RIGHT('" & codiBotigaextern & "',1) AS IdBalanzaMaestra, "
    sql = sql & "-1 AS IdBalanzaEsclava,DENSE_RANK() OVER (ORDER BY ven.num_tick,ven.data) + '" & maxIdTicket & "'  AS IdTicket,2 AS TipoVenta,0 AS EstadoLinea, "
    sql = sql & "isNull(tmp.idArticulo,'89992') AS IdArticulo,isNull(tmp.Descripcion,'Preu Directe Hit') AS Descripcion, "
    sql = sql & "isNull(tmp.Descripcion1,'Preu Directe Hit') AS Descripcion1,isNull(art.EsSumable,1) AS Comportamiento,0 AS Tara, "
    sql = sql & "ven.Quantitat AS Peso,NULL AS PesoRegalado,round(ven.import/ven.quantitat,2) AS Precio,0 AS PrecioSinIVA, "
    sql = sql & "round(ven.import/ven.quantitat,2) AS PrecioConIVASinDtoL,isNull(tmp.idIva,0) AS IdIVA, "
    sql = sql & "isNull(tmp.PorcentajeIVA,0) AS PorcentajeIVA,NULL AS Descuento, "
    sql = sql & "NULL AS ImporteSinIVASinDtoL,NULL AS ImporteConIVASinDtoL,NULL AS ImporteDelDescuento, "
    sql = sql & "isNull(ROUND(ven.Import,2),0) AS Importe, "
    sql = sql & "isNull(ROUND(CAST(ven.Import AS FLOAT)-((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100),2),0) AS ImporteSinIVAConDtoL, "
    sql = sql & "isNull(ROUND((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100,2),0) AS ImporteDelIVAConDtoL, "
    sql = sql & "isNull(ROUND(CAST(ven.Import AS FLOAT)-((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100),2),0) AS ImporteSinIVAConDtoLConDtoTotal, "
    sql = sql & "isNull(ROUND((CAST(ven.Import AS FLOAT)*CAST(iva.Iva AS INTEGER))/100,2),0) AS ImporteDelIVAConDtoLConDtoTotal, "
    sql = sql & "isNull(ROUND(ven.Import,2),0) AS ImporteConDtoTotal,0 AS TaraFija,0 AS TaraPorcentual,isNull(tmp.idArticulo,'89992') AS CodInterno, "
    sql = sql & "'' as EANScannerArticulo,0 AS IdClase,NULL AS NombreClase,0 AS IdElemAsociado, "
    sql = sql & "isNull(tmp.IdFamilia,0) AS IdFamilia,isNull(tmp.NombreFamilia,'Familia no asignada') AS NombreFamilia, "
    sql = sql & "isNull(tmp.IdSeccion,0) AS IdSeccion, "
    sql = sql & "isNull(tmp.NombreSeccion,'Seccio no asignada') AS NombreSeccion,isNull(tmp.IdSubFamilia,0) AS IdSubFamilia, "
    sql = sql & "isNull(tmp.NombreSubFamilia,'Subfamilia no asignada') AS NombreSubFamilia,isNull(tmp.IdDepartamento,0) AS IdDepartamento, "
    sql = sql & "isNull(tmp.NombreDepartamento,'Departament no asignat') AS NombreDepartamento,1 AS Modificado,'A' AS Operacion, "
    sql = sql & "'Comunicaciones' AS Usuario,ven.Data AS TimeStamp, "
    sql = sql & "'Casa Ametller S.L.' AS NombreEmpresa,CASE WHEN CHARINDEX('_',cli.nom)>0 THEN SUBSTRING(cli.nom,1,CAST(CHARINDEX('_',cli.nom)AS INTEGER)-1)  "
    sql = sql & "ELSE cli.Nom END AS NombreTienda,'T' AS Tipo,CASE WHEN dep.nom='Merma' THEN '17' ELSE isNull(dE.valor,'0000') END AS IdVendedor,isNull(dep.nom,'Dependenta sense asignar') AS NombreVendedor, "
    sql = sql & "CASE WHEN ven.Otros like '%CliBoti%' THEN isNull(tmp4.idCliente,'0') ELSE isNull(tmp2.idCliente,'0') END AS IdCliente, "
    sql = sql & "CASE WHEN ven.Otros like '%CliBoti%' THEN isNull(tmp4.Nombre,'0') ELSE isNull(tmp2.Nombre,'0') END AS NombreCliente, "
    sql = sql & "CASE WHEN ven.Otros like '%CliBoti%' THEN isNull(tmp4.Cif,'0') ELSE isNull(tmp2.Cif,'0') END AS DNICliente, "
    sql = sql & "CASE WHEN ven.Otros like '%CliBoti%' THEN isNull(tmp4.Dir,'0') ELSE isNull(tmp2.Dir,'0') END AS DireccionCliente, "
    sql = sql & "CASE WHEN ven.Otros like '%CliBoti%' THEN isNull(tmp4.Ciutat,'0') ELSE isNull(tmp2.Ciutat,'0') END AS PoblacionCliente,NULL AS ProvinciaCliente, "
    sql = sql & "CASE WHEN ven.Otros like '%CliBoti%' THEN isNull(tmp4.CP,'0') ELSE isNull(tmp2.CP,'0') END AS CPCliente, "
    sql = sql & "0 AS TelefonoCliente,NULL AS ImporteLineas,NULL AS PorcDescuento,NULL AS ImporteDescuento, "
    sql = sql & "NULL AS ImporteLineas2,NULL AS PorcDescuento2,NULL AS ImporteDescuento2, "
    sql = sql & "NULL AS ImporteLineas3,NULL AS PorcDescuento3,NULL AS ImporteDescuento3, "
    sql = sql & "NULL AS ImporteTotal3,NULL AS ImporteSinRedondeo,NULL AS ImporteDelRedondeo, "
    sql = sql & "NULL AS SerieLIdFinDeDia,0 AS SerieLTicketErroneo,0 AS ImporteDtoTotalSinIVA, "
    sql = sql & "ven.Data AS Fecha,'C' AS EstadoTicket,0 AS ImporteDevuelto,0 AS PuntosFidelidad,ven.num_tick AS NumTicket "
    sql = sql & "FROM " & tabla & " ven "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[clients] cli ON (ven.Botiga=cli.codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[dependentes] dep ON (ven.dependenta=dep.codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[dependentesExtes] dE on (dep.CODI=dE.id and dE.nom='CODI_DEP') "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[Articles] art ON (ven.Plu=art.Codi) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[ArticlesPropietats] artP1 ON (art.Codi=artP1.CodiArticle and artP1.Variable='CODI_PROD') "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[TipusIva] iva ON (art.TipoIva=iva.Tipus) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[ConstantsClient] cc ON (CAST(ven.otros as nvarchar(10))=CAST(cc.codi as nvarchar(10)) and cc.Variable='CodiClientOrigen' ) "
    sql = sql & "LEFT JOIN " & tablaTmp & " tmp ON (artP1.valor=tmp.IdArticulo) "
    sql = sql & "LEFT JOIN " & tablaTmp2 & " tmp2 ON (cc.valor=tmp2.IdCliente) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[clientsFinals] tmp3 "
    sql = sql & "ON (REPLACE(REPLACE(SUBSTRING(ven.Otros,CHARINDEX('CliBoti',ven.Otros),LEN(ven.otros)),']',''),']','')=tmp3.Id) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[clients] cli2 ON (tmp3.nom=cli2.Nom) "
    sql = sql & "LEFT JOIN [Fac_laforneria].[dbo].[ConstantsClient] cc2 ON (cli2.Codi=cc2.Codi and cc2.Variable='CodiClientOrigen') "
    sql = sql & "LEFT JOIN " & tablaTmp2 & " tmp4 ON (cc2.valor=tmp4.IdCliente) "
    sql = sql & "WHERE ven.Botiga='" & codiBotiga & "' and ven.data>'" & fecha_caracter & "' ORDER BY ven.data "
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    End If
    Set rsLin = Db.OpenResultset(sql)
    If Not rsLin.EOF Then
        idTicketAnt = rsLin("idTicket")
    End If
    Do While Not rsLin.EOF
'--------------------------------------------------------------------------------
'--INSERTANDO CABECERA
'--------------------------------------------------------------------------------
        If rsLin("IdTicket") <> idTicketAnt Then
            'Importe de ivas de los 3 tipos
            sql = "SELECT ("
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=1 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=1))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=2 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=2))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=3 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=3))/100,0),2) "
            sql = sql & ") as resultado "
            Set Rs = Db.OpenResultset(sql)
            If Not Rs.EOF Then
                ImporteTotalSinIVAConDtoL = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalSinIVAConDtoLConDtoTotal = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalDelIVAConDtoLConDtoTotal = Rs("resultado")
            End If
            ImporteEntregado = ImporteTotal
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR CABECERA!!! BORRADO DE CABECERA ACTUAL POR SI YA EXISTIERA
'----------------------------------------------------------------------------------------------------
            sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
            sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
            sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and idTicket='" & idTicketAnt & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
            'SQL INSERT Cabecera MYSQL
            'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_cabecera limit 1'') "
            sql = "INSERT INTO dat_ticket_cabecera "
            sql = sql & "(IdTicket, NumTicket, IdEmpresa, NombreEmpresa,IdTienda,NombreTienda,IdBalanzaMaestra,"
            sql = sql & "NombreBalanzaMaestra, IdBalanzaEsclava, Tipo,"
            sql = sql & "IdVendedor,NombreVendedor,IdCliente,NombreCliente,"
            sql = sql & "DNICliente,DireccionCliente,PoblacionCliente,"
            sql = sql & "ProvinciaCliente,CPCliente,TelefonoCliente,"
            sql = sql & "TipoVenta,ImporteLineas,PorcDescuento,ImporteDescuento,"
            sql = sql & "ImporteTotal,ImporteLineas2,PorcDescuento2,ImporteDescuento2,"
            sql = sql & "ImporteTotal2,ImporteLineas3,PorcDescuento3,ImporteDescuento3,"
            sql = sql & "ImporteTotal3,ImporteTotalSinIVAConDtoL,ImporteTotalSinIVAConDtoLConDtoTotal,"
            sql = sql & "ImporteDtoTotalSinIVA,ImporteTotalDelIVAConDtoLConDtoTotal,ImporteSinRedondeo,"
            sql = sql & "ImporteDelRedondeo,Fecha,NumLineas,SerieLIdFinDeDia,SerieLTicketErroneo,Modificado,"
            sql = sql & "Operacion,Usuario,TimeStamp,EstadoTicket,ImporteEntregado,ImporteDevuelto,PuntosFidelidad) values ("
            sql = sql & "'" & idTicketAnt & "','" & NumTicket & "','" & idEmpresa & "','" & NombreEmpresa & "',"
            sql = sql & "'" & IdTienda & "','" & NombreTienda & "','" & idBalanzaMaestra & "','-Balança " & idBalanzaMaestra & "',"
            sql = sql & "'" & idBalanzaEsclava & "','" & tipo & "','" & IdVendedor & "','" & Replace(NombreVendedor, "'", "''") & "',"
            sql = sql & "'" & idCliente & "','" & NombreCliente & "','" & DNICliente & "','" & DireccionCliente & "',"
            sql = sql & "'" & PoblacionCliente & "','" & ProvinciaCliente & "','" & CPCliente & "','" & TelefonoCliente & "',"
            sql = sql & "'" & tipoVenta & "','" & ImporteLineas & "','" & PorcDescuento & "','" & ImporteDescuento & "',"
            sql = sql & "'" & ImporteTotal & "','" & ImporteLineas2 & "','" & PorcDescuento2 & "','" & ImporteDescuento2 & "',"
            sql = sql & "'" & ImporteTotal2 & "','" & ImporteLineas3 & "','" & PorcDescuento3 & "','" & ImporteDescuento3 & "',"
            sql = sql & "'" & ImporteTotal3 & "','" & ImporteTotalSinIVAConDtoL & "','" & ImporteTotalSinIVAConDtoLConDtoTotal & "','" & ImporteDtoTotalSinIVA & "',"
            sql = sql & "'" & ImporteTotalDelIVAConDtoLConDtoTotal & "','" & ImporteSinRedondeo & "','" & ImporteDelRedondeo & "','" & Format(fecha, "yyyy-mm-dd hh:nn:ss") & ".000',"
            sql = sql & "'" & numLineas & "','" & SerieLIdFinDeDia & "','" & SerieLTicketErroneo & "','" & Modificado & "',"
            sql = sql & "'" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000','" & EstadoTicket & "',"
            sql = sql & "'" & ImporteEntregado & "','" & ImporteDevuelto & "','" & PuntosFidelidad & "')"
            InformaMiss "CABECERA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "CABECERA!"
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "INSERTANDO CABECERA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sql
            End If
            ImporteTotal = 0
            ImporteTotal2 = 0
            numLineas = 0
            idTicketAnt = rsLin("idTicket")
        End If
        'Campos lineas
        import = rsLin("Importe")
        IdLineaTicket = rsLin("IdLineaTicket")
        idEmpresa = rsLin("idEmpresa")
        IdTienda = rsLin("IdTienda")
        idBalanzaMaestra = rsLin("IdBalanzaMaestra")
        idBalanzaEsclava = rsLin("IdBalanzaEsclava")
        idTicket = rsLin("IdTicket")
        tipoVenta = rsLin("TipoVenta")
        EstadoLinea = rsLin("EstadoLinea")
        IdArticulo = rsLin("IdArticulo")
        DESCRIPCIoN = rsLin("Descripcion")
        Descripcion1 = rsLin("Descripcion1")
        Comportamiento = rsLin("Comportamiento")
        Tara = rsLin("Tara")
        Peso = rsLin("Peso")
        PesoRegalado = rsLin("PesoRegalado")
        precio = rsLin("Precio")
        PrecioSinIva = rsLin("PrecioSinIva")
        PrecioConIVASinDtoL = rsLin("PrecioConIVASinDtoL")
        IdIVA = rsLin("IdIVA")
        PorcentajeIva = rsLin("PorcentajeIva")
        descuento = rsLin("descuento")
        ImporteSinIVASinDtoL = rsLin("ImporteSinIVASinDtoL")
        ImporteConIVASinDtoL = rsLin("ImporteConIVASinDtoL")
        ImporteDelDescuento = rsLin("ImporteDelDescuento")
        importe = rsLin("Importe")
        ImporteSinIVAConDtoL = rsLin("ImporteSinIVAConDtoL")
        ImporteDelIVAConDtoL = rsLin("ImporteDelIVAConDtoL")
        ImporteSinIVAConDtoLConDtoTotal = rsLin("ImporteSinIVAConDtoLConDtoTotal")
        ImporteDelIVAConDtoLConDtoTotal = rsLin("ImporteDelIVAConDtoLConDtoTotal")
        ImporteConDtoTotal = rsLin("ImporteConDtoTotal")
        TaraFija = rsLin("TaraFija")
        TaraPorcentual = rsLin("TaraPorcentual")
        CodInterno = rsLin("CodInterno")
        EANScannerArticulo = rsLin("EANScannerArticulo")
        IdClase = rsLin("IdClase")
        NombreClase = rsLin("NombreClase")
        IdElemAsociado = rsLin("IdElemAsociado")
        IdFamilia = rsLin("IdFamilia")
        NombreFamilia = rsLin("NombreFamilia")
        IdSeccion = rsLin("IdSeccion")
        nombreSeccion = rsLin("NombreSeccion")
        IdSubFamilia = rsLin("IdSubFamilia")
        NombreSubFamilia = rsLin("NombreSubFamilia")
        IdDepartamento = rsLin("IdDepartamento")
        NombreDepartamento = rsLin("NombreDepartamento")
        Modificado = rsLin("Modificado")
        Operacion = rsLin("Operacion")
        usuario = rsLin("Usuario")
        TimeStamp = rsLin("TimeStamp")
        'Campos cabecera
        NombreEmpresa = rsLin("NombreEmpresa")
        NombreTienda = rsLin("NombreTienda")
        tipo = rsLin("tipo")
        IdVendedor = rsLin("IdVendedor")
        NombreVendedor = rsLin("NombreVendedor")
        idCliente = rsLin("IdCliente")
        NombreCliente = rsLin("NombreCliente")
        NombreCliente = Replace(NombreCliente, "'", "''")
        DNICliente = rsLin("DNICliente")
        DireccionCliente = rsLin("DireccionCliente")
        PoblacionCliente = rsLin("PoblacionCliente")
        ProvinciaCliente = rsLin("ProvinciaCliente")
        CPCliente = rsLin("CPCliente")
        TelefonoCliente = rsLin("TelefonoCliente")
        ImporteLineas = rsLin("ImporteLineas")
        PorcDescuento = rsLin("PorcDescuento")
        ImporteDescuento = rsLin("ImporteDescuento")
        ImporteLineas2 = rsLin("ImporteLineas2")
        PorcDescuento2 = rsLin("PorcDescuento2")
        ImporteDescuento2 = rsLin("ImporteDescuento2")
        ImporteLineas3 = rsLin("ImporteLineas3")
        PorcDescuento3 = rsLin("PorcDescuento3")
        ImporteDescuento3 = rsLin("ImporteDescuento3")
        ImporteTotal3 = rsLin("ImporteTotal3")
        ImporteSinRedondeo = rsLin("ImporteSinRedondeo")
        ImporteDelRedondeo = rsLin("ImporteDelRedondeo")
        SerieLIdFinDeDia = rsLin("SerieLIdFinDeDia")
        SerieLTicketErroneo = rsLin("SerieLTicketErroneo")
        ImporteDtoTotalSinIVA = rsLin("ImporteDtoTotalSinIVA")
        fecha = rsLin("fecha")
        EstadoTicket = rsLin("EstadoTicket")
        ImporteDevuelto = rsLin("ImporteDevuelto")
        PuntosFidelidad = rsLin("PuntosFidelidad")
        NumTicket = rsLin("numTicket")
        'ImporteTotal en ImporteTotal y ImporteTotal2
        ImporteTotal = ImporteTotal + Round(importe, 2)
        ImporteTotal2 = ImporteTotal2 + Round(importe, 2)
        numLineas = numLineas + 1
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR LINEA!!! BORRADO DE LINEAS DE TICKET PARA EL NUMERO DE TICKET ACTUAL
'----  ES POSIBLE QUE ALGUNA VEZ SE HAYA QUEDADO LA LINEA SIN CABECERA
'----------------------------------------------------------------------------------------------------
        If IdLineaTicket = 1 Then
            sqlSP = "delete FROM dat_ticket_linea "
            sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
            sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
            sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and idTicket='" & idTicket & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
        End If
'--------------------------------------------------------------------------------
'--INSERTANDO LINEA
'--------------------------------------------------------------------------------
        'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_linea limit 1'') ( "
        sql = "INSERT INTO dat_ticket_linea ( "
        sql = sql & "IdLineaTicket,IdEmpresa,IdTienda,IdBalanzaMaestra,IdBalanzaEsclava,IdTicket,"
        sql = sql & "TipoVenta,EstadoLinea,IdArticulo,Descripcion,Descripcion1,Comportamiento,Tara,"
        sql = sql & "Peso,PesoRegalado,Precio,PrecioSinIVA,PrecioConIVASinDtoL,IdIVA,PorcentajeIVA,"
        sql = sql & "Descuento,ImporteSinIVASinDtoL,ImporteConIVASinDtoL,ImporteDelDescuento,Importe,"
        sql = sql & "ImporteSinIVAConDtoL,ImporteDelIVAConDtoL,ImporteSinIVAConDtoLConDtoTotal,"
        sql = sql & "ImporteDelIVAConDtoLConDtoTotal,ImporteConDtoTotal,TaraFija,TaraPorcentual,"
        sql = sql & "CodInterno,EANScannerArticulo,IdClase,NombreClase,IdElemAsociado,IdFamilia,"
        sql = sql & "NombreFamilia,IdSeccion,NombreSeccion,IdSubFamilia,NombreSubFamilia,IdDepartamento,"
        sql = sql & "NombreDepartamento,Modificado,Operacion,Usuario,TimeStamp) "
        sql = sql & " values ('" & IdLineaTicket & "','" & idEmpresa & "','" & IdTienda & "','"
        sql = sql & idBalanzaMaestra & "','" & idBalanzaEsclava & "','" & idTicket & "','"
        sql = sql & tipoVenta & "','" & EstadoLinea & "','" & IdArticulo & "','" & Replace(DESCRIPCIoN, "'", "''") & "','" & Replace(Descripcion1, "'", "''") & "','" & CInt(Comportamiento) & "','" & Tara & "','"
        sql = sql & Peso & "','" & PesoRegalado & "','" & precio & "','" & PrecioSinIva & "','"
        sql = sql & PrecioConIVASinDtoL & "','" & IdIVA & "','" & PorcentajeIva & "','" & descuento & "','" & ImporteSinIVASinDtoL & "','" & ImporteConIVASinDtoL & "','"
        sql = sql & ImporteDelDescuento & "',' " & importe & "',' " & ImporteSinIVAConDtoL & "','"
        sql = sql & ImporteDelIVAConDtoL & "','" & ImporteSinIVAConDtoLConDtoTotal & "','"
        sql = sql & ImporteDelIVAConDtoLConDtoTotal & "','" & ImporteConDtoTotal & "','" & TaraFija & "','" & TaraPorcentual & "','"
        sql = sql & CodInterno & "','" & EANScannerArticulo & "','" & IdClase & "','" & NombreClase & "','" & IdElemAsociado & "','" & IdFamilia & "','"
        sql = sql & NombreFamilia & "','" & IdSeccion & "','" & nombreSeccion & "','" & IdSubFamilia & "','" & NombreSubFamilia & "','" & IdDepartamento & "','"
        sql = sql & NombreDepartamento & "','" & Modificado & "','" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000') "
        InformaMiss "LINEA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "INSERTANDO LINEA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            connMysql.Execute sql
        End If
        rsLin.MoveNext
'----------------------------------------------------------------------------------------------------
'----  ULTIMA CABECERA
'----------------------------------------------------------------------------------------------------
        If rsLin.EOF Then
            'Importe de ivas de los 3 tipos
            sql = "SELECT ("
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=1 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=1))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=2 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=2))/100,0),2) "
            sql = sql & "+"
            sql = sql & "ROUND(ISNULL(((SELECT SUM(import) FROM " & tabla & " v1 LEFT JOIN Articles a1 on (v1.Plu=a1.Codi) "
            sql = sql & "WHERE a1.TipoIva=3 and v1.Num_tick=" & idTicket & " and v1.Botiga=" & codiBotiga & ")* "
            sql = sql & "(SELECT Iva AS INTEGER FROM [TipusIva] where Tipus=3))/100,0),2) "
            sql = sql & ") as resultado "
            Set Rs = Db.OpenResultset(sql)
            If Not Rs.EOF Then
                ImporteTotalSinIVAConDtoL = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalSinIVAConDtoLConDtoTotal = Round(ImporteTotal - Rs("resultado"), 2)
                ImporteTotalDelIVAConDtoLConDtoTotal = Rs("resultado")
            End If
            ImporteEntregado = ImporteTotal
'----------------------------------------------------------------------------------------------------
'----  ANTES DE INSERTAR CABECERA!!! BORRADO DE CABECERA ACTUAL POR SI YA EXISTIERA
'----------------------------------------------------------------------------------------------------
            sqlSP = "delete FROM dat_ticket_cabecera where idEmpresa=1 "
            sqlSP = sqlSP & "and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "'  and idBalanzaEsclava=-1 "
            sqlSP = sqlSP & "and Usuario='Comunicaciones' and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
            sqlSP = sqlSP & "and nombreBalanzaMaestra like '-Balan%' and idTicket='" & idTicket & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sqlSP
            End If
            'SQL INSERT Cabecera MYSQL
            'sql = "INSERT INTO openquery(AMETLLER,''SELECT * FROM dat_ticket_cabecera limit 1'') "
            sql = "INSERT INTO dat_ticket_cabecera "
            sql = sql & "(IdTicket, NumTicket, IdEmpresa, NombreEmpresa,IdTienda,NombreTienda,IdBalanzaMaestra,"
            sql = sql & "NombreBalanzaMaestra, IdBalanzaEsclava, Tipo,"
            sql = sql & "IdVendedor,NombreVendedor,IdCliente,NombreCliente,"
            sql = sql & "DNICliente,DireccionCliente,PoblacionCliente,"
            sql = sql & "ProvinciaCliente,CPCliente,TelefonoCliente,"
            sql = sql & "TipoVenta,ImporteLineas,PorcDescuento,ImporteDescuento,"
            sql = sql & "ImporteTotal,ImporteLineas2,PorcDescuento2,ImporteDescuento2,"
            sql = sql & "ImporteTotal2,ImporteLineas3,PorcDescuento3,ImporteDescuento3,"
            sql = sql & "ImporteTotal3,ImporteTotalSinIVAConDtoL,ImporteTotalSinIVAConDtoLConDtoTotal,"
            sql = sql & "ImporteDtoTotalSinIVA,ImporteTotalDelIVAConDtoLConDtoTotal,ImporteSinRedondeo,"
            sql = sql & "ImporteDelRedondeo,Fecha,NumLineas,SerieLIdFinDeDia,SerieLTicketErroneo,Modificado,"
            sql = sql & "Operacion,Usuario,TimeStamp,EstadoTicket,ImporteEntregado,ImporteDevuelto,PuntosFidelidad) values ("
            sql = sql & "'" & idTicket & "','" & NumTicket & "','" & idEmpresa & "','" & NombreEmpresa & "',"
            sql = sql & "'" & IdTienda & "','" & NombreTienda & "','" & idBalanzaMaestra & "','-Balança " & idBalanzaMaestra & "',"
            sql = sql & "'" & idBalanzaEsclava & "','" & tipo & "','" & IdVendedor & "','" & NombreVendedor & "',"
            sql = sql & "'" & idCliente & "','" & NombreCliente & "','" & DNICliente & "','" & DireccionCliente & "',"
            sql = sql & "'" & PoblacionCliente & "','" & ProvinciaCliente & "','" & CPCliente & "','" & TelefonoCliente & "',"
            sql = sql & "'" & tipoVenta & "','" & ImporteLineas & "','" & PorcDescuento & "','" & ImporteDescuento & "',"
            sql = sql & "'" & ImporteTotal & "','" & ImporteLineas2 & "','" & PorcDescuento2 & "','" & ImporteDescuento2 & "',"
            sql = sql & "'" & ImporteTotal2 & "','" & ImporteLineas3 & "','" & PorcDescuento3 & "','" & ImporteDescuento3 & "',"
            sql = sql & "'" & ImporteTotal3 & "','" & ImporteTotalSinIVAConDtoL & "','" & ImporteTotalSinIVAConDtoLConDtoTotal & "','" & ImporteDtoTotalSinIVA & "',"
            sql = sql & "'" & ImporteTotalDelIVAConDtoLConDtoTotal & "','" & ImporteSinRedondeo & "','" & ImporteDelRedondeo & "','" & Format(fecha, "yyyy-mm-dd hh:nn:ss") & ".000',"
            sql = sql & "'" & numLineas & "','" & SerieLIdFinDeDia & "','" & SerieLTicketErroneo & "','" & Modificado & "',"
            sql = sql & "'" & Operacion & "','" & usuario & "','" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000','" & EstadoTicket & "',"
            sql = sql & "'" & ImporteEntregado & "','" & ImporteDevuelto & "','" & PuntosFidelidad & "')"
            InformaMiss "CABECERA " & idTicket & "-" & numLineas & "-->" & Format(TimeStamp, "yyyy-mm-dd hh:nn:ss") & ".000 "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "CABECERA!"
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "INSERTANDO CABECERA Ticket:" & idTicket & ",Linea:" & numLineas & "-->" & Now
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                connMysql.Execute sql
            End If
        End If
    Loop
'--------------------------------------------------------------------------------
'--SIGUIENTE TIENDA
'--------------------------------------------------------------------------------
    rsClients.MoveNext
    'Borramos tablas temp
    Db.OpenResultset ("DROP TABLE " & tablaTmp)
    Db.OpenResultset ("DROP TABLE " & tablaTmp2)
    Db.OpenResultset ("DROP TABLE " & tablaTmp3)
    Db.OpenResultset ("DROP TABLE " & tabla)
    html = html & "<p><b>Botiga: </b>" & codiBotiga & "</p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & desde & "</p>"
    html = html & "<p><b>Finalitzat: </b>" & Now() & "</p>"
Loop
If P3 <> "" Then sf_enviarMail "secrehit@hit.cat", Desti, " Vendes sincronitzades " & codiBotiga, html, "", ""

connMysql.Close
Set connMysql = Nothing

'Update de feinesafer nit para marcar cuando acaba proceso
If P4 = "Nit" Then
    sql = "Select count(id) num from feinesafer where tipus='SincroDbVendesAmetller' "
    Set Rs = Db.OpenResultset(sql)
    If Not Rs.EOF Then
        If Rs("num") <= 1 Then
            sql = "Update feinesafer set Param3='Fi " & Now() & "' where tipus='SincroDbVendesAmetllerNit' and param2 like '%Si%' "
            Set Rs = Db.OpenResultset(sql)
        End If
    End If
End If

InformaMiss "FIN SINCRO_VENDES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "FIN SINCRO_VENDES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.Close
End If
Exit Function

norVendes:
'----------------------------------------------------------------------------------------------------
'----  SI HAY ERROR: BORRADO ULTIMAS LINEAS DE TICKET HUERFANAS SIN CABECERA
'----------------------------------------------------------------------------------------------------
    html = "<p><h3>Resum Vendes Ametller </h3></p>"
    html = html & "<p><b>Botiga: </b>" & botiguesCad & "</p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & desde & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    If P3 <> "" Then sf_enviarMail "secrehit@hit.cat", Desti, "ERROR! Sincronitzacio de vendes de " & codiBotiga & " ha fallat", html, "", ""
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    html = html & "<p><b>Ultima sqlSP:</b>" & sqlSP & "</p>"
    
    sf_enviarMail "secrehit@hit.cat", EmailGuardia, "ERROR! Sincronitzacio de vendes de " & codiBotiga & " ha fallat", html, "", ""
        
    Set connMysql = New ADODB.Connection
    connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
    connMysql.ConnectionTimeout = 1000 '16 min
    connMysql.Open

    sqlSP = "delete FROM dat_ticket_linea "
    sqlSP = sqlSP & "where idEmpresa=1 and idBalanzaMaestra='" & Right(codiBotigaextern, 1) & "' "
    sqlSP = sqlSP & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' "
    sqlSP = sqlSP & "and Operacion='A' and idTienda='" & Left(codiBotigaextern, pos) & "' "
    sqlSP = sqlSP & "and idTicket='" & idTicket & "' "
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sqlSP & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "Error:" & err.Description & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        connMysql.Execute sqlSP
    End If
    
    connMysql.Close
    Set connMysql = Nothing
    
    'Borramos tablas temporales
    If ExisteixTaula(tablaTmp3) Then Db.OpenResultset ("DROP TABLE " & tablaTmp3)
    If ExisteixTaula(tablaTmp2) Then Db.OpenResultset ("DROP TABLE " & tablaTmp2)
    If ExisteixTaula(tablaTmp) Then Db.OpenResultset ("DROP TABLE " & tablaTmp)
    If ExisteixTaula(tabla) Then Db.OpenResultset ("DROP TABLE " & tabla)
    
    ExecutaComandaSql "Delete from  " & TaulaCalculsEspecials & " Id = '" & idTasca & "' "
    
    Informa "Error : " & err.Description
    
End Function



Function SincroDbAmetllerArticles(p1, idTasca) As Boolean
'*************************************************************************************
'SincroDbAmetllerArticles
'Sincroniza articles entre servidor Ametller y servidor Hit
'La variable CODI_PROD de la tabla articlesPropietats liga el codigo articulo de Hit con
'el codigo articulo de Ametller.
'Proceso:
'1.Se recoge fecha ultima modificacion de la tabla sincro_laforneria.
'2.Se abre un bucle mirando los articulos de ametller.
'3.En cada articulo se mira si ya existe en hit. Si existe actualiza sino inserta.
'4.Se inserta linea en MissatgesAEnviar.
'5.Se regeneran secciones para los teclados (articlesPropietats,SECCIO).
'   Se utilizaba para el proceso SINCRO_TECLATSTPV del SQL SERVER. Este proceso no esta en VB porque no se usa.
'6.Se borran(primero se copian a Articles_Zombis y luego se borran) articulos antiguos que no existan en ametller.
'*************************************************************************************
Dim codiFamilia As String, nomFamilia As String, codiSubFamilia As String, nomSubFamilia As String, codiArticle As String, descArticle As String
Dim unitatsArticle As String, pIvaArticle As String, preuArticle As Double, preuMArticle As Double, codiNum As Integer, iva As String, insArt As Integer, updArt As Integer
Dim codiBarres As String, fecMod As Date, ultimaMod As String, diff As Long, nombreSeccion As String, nombreSeccionAnt As String, regenSeccions As Integer
Dim idTabla As String, debugSincro As Boolean, sql As String, sql2 As String, sql3 As String, sqlSP As String, fecha As Date
Dim tablaTmp As String, Rs As rdoResultset, Rs2 As rdoResultset, html As String, tablaTmp2 As String, tablaTmp3 As String
Dim connMysql As ADODB.Connection, dataIni As Date
dataIni = Now()
codiNum = 0
insArt = 0
updArt = 0
    
On Error GoTo norArt

'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
'Direccion email a la que se enviara resultado/error
If p1 = "" Then p1 = EmailGuardia
debugSincro = False 'Solo impresion o ejecucion completa
'Crear txt log
If debugSincro = True Then
    p1 = EmailGuardia
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\Articles_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
End If
'--------------------------------------------------------------
InformaMiss "INICIO SINCRO_ARTICLES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "INICIO SINCRO_ARTICLES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'--ULTIMA FECHA MODIFICACION
'--------------------------------------------------------------------------------
InformaMiss "ULTIMA FECHA MODIFICACION-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "ULTIMA FECHA MODIFICACION-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
sql = "SELECT fecha FROM [Fac_LaForneria].[dbo].[sincro_laforneria] WHERE Variable='FECMOD' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    fecMod = Rs("fecha")
Else
    'Comprobamos que existe tabla sincro_laforneria
    tablaTmp = "[Fac_laforneria].[dbo].[sincro_laforneria]"
    sql = "SELECT object_id FROM sys.objects with (nolock) "
    sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp & "','[',''),']','') AND type='U' "
    Set Rs = Db.OpenResultset(sql)
    If Rs.EOF Then
        sql = "CREATE TABLE [Fac_laforneria].[dbo].[sincro_laforneria]("
        sql = sql & "       [variable] [varchar](255) NULL,"
        sql = sql & "       [fecha] [datetime] NULL,"
        sql = sql & "       [p1] [bit] NULL,"
        sql = sql & "       [p2] [nvarchar] (255) NULL,"
        sql = sql & "       [p3] [nvarchar] (255) NULL,"
        sql = sql & "       [p4] [nvarchar] (255) NULL,"
        sql = sql & "       [p5] [nvarchar] (255) NULL"
        sql = sql & "   ) ON [PRIMARY]"
        Set Rs = Db.OpenResultset(sql)
        sql = "INSERT INTO [Fac_laforneria].[dbo].[sincro_laforneria] values ("
        sql = sql & "'FECMOD',GETDATE(),0,GETDATE(),NULL,NULL,NULL) "
        Set Rs = Db.OpenResultset(sql)
        sql = "SELECT fecha FROM [Fac_LaForneria].[dbo].[sincro_laforneria] WHERE Variable='FECMOD' "
        Set Rs = Db.OpenResultset(sql)
        If Not Rs.EOF Then fecMod = Rs("fecha")
    End If
End If
'--------------------------------------------------------------------------------
html = "<p><h3>Resum Articles Ametller</h3></p>"
'--------------------------------------------------------------------------------
'--BUCLE ARTICULOS
'--------------------------------------------------------------------------------
InformaMiss "BUCLE ARTICULOS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "BUCLE ARTICULOS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
sql = "'SELECT art.IdFamilia AS CodiFamilia, fam.NombreFamilia AS Familia, art.IdSubfamilia AS CodiSubfamilia, sfm.NombreSubfamilia AS Subfamilia,"
sql = sql & "art.IdArticulo AS CodiArticle, art.Descripcion AS Article, IF(art.IdTipo=2,''UNI'',''KG'') AS Unitats, PorcentajeIVA AS pctIVA,"
sql = sql & "art.PrecioSinIva AS Preu, art.PrecioConIva AS PreuMajor,art.Timestamp AS fecMod,sec.nombreSeccion "
sql = sql & "FROM dat_articulo art "
sql = sql & "LEFT JOIN dat_iva AS iva ON iva.IdEmpresa=art.IdEmpresa AND iva.IdIVA=art.IdIVA "
sql = sql & "LEFT JOIN dat_familia AS fam ON fam.IdEmpresa=art.IdEmpresa AND fam.IdFamilia=art.IdFamilia "
sql = sql & "LEFT JOIN dat_subfamilia AS sfm on sfm.IdEmpresa=art.IdEmpresa AND sfm.IdFamilia=art.IdFamilia AND sfm.IdSubfamilia=art.IdSubfamilia "
sql = sql & "LEFT JOIN dat_seccion sec on (art.idSeccion=sec.idSeccion) "
sql = sql & "WHERE art.IdEmpresa=1 AND art.IdArticulo<90000 "
sql = sql & "and art.Timestamp>''" & Year(fecMod) & "-" & Month(fecMod) & "-" & Day(fecMod) & " "
sql = sql & DatePart("H", fecMod) & ":" & DatePart("n", fecMod) & ":" & DatePart("s", fecMod) & "'' '"
sql2 = "SELECT * FROM openquery(AMETLLER," & sql & ") "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql2)
Db.QueryTimeout = 60
Do While Not Rs.EOF
    If IsNull(Rs("CodiFamilia")) = False Then codiFamilia = Rs("CodiFamilia")
    If IsNull(Rs("Familia")) = False Then nomFamilia = Rs("Familia")
    If IsNull(Rs("CodiSubfamilia")) = False Then codiSubFamilia = Rs("CodiSubfamilia")
    If IsNull(Rs("Subfamilia")) = False Then nomSubFamilia = Rs("Subfamilia")
    If IsNull(Rs("CodiArticle")) = False Then codiArticle = Rs("CodiArticle")
    If IsNull(Rs("Article")) = False Then descArticle = Rs("Article")
    If IsNull(Rs("Unitats")) = False Then unitatsArticle = Rs("Unitats")
    If IsNull(Rs("pctIVA")) = False Then pIvaArticle = Rs("pctIVA")
    If IsNull(Rs("Preu")) = False Then preuMArticle = Rs("Preu")
    If IsNull(Rs("PreuMajor")) = False Then preuArticle = Rs("PreuMajor")
    If IsNull(Rs("fecmod")) = False Then
        ultimaMod = Rs("fecmod")
        'Se borra del tiempo mas alla de los segundos
        ultimaMod = Mid(ultimaMod, 1, Len(ultimaMod) - 8)
        ultimaMod = CDate(ultimaMod)
    End If
    If IsNull(Rs("nombreSeccion")) = False Then nombreSeccion = Rs("nombreSeccion")
    'Quitamos comillas
    descArticle = Replace(descArticle, "'", " ")
    nomSubFamilia = Replace(nomSubFamilia, "'", " ")
    'Codi producte
    codiNum = 0
    sql = "SELECT isNull(Codiarticle,'0') codi FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] "
    sql = sql & "WHERE Variable='CODI_PROD' and Valor='" & codiArticle & "' "
    Set Rs2 = Db.OpenResultset(sql)
    If Not Rs2.EOF Then codiNum = Rs2("codi")
    'Diferencia fechas para saber si hay que actualizar o insertar
    diff = DateDiff("s", fecMod, ultimaMod)
    If IsNull(diff) = True Then diff = 1
    If diff > 0 Then 'Se debe actualizar o insertar
        'Sino existe el codigo producto de este articulo se inserta.
        If codiNum = 0 Then
            'Proximo codigo article disponible
            sql = "SELECT top 1 CAST(Codi as integer)+1 num from [Fac_LaForneria].[dbo].[articles] order by codi desc"
            Set Rs2 = Db.OpenResultset(sql)
            If Not Rs2.EOF Then codiNum = Rs2("num")
            'tipusIva, si tipo 0%, aplicamos tipo 4%
            If CDbl(pIvaArticle) = 0 Then
                iva = 4
            Else
                sql = "SELECT top 1 tipus from [Fac_LaForneria].[dbo].[tipusIva] WITH (NOLOCK) WHERE Iva like '" & CDbl(pIvaArticle) & "' "
                Set Rs2 = Db.OpenResultset(sql)
                If Not Rs2.EOF Then iva = Rs2("tipus")
            End If
            'esSumable articles
            If unitatsArticle = "UNI" Then
                unitatsArticle = 1
            Else
                unitatsArticle = 0
            End If
            'insert articles
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[articles] ("
            sql = sql & "[Codi],[NOM],[PREU],[PreuMajor],[Desconte],[EsSumable],[Familia],[CodiGenetic],[TipoIva],[NoDescontesEspecials]) "
            sql = sql & "VALUES ("
            sql = sql & "'" & codiNum & "','" & descArticle & "','" & preuArticle & "','" & preuMArticle & "',"
            sql = sql & "'0','" & unitatsArticle & "','" & nomSubFamilia & "','" & codiNum & "','" & iva & "','0')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'insert articlesPropietats codi_prod
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[articlesPropietats] ("
            sql = sql & "[CodiArticle],[Variable],[Valor])"
            sql = sql & "VALUES ('" & codiNum & "','CODI_PROD','" & codiArticle & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'Limpieza de codigos de barras para este articulo
            sql = "DELETE FROM [Fac_LaForneria].[dbo].[CodisBarres] WHERE Producte='" & codiNum & "'"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'Insert de todos los codigos de barras para este articulo
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[codisbarres] "
            sql = sql & "SELECT * FROM openquery(AMETLLER,'SELECT EANScanner,''" & codiNum & "'' "
            sql = sql & "FROM dat_articulo_eanscanner WHERE IdEmpresa=1 AND IdArticulo=''" & codiArticle & "''')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'Si hay seccion, se regeneran al final
            If IsNull(nombreSeccion) = False Then regenSeccions = 1
            insArt = 1
'--------------------------------------------------------------------------------
InformaMiss "ARTICULO INSERTADO " & codiNum & "," & codiArticle & "," & descArticle & ",Preu " & preuArticle & ",PreuM " & preuMArticle & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "ARTICULO INSERTADO " & codiNum & "," & codiArticle & "," & descArticle & ",Preu " & preuArticle & ",PreuM " & preuMArticle & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
        Else 'Actualizacion articulo
            'tipusIva
            sql = "SELECT top 1 tipus from [Fac_LaForneria].[dbo].[tipusIva] WITH (NOLOCK) WHERE Iva like '" & CDbl(pIvaArticle) & "' "
            Set Rs2 = Db.OpenResultset(sql)
            If Not Rs2.EOF Then iva = Rs2("tipus")
            'esSumable articles
            If unitatsArticle = "UNI" Then
                unitatsArticle = 1
            Else
                unitatsArticle = 0
            End If
            'Insert articlesHistorial para saber que articulos se han ido modificando
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[articlesHistorial]( "
            sql = sql & "[Codi],[NOM],[PREU],[PreuMajor],[Desconte],"
            sql = sql & "[EsSumable],[Familia],[CodiGenetic],[TipoIva],"
            sql = sql & "[NoDescontesEspecials],[fechaModif],[usuarioModif]) "
            sql = sql & "VALUES ("
            sql = sql & "'" & codiNum & "','" & descArticle & "','" & preuArticle & "','" & preuMArticle & "',"
            sql = sql & "'0','" & unitatsArticle & "','" & nomSubFamilia & "','" & codiNum & "','" & iva & "','0',"
            sql = sql & "GETDATE(),'SINCRO_LAFORNERIA')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'Update articles
            sql = "UPDATE [Fac_LaForneria].[dbo].[articles] "
            sql = sql & "SET [NOM]='" & descArticle & "',"
            sql = sql & "[PREU]='" & preuArticle & "',"
            sql = sql & "[PreuMajor]='" & preuMArticle & "',"
            sql = sql & "[Desconte]='0',"
            sql = sql & "[EsSumable]='" & unitatsArticle & "',"
            sql = sql & "[Familia]='" & nomSubFamilia & "',"
            sql = sql & "[TipoIva]='" & iva & "',"
            sql = sql & "[NoDescontesEspecials]='0' "
            sql = sql & "WHERE [Codi]='" & codiNum & "'"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'Limpieza de codigos de barras para este articulo
            sql = "DELETE FROM [Fac_LaForneria].[dbo].[CodisBarres] WHERE Producte='" & codiNum & "'"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'Insert de todos los codigos de barras para este articulo
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[codisbarres] "
            sql = sql & "SELECT * FROM openquery(AMETLLER,'SELECT EANScanner,''" & codiNum & "'' "
            sql = sql & "FROM dat_articulo_eanscanner WHERE IdEmpresa=1 AND IdArticulo=''" & codiArticle & "''')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'Si cambia el nombre de la seccion, regeneramos teclados
            'nombreSeccionAnt = ""
            'Sql = "SELECT isNull(Valor,0) seccio from [Fac_LaForneria].[dbo].[ArticlesPropietats] "
            'Sql = Sql & "WHERE Variable='SECCIO' and CodiArticle='" & codiNum & "'"
            'Set Rs2 = Db.OpenResultset(Sql)
            'If Not Rs2.EOF Then nombreSeccionAnt = Rs2("seccio")
            'If nombreSeccion <> nombreSeccionAnt Then regenSeccions = 1
            updArt = 1
'--------------------------------------------------------------------------------
InformaMiss "ARTICULO ACTUALIZADO" & codiNum & "," & codiArticle & "," & descArticle & ",Preu " & preuArticle & ",PreuM " & preuMArticle & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "ARTICULO ACTUALIZADO " & codiNum & "," & codiArticle & "," & descArticle & ",Preu " & preuArticle & ",PreuM " & preuMArticle & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
        End If
    End If
    Rs.MoveNext
Loop
'--------------------------------------------------------------------------------
'Si se han insertado/actualizado articulos se pone missatge
If insArt > 0 Or updArt > 0 Then
    'MissatgesAEnviar
    sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) "
    sql = sql & "VALUES ('Articles','')"
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        Set Rs = Db.OpenResultset(sql)
    End If
    sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) "
    sql = sql & "VALUES ('CodisBarres','')"
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        Set Rs = Db.OpenResultset(sql)
    End If
    sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) "
    sql = sql & "VALUES ('Families','')"
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        Set Rs = Db.OpenResultset(sql)
    End If
End If
'--------------------------------------------------------------------------------
'If regenSeccions = 1 Then
'    'Regenerem seccions
'    Sql = "UPDATE [Fac_LaForneria].[dbo].[sincro_laforneria] "
'    Sql = Sql & "SET [p1]='1' WHERE variable='FECMOD' "
'    If debugSincro = True Then
'        Txt.WriteLine "--------------------------------------------------------------------------------"
'        Txt.WriteLine "SQL:" & Sql & "-->" & Now
'       Txt.WriteLine "--------------------------------------------------------------------------------"
'    Else
'        Set Rs = Db.OpenResultset(Sql)
'    End If
'    Sql = "DELETE FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] WHERE Variable='SECCIO' "
'    If debugSincro = True Then
'        Txt.WriteLine "--------------------------------------------------------------------------------"
'        Txt.WriteLine "SQL:" & Sql & "-->" & Now
'        Txt.WriteLine "--------------------------------------------------------------------------------"
'    Else
'        Set Rs = Db.OpenResultset(Sql)
'    End If
'    Sql = "INSERT INTO [Fac_LaForneria].[dbo].[ArticlesPropietats] "
'    Sql = Sql & "select a.codi,'SECCIO',s.nombreSeccion from ("
'    Sql = Sql & "SELECT * FROM openquery(AMETLLER,' "
'    Sql = Sql & "SELECT sec.nombreSeccion as nombreSeccion,art.idArticulo as idArticulo "
'    Sql = Sql & "FROM dat_articulo art "
'    Sql = Sql & "LEFT JOIN dat_seccion sec on (art.idSeccion=sec.idSeccion) "
'    Sql = Sql & "where art.idArticulo<90000')) s "
'    Sql = Sql & "LEFT JOIN [Fac_LaForneria].[dbo].[ArticlesPropietats] ap ON (CAST(ap.Valor as nvarchar(255))=cast(s.idArticulo as nvarchar(255))) "
'    Sql = Sql & "LEFT JOIN [Fac_LaForneria].[dbo].[Articles] a ON (ap.CodiArticle=a.Codi) "
'    Sql = Sql & "WHERE s.nombreSeccion is not null and a.Codi is not null"
'    If debugSincro = True Then
'        Txt.WriteLine "--------------------------------------------------------------------------------"
'        Txt.WriteLine "SQL:" & Sql & "-->" & Now
'        Txt.WriteLine "--------------------------------------------------------------------------------"
'    Else
'        Set Rs = Db.OpenResultset(Sql)
'    End If
'End If
'--------------------------------------------------------------------------------
InformaMiss "BORRADO ARTICULO ANTIGUOS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "BORRADO ARTICULO ANTIGUOS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'Comparar con tabla dat_articulo mysql, si no existe en esta tabla se copia a articles_zombis y se borra el articulo
sql = "INSERT INTO [Fac_LaForneria].[dbo].[Articles_Zombis] "
sql = sql & "SELECT GETDATE(),a.codi,a.NOM,a.PREU,a.PreuMajor,"
sql = sql & "a.Desconte,a.EsSumable,a.Familia,a.CodiGenetic,a.TipoIva,"
sql = sql & "a.NoDescontesEspecials FROM [Fac_LaForneria].[dbo].[Articles] a WITH (NOLOCK) "
sql = sql & "LEFT JOIN [Fac_LaForneria].[dbo].[ArticlesPropietats] ap WITH (NOLOCK)ON (a.Codi=ap.CodiArticle) "
sql = sql & "WHERE Variable is null or (Variable='CODI_PROD' AND Valor NOT IN( "
sql = sql & "SELECT * FROM openquery(AMETLLER,'SELECT IdArticulo FROM dat_articulo "
sql = sql & "WHERE IdEmpresa=1 AND IdArticulo<90000'))) "
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "SQL:" & sql & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
Else
    Set Rs = Db.OpenResultset(sql)
End If
'Borrado articles que no existan en dat_articulo
sql = "DELETE FROM [Fac_LaForneria].[dbo].[Articles] "
sql = sql & "WHERE Codi in ( "
sql = sql & "SELECT ISNULL(CodiArticle,'') FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] WITH (NOLOCK) "
sql = sql & "WHERE Variable is null or (Variable='CODI_PROD' AND Valor NOT IN( "
sql = sql & "SELECT * FROM openquery(AMETLLER,'SELECT IdArticulo FROM dat_articulo "
sql = sql & "WHERE IdEmpresa=1 AND IdArticulo<90000'))) )"
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "SQL:" & sql & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
Else
    Set Rs = Db.OpenResultset(sql)
End If
'Borrado articlesPropietats que no existan en dat_articulo
sql = "DELETE FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] "
sql = sql & "WHERE Variable is null or (Variable='CODI_PROD' AND Valor NOT IN( "
sql = sql & "SELECT * FROM openquery(AMETLLER,'SELECT IdArticulo FROM dat_articulo "
sql = sql & "WHERE IdEmpresa=1 AND IdArticulo<90000'))) "
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "SQL:" & sql & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
Else
    Set Rs = Db.OpenResultset(sql)
End If
'Borrado Articles que no tengan CODI_PROD en articlesPropietats
sql = "DELETE [Fac_LaForneria].[dbo].[Articles] WHERE codi NOT IN "
sql = sql & "(SELECT codiarticle FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] WITH (NOLOCK) "
sql = sql & "WHERE Variable='CODI_PROD' ) "
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "SQL:" & sql & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
Else
    Set Rs = Db.OpenResultset(sql)
End If
'Borrado codigos de barras a 0
sql = "DELETE FROM codisBarres WHERE codi like '0000%' "
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "SQL:" & sql & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
Else
    Set Rs = Db.OpenResultset(sql)
End If
'----------------------------------------------------------------------------------------------------
connMysql.Close
Set connMysql = Nothing
'----------------------------------------------------------------------------------------------------
'--UPDATE SINCRO_LAFORNERIA
'----------------------------------------------------------------------------------------------------
'Se indica que el proceso actual ha acabado
Set Rs = Db.OpenResultset("Delete from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='ARTICLES'")
Set Rs = Db.OpenResultset("Insert into [Fac_LaForneria].[dbo].[sincro_laforneria] (variable,fecha,p1) values ('ARTICLES',getdate(),1)")
'Se mira si es el ultimo proceso activo, si es asi se actualiza la fecha de la variable fecmod
'que es por la que se rigen los procesos.
sql = "Select COUNT(variable) num from [Fac_LaForneria].[dbo].[sincro_laforneria] where p1=1 and variable<>'FECMOD' "
sql = sql & "and fecha>=(select convert(datetime,p2,103) from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='fecmod') "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    'Si se han completado los cinco procesos se actualiza
    If Rs("num") = 5 Then
        Set Rs2 = Db.OpenResultset("Update [Fac_LaForneria].[dbo].[sincro_laforneria] Set [fecha] = [P2] WHERE [Variable]='FECMOD'")
        Set Rs2 = Db.OpenResultset("Update feinesafer set Param3='Fi " & Now() & "' where tipus='SincroDbAmetllerHorari' and param2 like '%Si%' ")
    End If
End If
'----------------------------------------------------------------------------------------------------
InformaMiss "FIN SINCRO_ARTICLES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "FIN SINCRO_ARTICLES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.Close
End If
Exit Function


norArt:
'----------------------------------------------------------------------------------------------------
'----  SI HAY ERROR
'----------------------------------------------------------------------------------------------------
    'Se indica que el proceso actual ha fallado
    Set Rs = Db.OpenResultset("Delete from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='ARTICLES'")
    Set Rs = Db.OpenResultset("Insert into [Fac_LaForneria].[dbo].[sincro_laforneria] (variable,fecha,p1) values ('ARTICLES',getdate(),0)")
    
    html = "<p><h3>Resum Articles Ametller </h3></p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & dataIni & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    html = html & "<p><b>Ultima sqlSP:</b>" & sqlSP & "</p>"
    
    sf_enviarMail "secrehit@hit.cat", p1, "ERROR! Sincronitzacio d'articles ha fallat", html, "", ""
            
    ExecutaComandaSql "Delete from  " & TaulaCalculsEspecials & " Id = '" & idTasca & "' "
    
    Informa "Error : " & err.Description
    
End Function

Function SincroDbAmetllerBotigues(p1, idTasca) As Boolean
'*************************************************************************************
'SincroDbAmetllerBotigues
'Sincroniza botigues entre servidor Ametller y servidor Hit
'La variables CodiContable y CodiClientOrigen (contienen lo mismo) de la tabla constantsClient
'ligan el codigo tienda de Hit con el codigo tienda de Ametller.
'Proceso:
'1.Se recoge fecha ultima modificacion de la tabla sincro_laforneria.
'2.Se abre un bucle mirando las tiendas ametller, segun los codigos
'    de la variable llistatBot y la ultima fecha de modificacion.
'3.En cada tienda se mira si ya existe en hit. Si existe actualiza sino inserta.
'4.Se insertan lineas (Tpv_Configuracio_ y Clients) en MissatgesAEnviar.
'5.Se regeneran tarifas y promociones de las tiendas
'6.Se insertan lineas (Tpv_Configuracio_ y ProductesPromocionats) en MissatgesAEnviar.
'*************************************************************************************
'Proceso para nuevas tiendas:
'1. añadir codigo tienda a la variable llistaBot.
'2. modificar [SQL BUCLE TIENDAS] para que no tenga en cuenta la fecha de modificacion.
'3. se generaran licencias 9999 para las tiendas nuevas en la tabla paramshw
'4. vincular licencias (segun ficheros txt de las nuevas licencias) con clientes.
'5. volver a pasar este proceso para el correcto envio de datos.
'6. modificar [SQL BUCLE TIENDAS] para que vuelva a tener en cuenta la fecha de modificacion.
'*************************************************************************************
Dim codiBotiga As String, codiBotigaOriginal As String, nomBotiga As String, codiBalanza As String, nomBalanza As String, idTarifa As String
Dim nomTarifa As String, codiExist As String, cifBotiga As String, adresaBotiga As String, cpBotiga As String, ciutatBotiga As String
Dim provinciaBotiga As String, codiNum As Integer, fecMod As Date, ultimaMod As String, diff As Long, insBot As Integer, updBot As Integer
Dim idTabla As String, debugSincro As Boolean, sql As String, sql2 As String, sql3 As String, sqlSP As String, fecha As Date
Dim tablaTmp As String, Rs As rdoResultset, Rs2 As rdoResultset, html As String, tablaTmp2 As String, tablaTmp3 As String
Dim connMysql As ADODB.Connection, dataIni As Date, llistaBot As String
dataIni = Now()
codiNum = 0
insBot = 0
updBot = 0
'¡¡¡Lista de tiendas que se comprobaran si hace falta actualizar/insertar!!!
llistaBot = "51,57,59,60,61,63,64,65,68,70,71,72,67,901,73,11,74,76,201,77,66"
On Error GoTo norBot

'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
'Direccion email a la que se enviara resultado/error
If p1 = "" Then p1 = EmailGuardia
debugSincro = False 'Solo impresion o ejecucion completa
'Crear txt log
If debugSincro = True Then
    p1 = EmailGuardia
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\Botigues_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
End If
'--------------------------------------------------------------
InformaMiss "INICIO SINCRO_BOTIGUES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "INICIO SINCRO_BOTIGUES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'--ULTIMA FECHA MODIFICACION
'--------------------------------------------------------------------------------
InformaMiss "ULTIMA FECHA MODIFICACION-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "ULTIMA FECHA MODIFICACION-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
sql = "SELECT fecha FROM [Fac_LaForneria].[dbo].[sincro_laforneria] WHERE Variable='FECMOD' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    fecMod = Rs("fecha")
Else
    'Comprobamos que existe tabla sincro_laforneria
    tablaTmp = "[Fac_laforneria].[dbo].[sincro_laforneria]"
    sql = "SELECT object_id FROM sys.objects with (nolock) "
    sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp & "','[',''),']','') AND type='U' "
    Set Rs = Db.OpenResultset(sql)
    If Rs.EOF Then
        sql = "CREATE TABLE [Fac_laforneria].[dbo].[sincro_laforneria]("
        sql = sql & "       [variable] [varchar](255) NULL,"
        sql = sql & "       [fecha] [datetime] NULL,"
        sql = sql & "       [p1] [bit] NULL,"
        sql = sql & "       [p2] [nvarchar] (255) NULL,"
        sql = sql & "       [p3] [nvarchar] (255) NULL,"
        sql = sql & "       [p4] [nvarchar] (255) NULL,"
        sql = sql & "       [p5] [nvarchar] (255) NULL"
        sql = sql & "   ) ON [PRIMARY]"
        Set Rs = Db.OpenResultset(sql)
        sql = "INSERT INTO [Fac_laforneria].[dbo].[sincro_laforneria] values ("
        sql = sql & "'FECMOD',GETDATE(),0,GETDATE(),NULL,NULL,NULL) "
        Set Rs = Db.OpenResultset(sql)
        sql = "SELECT fecha FROM [Fac_LaForneria].[dbo].[sincro_laforneria] WHERE Variable='FECMOD' "
        Set Rs = Db.OpenResultset(sql)
        If Not Rs.EOF Then fecMod = Rs("fecha")
    End If
End If
'--------------------------------------------------------------------------------
html = "<p><h3>Resum Botigues Ametller</h3></p>"
'--------------------------------------------------------------------------------
'--BUCLE TIENDAS
'--------------------------------------------------------------------------------
InformaMiss "BUCLE TIENDAS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "BUCLE TIENDAS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'[SQL BUCLE TIENDAS]
sql = "'SELECT b.idBalanza AS CodiBalanza,t.idTienda AS CodiBotiga,TRIM(t.Nombre) AS Botiga,t.idTarifa AS CodiTarifa,"
sql = sql & "TRIM(t.CIF_VAT) AS Cif,TRIM(t.Direccion) AS Dir,TRIM(t.CodPostal) AS CP,TRIM(t.Poblacion) AS Ciutat,"
sql = sql & "TRIM(t.Provincia) AS Provincia,TRIM(b.Nombre) as Balanza,t.TimeStamp AS fecMod "
sql = sql & "FROM dat_tienda t "
sql = sql & "LEFT JOIN dat_balanza b on (b.idTienda=t.idTienda) "
sql = sql & "WHERE t.IdEmpresa=1 AND t.idTienda in (" & llistaBot & ") "
sql = sql & "and b.nombre like ''%hit%'' "      ' original comentar linia per desactivar control de temps.
'Sql = Sql & "and b.nombre like ''%hit%'' ' "      ' descomentar linia per desactivar control de temps.
'Se mira la fecha de modificacion de la tienda!
sql = sql & "and t.Timestamp>''" & Year(fecMod) & "-" & Month(fecMod) & "-" & Day(fecMod) & " "   ' comentar linia per desactivar control de temps.
sql = sql & DatePart("H", fecMod) & ":" & DatePart("n", fecMod) & ":" & DatePart("s", fecMod) & "'' '"  ' comentar linia per desactivar control de temps.
sql2 = "SELECT * FROM openquery(AMETLLER," & sql & ") "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql2)
Db.QueryTimeout = 60
Do While Not Rs.EOF
    If IsNull(Rs("CodiBalanza")) = False Then codiBalanza = Rs("CodiBalanza")
    If IsNull(Rs("CodiBotiga")) = False Then codiBotiga = Rs("CodiBotiga")
    If IsNull(Rs("Botiga")) = False Then nomBotiga = Rs("Botiga")
    If IsNull(Rs("CodiTarifa")) = False Then idTarifa = Rs("CodiTarifa")
    If IsNull(Rs("Cif")) = False Then cifBotiga = Rs("Cif")
    If IsNull(Rs("Dir")) = False Then adresaBotiga = Rs("Dir")
    If IsNull(Rs("CP")) = False Then cpBotiga = Rs("CP")
    If IsNull(Rs("Ciutat")) = False Then ciutatBotiga = Rs("Ciutat")
    If IsNull(Rs("Provincia")) = False Then provinciaBotiga = Rs("Provincia")
    If IsNull(Rs("Balanza")) = False Then nomBalanza = Rs("Balanza")
    If IsNull(Rs("fecmod")) = False Then
        ultimaMod = Rs("fecmod")
        'Se borra del tiempo mas alla de los segundos
        ultimaMod = Mid(ultimaMod, 1, Len(ultimaMod) - 8)
        ultimaMod = CDate(ultimaMod)
    End If
    'El proceso mira la diferencia de fechas de la ultima vez que se ejecuto el proceso
    'y el timestamp de la tienda en mysql. Por ahora se obliga siempre a actualizar tiendas
    'o insertar si es necesario con la siguiente instruccion:
    ultimaMod = Now()
    'Esto hace que se reenvie siempre la configuracion del tpv a las tiendas
    nomBotiga = nomBotiga & "_" & nomBalanza
    nomBotiga = Replace(nomBotiga, "'", " ")
    adresaBotiga = Replace(adresaBotiga, "'", " ")
    codiBotigaOriginal = codiBotiga
    'CodiClient+CodiBalança
    codiBotiga = codiBotiga & codiBalanza
    sql = "SELECT TarifaNom from TarifesEspecials where TarifaCodi='" & idTarifa & "'"
    Set Rs2 = Db.OpenResultset(sql)
    If Not Rs2.EOF Then nomTarifa = Rs2("TarifaNom")
    'Codi client
    codiNum = 0
    sql = "SELECT isNull(Codi,'0') codi FROM [Fac_LaForneria].[dbo].[Clients] WHERE Codi='" & codiBotiga & "' "
    Set Rs2 = Db.OpenResultset(sql)
    If Not Rs2.EOF Then codiNum = Rs2("codi")
    'Diferencia fechas para saber si hay que actualizar o insertar
    diff = DateDiff("s", fecMod, ultimaMod)
    If IsNull(diff) = True Then diff = 1
    If diff > 0 Then 'Se debe actualizar o insertar
        'Sino existe el codigo de esta tienda, se inserta tienda
        If codiNum = 0 Then
            'insert clients
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[clients] ("
            sql = sql & "[Codi],[Nom],[Nif],[Adresa],[Ciutat],[Cp],[Lliure],[Nom Llarg],[Tipus Iva],[Preu Base],"
            sql = sql & "[Desconte ProntoPago],[Desconte 1],[Desconte 2],[Desconte 3],[Desconte 4],[Desconte 5],[AlbaraValorat]) "
            sql = sql & "VALUES ('" & codiBotiga & "','" & nomBotiga & "','" & cifBotiga & "','" & adresaBotiga & "','" & ciutatBotiga & "',"
            sql = sql & "'" & cpBotiga & "','','" & nomBotiga & "','2','2','0','0','0','0','0','" & idTarifa & "',NULL)"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'insert constantsClients
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','CodiContable','" & codiBotiga & "')  "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','CodiClientOrigen','" & codiBotiga & "') "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'insert paramsHW llicencia
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsHw] ([Tipus],[Codi],[Valor1],[Valor2],[Valor3],[Valor4],[Descripcio]) "
            sql = sql & "VALUES ('1','" & codiBotiga & "','" & codiBotiga & "','','','','La Llicencia : " & codiBotiga & " Correspon a la botiga " & codiBotiga & "') "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'insert paramsTPV
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "', 'SeleccionarTarifa','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Nom','" & nomBotiga & "')    "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','CodiBotiga','" & codiNum & "')   "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Preu Base',2)     "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Nif','" & cifBotiga & "')    "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Adresa','" & adresaBotiga & "') "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Ciutat','" & ciutatBotiga & "') "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Cp','" & cpBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Lliure','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Nom Llarg','" & nomBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Tipus Iva',1)    "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','BotonsPreu','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Capselera_1','Casa Ametller '+'" & cifBotiga & "'+' '+'" & adresaBotiga & "'+' '+'" & cpBotiga & "'+ ' '+'" & ciutatBotiga & "'+ '('+'" & provinciaBotiga & "'+')') "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Capselera_2','Gràcies per la seva visita')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_0_Imatge','aguacate.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_0_Tag','01')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_0_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_0_Texte','Fruita')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_1_Imatge','calabacin.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_1_Tag','02')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_1_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_1_Texte','Verdura')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_2_Imatge','apio.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_2_Tag','03')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_2_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_2_Texte','Hortalises')    "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_3_Imatge','Zumos.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_3_Tag','04')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_3_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_3_Texte','Tomaquets')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_4_Imatge','fresitas.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_4_Tag','05')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_4_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_4_Texte','Amanida')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_5_Imatge','pimientos2.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_5_Tag','06')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_5_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_5_Texte','Patata I Ceba I Alls')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_6_Imatge','platanos.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_6_Tag','07')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_6_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_6_Texte','Bolets I Olives')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_7_Imatge','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_7_Tag','08')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_7_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_7_Texte','Fruits Secs I Vins')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','EditarTeclats','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','EntrarCanviEnMonedes','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','UnSolOperari','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DetallIva','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','RebutjaCarregaTeclats','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','SempreTicket','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','ZAmagada','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'TPVEQUIVALENTS
            'Pone las tiendas asociadas
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[TpvEquivalents] ([Tipus],[Valor1],[Valor2],[Valor3],[Valor4],[Valor5]) "
            sql = sql & "VALUES ('Preus','" & codiBotiga & "','',null,null,null)    "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            insBot = 1
'------------------------------------------------------------------------------------------
InformaMiss "TIENDA INSERTADA " & codiBotiga & "," & nomBotiga & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "TIENDA INSERTADA " & codiBotiga & "," & nomBotiga & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
        Else
            'Update clients
            sql = "UPDATE [Fac_LaForneria].[dbo].[clients] "
            sql = sql & "SET [NOM]='" & nomBotiga & "',"
            sql = sql & "   [Nif]='" & cifBotiga & "',"
            sql = sql & "   [Adresa]='" & adresaBotiga & "',"
            sql = sql & "   [Ciutat]='" & ciutatBotiga & "',"
            sql = sql & "   [Cp]='" & cpBotiga & "',"
            sql = sql & "   [Lliure]='',"
            sql = sql & "   [Nom Llarg]='" & nomBotiga & "',"
            sql = sql & "   [Tipus Iva]='2',"
            sql = sql & "   [Preu Base]='2',"
            sql = sql & "   [Desconte ProntoPago]='0',"
            sql = sql & "   [Desconte 5]='" & idTarifa & "',"
            sql = sql & "   [AlbaraValorat]=NULL "
            sql = sql & "WHERE [Codi]='" & codiBotiga & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'Reinsertamos datos de constantsclient!
            sql = "DELETE FROM [Fac_LaForneria].[dbo].[constantsclient] WHERE Codi='" & codiBotiga & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'insert constantsClients
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','CodiContable','" & codiBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','CodiClientOrigen','" & codiBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'insert paramsHW
            sql = "SELECT Codi FROM ParamsHw WHERE valor1='" & codiBotiga & "' "
            Set Rs2 = Db.OpenResultset(sql)
            If Rs2.EOF Then
'------------------------------------------------------------------------------------------
InformaMiss "NO SE HA ENCONTRADO LICENCIA PARA TIENDA " & codiBotiga & "," & nomBotiga & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "NO SE HA ENCONTRADO LICENCIA PARA TIENDA " & codiBotiga & "," & nomBotiga & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'------------------------------------------------------------------------------------------
                'Se inserta linea en tabla ParamsHw pero con una licencia ficticia que luego habra que cambiar!
                sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsHw] ([Tipus],[Codi],[Valor1],[Valor2],[Valor3],[Valor4],[Descripcio]) "
                sql = sql & "VALUES ('1','9999','" & codiBotiga & "','','','','La Llicencia : " & codiBotiga & " Correspon a la botiga " & codiBotiga & "')"
                If debugSincro = True Then
                    Txt.WriteLine "--------------------------------------------------------------------------------"
                    Txt.WriteLine "SQL:" & sql & "-->" & Now
                    Txt.WriteLine "--------------------------------------------------------------------------------"
                Else
                    Set Rs2 = Db.OpenResultset(sql)
                End If
            End If
            'Se reinsertan paramsTPV
            sql = "DELETE FROM [Fac_LaForneria].[dbo].[ParamsTpv] WHERE CodiClient='" & codiBotiga & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'insert paramsTPV
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "', 'SeleccionarTarifa','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Nom','" & nomBotiga & "')    "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','CodiBotiga','" & codiNum & "')   "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Preu Base',2)     "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Nif','" & cifBotiga & "')    "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Adresa','" & adresaBotiga & "') "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Ciutat','" & ciutatBotiga & "') "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Cp','" & cpBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Lliure','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Nom Llarg','" & nomBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Tipus Iva',1)    "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','BotonsPreu','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Capselera_1','Casa Ametller '+'" & cifBotiga & "'+' '+'" & adresaBotiga & "'+' '+'" & cpBotiga & "'+ ' '+'" & ciutatBotiga & "'+ '('+'" & provinciaBotiga & "'+')') "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','Capselera_2','Gràcies per la seva visita')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_0_Imatge','aguacate.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_0_Tag','01')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_0_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_0_Texte','Fruita')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_1_Imatge','calabacin.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_1_Tag','02')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_1_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_1_Texte','Verdura')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_2_Imatge','apio.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_2_Tag','03')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_2_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_2_Texte','Hortalises')    "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_3_Imatge','Zumos.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_3_Tag','04')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_3_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_3_Texte','Tomaquets')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_4_Imatge','fresitas.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_4_Tag','05')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_4_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_4_Texte','Amanida')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_5_Imatge','pimientos2.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_5_Tag','06')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_5_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_5_Texte','Patata I Ceba I Alls')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_6_Imatge','platanos.bmp')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_6_Tag','07')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_6_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_6_Texte','Bolets I Olives')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_7_Imatge','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_7_Tag','08')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_7_Tarifa','')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DosNivells_7_Texte','Fruits Secs I Vins')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','EditarTeclats','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','EntrarCanviEnMonedes','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','UnSolOperari','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','DetallIva','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','RebutjaCarregaTeclats','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','SempreTicket','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "','ZAmagada','Si')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
'--------------------------------------------------------------------------------
InformaMiss "TIENDA ACTUALIZADA " & codiBotiga & "," & nomBotiga & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "TIENDA ACTUALIZADA " & codiBotiga & "," & nomBotiga & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
            updBot = 1
        End If
'--------------------------------------------------------------------------------
        'MissatgesAEnviar Tpv_Configuracio
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) "
        sql = sql & "VALUES ('Tpv_Configuracio_',(select top 1 codi from ParamsHw where Valor1 = '" & codiBotiga & "')) "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
    End If
    Rs.MoveNext
Loop
'---------------------------------------------------------------------------------------------
If insBot > 0 Or updBot > 0 Then
    'MissatgesAEnviar
    sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) "
    sql = sql & "VALUES ('Clients','')  "
End If
'---------------------------------------------------------------------------------------------
'--REGENERA TARIFAS Y PROMOCIONES SIEMPRE!
'--------------------------------------------------------------------------------------------
InformaMiss "REGENERACION DE TARIFAS Y PROMOCIONES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "REGENERACION DE TARIFAS Y PROMOCIONES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------------------
'Bucle tiendas
sql = "'SELECT b.idBalanza AS CodiBalanza,t.idTienda AS CodiBotiga,TRIM(t.Nombre) AS Botiga,t.idTarifa AS CodiTarifa, "
sql = sql & "TRIM(b.Nombre) as Balanza,t.TimeStamp AS fecMod FROM dat_tienda t "
sql = sql & "LEFT JOIN dat_balanza b on (b.idTienda=t.idTienda) "
sql = sql & "WHERE t.IdEmpresa=1 AND t.idTienda in (" & llistaBot & ") '"
sql2 = "SELECT * FROM openquery(AMETLLER," & sql & ") "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql2)
Db.QueryTimeout = 60
Do While Not Rs.EOF
    If IsNull(Rs("CodiBalanza")) = False Then codiBalanza = Rs("CodiBalanza")
    If IsNull(Rs("CodiBotiga")) = False Then codiBotiga = Rs("CodiBotiga")
    If IsNull(Rs("Botiga")) = False Then nomBotiga = Rs("Botiga")
    If IsNull(Rs("CodiTarifa")) = False Then idTarifa = Rs("CodiTarifa")
    If IsNull(Rs("Balanza")) = False Then nomBalanza = Rs("Balanza")
    
    If codiBotiga = "76" Then
     codiBotiga = "76"
    End If
    
    If IsNull(Rs("fecmod")) = False Then
        ultimaMod = Rs("fecmod")
        'Se borra del tiempo mas alla de los segundos
        ultimaMod = Mid(ultimaMod, 1, Len(ultimaMod) - 8)
        ultimaMod = CDate(ultimaMod)
    End If
    ultimaMod = Now()
    nomBotiga = nomBotiga & "_" & nomBalanza
    nomBotiga = Replace(nomBotiga, "'", " ")
    codiBotigaOriginal = codiBotiga
    'CodiClient+CodiBalança
    codiBotiga = codiBotiga & codiBalanza
    'nomTarifa
    sql = "SELECT TarifaNom from TarifesEspecials where TarifaCodi='" & idTarifa & "'"
    Set Rs2 = Db.OpenResultset(sql)
    If Not Rs2.EOF Then nomTarifa = Rs2("TarifaNom")
    'Codi client
    codiNum = 0
    sql = "SELECT isNull(Codi,'0') codi FROM [Fac_LaForneria].[dbo].[Clients] WHERE Codi='" & codiBotiga & "' "
    Set Rs2 = Db.OpenResultset(sql)
    If Not Rs2.EOF Then codiNum = Rs2("codi")
    'Diferencia fechas para saber si hay que actualizar o insertar
    diff = DateDiff("s", fecMod, ultimaMod)
    If IsNull(diff) = True Then diff = 1
    If diff > 0 Then 'Se debe actualizar o insertar
'--------------------------------------------------------------------------------------------
InformaMiss "REGENERACION DE TARIFAS Y PROMOCIONES " & codiBotiga & ",TARIFA " & idTarifa & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "REGENERACION DE TARIFAS Y PROMOCIONES " & codiBotiga & ",TARIFA " & idTarifa & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'------------------------------------------------------------------------------------------
'------ TARIFAS
'------------------------------------------------------------------------------------------
        'Borrando tarifas para la tienda actual
        sql = "DELETE from ParamsTpv where Variable = 'Tarifa' and CodiClient  = '" & codiBotiga & "' "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'METEMOS SOLO LAS TARIFAS QUE TENGAN VINCULADO ALGUN ARTICULO
        sql = "'SELECT distinct dat_tarifa.IdTarifa,dat_tarifa.NombreTarifa, dat_tarifa_tienda.IdTienda, idbalanza from dat_tarifa_tienda left join dat_tarifa "
        sql = sql & "on dat_tarifa_tienda.IdTarifa = dat_Tarifa.IdTarifa "
        sql = sql & "LEFT JOIN dat_balanza b on (b.idTienda=dat_tarifa_tienda.idTienda) "
        sql = sql & "where dat_tarifa_tienda.IdTienda = ''" & codiBotigaOriginal & "'' and dat_tarifa.idArticulo<>0 ' "
        sql2 = "INSERT INTO [fac_laforneria].dbo.[paramstpv] "
        sql2 = sql2 & "SELECT cast(idTienda as varchar) + cast(idbalanza as varchar), 'Tarifa', '1,' + left(cast(IdTarifa as varchar) + NombreTarifa,20) "
        sql2 = sql2 & "FROM openquery(AMETLLER, " & sql & ") "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql2 & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql2)
        End If
        'METEMOS LAS MISMAS TARIFAS PERO CON UNA O DELANTE PARA POSIBLES OFERTAS
        sql = "'SELECT distinct dat_tarifa.IdTarifa,dat_tarifa.NombreTarifa, dat_tarifa_tienda.IdTienda, idbalanza from dat_tarifa_tienda left join dat_tarifa "
        sql = sql & "on dat_tarifa_tienda.IdTarifa = dat_Tarifa.IdTarifa "
        sql = sql & "LEFT JOIN dat_balanza b on (b.idTienda=dat_tarifa_tienda.idTienda) "
        sql = sql & "where dat_tarifa_tienda.IdTienda = ''" & codiBotigaOriginal & "'' and dat_tarifa.idArticulo<>0 ' "
        sql2 = "INSERT INTO [fac_laforneria].dbo.[paramstpv] "
        sql2 = sql2 & "SELECT cast(idTienda as varchar) + cast(idbalanza as varchar), 'Tarifa', '1,' + left('O_' + cast(cast(IdTarifa as integer)+5000 as varchar) + NombreTarifa,20) "
        sql2 = sql2 & "FROM openquery(AMETLLER, " & sql & " )"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql2 & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql2)
        End If
        'Insert Tarifa Principal como tarifa oferta
        sql = "SELECT TarifaNom from TarifesEspecials where TarifaCodi=cast('" & idTarifa & "' as integer)+5000 group by TarifaNom "
        Set Rs2 = Db.OpenResultset(sql)
        If Not Rs2.EOF Then
            nomTarifa = Rs2("TarifaNom")
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[ParamsTpv] ([CodiClient],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiBotiga & "', 'Tarifa','1," & nomTarifa & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
        End If
'------------------------------------------------------------------------------------------
        'Borrando tarifas sin articulos
        sql = "DELETE from [Fac_LaForneria].[dbo].[ParamsTpv] where CodiClient = '" & codiBotiga & "' "
        sql = sql & "and Variable = 'Tarifa' and SUBSTRING(Valor,3,LEN(valor)) not in "
        sql = sql & "(select ltrim(rtrim(tarifanom)) from [Fac_LaForneria].[dbo].[TarifesEspecials]) "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
'------------------------------------------------------------------------------------------
'--Promocions
'------------------------------------------------------------------------------------------
        'Borrando Productes Promocionats
        sql = "DELETE FROM [Fac_LaForneria].[dbo].[ProductesPromocionats] where client='" & codiBotiga & "' "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Insertando Productes Promocionats de Articles
        sql = " INSERT INTO [Fac_LaForneria].[dbo].[ProductesPromocionats] "
        sql = sql & "SELECT NEWID(),'01/01/2007 00:00:00.000','01/01/2007 23:59:00.000', "
        sql = sql & "t.CodiArticle,t.Cantidad,'-1','-1',t.Precio,'" & codiBotiga & "' FROM ( "
        sql = sql & "SELECT s.idTienda,ap.CodiArticle,s.Cantidad,s.Precio "
        sql = sql & "FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] ap "
        sql = sql & "LEFT JOIN (SELECT idTienda,idArticulo,Cantidad,Precio FROM openquery(AMETLLER,"
        sql = sql & "'SELECT ''" & codiBotiga & "'' as idTienda,s.idArticulo,s.Cantidad,s.Precio "
        sql = sql & "FROM dat_articulo_segmentos s where Cantidad <> 0 and Precio <> 0 "
        sql = sql & "')) s on (ap.Valor=s.idArticulo and ap.Variable='CODI_PROD') "
        sql = sql & "WHERE s.idArticulo is not null) t "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
'------------------------------------------------------------------------------------------
        'Delete Productes Promocionats de Tarifa Tienda
        sql = "DELETE FROM [Fac_LaForneria].[dbo].[ProductesPromocionats] "
        sql = sql & "WHERE Client='" & codiBotiga & "' and D_Producte in ( "
        sql = sql & "SELECT ap.CodiArticle FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] ap "
        sql = sql & "LEFT JOIN (SELECT idArticulo FROM openquery(AMETLLER,"
        sql = sql & "'SELECT s.idArticulo FROM dat_tarifa_segmentos s "
        sql = sql & "LEFT JOIN dat_tarifa_tienda t on (s.idTarifa=t.idTarifa) "
        sql = sql & "WHERE t.idTienda is not null and t.idTienda= ''" & codiBotigaOriginal & "'' "
        sql = sql & "')) s on (ap.Valor=s.idArticulo and ap.Variable='CODI_PROD') "
        sql = sql & "WHERE s.idArticulo is not null ) "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Delete Productes Promocionats de Tarifes Client
        sql = "DELETE FROM [Fac_LaForneria].[dbo].[ProductesPromocionats] "
        sql = sql & "WHERE Client='" & codiBotiga & "' and D_Producte in ( "
        sql = sql & "SELECT ap.CodiArticle FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] ap "
        sql = sql & "LEFT JOIN (SELECT * FROM openquery(AMETLLER,"
        sql = sql & "'select s.idArticulo from dat_tienda t "
        sql = sql & "left join dat_tarifa_segmentos s on (s.idtarifa=t.idtarifa) "
        sql = sql & "WHERE t.idTienda is not null and t.idTienda=''" & codiBotigaOriginal & "'' "
        sql = sql & "')) s on (ap.Valor=s.idArticulo and ap.Variable='CODI_PROD') "
        sql = sql & "WHERE s.idArticulo is not null) "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Delete Productes Promocionats de articles sense tarifa
        sql = "DELETE FROM [Fac_LaForneria].[dbo].[ProductesPromocionats] "
        sql = sql & "WHERE Client='" & codiBotiga & "' and D_Producte in ( "
        sql = sql & "SELECT ap.CodiArticle FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] ap "
        sql = sql & "LEFT JOIN (SELECT * FROM openquery(AMETLLER,"
        sql = sql & "'select s.idArticulo from dat_articulo_segmentos s left join dat_articulo a on s.idArticulo=a.idArticulo "
        sql = sql & "where s.idEmpresa=1 and a.idEmpresa=1')) s on (ap.Valor=s.idArticulo and ap.Variable='CODI_PROD') "
        sql = sql & "WHERE s.idArticulo is not null )"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
'------------------------------------------------------------------------------------------
        'Insertando Productes Promocionats de Tarifa Tienda
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[ProductesPromocionats] "
        sql = sql & "SELECT NEWID(),'01/01/2007 00:00:00.000','01/01/2007 23:59:00.000',"
        sql = sql & "t.CodiArticle,t.Cantidad,'-1','-1',t.Precio,'" & codiBotiga & "' FROM ( "
        sql = sql & "SELECT s.idTienda,ap.CodiArticle,s.Cantidad,s.Precio "
        sql = sql & "FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] ap LEFT JOIN "
        sql = sql & "(SELECT idTienda,idArticulo,idTarifa,Cantidad,Precio FROM openquery(AMETLLER,"
        sql = sql & "'SELECT t.idTienda,s.idArticulo,s.idTarifa,s.Cantidad,s.Precio "
        sql = sql & "FROM dat_tarifa_segmentos s LEFT JOIN dat_tarifa_tienda t on (s.idTarifa=t.idTarifa) "
        sql = sql & "WHERE t.idTienda is not null and t.idTienda=''" & codiBotigaOriginal & "'' "
        sql = sql & "')) s on (ap.Valor=s.idArticulo and ap.Variable='CODI_PROD') "
        sql = sql & "WHERE s.idArticulo is not null) t "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Insertando Productes Promocionats de Tarifes Client
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[ProductesPromocionats] "
        sql = sql & "SELECT NEWID(),'01/01/2007 00:00:00.000','01/01/2007 23:59:00.000',"
        sql = sql & "t.CodiArticle,t.Cantidad,'-1','-1',t.Precio, '" & codiBotiga & "' FROM ( "
        sql = sql & "SELECT s.idTienda,ap.CodiArticle,s.Cantidad,s.Precio "
        sql = sql & "FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] ap "
        sql = sql & "LEFT JOIN (SELECT * FROM openquery(AMETLLER,"
        sql = sql & "'select t.idtienda,s.idArticulo,s.idTarifa,s.Cantidad,s.Precio "
        sql = sql & "from dat_tienda t left join dat_tarifa_segmentos s on (s.idtarifa=t.idtarifa) "
        sql = sql & "WHERE t.idTienda is not null and t.idTienda=''" & codiBotigaOriginal & "'' "
        sql = sql & "')) s on (ap.Valor=s.idArticulo and ap.Variable='CODI_PROD') "
        sql = sql & "WHERE s.idArticulo is not null ) t"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Insertando Productes Promocionats de articles sense tarifa
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[ProductesPromocionats] "
        sql = sql & "SELECT NEWID(),'01/01/2007 00:00:00.000','01/01/2007 23:59:00.000',"
        sql = sql & "t.CodiArticle,t.Cantidad,'-1','-1',t.Precio,'" & codiBotiga & "' FROM ( "
        sql = sql & "SELECT ap.CodiArticle,s.Cantidad,s.Precio "
        sql = sql & "FROM [Fac_LaForneria].[dbo].[ArticlesPropietats] ap "
        sql = sql & "LEFT JOIN (SELECT * FROM openquery(AMETLLER,"
        sql = sql & "'select s.idArticulo,s.Cantidad,s.Precio  "
        sql = sql & "from dat_articulo_segmentos s left join dat_articulo a on s.idArticulo=a.idArticulo "
        sql = sql & "where s.idEmpresa=1 and a.idEmpresa=1')) s on (ap.Valor=s.idArticulo and ap.Variable='CODI_PROD') "
        sql = sql & "WHERE s.idArticulo is not null ) t"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
'------------------------------------------------------------------------------------------
        'Revision de promocions duplicadas
        'Eliminamos duplicados de articulos+tienda en el qual las promociones sean las menos provechosas para el cliente
        sql = "DELETE FROM ProductesPromocionatsbk WHERE ID IN ( "
        sql = sql & "SELECT pp.id FROM ProductesPromocionatsbk pp LEFT JOIN( "
        sql = sql & "select D_producte,MIN(D_Quantitat*S_Preu) minimo from ProductesPromocionatsbk  "
        sql = sql & "group by D_Producte,client having COUNT(d_producte)>=2 ) ppm "
        sql = sql & "ON (pp.D_Producte=ppm.D_Producte) WHERE (pp.D_Quantitat*pp.S_Preu)>ppm.minimo) "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Eliminamos duplicados de articulos+tienda que contengan la misma promocion
        sql = "DELETE FROM ProductesPromocionatsbk WHERE ID IN ( "
        sql = sql & "SELECT pp.id FROM ProductesPromocionatsbk pp LEFT JOIN( "
        sql = sql & "select D_producte,MIN(id) minimo from ProductesPromocionatsbk "
        sql = sql & "group by D_Producte,client having COUNT(d_producte)>=2 "
        sql = sql & ") ppm ON (pp.D_Producte=ppm.D_Producte) WHERE pp.id<>ppm.minimo) "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        insBot = 1
'------------------------------------------------------------------------------------------
        'missatgesaenviar
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) "
        sql = sql & "VALUES ('Tpv_Configuracio_',(select codi from ParamsHw where Valor1 = '" & codiBotiga & "')) "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
    End If
    Rs.MoveNext
Loop
'----------------------------------------------------------------------------------------------------
'Missatgesaenviar
sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) "
sql = sql & "VALUES ('ProductesPromocionats','')"
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "SQL:" & sql & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
Else
    Set Rs2 = Db.OpenResultset(sql)
End If
'--------------------------------------------------------------------------------
connMysql.Close
Set connMysql = Nothing
'----------------------------------------------------------------------------------------------------
'--UPDATE SINCRO_LAFORNERIA
'----------------------------------------------------------------------------------------------------
'Se indica que el proceso actual ha acabado
Set Rs = Db.OpenResultset("Delete from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='BOTIGUES'")
Set Rs = Db.OpenResultset("Insert into [Fac_LaForneria].[dbo].[sincro_laforneria] (variable,fecha,p1) values ('BOTIGUES',getdate(),1)")
'Se mira si es el ultimo proceso activo, si es asi se actualiza la fecha de la variable fecmod
'que es por la que se rigen los procesos.
sql = "Select COUNT(variable) num from [Fac_LaForneria].[dbo].[sincro_laforneria] where p1=1 and variable<>'FECMOD' "
sql = sql & "and fecha>=(select convert(datetime,p2,103) from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='fecmod') "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    'Si se han completado los cinco procesos se actualiza
    If Rs("num") = 5 Then
        Set Rs2 = Db.OpenResultset("Update [Fac_LaForneria].[dbo].[sincro_laforneria] Set [fecha] = [P2] WHERE [Variable]='FECMOD'")
        Set Rs2 = Db.OpenResultset("Update feinesafer set Param3='Fi " & Now() & "' where tipus='SincroDbAmetllerHorari' and param2 like '%Si%' ")
    End If
End If
'----------------------------------------------------------------------------------------------------
InformaMiss "FIN SINCRO_BOTIGUES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "FIN SINCRO_BOTIGUES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.Close
End If
Exit Function

norBot:
'----------------------------------------------------------------------------------------------------
'----  SI HAY ERROR
'----------------------------------------------------------------------------------------------------
'Se indica que el proceso actual ha fallado
    Set Rs = Db.OpenResultset("Delete from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='BOTIGUES'")
    Set Rs = Db.OpenResultset("Insert into [Fac_LaForneria].[dbo].[sincro_laforneria] (variable,fecha,p1) values ('BOTIGUES',getdate(),0)")
    html = "<p><h3>Resum Botigues Ametller </h3></p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & dataIni & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    html = html & "<p><b>Ultima sqlSP:</b>" & sqlSP & "</p>"
    
    sf_enviarMail "secrehit@hit.cat", p1, "ERROR! Sincronitzacio de botigues ha fallat", html, "", ""
            
    ExecutaComandaSql "Delete from  " & TaulaCalculsEspecials & " Id = '" & idTasca & "' "
    
    Informa "Error : " & err.Description
    
End Function


Function SincroDbAmetllerClients(p1, idTasca) As Boolean
'*************************************************************************************
'SincroDbAmetllerClients
'Sincroniza clientes entre servidor Ametller y servidor Hit
'La variables CodiContable y CodiClientOrigen (contienen lo mismo) de la tabla constantsClient
'ligan el codigo cliente de Hit con el codigo cliente de Ametller.
'Proceso:
'1.Se recoge fecha ultima modificacion de la tabla sincro_laforneria.
'2.Se abre un bucle mirando los tiendas de ametller para que tambien esten como clientes.
'3.En cada tienda se mira si ya existe en hit. Si existe actualiza sino inserta.
'4.Se abre un bucle mirando los clientes de ametller.
'5.En cada cliente se mira si ya existe en hit. Si existe actualiza sino inserta.
'6.Se inserta linea en MissatgesAEnviar.
'*************************************************************************************
Dim codiBotiga As String, nomBotiga As String, idTarifa As String
Dim codiExist As String, cifBotiga As String, adresaBotiga As String, cpBotiga As String, ciutatBotiga As String
Dim codiNum As Integer, fecMod As Date, ultimaMod As String, diff As Long, insCli As Integer, updCli As Integer
Dim idTabla As String, debugSincro As Boolean, sql As String, sql2 As String, sql3 As String, sqlSP As String, fecha As Date
Dim tablaTmp As String, Rs As rdoResultset, Rs2 As rdoResultset, html As String, tablaTmp2 As String, tablaTmp3 As String
Dim connMysql As ADODB.Connection, dataIni As Date, idCF As String
dataIni = Now()
codiNum = 0
insCli = 0
updCli = 0

On Error GoTo norCli

'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
'Direccion email a la que se enviara resultado/error
If p1 = "" Then p1 = EmailGuardia
debugSincro = False 'Solo impresion o ejecucion completa
'Crear txt log
If debugSincro = True Then
    p1 = EmailGuardia
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\Clients_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
End If
'--------------------------------------------------------------
InformaMiss "INICIO SINCRO_CLIENTS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "INICIO SINCRO_CLIENTS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'--ULTIMA FECHA MODIFICACION
'--------------------------------------------------------------------------------
InformaMiss "ULTIMA FECHA MODIFICACION-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "ULTIMA FECHA MODIFICACION-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
sql = "SELECT fecha FROM [Fac_LaForneria].[dbo].[sincro_laforneria] WHERE Variable='FECMOD' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    fecMod = Rs("fecha")
Else
    'Comprobamos que existe tabla sincro_laforneria
    tablaTmp = "[Fac_laforneria].[dbo].[sincro_laforneria]"
    sql = "SELECT object_id FROM sys.objects with (nolock) "
    sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp & "','[',''),']','') AND type='U' "
    Set Rs = Db.OpenResultset(sql)
    If Rs.EOF Then
        sql = "CREATE TABLE [Fac_laforneria].[dbo].[sincro_laforneria]("
        sql = sql & "       [variable] [varchar](255) NULL,"
        sql = sql & "       [fecha] [datetime] NULL,"
        sql = sql & "       [p1] [bit] NULL,"
        sql = sql & "       [p2] [nvarchar] (255) NULL,"
        sql = sql & "       [p3] [nvarchar] (255) NULL,"
        sql = sql & "       [p4] [nvarchar] (255) NULL,"
        sql = sql & "       [p5] [nvarchar] (255) NULL"
        sql = sql & "   ) ON [PRIMARY]"
        Set Rs = Db.OpenResultset(sql)
        sql = "INSERT INTO [Fac_laforneria].[dbo].[sincro_laforneria] values ("
        sql = sql & "'FECMOD',GETDATE(),0,GETDATE(),NULL,NULL,NULL) "
        Set Rs = Db.OpenResultset(sql)
        sql = "SELECT fecha FROM [Fac_LaForneria].[dbo].[sincro_laforneria] WHERE Variable='FECMOD' "
        Set Rs = Db.OpenResultset(sql)
        If Not Rs.EOF Then fecMod = Rs("fecha")
    End If
End If
'--------------------------------------------------------------------------------
html = "<p><h3>Resum clients Ametller</h3></p>"
'--------------------------------------------------------------------------------
'--BUCLE TIENDAS
'--------------------------------------------------------------------------------
InformaMiss "BUCLE TIENDAS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "BUCLE TIENDAS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
sql = "SELECT * FROM openquery(AMETLLER,' "
sql = sql & "SELECT IdTienda AS CodiBotiga, TRIM(Nombre) AS Botiga, IdTarifa AS CodiTarifa,"
sql = sql & "TRIM(CIF_VAT) AS Cif, TRIM(Direccion) AS Dir, TRIM(CodPostal) AS CP, TRIM(Poblacion) AS Ciutat, "
sql = sql & "TimeStamp AS fecMod FROM dat_tienda WHERE IdEmpresa=1')"
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
Do While Not Rs.EOF
    If IsNull(Rs("CodiBotiga")) = False Then codiBotiga = Rs("CodiBotiga")
    If IsNull(Rs("Botiga")) = False Then nomBotiga = Rs("Botiga")
    If IsNull(Rs("CodiTarifa")) = False Then idTarifa = Rs("CodiTarifa")
    If IsNull(Rs("Cif")) = False Then cifBotiga = Rs("Cif")
    If IsNull(Rs("Dir")) = False Then adresaBotiga = Rs("Dir")
    If IsNull(Rs("CP")) = False Then cpBotiga = Rs("CP")
    If IsNull(Rs("Ciutat")) = False Then ciutatBotiga = Rs("Ciutat")
    If IsNull(Rs("fecmod")) = False Then
        ultimaMod = Rs("fecmod")
        'Se borra del tiempo mas alla de los segundos
        ultimaMod = Mid(ultimaMod, 1, Len(ultimaMod) - 8)
        ultimaMod = CDate(ultimaMod)
    End If
    nomBotiga = Replace(nomBotiga, "'", " ")
    adresaBotiga = Replace(adresaBotiga, "'", " ")
    ciutatBotiga = Replace(ciutatBotiga, "'", " ")
    'Codi client, seleccionem id si existeix client i mirem de no matxacar botigues
    codiNum = 0
    sql = "SELECT isNull(Codi,'0') codi FROM [Fac_LaForneria].[dbo].[ConstantsClient] "
    sql = sql & "WHERE Variable='CodiContable' and Valor='" & codiBotiga & "' and NOT EXISTS ("
    sql = sql & "SELECT Valor1 FROM [Fac_LaForneria].[dbo].[ParamsHw] WHERE Valor1='" & codiBotiga & "') "
    Set Rs2 = Db.OpenResultset(sql)
    If Not Rs2.EOF Then codiNum = Rs2("codi")
    'Diferencia fechas para saber si hay que actualizar o insertar
    diff = DateDiff("s", fecMod, ultimaMod)
    If IsNull(diff) = True Then diff = 1
    If diff > 0 Then 'Se debe actualizar o insertar
        'Sino existe el codigo de esta tienda, se inserta tienda
        If codiNum = 0 Then
            'codiClient
            sql = "SELECT top 1 CAST(Codi as integer)+1 codi from [Fac_LaForneria].[dbo].[clients] where codi<9000 order by codi desc"
            Set Rs2 = Db.OpenResultset(sql)
            If Not Rs2.EOF Then codiNum = Rs2("codi")
            'insert clients
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[clients] ("
            sql = sql & "[Codi],[Nom],[Nif],[Adresa],[Ciutat],[Cp],[Lliure],[Nom Llarg],[Tipus Iva],[Preu Base],"
            sql = sql & "[Desconte ProntoPago],[Desconte 1],[Desconte 2],[Desconte 3],[Desconte 4],[Desconte 5],[AlbaraValorat]) "
            sql = sql & "VALUES ('" & codiNum & "','" & nomBotiga & "','" & cifBotiga & "','" & adresaBotiga & "','" & ciutatBotiga & "','" & cpBotiga & "','','" & nomBotiga & "','2','2',"
            sql = sql & "'0','0','0','0','0','" & idTarifa & "',NULL)"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'insert constantsClients
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiNum & "','CodiContable','" & codiBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiNum & "','CodiClientOrigen','" & codiBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            insCli = 1
'--------------------------------------------------------------------------------
InformaMiss "CLIENTE TIENDA INSERTADO " & codiBotiga & "," & nomBotiga & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CLIENTE TIENDA INSERTADO " & codiBotiga & "," & nomBotiga & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
        Else
            'Update clients
            sql = "UPDATE [Fac_LaForneria].[dbo].[clients] "
            sql = sql & "SET [NOM]='" & nomBotiga & "',"
            sql = sql & "   [Nif]='" & cifBotiga & "',"
            sql = sql & "   [Adresa]='" & adresaBotiga & "',"
            sql = sql & "   [Ciutat]='" & ciutatBotiga & "',"
            sql = sql & "   [Cp]='" & cpBotiga & "',"
            sql = sql & "   [Lliure]='',"
            sql = sql & "   [Nom Llarg]='" & nomBotiga & "',"
            sql = sql & "   [Tipus Iva]='2',"
            sql = sql & "   [Preu Base]='2',"
            sql = sql & "   [Desconte ProntoPago]='0',"
            sql = sql & "   [Desconte 5]='" & idTarifa & "',"
            sql = sql & "   [AlbaraValorat]=NULL "
            sql = sql & "WHERE [Codi] in ( "
            sql = sql & "   SELECT [Codi] FROM [Fac_LaForneria].[dbo].[ConstantsClient] "
            sql = sql & "   WHERE [Variable]='CodiContable' and [Valor]='" & codiBotiga & "') "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            updCli = 1
'--------------------------------------------------------------------------------
InformaMiss "CLIENTE TIENDA ACTUALIZADO " & codiBotiga & "," & nomBotiga & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CLIENTE TIENDA ACTUALIZADO " & codiBotiga & "," & nomBotiga & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
        End If
    End If
    Rs.MoveNext
Loop
'--------------------------------------------------------------------------------
'--BUCLE CLIENTES
'--------------------------------------------------------------------------------
InformaMiss "BUCLE CLIENTES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "BUCLE CLIENTES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
sql = "SELECT * FROM openquery(AMETLLER,'"
sql = sql & "SELECT idCliente AS CodiBotiga, Nombre as botiga, Ofertas AS CodiTarifa,"
sql = sql & "TRIM(DNI) AS Cif, TRIM(Direccion) AS Dir, TRIM(CodPostal) AS CP, TRIM(Poblacion) AS Ciutat,TimeStamp AS fecMod "
sql = sql & "FROM dat_cliente WHERE IdEmpresa=1')"
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
Do While Not Rs.EOF
    If IsNull(Rs("CodiBotiga")) = False Then codiBotiga = Rs("CodiBotiga")
    If IsNull(Rs("Botiga")) = False Then nomBotiga = Rs("Botiga")
    If IsNull(Rs("CodiTarifa")) = False Then idTarifa = Rs("CodiTarifa")
    If IsNull(Rs("Cif")) = False Then cifBotiga = Rs("Cif")
    If IsNull(Rs("Dir")) = False Then adresaBotiga = Rs("Dir")
    If IsNull(Rs("CP")) = False Then cpBotiga = Rs("CP")
    If IsNull(Rs("Ciutat")) = False Then ciutatBotiga = Rs("Ciutat")
    If IsNull(Rs("fecmod")) = False Then
        ultimaMod = Rs("fecmod")
        'Se borra del tiempo mas alla de los segundos
        ultimaMod = Mid(ultimaMod, 1, Len(ultimaMod) - 8)
        ultimaMod = CDate(ultimaMod)
    End If
    nomBotiga = Replace(nomBotiga, "'", " ")
    adresaBotiga = Replace(adresaBotiga, "'", " ")
    'Codi client, seleccionem id si existeix client i mirem de no matxacar botigues
    codiNum = 0
    sql = "SELECT isNull(Codi,'0') codi FROM [Fac_LaForneria].[dbo].[ConstantsClient] "
    sql = sql & "WHERE Variable='CodiContable' and Valor='" & codiBotiga & "' "
    Set Rs2 = Db.OpenResultset(sql)
    If Not Rs2.EOF Then codiNum = Rs2("codi")
    'Diferencia fechas para saber si hay que actualizar o insertar
    diff = DateDiff("s", fecMod, ultimaMod)
    If IsNull(diff) = True Then diff = 1
    If diff > 0 Then 'Se debe actualizar o insertar
        'Sino existe el codigo de esta tienda, se inserta tienda
        If codiNum = 0 Then
            'codiClient
            sql = "SELECT top 1 CAST(Codi as integer)+1 codi from [Fac_LaForneria].[dbo].[clients] where codi<9000 order by codi desc "
            Set Rs2 = Db.OpenResultset(sql)
            If Not Rs2.EOF Then codiNum = Rs2("codi")
            'insert clients
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[clients] ("
            sql = sql & "[Codi],[Nom],[Nif],[Adresa],[Ciutat],[Cp],[Lliure],[Nom Llarg],[Tipus Iva],[Preu Base],"
            sql = sql & "[Desconte ProntoPago],[Desconte 1],[Desconte 2],[Desconte 3],[Desconte 4],[Desconte 5],[AlbaraValorat]) "
            sql = sql & "VALUES ('" & codiNum & "','" & nomBotiga & "','" & cifBotiga & "','" & adresaBotiga & "','" & ciutatBotiga & "','" & cpBotiga & "','','" & nomBotiga & "','2','1',"
            sql = sql & "'0','0','0','0','0','" & idTarifa & "',NULL)"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'insert constantsClients
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiNum & "','CodiContable','" & codiBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiNum & "','CodiClientOrigen','" & codiBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiNum & "','AlbaransValorats','AlbaransValorats')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiNum & "','COPIES_ALB','1')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiNum & "','EsClient','EsClient')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiNum & "','NomClientFactura','FISCAL')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "SELECT 'CliBoti_000_'+CAST(NEWID() as nvarchar(255)) id "
            Set Rs2 = Db.OpenResultset(sql)
            If Not Rs2.EOF Then idCF = Rs2("id")
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[constantsclient] ([Codi],[Variable],[Valor]) "
            sql = sql & "VALUES ('" & codiNum & "','CFINAL','" & idCF & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            sql = "INSERT INTO [Fac_LaForneria].[dbo].[clientsFinals] ([Id],[idExterna],[Nom],[Telefon],[Adreca],[emili],[Descompte],[Altres],[Nif]) "
            sql = sql & "VALUES ('" & idCF & "',NULL,'" & nomBotiga & "',NULL,NULL,NULL,NULL,NULL,NULL)"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            insCli = 1
'--------------------------------------------------------------------------------
InformaMiss "CLIENTE INSERTADO " & codiBotiga & "," & nomBotiga & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CLIENTE INSERTADO " & codiBotiga & "," & nomBotiga & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
        Else
            'Update clients
            sql = "UPDATE [Fac_LaForneria].[dbo].[clients] "
            sql = sql & "SET [NOM]='" & nomBotiga & "',"
            sql = sql & "   [Nif]='" & cifBotiga & "',"
            sql = sql & "   [Adresa]='" & adresaBotiga & "',"
            sql = sql & "   [Ciutat]='" & ciutatBotiga & "',"
            sql = sql & "   [Cp]='" & cpBotiga & "',"
            sql = sql & "   [Lliure]='',"
            sql = sql & "   [Nom Llarg]='" & nomBotiga & "',"
            sql = sql & "   [Tipus Iva]='2',"
            sql = sql & "   [Preu Base]='1',"
            sql = sql & "   [Desconte ProntoPago]='0',"
            sql = sql & "   [Desconte 5]='" & idTarifa & "',"
            sql = sql & "   [AlbaraValorat]=NULL "
            sql = sql & "WHERE [Codi] in ("
            sql = sql & "   SELECT [Codi] FROM [Fac_LaForneria].[dbo].[ConstantsClient] "
            sql = sql & "   WHERE [Variable]='CodiContable' and [Valor]='" & codiBotiga & "')"
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            'Update clientsFinals
            sql = "UPDATE [Fac_LaForneria].[dbo].[clientsFinals] SET Nom='" & nomBotiga & "' "
            sql = sql & "WHERE Id in (SELECT valor FROM [Fac_LaForneria].[dbo].[ConstantsClient] WHERE "
            sql = sql & "Variable='CFINAL' and Codi='" & codiNum & "') "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            updCli = 1
'--------------------------------------------------------------------------------
InformaMiss "CLIENTE ACTUALIZADO " & codiBotiga & "," & nomBotiga & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CLIENTE ACTUALIZADO " & codiBotiga & "," & nomBotiga & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
        End If
    End If
    Rs.MoveNext
Loop
'--------------------------------------------------------------------------------
If insCli > 0 Or updCli > 0 Then
    'Missatgeaenviar
    sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) VALUES ('Clients','')"
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        Set Rs2 = Db.OpenResultset(sql)
    End If
    sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) SELECT 'ClientsFinals' as tipus, id param FROM clientsfinals"
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        Set Rs2 = Db.OpenResultset(sql)
    End If
End If
'--------------------------------------------------------------------------------
connMysql.Close
Set connMysql = Nothing
'----------------------------------------------------------------------------------------------------
'--UPDATE SINCRO_LAFORNERIA
'----------------------------------------------------------------------------------------------------
'Se indica que el proceso actual ha acabado
Set Rs = Db.OpenResultset("Delete from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='CLIENTS'")
Set Rs = Db.OpenResultset("Insert into [Fac_LaForneria].[dbo].[sincro_laforneria] (variable,fecha,p1) values ('CLIENTS',getdate(),1)")
'Se mira si es el ultimo proceso activo, si es asi se actualiza la fecha de la variable fecmod
'que es por la que se rigen los procesos.
sql = "Select COUNT(variable) num from [Fac_LaForneria].[dbo].[sincro_laforneria] where p1=1 and variable<>'FECMOD' "
sql = sql & "and fecha>=(select convert(datetime,p2,103) from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='fecmod') "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    'Si se han completado los cinco procesos se actualiza
    If Rs("num") = 5 Then
        Set Rs2 = Db.OpenResultset("Update [Fac_LaForneria].[dbo].[sincro_laforneria] Set [fecha] = [P2] WHERE [Variable]='FECMOD'")
        Set Rs2 = Db.OpenResultset("Update feinesafer set Param3='Fi " & Now() & "' where tipus='SincroDbAmetllerHorari' and param2 like '%Si%' ")
    End If
End If
'----------------------------------------------------------------------------------------------------
InformaMiss "FIN SINCRO_CLIENTS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "FIN SINCRO_CLIENTS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.Close
End If
Exit Function

norCli:
'----------------------------------------------------------------------------------------------------
'----  SI HAY ERROR
'----------------------------------------------------------------------------------------------------
'Se indica que el proceso actual ha fallado
    Set Rs = Db.OpenResultset("Delete from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='CLIENTS'")
    Set Rs = Db.OpenResultset("Insert into [Fac_LaForneria].[dbo].[sincro_laforneria] (variable,fecha,p1) values ('CLIENTS',getdate(),0)")
    html = "<p><h3>Resum Clients Ametller </h3></p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & dataIni & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    html = html & "<p><b>Ultima sqlSP:</b>" & sqlSP & "</p>"
    
    sf_enviarMail "secrehit@hit.cat", p1, "ERROR! Sincronitzacio de clients ha fallat", html, "", ""
            
    ExecutaComandaSql "Delete from  " & TaulaCalculsEspecials & " Id = '" & idTasca & "' "
    
    Informa "Error : " & err.Description
    
End Function


Function SincroDbAmetllerDependentes(p1, idTasca) As Boolean
'*************************************************************************************
'SincroDbAmetllerDependentes
'Sincroniza dependientas entre servidor Ametller y servidor Hit
'La variable CODI_DEP de la tabla dependentesExtes liga el codigo dependienta de Hit con
'el codigo de dependienta de Ametller.
'Proceso:
'1.Crea tablas temporales con datos de mysql ametller.
'2.Actualiza las dependientas en las que ha encontrado diferencias entre la
'   tabla temporal y la tabla dependentes/dependentesExtes.
'3.Se vacia la tabla temporal y se vuelve a rellenar solo con los nuevos.
'4.Se insertan los nuevos.
'5.Se borran tablas temporales y se envia tarea a MissatgesAEnviar si se ha insertado/actualizado algo.
'*************************************************************************************
Dim codiDep As String, NomDep As String, telfDep As String, emailDep As String, tipusDep As String, codiTargeta As String
Dim codiNum As Integer, updDep As Integer, insDep As Integer, fecMod As Date, ultimaMod As String, diff As Long, tmp As String, proc As String
Dim idTabla As String, debugSincro As Boolean, sql As String, sql2 As String, sql3 As String, sqlSP As String
Dim tablaTmp As String, Rs As rdoResultset, Rs2 As rdoResultset, html As String, tablaTmp2 As String, tablaTmp3 As String
Dim connMysql As ADODB.Connection, dataIni As Date
dataIni = Now()
codiNum = 0
insDep = 0
    
On Error GoTo norDep

'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
If p1 = "" Then p1 = EmailGuardia
debugSincro = False 'Solo impresion o ejecucion completa
'Crear txt log
If debugSincro = True Then
    p1 = EmailGuardia
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\Dependentes_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
End If
'--------------------------------------------------------------
InformaMiss "INICIO SINCRO_DEPENDENTES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "INICIO SINCRO_DEPENDENTES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'--TABLAS TEMPORALES
'--------------------------------------------------------------------------------
'Creamos tablas temporales
InformaMiss "CREANDO TABLAS TEMPORALES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "CREANDO TABLAS TEMPORALES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'Si existe sincro_dependentesTmp la borramos y volvemos a generar
tablaTmp = "[Fac_laforneria].[dbo].[sincro_dependentesTmp]"
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp)
sql = " SELECT * INTO " & tablaTmp & " FROM OPENQUERY(AMETLLER,' "
sql = sql & "SELECT a.IdArticulo AS CodiTreballador,a.Descripcion AS Treballador, "
sql = sql & "CONCAT(REPEAT(''0'',4-LENGTH(REPLACE(FORMAT(a.PrecioConIVA*100,0),'','',''''))), "
sql = sql & "REPLACE(FORMAT(a.PrecioConIVA*100,0),'','','''')) AS PIN, e.job_title_code "
sql = sql & "AS CodiLlocDeTreball, j.jobtit_name AS LlocDeTreball,emp_mobile AS Mobil, "
sql = sql & "emp_work_email AS Email,ean.EANScanner AS CodiTargeta "
sql = sql & "FROM dat_articulo a "
sql = sql & "JOIN casaametller_orangehrm.hs_hr_employee e ON e.employee_id=a.IdArticulo "
sql = sql & "LEFT JOIN casaametller_orangehrm.hs_hr_job_title j ON j.jobtit_code=e.job_title_code "
sql = sql & "LEFT JOIN dat_articulo_eanscanner ean on ean.idArticulo=a.IdArticulo "
sql = sql & "Where a.IdArticulo > 90000 order by Treballador') "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
'Si existe sincro_dependentesTmp2 la borramos y volvemos a generar
tablaTmp2 = "[Fac_laforneria].[dbo].[sincro_dependentesTmp2]"
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp2 & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp2)
sql = "SELECT * INTO " & tablaTmp2 & " FROM " & tablaTmp
Set Rs = Db.OpenResultset(sql)
sql = "DELETE FROM " & tablaTmp2
Set Rs = Db.OpenResultset(sql)
'Insertamos en segunta tabla temp solo los actualizables
sql = "INSERT INTO " & tablaTmp2
sql = sql & "SELECT t.CODI,t.Treballador,t.PIN,t.CodiLlocDeTreball,t.tipusTreballador,t.Mobil,t.Email,t.CodiTargeta "
sql = sql & "From ( "
sql = sql & "SELECT d.CODI,s.*,CASE WHEN s.LlocDeTreball='Cap de zona' THEN 'ADMINISTRACIO' "
sql = sql & "WHEN s.LlocDeTreball='Encarregat de botiga' THEN 'RESPONSABLE' "
sql = sql & "WHEN s.LlocDeTreball='Tecnic' THEN 'TECNIC' "
sql = sql & "WHEN s.LlocDeTreball='Venedor' THEN 'DEPENDENTA' "
sql = sql & "WHEN s.LlocDeTreball='Xofer' THEN 'REPARTIDOR' "
sql = sql & "WHEN s.LlocDeTreball='' THEN 'DEPENDENTA' "
sql = sql & "Else 'DEPENDENTA' END as tipusTreballador "
sql = sql & "FROM " & tablaTmp & " s "
sql = sql & "left join dependentesExtes de1 on (s.CodiTreballador=de1.valor and de1.nom='CODI_DEP') "
sql = sql & "left join dependentesExtes de2 on (de1.id=de2.id and de2.nom='EMAIL') "
sql = sql & "left join dependentesExtes de3 on (de1.id=de3.id and de3.nom='TIPUSTREBALLADOR') "
sql = sql & "left join dependentesExtes de4 on (de1.id=de4.id and de4.nom='PASSWORD') "
sql = sql & "left join dependentesExtes de5 on (de1.id=de5.id and de5.nom='CODI_TARGETA') "
sql = sql & "left join dependentes d on (de1.id=d.CODI) "
sql = sql & "WHERE de2.valor<>s.Email or de3.valor<> "
sql = sql & "(SELECT CASE WHEN s2.LlocDeTreball='Cap de zona' THEN 'ADMINISTRACIO' "
sql = sql & "WHEN s2.LlocDeTreball='Encarregat de botiga' THEN 'RESPONSABLE' "
sql = sql & "WHEN s2.LlocDeTreball='Tecnic' THEN 'TECNIC' "
sql = sql & "WHEN s2.LlocDeTreball='Venedor' THEN 'DEPENDENTA' "
sql = sql & "WHEN s2.LlocDeTreball='Xofer' THEN 'REPARTIDOR' "
sql = sql & "WHEN s2.LlocDeTreball='' THEN 'DEPENDENTA' "
sql = sql & "Else 'DEPENDENTA' END as tipusTreballador from " & tablaTmp & " s2 where "
sql = sql & "s2.CodiTreballador=s.CodiTreballador) "
sql = sql & "or d.MEMO<>s.Treballador or d.NOM<>s.Treballador "
sql = sql & "or d.TELEFON<>s.Mobil or de4.valor<>CONVERT(varchar(10),s.PIN) or de5.valor<>s.CodiTargeta "
sql = sql & ")t "
Set Rs = Db.OpenResultset(sql)
'--------------------------------------------------------------------------------
html = "<p><h3>Resum Dependentes Ametller</h3></p>"
'--------------------------------------------------------------------------------
'--UPDATE ACTUALIZABLES
'--------------------------------------------------------------------------------
InformaMiss "UPDATE DEPENDIENTAS ACTUALIZABLES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "UPDATE DEPENDIENTAS ACTUALIZABLES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
sql = "SELECT COUNT(CodiTreballador) as num FROM " & tablaTmp2
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    updDep = Rs("num")
    If updDep >= 1 Then
        'Tabla Dependentes
        sql = "UPDATE dependentes SET NOM=t.Treballador,MEMO=t.Treballador,TELEFON=t.Mobil "
        sql = sql & "From (SELECT s.CodiTreballador,s.Treballador,s.Mobil FROM " & tablaTmp2 & " s) t "
        sql = sql & "Where Dependentes.Codi = t.CodiTreballador"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Tabla DependentesExtes
        'TipusTreballador
        sql = "UPDATE dependentesExtes SET valor=t.LlocDeTreball From ("
        sql = sql & "SELECT s.CodiTreballador,s.LlocDeTreball FROM " & tablaTmp2 & " s ) t "
        sql = sql & "WHERE dependentesExtes.id=t.CodiTreballador and nom='TIPUSTREBALLADOR' "
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Password
        sql = "UPDATE dependentesExtes SET valor=t.PIN From ("
        sql = sql & "SELECT s.CodiTreballador,s.PIN FROM " & tablaTmp2 & " s ) t "
        sql = sql & "WHERE dependentesExtes.id=t.CodiTreballador and nom='PASSWORD'"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Email
        sql = "UPDATE dependentesExtes SET valor=t.Email From ("
        sql = sql & "SELECT s.CodiTreballador,s.Email FROM " & tablaTmp2 & " s ) t "
        sql = sql & "WHERE dependentesExtes.id=t.CodiTreballador and nom='EMAIL'"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'CodiTargeta
        sql = "UPDATE dependentesExtes SET valor=t.CodiTargeta From ("
        sql = sql & "SELECT s.CodiTreballador,s.CodiTargeta FROM " & tablaTmp2 & " s ) t "
        sql = sql & "WHERE dependentesExtes.id=t.CodiTreballador and nom='CODI_TARGETA'"
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
    End If
End If
'--------------------------------------------------------------------------------
InformaMiss "ACTUALIZADAS " & updDep & "DEPENDIENTAS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "ACTUALIZADAS " & updDep & "DEPENDIENTAS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
InformaMiss "NUEVAS DEPENDIENTAS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "NUEVAS DEPENDIENTAS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'Insertamos en segunta tabla temp solo los nuevos
sql = "DELETE FROM " & tablaTmp2
Set Rs = Db.OpenResultset(sql)
sql = "INSERT INTO " & tablaTmp2
sql = sql & "SELECT t.CodiTreballador,t.Treballador,t.PIN,t.CodiLlocDeTreball,t.tipusTreballador, "
sql = sql & "t.Mobil , t.Email, t.codiTargeta From ( "
sql = sql & "SELECT s.*,CASE WHEN s.LlocDeTreball='Cap de zona' THEN 'ADMINISTRACIO' "
sql = sql & "WHEN s.LlocDeTreball='Encarregat de botiga' THEN 'RESPONSALBE' "
sql = sql & "WHEN s.LlocDeTreball='Tecnic' THEN 'TECNIC' "
sql = sql & "WHEN s.LlocDeTreball='Venedor' THEN 'DEPENDENTA' "
sql = sql & "WHEN s.LlocDeTreball='Xofer' THEN 'REPARTIDOR' "
sql = sql & "WHEN s.LlocDeTreball='' THEN 'DEPENDENTA' "
sql = sql & "ELSE 'DEPENDENTA' END as tipusTreballador "
sql = sql & "FROM " & tablaTmp & " s "
sql = sql & "WHERE s.CodiTreballador not in (SELECT valor FROM dependentesExtes WHERE nom='CODI_DEP') "
sql = sql & "and s.treballador not in (select nom from Dependentes where NOM=s.treballador) "
sql = sql & ")t "
Set Rs = Db.OpenResultset(sql)
sql = "SELECT COUNT(CodiTreballador) as num FROM " & tablaTmp2
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    insDep = Rs("num")
    If insDep >= 1 Then
        'Proximo codigo dependienta disponible
        sql = "SELECT top 1 CAST(Codi as integer)+1 num from [Fac_LaForneria].[dbo].[dependentes] where codi<9999 order by codi desc"
        Set Rs2 = Db.OpenResultset(sql)
        If Not Rs2.EOF Then codiDep = Rs2("num")
        'Tabla Dependentes
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[dependentes]( "
        sql = sql & "[Codi] , [nom], [Memo], [Telefon], [ADREÇA], [Icona], [Hi Editem Horaris], [Tid]) "
        sql = sql & "SELECT " & codiDep & "+RANK() OVER (ORDER BY CodiTreballador),Treballador,Treballador,Mobil, "
        sql = sql & "NULL,NULL,1,NULL FROM " & tablaTmp2
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Tabla dependentesExtes
        'Email
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[dependentesExtes]([Id] , [nom], [Valor]) "
        sql = sql & "SELECT " & codiDep & "+RANK() OVER (ORDER BY CodiTreballador),'EMAIL',Email From " & tablaTmp2
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'TipusTreballador
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[dependentesExtes]([Id] , [nom], [Valor]) "
        sql = sql & "SELECT " & codiDep & "+RANK() OVER (ORDER BY CodiTreballador),'TIPUSTREBALLADOR',LlocDeTreball From " & tablaTmp2
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'NivellSeguretat
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[dependentesExtes]([Id] , [nom], [Valor]) "
        sql = sql & "SELECT " & codiDep & "+RANK() OVER (ORDER BY CodiTreballador),'NIVELLSEGURETAT','0' From " & tablaTmp2
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'CodiDep
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[dependentesExtes]([Id] , [nom], [Valor]) "
        sql = sql & "SELECT " & codiDep & "+RANK() OVER (ORDER BY CodiTreballador),'CODI_DEP',CodiTreballador From " & tablaTmp2
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'Password
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[dependentesExtes]([Id] , [nom], [Valor]) "
        sql = sql & "SELECT " & codiDep & "+RANK() OVER (ORDER BY CodiTreballador),'PASSWORD',PIN From " & tablaTmp2
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
        'CodiTargeta
        sql = "INSERT INTO [Fac_LaForneria].[dbo].[dependentesExtes]([Id] , [nom], [Valor]) "
        sql = sql & "SELECT " & codiDep & "+RANK() OVER (ORDER BY CodiTreballador),'CODI_TARGETA',CodiTargeta From " & tablaTmp2
        If debugSincro = True Then
            Txt.WriteLine "--------------------------------------------------------------------------------"
            Txt.WriteLine "SQL:" & sql & "-->" & Now
            Txt.WriteLine "--------------------------------------------------------------------------------"
        Else
            Set Rs2 = Db.OpenResultset(sql)
        End If
    End If
End If
'--------------------------------------------------------------------------------
InformaMiss "INSERTADAS" & insDep & "DEPENDIENTAS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "INSERTADAS " & insDep & "DEPENDIENTAS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
InformaMiss "BORRADO TABLAS TEMPORALES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "BORRADO TABLAS TEMPORALES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
If ExisteixTaula(tablaTmp) Then Db.OpenResultset ("DROP TABLE " & tablaTmp)
If ExisteixTaula(tablaTmp2) Then Db.OpenResultset ("DROP TABLE " & tablaTmp2)
'--------------------------------------------------------------------------------
If insDep > 0 Or updDep > 0 Then
    'Missatgeaenviar
    sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) VALUES ('Dependentes','')"
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        Set Rs2 = Db.OpenResultset(sql)
    End If
End If
'----------------------------------------------------------------------------------------------------
connMysql.Close
Set connMysql = Nothing
'----------------------------------------------------------------------------------------------------
'--UPDATE SINCRO_LAFORNERIA
'----------------------------------------------------------------------------------------------------
'Se indica que el proceso actual ha acabado
Set Rs = Db.OpenResultset("Delete from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='DEPENDENTES'")
Set Rs = Db.OpenResultset("Insert into [Fac_LaForneria].[dbo].[sincro_laforneria] (variable,fecha,p1) values ('DEPENDENTES',getdate(),1)")
'Se mira si es el ultimo proceso activo, si es asi se actualiza la fecha de la variable fecmod
'que es por la que se rigen los procesos.
sql = "Select COUNT(variable) num from [Fac_LaForneria].[dbo].[sincro_laforneria] where p1=1 and variable<>'FECMOD' "
sql = sql & "and fecha>=(select convert(datetime,p2,103) from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='fecmod') "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    'Si se han completado los cinco procesos se actualiza
    If Rs("num") = 5 Then
        Set Rs2 = Db.OpenResultset("Update [Fac_LaForneria].[dbo].[sincro_laforneria] Set [fecha] = [P2] WHERE [Variable]='FECMOD'")
        Set Rs2 = Db.OpenResultset("Update feinesafer set Param3='Fi " & Now() & "' where tipus='SincroDbAmetllerHorari' and param2 like '%Si%' ")
    End If
End If
'----------------------------------------------------------------------------------------------------
InformaMiss "FIN SINCRO_DEPENDENTES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "FIN SINCRO_DEPENDENTES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.Close
End If
Exit Function

norDep:
'----------------------------------------------------------------------------------------------------
'----  SI HAY ERROR
'----------------------------------------------------------------------------------------------------
'Se indica que el proceso actual ha fallado
    Set Rs = Db.OpenResultset("Delete from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='DEPENDENTES'")
    Set Rs = Db.OpenResultset("Insert into [Fac_LaForneria].[dbo].[sincro_laforneria] (variable,fecha,p1) values ('DEPENDENTES',getdate(),0)")
    html = "<p><h3>Resum Dependentes Ametller </h3></p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & dataIni & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    html = html & "<p><b>Ultima sqlSP:</b>" & sqlSP & "</p>"
    
    sf_enviarMail "secrehit@hit.cat", p1, "ERROR! Sincronitzacio de dependentes ha fallat", html, "", ""
         
    'Borramos tablas temporales
    If ExisteixTaula(tablaTmp) Then Db.OpenResultset ("DROP TABLE " & tablaTmp)
    If ExisteixTaula(tablaTmp2) Then Db.OpenResultset ("DROP TABLE " & tablaTmp2)
   
    ExecutaComandaSql "Delete from  " & TaulaCalculsEspecials & " Id = '" & idTasca & "' "
    
    Informa "Error : " & err.Description
    
End Function

Function SincroDbAmetllerTarifes(p1, idTasca) As Boolean
'*************************************************************************************
'SincroDbAmetllerTarifes
'Sincroniza tarifas entre servidor Ametller y servidor Hit
'La variable CODI_PROD de la tabla articlesPropietats liga el codigo articulo de Hit con
'el codigo articulo de Ametller y este se repite en la tabla tarifesEspecials que permite
'saber si existe o no tarifa para ese codigo articulo
'Proceso:
'1.Se recoge fecha ultima modificacion de la tabla sincro_laforneria.
'2.Se abre un bucle mirando las tarifas de ametller.
'3.En cada articulo se mira si ya existe en hit y si ademas existe una tarifa asociada a ese
'   codigo de articulo. Si existe actualiza sino inserta. Tambien se revisan las tarifas oferta.
'4.Se inserta linea en MissatgesAEnviar.
'*************************************************************************************
Dim codiTarifa As String, nomTarifa As String, codiTarifaO As String, nomTarifaO As String, tipusTarifa As String
Dim codiArticleOrigen As String, codiArticle As String, preuArticle As String, preuMArticle As String, preuOferta As String
Dim fecMod As Date, ultimaMod As String, diff As Long, insTar As Integer, updTar As Integer, preuOferta2 As Double
Dim idTabla As String, debugSincro As Boolean, sql As String, sql2 As String, sql3 As String, sqlSP As String, fecha As Date
Dim tablaTmp As String, Rs As rdoResultset, Rs2 As rdoResultset, html As String, tablaTmp2 As String, tablaTmp3 As String
Dim connMysql As ADODB.Connection, dataIni As Date
dataIni = Now()
codiArticle = 0
insTar = 0
updTar = 0

On Error GoTo norTar

'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
'Direccion email a la que se enviara resultado/error
If p1 = "" Then p1 = EmailGuardia
debugSincro = False 'Solo impresion o ejecucion completa
'Crear txt log
If debugSincro = True Then
    p1 = EmailGuardia
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\Tarifes_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
End If
'--------------------------------------------------------------
InformaMiss "INICIO SINCRO_CLIENTS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "INICIO SINCRO_CLIENTS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
'--ULTIMA FECHA MODIFICACION
'--------------------------------------------------------------------------------
InformaMiss "ULTIMA FECHA MODIFICACION-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "ULTIMA FECHA MODIFICACION-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
sql = "SELECT fecha FROM [Fac_LaForneria].[dbo].[sincro_laforneria] WHERE Variable='FECMOD' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    fecMod = Rs("fecha")
Else
    'Comprobamos que existe tabla sincro_laforneria
    tablaTmp = "[Fac_laforneria].[dbo].[sincro_laforneria]"
    sql = "SELECT object_id FROM sys.objects with (nolock) "
    sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp & "','[',''),']','') AND type='U' "
    Set Rs = Db.OpenResultset(sql)
    If Rs.EOF Then
        sql = "CREATE TABLE [Fac_laforneria].[dbo].[sincro_laforneria]("
        sql = sql & "       [variable] [varchar](255) NULL,"
        sql = sql & "       [fecha] [datetime] NULL,"
        sql = sql & "       [p1] [bit] NULL,"
        sql = sql & "       [p2] [nvarchar] (255) NULL,"
        sql = sql & "       [p3] [nvarchar] (255) NULL,"
        sql = sql & "       [p4] [nvarchar] (255) NULL,"
        sql = sql & "       [p5] [nvarchar] (255) NULL"
        sql = sql & "   ) ON [PRIMARY]"
        Set Rs = Db.OpenResultset(sql)
        sql = "INSERT INTO [Fac_laforneria].[dbo].[sincro_laforneria] values ("
        sql = sql & "'FECMOD',GETDATE(),0,GETDATE(),NULL,NULL,NULL) "
        Set Rs = Db.OpenResultset(sql)
        sql = "SELECT fecha FROM [Fac_LaForneria].[dbo].[sincro_laforneria] WHERE Variable='FECMOD' "
        Set Rs = Db.OpenResultset(sql)
        If Not Rs.EOF Then fecMod = Rs("fecha")
    End If
End If
'--------------------------------------------------------------------------------
html = "<p><h3>Resum tarifes Ametller</h3></p>"
'--------------------------------------------------------------------------------
'--BUCLE TARIFAS
'--------------------------------------------------------------------------------
InformaMiss "BUCLE TARIFAS-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "BUCLE TARIFAS-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'--------------------------------------------------------------------------------
sql = "'SELECT trf.IdTarifa AS CodiTarifa, trf.NombreTarifa AS Tarifa,trf.IdArticulo AS CodiArticle, "
sql = sql & "trf.PrecioSinIva AS Preu,trf.PrecioConIva AS PreuMajor,trf.tipoTarifa AS TipoTarifa, "
sql = sql & "trf.PrecioOferta AS PrecioOferta,trf.Timestamp AS fecMod FROM dat_tarifa AS trf "
sql = sql & "WHERE trf.IdEmpresa=1 AND trf.IdArticulo<>0 "
sql = sql & "and trf.Timestamp>''" & Year(fecMod) & "-" & Month(fecMod) & "-" & Day(fecMod) & " "
sql = sql & DatePart("H", fecMod) & ":" & DatePart("n", fecMod) & ":" & DatePart("s", fecMod) & "'' '"
sql2 = "SELECT * FROM openquery(AMETLLER," & sql & ") "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql2)
Db.QueryTimeout = 60
Do While Not Rs.EOF
    If IsNull(Rs("CodiTarifa")) = False Then codiTarifa = Rs("CodiTarifa")
    If codiTarifa = "2074" Then
        InformaMiss "TARIFA CHUNGA"
    End If
    If IsNull(Rs("Tarifa")) = False Then nomTarifa = Rs("Tarifa")
    If IsNull(Rs("CodiArticle")) = False Then codiArticleOrigen = Rs("CodiArticle")
    If IsNull(Rs("Preu")) = False Then preuMArticle = Rs("Preu")
    If IsNull(Rs("PreuMajor")) = False Then preuArticle = Rs("PreuMajor")
    If IsNull(Rs("TipoTarifa")) = False Then tipusTarifa = Rs("TipoTarifa")
    If IsNull(Rs("PrecioOferta")) = False Then preuOferta2 = Rs("PrecioOferta")
    If IsNull(Rs("fecmod")) = False Then
        ultimaMod = Rs("fecmod")
        'Se borra del tiempo mas alla de los segundos
        ultimaMod = Mid(ultimaMod, 1, Len(ultimaMod) - 8)
        ultimaMod = CDate(ultimaMod)
    End If
    nomTarifa = Replace(nomTarifa, "'", " ")
    codiTarifaO = CInt(codiTarifa) + 5000
    nomTarifaO = "O_" + CStr(codiTarifaO) & CStr(nomTarifa)
    nomTarifa = Left(CStr(codiTarifa) & CStr(nomTarifa), 20) 'Reduccion temporal nomTarifa, limite 20 caracteres
    nomTarifaO = Left(CStr(nomTarifaO), 20)
    If IsNull(preuArticle) = True Then preuArticle = 0
    If IsNull(preuMArticle) = True Then preuMArticle = 0
    If IsNull(preuOferta2) = True Then preuOferta2 = 0
    codiArticle = 0
    sql = "SELECT top 1 codiArticle from [Fac_LaForneria].[dbo].[articlesPropietats] WHERE valor='" & codiArticleOrigen & "' AND variable='CODI_PROD' "
    Set Rs2 = Db.OpenResultset(sql)
    If Not Rs2.EOF Then codiArticle = Rs2("codiArticle")
    'Diferencia fechas para saber si hay que actualizar o insertar
    diff = DateDiff("s", fecMod, ultimaMod)
    If IsNull(diff) = True Then diff = 1
    If diff > 0 Then 'Se debe actualizar o insertar
        'Sino existe el codigo de esta tienda, se inserta tienda
        'If codiArticle = 0 Then
            sql = "SELECT tarifaCodi FROM [Fac_LaForneria].[dbo].[tarifesEspecials] WHERE tarifaCodi='" & codiTarifa & "' AND codi='" & codiArticle & "' "
            Set Rs2 = Db.OpenResultset(sql)
            If Rs2.EOF Then
                'Insert tarifesEspecials
                sql = "INSERT INTO [Fac_LaForneria].[dbo].[tarifesEspecials]([TarifaCodi],[TarifaNom],[Codi],[PREU],[PreuMajor]) "
                sql = sql & "VALUES ('" & codiTarifa & "','" & nomTarifa & "','" & codiArticle & "','" & preuArticle & "','" & preuMArticle & "')"
                If debugSincro = True Then
                    Txt.WriteLine "--------------------------------------------------------------------------------"
                    Txt.WriteLine "SQL:" & sql & "-->" & Now
                    Txt.WriteLine "--------------------------------------------------------------------------------"
                Else
                    Set Rs2 = Db.OpenResultset(sql)
                End If
                If preuOferta2 > 0 Then
                    'OFERTA DIFERENTE A 0, insertamos tarifa de oferta
                    sql = "INSERT INTO [Fac_LaForneria].[dbo].[tarifesEspecials]([TarifaCodi],[TarifaNom],[Codi],[PREU],[PreuMajor]) "
                    sql = sql & "VALUES ('" & codiTarifaO & "','" & nomTarifaO & "','" & codiArticle & "','" & preuOferta2 & "',0)"
                    If debugSincro = True Then
                        Txt.WriteLine "--------------------------------------------------------------------------------"
                        Txt.WriteLine "SQL:" & sql & "-->" & Now
                        Txt.WriteLine "--------------------------------------------------------------------------------"
                    Else
                        Set Rs2 = Db.OpenResultset(sql)
                    End If
                Else
                    'OFERTA IGUAL a 0, eliminamos tarifa de oferta
                    sql = "DELETE [Fac_LaForneria].[dbo].[tarifesEspecials] where TarifaCodi = '" & codiTarifaO & "' and codi = '" & codiArticle & "' "
                    If debugSincro = True Then
                        Txt.WriteLine "--------------------------------------------------------------------------------"
                        Txt.WriteLine "SQL:" & sql & "-->" & Now
                        Txt.WriteLine "--------------------------------------------------------------------------------"
                    Else
                        Set Rs2 = Db.OpenResultset(sql)
                    End If
                End If
                insTar = 1
'------------------------------------------------------------------------------------------
InformaMiss "TARIFA INSERTADA " & codiTarifa & "," & nomTarifa & "," & codiArticle & "," & preuArticle & "," & preuMArticle & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "TARIFA INSERTADA " & codiTarifa & "," & nomTarifa & "," & codiArticle & "," & preuArticle & "," & preuMArticle & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'------------------------------------------------------------------------------------------
            'End If
            Else
            'Update tarifa
            sql = "UPDATE [Fac_LaForneria].[dbo].[tarifesEspecials] "
            sql = sql & "SET [TarifaNom]='" & nomTarifa & "',"
            sql = sql & "   [PREU]='" & preuArticle & "',"
            sql = sql & "   [PreuMajor]='" & preuMArticle & "' "
            sql = sql & "WHERE tarifaCodi='" & codiTarifa & "' AND codi='" & codiArticle & "' "
            If debugSincro = True Then
                Txt.WriteLine "--------------------------------------------------------------------------------"
                Txt.WriteLine "SQL:" & sql & "-->" & Now
                Txt.WriteLine "--------------------------------------------------------------------------------"
            Else
                Set Rs2 = Db.OpenResultset(sql)
            End If
            If preuOferta2 > 0 Then
                sql = "SELECT tarifaCodi FROM [Fac_LaForneria].[dbo].[tarifesEspecials] WHERE tarifaCodi='" & codiTarifaO & "' AND codi='" & codiArticle & "' "
                Set Rs2 = Db.OpenResultset(sql)
                If Rs2.EOF Then
                    'NUEVA OFERTA DIFERENTE A 0, instertamos tarifa de oferta
                    sql = "INSERT INTO [Fac_LaForneria].[dbo].[tarifesEspecials]([TarifaCodi],[TarifaNom],[Codi],[PREU],[PreuMajor]) "
                    sql = sql & "VALUES ('" & codiTarifaO & "','" & nomTarifaO & "','" & codiArticle & "','" & preuOferta2 & "',0)"
                    If debugSincro = True Then
                        Txt.WriteLine "--------------------------------------------------------------------------------"
                        Txt.WriteLine "SQL:" & sql & "-->" & Now
                        Txt.WriteLine "--------------------------------------------------------------------------------"
                    Else
                        Set Rs2 = Db.OpenResultset(sql)
                    End If
                Else
                    'UPDATE OFERTA DIFERENTE A 0, updatamos tarifa de oferta
                    sql = "UPDATE [Fac_LaForneria].[dbo].[tarifesEspecials] "
                    sql = sql & "SET [TarifaNom]='" & nomTarifaO & "',"
                    sql = sql & "[PREU]='" & preuOferta2 & "',"
                    sql = sql & "[PreuMajor]=0 "
                    sql = sql & "WHERE tarifaCodi='" & codiTarifaO & "' AND codi='" & codiArticle & "' "
                    If debugSincro = True Then
                        Txt.WriteLine "--------------------------------------------------------------------------------"
                        Txt.WriteLine "SQL:" & sql & "-->" & Now
                        Txt.WriteLine "--------------------------------------------------------------------------------"
                    Else
                        Set Rs2 = Db.OpenResultset(sql)
                    End If
                End If
            Else
                'OFERTA IGUAL a 0, eliminamos tarifa de oferta
                sql = "DELETE [Fac_LaForneria].[dbo].[tarifesEspecials] where TarifaCodi = '" & codiTarifaO & "' and codi = '" & codiArticle & "' "
                If debugSincro = True Then
                    Txt.WriteLine "--------------------------------------------------------------------------------"
                    Txt.WriteLine "SQL:" & sql & "-->" & Now
                    Txt.WriteLine "--------------------------------------------------------------------------------"
                Else
                    Set Rs2 = Db.OpenResultset(sql)
                End If
            End If
            updTar = 1
'------------------------------------------------------------------------------------------
InformaMiss "TARIFA ACTUALIZADA " & codiTarifa & "," & nomTarifa & "," & codiArticle & "," & preuArticle & "," & preuMArticle & "-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "TARIFA ACTUALIZADA " & codiTarifa & "," & nomTarifa & "," & codiArticle & "," & preuArticle & "," & preuMArticle & "-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
End If
'------------------------------------------------------------------------------------------
        End If
    End If
    Rs.MoveNext
Loop
'--------------------------------------------------------------------------------
If insTar > 0 Or updTar > 0 Then
    'Missatgeaenviar
    sql = "INSERT INTO [Fac_LaForneria].[dbo].[MissatgesAEnviar]([Tipus],[Param]) SELECT DISTINCT 'Tarifa', TarifaCodi from TarifesEspecials"
    If debugSincro = True Then
        Txt.WriteLine "--------------------------------------------------------------------------------"
        Txt.WriteLine "SQL:" & sql & "-->" & Now
        Txt.WriteLine "--------------------------------------------------------------------------------"
    Else
        Set Rs2 = Db.OpenResultset(sql)
    End If
End If
'--------------------------------------------------------------------------------
connMysql.Close
Set connMysql = Nothing
sql = "select id from feinesafer where tipus like '%SincroDbAmetller%' "
'----------------------------------------------------------------------------------------------------
'--UPDATE SINCRO_LAFORNERIA
'----------------------------------------------------------------------------------------------------
'Se indica que el proceso actual ha acabado
Set Rs = Db.OpenResultset("Delete from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='TARIFES'")
Set Rs = Db.OpenResultset("Insert into [Fac_LaForneria].[dbo].[sincro_laforneria] (variable,fecha,p1) values ('TARIFES',getdate(),1)")
'Se mira si es el ultimo proceso activo, si es asi se actualiza la fecha de la variable fecmod
'que es por la que se rigen los procesos.
sql = "Select COUNT(variable) num from [Fac_LaForneria].[dbo].[sincro_laforneria] where p1=1 and variable<>'FECMOD' "
sql = sql & "and fecha>=(select convert(datetime,p2,103) from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='fecmod') "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then
    'Si se han completado los cinco procesos se actualiza
    If Rs("num") = 5 Then
        Set Rs2 = Db.OpenResultset("Update [Fac_LaForneria].[dbo].[sincro_laforneria] Set [fecha] = [P2] WHERE [Variable]='FECMOD'")
        Set Rs2 = Db.OpenResultset("Update feinesafer set Param3='Fi " & Now() & "' where tipus='SincroDbAmetllerHorari' and param2 like '%Si%' ")
    End If
End If
'----------------------------------------------------------------------------------------------------
InformaMiss "FIN SINCRO_TARIFES-->" & Now
If debugSincro = True Then
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.WriteLine "FIN SINCRO_TARIFES-->" & Now
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt.Close
End If
Exit Function

norTar:
'----------------------------------------------------------------------------------------------------
'----  SI HAY ERROR
'----------------------------------------------------------------------------------------------------
'Se indica que el proceso actual ha fallado
    Set Rs = Db.OpenResultset("Delete from [Fac_LaForneria].[dbo].[sincro_laforneria] where variable='TARIFES'")
    Set Rs = Db.OpenResultset("Insert into [Fac_LaForneria].[dbo].[sincro_laforneria] (variable,fecha,p1) values ('TARIFES',getdate(),0)")
    html = "<p><h3>Resum Tarifes Ametller </h3></p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & dataIni & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    html = html & "<p><b>Ultima sqlSP:</b>" & sqlSP & "</p>"
    
    sf_enviarMail "secrehit@hit.cat", p1, "ERROR! Sincronitzacio de tarifes ha fallat", html, "", ""
            
    ExecutaComandaSql "Delete from  " & TaulaCalculsEspecials & " Id = '" & idTasca & "' "
    
    Informa "Error : " & err.Description
    
End Function

Function SincroDbVendesIdentAmetller2(idTasca) As Boolean
Dim botiguesCad As String, mes As String, anyo As String, desde As String, debugSincro As Boolean
Dim codiBotiga As String, codiBotigaextern As String, tabla As String, fecha, fecha_caracter
Dim sql As String, sql2 As String, sql3 As String, sqlSP As String, numCab As Integer, maxIdTicket As String
Dim parametros As String, tablaTmp As String, Rs As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset
Dim tablaTmp2 As String, tablaTmp3 As String, idClientFinal
Dim idCliente, NombreCliente, idDep, NombreDep, NumTick, botiga, data
Dim connMysql As ADODB.Connection

'ACTUALIZA VENDAS IDENTIFICADAS HASTA 03-2013
On Error GoTo norVendes
'Conexion servidor MYSQL ametller
Set connMysql = New ADODB.Connection
connMysql.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=88.87.154.9;Port=3307;Database=sys_datos;User=hituser;Password=aM3fP6x8;Option=3;"
connMysql.ConnectionTimeout = 1000 '16 min
connMysql.Open
Set obj_FSO = CreateObject("Scripting.FileSystemObject")
Set Txt = obj_FSO.CreateTextFile(AppPath & "\Tmp\Vendes_" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
Set obj_FSO2 = CreateObject("Scripting.FileSystemObject")
Set Txt2 = obj_FSO2.CreateTextFile(AppPath & "\Tmp\Vendes_2" & Day(Date) & "-" & Month(Date) & "-" & Year(Date) & ".txt", True)
'--------------------------------------------------------------------------------
'--TABLAS TEMPORALES
'--------------------------------------------------------------------------------
'Creamos tablas temporales de las quales podemos obtener datos de familia, secciones, etc
nTmp = Now
tablaTmp2 = "[Fac_laforneria].[dbo].[sincro_vendesTmpClients_" & nTmp & "]"
tablaTmp3 = "[Fac_laforneria].[dbo].[sincro_vendesTmpDependentes_" & nTmp & "]"
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmpClients la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp2 & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp2)
sql = " SELECT * INTO " & tablaTmp2 & " FROM OPENQUERY(AMETLLER,'"
sql = sql & "SELECT idCliente, Nombre,TRIM(DNI) AS Cif, TRIM(Direccion) AS Dir,"
sql = sql & "TRIM(CodPostal) AS CP, TRIM(Poblacion) AS Ciutat,TimeStamp AS fecMod "
sql = sql & "FROM dat_cliente WHERE IdEmpresa=1') "
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
'Si existe sincro_vendesTmpClients la borramos y volvemos a generar
sql = "SELECT object_id FROM sys.objects with (nolock) "
sql = sql & "WHERE name=REPLACE(REPLACE('" & tablaTmp3 & "','[',''),']','') AND type='U' "
Set Rs = Db.OpenResultset(sql)
If Not Rs.EOF Then Db.OpenResultset ("DROP TABLE " & tablaTmp2)
sql = "SELECT * INTO " & tablaTmp3 & " FROM OPENQUERY(AMETLLER,' "
sql = sql & "SELECT IdArticulo AS CodiTreballador, Descripcion AS Treballador, "
sql = sql & "CONCAT(REPEAT(''0'',4-LENGTH(REPLACE(FORMAT(PrecioConIVA*100,0),'','',''''))), "
sql = sql & "REPLACE(FORMAT(PrecioConIVA*100,0),'','','''')) AS PIN, e.job_title_code "
sql = sql & "AS CodiLlocDeTreball, j.jobtit_name AS LlocDeTreball,emp_mobile AS Mobil, "
sql = sql & "emp_work_email AS Email "
sql = sql & "From dat_articulo "
sql = sql & "     JOIN casaametller_orangehrm.hs_hr_employee e ON e.employee_id=IdArticulo "
sql = sql & "     LEFT JOIN casaametller_orangehrm.hs_hr_job_title j ON j.jobtit_code=e.job_title_code "
sql = sql & "Where IdArticulo > 90000 order by treballador ')"
Db.QueryTimeout = 0
Set Rs = Db.OpenResultset(sql)
Db.QueryTimeout = 60
'--------------------------------------------------------------------------------
sql = " select cf.id,cf.Nom,c.Codi,c.Nom,cc.codi,cc.valor,tmp2.idCliente,tmp2.Nombre,tmp3.codiTreballador,tmp3.Treballador from ( "
sql = sql & "select otros from [V_Venut_2012-10] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2012-11] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2012-12] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2013-01] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2013-02] where Otros like '%cli%' group by otros Union "
sql = sql & "select otros from [V_Venut_2013-03] where Otros like '%cli%' group by otros ) t "
sql = sql & "left join ClientsFinals cf on REPLACE(REPLACE(SUBSTRING(t.otros,CHARINDEX('CliBoti',t.otros),LEN(t.otros)),']',''),']','')=cf.Id "
sql = sql & "left join clients c on cf.nom=c.Nom "
sql = sql & "left join constantsclient cc on c.codi=cc.codi and cc.Variable='CodiClientOrigen' "
sql = sql & "LEFT JOIN " & tablaTmp2 & " tmp2 ON (cc.valor=tmp2.IdCliente) "
sql = sql & "left join Dependentes d on cf.Nom=d.NOM "
sql = sql & "left join DependentesExtes de on d.CODI=de.id and de.nom='CODI_DEP' "
sql = sql & "LEFT JOIN " & tablaTmp3 & " tmp3 ON (de.valor=tmp3.CodiTreballador) "
sql = sql & "Where tmp2.idCliente Is Not Null Or tmp3.CodiTreballador Is Not Null "
sql = sql & "group by cf.id,cf.Nom,c.Codi,c.Nom,cc.codi,cc.valor,tmp2.idCliente,tmp2.Nombre,tmp3.codiTreballador,tmp3.Treballador "
Txt2.WriteLine sql
'Select de todos los clientes identificados de vendas desde febrero 2012
Set Rs = Db.OpenResultset(sql)
Do While Not Rs.EOF
    idClientFinal = Rs("id")
    idCliente = Rs("idCliente")
    NombreCliente = Rs("Nombre")
    If NombreCliente <> "" Then NombreCliente = Replace(NombreCliente, "'", "''")
    idDep = Rs("codiTreballador")
    NombreDep = Rs("treballador")
    If NombreDep <> "" Then NombreDep = Replace(NombreDep, "'", "''")
    Txt.WriteLine "---------------------------------CLIENT/TREBALLADOR A MODIFICAR-----------------------------------------------"
    Txt.WriteLine "idClient:" & idCliente
    Txt.WriteLine "Client:" & NombreCliente
    Txt.WriteLine "idTreballador:" & idDep
    Txt.WriteLine "Treballador:" & NombreDep
    sql = "select t.Num_tick,t.Botiga,t.data from (select Num_tick,Botiga,data from [V_Venut_2012-02] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2012-10] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2012-11] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2012-12] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2013-01] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2013-02] where Otros like '%" & idClientFinal & "%' Union "
    sql = sql & "select Num_tick,Botiga,data from [V_Venut_2013-03] where Otros like '%" & idClientFinal & "%' ) t group by t.Num_tick,t.Botiga,t.data "
    Txt2.WriteLine sql
    Set Rs2 = Db.OpenResultset(sql)
    Do While Not Rs2.EOF
        sql = ""
        NumTick = Rs2("num_tick")
        botiga = Rs2("botiga")
        data = Rs2("data")
        Txt.WriteLine "Numero de ticket HIT:" & NumTick & " Botiga:" & Left(botiga, 2) & " Balança:" & Right(botiga, 1) & " Dia:" & data
        If botiga = "518" Then
            botiga = 1061
        End If
        If idCliente <> "" Then
            sql = "update dat_ticket_cabecera set idCliente='" & idCliente & "',"
            sql = sql & "NombreCliente='" & NombreCliente & "' where NumTicket='" & NumTick & "' "
            sql = sql & "and idTienda='" & Left(botiga, 2) & "' and idBalanzaMaestra='" & Right(botiga, 1) & "' "
            sql = sql & "and NombreBalanzaMaestra like '-Balan%' and idEmpresa=1 "
            sql = sql & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' and Operacion='A'"
            Txt2.WriteLine sql
            connMysql.Execute sql
        ElseIf idDep <> "" Then
            sql = "update dat_ticket_cabecera set idCliente='" & idDep & "',"
            sql = sql & "NombreCliente='" & NombreDep & "' where NumTicket='" & NumTick & "' "
            sql = sql & "and idTienda='" & Left(botiga, 2) & "' and idBalanzaMaestra='" & Right(botiga, 1) & "' "
            sql = sql & "and NombreBalanzaMaestra like '-Balan%' and idEmpresa=1 "
            sql = sql & "and idBalanzaEsclava=-1 and Usuario='Comunicaciones' and Operacion='A'"
            Txt2.WriteLine sql
            connMysql.Execute sql
        End If
        Rs2.MoveNext
    Loop
    Txt.WriteLine "--------------------------------------------------------------------------------"
    Txt2.WriteLine "--next"
    Rs.MoveNext
Loop
 
connMysql.Close
Set connMysql = Nothing
     
Txt.WriteLine "--------------------------------------------------------------------------------"
Txt2.WriteLine "--------------------------------------------------------------------------------"
Txt.Close
Txt2.Close
'Borramos tablas temporales
Db.OpenResultset ("DROP TABLE " & tablaTmp3)
Db.OpenResultset ("DROP TABLE " & tablaTmp2)
         
Exit Function



norVendes:
'----------------------------------------------------------------------------------------------------
'----  SI HAY ERROR: BORRADO ULTIMAS LINEAS DE TICKET HUERFANAS SIN CABECERA
'----------------------------------------------------------------------------------------------------
    html = "<p><h3>Resum Vendes Ametller </h3></p>"
    html = html & "<p><b>Botiga: </b>" & botiguesCad & "</p>"
    html = html & "<p><b>Data sincronitzacio: </b>" & desde & "</p>"
    html = html & "<p><b>NO ha finalitzat: </b>" & Now() & "</p>"
    If P3 <> "" Then sf_enviarMail "secrehit@hit.cat", Desti, "ERROR! Sincronitzacio de vendes de " & codiBotiga & " ha fallat", html, "", ""
    html = html & "<p><b>ERROR:</b>" & err.Number & " - " & err.Description & "</p>"
    html = html & "<p><b>Ultima sql:</b>" & sql & "</p>"
    html = html & "<p><b>Ultima sqlSP:</b>" & sqlSP & "</p>"
    
    sf_enviarMail "secrehit@hit.cat", EmailGuardia, "ERROR! Sincronitzacio de vendes de " & codiBotiga & " ha fallat", html, "", ""
        
    'Borramos tablas temporales
    Db.OpenResultset ("DROP TABLE " & tablaTmp3)
    Db.OpenResultset ("DROP TABLE " & tablaTmp2)
    
    ExecutaComandaSql "Delete from  " & TaulaCalculsEspecials & " Id = '" & idTasca & "' "
    
    Informa "Error : " & err.Description
    
End Function


Sub ExportaMURANO_CaixaBotiga(nEmpresa As Double, botiga As String, fecha As Date, idCalcul As String)
'NO SE USA 05/03/2020
    Dim D As Date, sql As String, rsCtb As rdoResultset
    Dim import As Double, tipoIva, PctIva, Base, Quota, CuentaVentas
    Dim T1 As Double, T2 As Double, T3 As Double, T4 As Double, TR1 As Double, TR2 As Double, TR3 As Double, TR4 As Double
    Dim cCtble, rsCodi As rdoResultset
    Dim rsHist As rdoResultset
    Dim rsNA As rdoResultset, rsSage As rdoResultset
    Dim rsCaixes As rdoResultset, rsCCBanc As rdoResultset, ccBanc As String
    Dim CcVentas As String
    Dim Di As Date, Df As Date
    Dim iTargeta As Double, iTkRs As Double, importZ As Double, import43 As Double
    Dim numAsiento As String
    Dim Motiu As String, nifClienteContado As String
    Dim rsPrimerTick As rdoResultset, rsUltimTick As rdoResultset, rsBotiga As rdoResultset
    Dim primerTick As String, ultimTick As String
    Dim iva1 As Double, baseIva1 As Double, iva2 As Double, baseIva2 As Double, iva3 As Double, baseIva3 As Double
    
    Dim nCierres As Integer
    Dim importZTotal As Double, importVTotal As Double
  
    Dim asientosCaixa As String
    asientosCaixa = "0"
    
    On Error GoTo noExportat
    
    If nEmpresa = "0" Then
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampNif' ")
    Else
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampNif' ")
    End If
    
    Dim nifEmpMurano As String, rsEmpMurano As rdoResultset
    
    If Not rsCtb.EOF Then nifEmpMurano = rsCtb("valor")
    Set rsEmpMurano = Db.OpenResultset("select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'")
    If Not rsEmpMurano.EOF Then
        EmpresaMurano = rsEmpMurano("CodigoEmpresa")
    Else
        GoTo noExportat
    End If

    nifClienteContado = "22222222J"

    D = fecha
    
    Set rsNA = Db.OpenResultset("select isnull(max(Asiento), 0) + 1 NumAsiento from " & TaulaHistoricoMURANO(D) & " where CodigoEmpresa = " & EmpresaMurano & " and year(fechaAsiento) = " & Year(D))
    If Not rsNA.EOF Then numAsiento = rsNA("NumAsiento")
    
    Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(D) & " where month(FechaAsiento)=" & Month(D) & " and day(FechaAsiento)=" & Day(D) & " and Param1 = '" & botiga & "' and TipoExportacion='CAIXA' and CodigoEmpresa=" & EmpresaMurano)
    While Not rsHist.EOF
        ExecutaComandaSql "Delete from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & rsHist("Asiento") & " and (statusTraspasadoIME=0 or statusTraspasadoIME=2) and CodigoEmpresa=" & EmpresaMurano & " and ejercicio=" & Year(D)
        rsHist.MoveNext
    Wend
    
    Set rsSage = Db.OpenResultset("select * from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & numAsiento & " and statusTraspasadoIME=1 and CodigoEmpresa=" & EmpresaMurano & " and ejercicio=" & Year(D))
    If Not rsSage.EOF Then GoTo noExportat  'Si ya se ha traspasado correctamente a MURANO no volvemos a traspasarla
    
    TipusDeIva T1, T2, T3, T4, TR1, TR2, TR3, TR4, D
    
    CuentaVentas = cVentaMercaderies(botiga, False)
    CcVentas = cVentaMercaderies(botiga, False)
    
    cCtble = botiga
    Set rsCodi = Db.OpenResultset("SELECT Valor FROM " & tablaConstantsClient() & " WHERE  codi = " & botiga & " AND variable = 'CodiContable' ")
    If Not rsCodi.EOF Then If Not IsNull(rsCodi("Valor")) And (Len(rsCodi("Valor")) > 0) And IsNumeric(rsCodi("Valor")) Then cCtble = CDbl(rsCodi("Valor"))
    
    InformaMiss "MURANO Vendes Botiga: " & BotigaCodiNom(botiga) & " Dia: " & D, True
    
    nCierres = 0
    importZTotal = 0
    sql = "select distinct data, tipus_moviment "
    sql = sql & "from [" & NomTaulaMovi(D) & "] where "
    sql = sql & "botiga='" & botiga & "' and day(data)=" & Day(D) & " and (tipus_moviment='Wi' or tipus_moviment='W') order by data"
    Set rsCaixes = Db.OpenResultset(sql)
    While Not rsCaixes.EOF
        If rsCaixes("tipus_moviment") = "Wi" Then
            Di = rsCaixes("data")
            
            Set rsPrimerTick = Db.OpenResultset("select top 1 Num_Tick from [" & NomTaulaVentas(D) & "] where botiga='" & botiga & "' and day(data)=" & Day(D) & " and data >= convert(datetime, '" & Day(Di) & "/" & Month(Di) & "/" & Year(Di) & " " & DatePart("h", Di) & ":" & DatePart("n", Di) & ":" & DatePart("s", Di) & "', 103) order by data")
            If Not rsPrimerTick.EOF Then primerTick = rsPrimerTick("Num_Tick")

            rsCaixes.MoveNext
            If Not rsCaixes.EOF Then
                If rsCaixes("tipus_moviment") = "W" Then
                    nCierres = nCierres + 1
                    Df = rsCaixes("data")
        
                    Set rsUltimTick = Db.OpenResultset("select top 1 Num_Tick from [" & NomTaulaVentas(D) & "] where botiga='" & botiga & "' and day(data)=" & Day(D) & " and data <= convert(datetime, '" & Day(Df) & "/" & Month(Df) & "/" & Year(Df) & " " & DatePart("h", Df) & ":" & DatePart("n", Df) & ":" & DatePart("s", Df) & "', 103) order by data desc")
                    If Not rsUltimTick.EOF Then ultimTick = rsUltimTick("Num_tick")
            
                    sql = "Select a.tipoiva, v.botiga, sum(v.import) as import "
                    sql = sql & "from ( "
                    sql = sql & "Select vv.Botiga, vv.Plu ,sum(vv.import) Import "
                    sql = sql & "From [" & NomTaulaVentas(D) & "] vv "
                    sql = sql & "where vv.data between CONVERT(datetime, '" & Di & "', 103) and convert(datetime,'" & Df & "', 103) and vv.botiga = '" & botiga & "' "
                    sql = sql & "group by vv.botiga, vv.Plu) v "
                    sql = sql & "Left Join "
                    sql = sql & "(select Aa.codi, aa.TipoIva "
                    sql = sql & "From (select Familia,codi,tipoiva from articles union select Familia,codi,tipoiva from articles_zombis) aa ) a on a.codi = v.plu "
                    sql = sql & "group by botiga, tipoiva "
                    sql = sql & "order by botiga, tipoiva "
                    Set rsCtb = Db.OpenResultset(sql)
                    
                    ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            
                    import43 = 0
                    iva1 = 0: baseIva1 = 0:  iva2 = 0: baseIva2 = 0: iva3 = 0: baseIva3 = 0
                    While Not rsCtb.EOF
                        import = rsCtb("Import")
                        tipoIva = rsCtb("TipoIva")
                        
                        InformaMiss "MURANO Vendes Botiga: " & BotigaCodiNom(botiga) & " Dia: " & D, True
                        
                        Select Case tipoIva
                            Case 1: PctIva = T1
                            Case 2: PctIva = T2
                            Case 3: PctIva = T3
                        End Select
                        Base = Round(import / (1 + (PctIva / 100)), 2)
                        Quota = Round(import - Base, 2)
                        import = Base + Quota
                        
                        Select Case tipoIva
                            Case 1: iva1 = iva1 + Quota: baseIva1 = baseIva1 + Base
                            Case 2: iva2 = iva2 + Quota: baseIva2 = baseIva2 + Base
                            Case 3: iva3 = iva3 + Quota: baseIva3 = baseIva3 + Base
                        End Select
                        
                        import43 = import43 + import

                        AsientoAddMURANO_TS "", 0, numAsiento, D, ("477" & Right("000000000000", nDigitos - 3)) + PctIva, "", "", "", 0, Quota, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Ventas " & BotigaCodiNom(botiga), ""
                        AsientoAddMURANO_TS "", 0, numAsiento, D, CcVentas, "", "", "", 0, Base, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Ventas " & BotigaCodiNom(botiga), ""
                        
                        InformaMiss "Ventas " & D, True
                        DoEvents
            
                        rsCtb.MoveNext
                    Wend
                    
                    AsientoAddMURANO_TS "", 0, numAsiento, D, ("43" & Right("000000000000", nDigitos - 2)) + cCtble, "E", BotigaCodiNom(botiga), primerTick, import43, 0, iva1, baseIva1, iva2, baseIva2, iva3, baseIva3, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Ventas " & BotigaCodiNom(botiga), "", BotigaCodiNom(botiga), primerTick, ultimTick, "B"
                    
                    asientosCaixa = asientosCaixa + "," + numAsiento
                    
                    'Metálico
                    'import = total de ventas - cobros tarjeta - cobros cheques - descuadre
                    InformaMiss "MURANO metàl·lic Botiga: " & BotigaCodiNom(botiga) & " Dia: " & D, True
                    numAsiento = numAsiento + 1
                    
                    importZ = 0
                    iTargeta = 0
                    iTkRs = 0
                    
                    sql = "select import "
                    sql = sql & "from [" & NomTaulaMovi(D) & "] "
                    sql = sql & "where botiga='" & botiga & "' and Data between convert(datetime,'" & Di & "', 103) and convert(datetime,'" & Df & "', 103) and tipus_moviment='Z'"
                    Set rsCtb = Db.OpenResultset(sql)
                    If Not rsCtb.EOF Then importZ = rsCtb("import")
                    'importZTotal = importZTotal + importZ
                    
                    sql = "Select c.codi botigaCodi, abs(sum(import)) as import "
                    sql = sql & "from [" & NomTaulaMovi(D) & "] v "
                    sql = sql & "left join clients c on v.botiga = c.codi "
                    sql = sql & "where Data between convert(datetime,'" & Di & "', 103) and convert(datetime,'" & Df & "', 103) And botiga = '" & botiga & "' and motiu like 'Pagat Targeta%' "
                    sql = sql & "Group By c.codi"
                    Set rsCtb = Db.OpenResultset(sql)
                    If Not rsCtb.EOF Then iTargeta = rsCtb("import")
                    
                    sql = "Select c.codi botigaCodi, abs(sum(import)) as import "
                    sql = sql & "from [" & NomTaulaMovi(D) & "] v "
                    sql = sql & "left join clients c on v.botiga = c.codi "
                    sql = sql & "where Data between convert(datetime,'" & Di & "', 103) and convert(datetime,'" & Df & "', 103) And botiga = '" & botiga & "' and motiu like 'Pagat TkRs%' "
                    sql = sql & "Group By c.codi"
                    Set rsCtb = Db.OpenResultset(sql)
                    If Not rsCtb.EOF Then iTkRs = rsCtb("import")
                    
                    import = importZ - iTargeta - iTkRs
                
                    ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
                    AsientoAddMURANO_TS "", 0, numAsiento, D, "43" & Right("000000000000" & cCtble, nDigitos - 2), "B", "", "", 0, import, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "COBRO EN METALICO CAJA REGISTRA " & BotigaCodiNom(botiga), ""
                    AsientoAddMURANO_TS "", 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", import, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "COBRO EN METALICO CAJA REGISTRA " & BotigaCodiNom(botiga), ""
                    
                    asientosCaixa = asientosCaixa + "," + numAsiento
                    
                    rsCaixes.MoveNext
                    numAsiento = numAsiento + 1
                End If 'Hay 2 Wi seguidos
            End If 'NO hay más cajas
        Else 'NO és Wi
            rsCaixes.MoveNext
        End If
    Wend
    rsCtb.Close
    
    If nCierres < 1 Then
        ExecutaComandaSql "insert into debug (data, str) values (getdate(), 'MURANO FALTEN CAIXES DE [" & botiga & "] " & BotigaCodiNom(botiga) & " DIA " & D & "')"
        
        'ELIMINAR TODOS LOS ASIENTOS ANTERIORES!! PARA PODER VOLVER A EXPORTAR LA CAJA ENTERA
        ExecutaComandaSql "delete from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where asiento in (" & asientosCaixa & ")"
        
        ExecutaComandaSql "insert into feinesafer (Id, Tipus, ciclica, param1, param2) values (newid(), 'SincronitzaCaixaPendentMURANO', 0, '[" & botiga & "]', '[" & Right("00" & Day(D), 2) & "/" & Right("00" & Month(D), 2) & "/" & Year(D) & "]')"
        GoTo noExportat
    End If
    
    'Moviments ENTRADA/SORTIDA
    InformaMiss "MURANO Moviments ENTRADA/SORTIDA Botiga: " & BotigaCodiNom(botiga) & " Dia: " & D, True
    
    sql = "Select motiu, tipus_moviment, c.codi botigaCodi, sum(import) as import "
    sql = sql & "from [" & NomTaulaMovi(D) & "] v "
    sql = sql & "left join clients c on v.botiga = c.codi "
    sql = sql & "where tipus_moviment in ('O','A') and day(Data) = " & Day(D) & " And botiga = '" & botiga & "' and Import<>0 "
    sql = sql & "Group By motiu, tipus_moviment, c.codi "
    
    Set rsCtb = Db.OpenResultset(sql)
    While Not rsCtb.EOF
        import = Format(rsCtb("Import"), "0.0#")
        Motiu = rsCtb("motiu") & " " & BotigaCodiNom(botiga)

        If rsCtb("motiu") = "Entrega Diària" Or UCase(rsCtb("motiu")) = UCase("Entrega Diaria") Or rsCtb("motiu") = "Sortida de Canvi" Then
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"

            AsientoAddMURANO_TS "", 0, numAsiento, D, "5701" & Right("000000000000" & cCtble, nDigitos - 4), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "SALIDA DE CAJA A CAJA FUERTE " & BotigaCodiNom(botiga), ""
            AsientoAddMURANO_TS "", 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "SALIDA DE CAJA A CAJA FUERTE " & BotigaCodiNom(botiga), ""
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
            
        ElseIf Left(rsCtb("motiu"), 13) = "Pagat Targeta" Then
            'Set rsCCBanc = Db.OpenResultset("select valor from ConstantsClient where Variable='EmpresaVendesCC' and Codi='" & botiga & "'")
            'If Not rsCCBanc.EOF Then ccBanc = rsCCBanc("valor")
            
            'ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", " & numAsiento & ", 'CAIXA')"
            'AsientoAddMURANO 0, numAsiento, D, ("43" & Right("000000000000" & cCtble, nDigitos - 2)), "", "Cobros Tarjeta " & BotigaCodiNom(botiga), 0, import
            'AsientoAddMURANO 0, numAsiento, D, ("572" & Right("000000000000" & ccBanc, nDigitos - 3)), "", "Cobros Tarjeta " & BotigaCodiNom(botiga), import, 0
            'numAsiento = numAsiento + 1

        ElseIf Left(rsCtb("motiu"), 10) = "Pagat TkRs" Then
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS "", 0, numAsiento, D, ("43" & Right("000000000000" & cCtble, nDigitos - 2)), "B", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Cobros ticket restaurant " & BotigaCodiNom(botiga), ""
            AsientoAddMURANO_TS "", 0, numAsiento, D, ("44" & Right("000000000000", nDigitos - 2)), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Cobros ticket restaurant " & BotigaCodiNom(botiga), ""
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
            
        ElseIf Left(rsCtb("motiu"), 9) = "Excs.TkRs" Then
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS "", 0, numAsiento, D, ("768" & Right("000000000000", nDigitos - 3)), "B", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Exc. cobro CHQ RTE", ""
            AsientoAddMURANO_TS "", 0, numAsiento, D, ("44" & Right("000000000000", nDigitos - 2)), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Exc. cobro CHQ RTE", ""
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
            
        ElseIf Left(rsCtb("motiu"), 6) = "Albara" Then 'ALBARANES
            InformaMiss "MURANO Moviments Botiga Albara " & D, True

        ElseIf Left(rsCtb("motiu"), 2) = "D." Then 'DEIXA A DEURE
            InformaMiss "MURANO Moviments Deixen deute " & D, True
            
        ElseIf Left(rsCtb("motiu"), 2) = "P." Then 'PAGA DEUTE
            InformaMiss "MURANO Moviments Paguen deute " & D, True
            
        'SOTIDES DE DINERS
        Else
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"

            AsientoAddMURANO_TS "", 0, numAsiento, D, "629005" & Right("000000000000" & cCtble, nDigitos - 6), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "GASTOS " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            AsientoAddMURANO_TS "", 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "GASTOS " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
        End If
        
        DoEvents
        rsCtb.MoveNext
    Wend
    rsCtb.Close
    
    'Moviments DESCUADRE
    InformaMiss "MURANO Moviments DESCUADRE Botiga: " & BotigaCodiNom(botiga) & " Dia: " & D, True
    
    sql = "Select data, c.codi botigaCodi, cast(import as nvarchar) import "
    sql = sql & "from [" & NomTaulaMovi(D) & "] v "
    sql = sql & "left join clients c on v.botiga = c.codi "
    sql = sql & "where tipus_moviment = 'J' and day(Data) = " & Day(D) & " And botiga = '" & botiga & "' "
    sql = sql & "Order By data "
    
    Set rsCtb = Db.OpenResultset(sql)
    While Not rsCtb.EOF
        import = rsCtb("import")

        If import < 0 Then 'DESCUADRE NEGATIVO
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS "", 0, numAsiento, D, "659" & Right("000000000000" & cCtble, nDigitos - 3), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            AsientoAddMURANO_TS "", 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
        Else 'DESCUADRE POSITIVO
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS "", 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            AsientoAddMURANO_TS "", 0, numAsiento, D, "659" & Right("000000000000" & cCtble, nDigitos - 3), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""

            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
        End If
        rsCtb.MoveNext
    Wend
    rsCtb.Close
    
noExportat:
    
End Sub



