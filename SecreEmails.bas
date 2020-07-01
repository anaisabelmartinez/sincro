Attribute VB_Name = "SecreEmails"
Option Explicit

Function calculaCompras(codiBot As Double, fecha As Date, Optional familia As String) As Double
'Familia: excluimos la familia
'SIN ***
    Dim facTabData As String, facTabIva As String
    Dim compres As Double
    Dim sql As String
    Dim rsCompras As rdoResultset
    
On Error GoTo errFac
    'Compras
    'Buscamos la factura
    facTabData = "Facturacio_"
    facTabIva = "Facturacio_"
    If Month(fecha) = 12 Then
        facTabData = facTabData & Year(fecha) + 1 & "-01_Data"
        facTabIva = facTabIva & Year(fecha) + 1 & "-01_Iva"
    Else
        facTabData = facTabData & Year(fecha) & "-" & Right("0" & Month(fecha) + 1, 2) & "_Data"
        facTabIva = facTabIva & Year(fecha) & "-" & Right("0" & Month(fecha) + 1, 2) & "_Iva"
    End If
    
    compres = 0
    sql = "select isnull(sum(import), 0) compras "
    sql = sql & "from [Facturacio_" & Year(fecha) & "-" & Right("0" & Month(fecha), 2) & "_Data] d "
    If familia <> "" Then
        sql = sql & "left join articles a on d.producte=a.codi "
        sql = sql & "left join families f3 on a.familia = f3.nom "
        sql = sql & "left join families f2 on f3.pare = f2.nom "
        sql = sql & "left join families f1 on f2.pare = f1.nom "
    End If
    sql = sql & "left join [Facturacio_" & Year(fecha) & "-" & Right("0" & Month(fecha), 2) & "_Iva] i on d.idFactura=i.idfactura "
    sql = sql & "where i.serie not like '%RE%' and day(d.data)=" & Day(fecha) & " and month(d.data)=" & Month(fecha) & " and year(d.data)=" & Year(fecha) & " and d.client = " & codiBot & " and d.producteNom not like '***%' "
    If familia <> "" Then
        sql = sql & "and (f1.nom not like '%" & familia & "%' and f2.nom not like '%" & familia & "%' and f3.nom not like '%" & familia & "%') "
    End If
    Set rsCompras = Db.OpenResultset(sql)
    
    If (rsCompras.EOF Or rsCompras("compras") = 0) And ExisteixTaula(facTabIva) Then
        sql = "select isnull(sum(import), 0) compras "
        sql = sql & "from [" & facTabData & "] d "
        If familia <> "" Then
            sql = sql & "left join articles a on d.producte=a.codi "
            sql = sql & "left join families f3 on a.familia = f3.nom "
            sql = sql & "left join families f2 on f3.pare = f2.nom "
            sql = sql & "left join families f1 on f2.pare = f1.nom "
        End If

        sql = sql & "left join [" & facTabIva & "] i on d.idFactura=i.idfactura "
        sql = sql & "where i.serie not like '%RE%' and day(d.data)=" & Day(fecha) & " and month(d.data)=" & Month(fecha) & " and year(d.data)=" & Year(fecha) & " and d.client = " & codiBot & " and d.producteNom not like '***%' "
        If familia <> "" Then
            sql = sql & "and (f1.nom not like '%" & familia & "%' and f2.nom not like '%" & familia & "%' and f3.nom not like '%" & familia & "%') "
        End If
        
        Set rsCompras = Db.OpenResultset(sql)
    End If
    If Not rsCompras.EOF Then compres = rsCompras("compras")
    
errFac:
    If compres = 0 Then 'servit
        sql = "select isnull(sum(s.quantitatServida * (a.preumajor - (a.preumajor * case when a.desconte=1 then cast(c.[Desconte 1] as float)/100 when a.desconte=2 then cast(c.[Desconte 2] as float)/100 when a.desconte=3 then cast(c.[Desconte 3] as float)/100 when a.desconte=4 then cast(c.[Desconte 4] as float)/100 else 0 end ))), 0) compras "
        sql = sql & "from " & DonamTaulaServit(fecha) & " s "
        sql = sql & "left join articles a on s.codiarticle = a.codi "
        If familia <> "" Then
            sql = sql & "left join families f3 on a.familia = f3.nom "
            sql = sql & "left join families f2 on f3.pare = f2.nom "
            sql = sql & "left join families f1 on f2.pare = f1.nom "
        End If
        
        sql = sql & "left join clients c on s.client=c.codi "
        sql = sql & "where client = " & codiBot & " and a.nom not like '***%' "
        If familia <> "" Then
            sql = sql & "and (f1.nom not like '%" & familia & "%' and f2.nom not like '%" & familia & "%' and f3.nom not like '%" & familia & "%') "
        End If
        
        Set rsCompras = Db.OpenResultset(sql)
        If Not rsCompras.EOF Then
            compres = rsCompras("compras")
        End If
    End If
    
    calculaCompras = compres
End Function




Function calculaComprasAsterisco(codiBot As Double, fecha As Date) As Double

'CON ***

    Dim facTabData As String, facTabIva As String
    Dim compres As Double
    Dim sql As String
    Dim rsCompras As rdoResultset
    
On Error GoTo errFac
    'Compras
    'Buscamos la factura
    facTabData = "Facturacio_"
    facTabIva = "Facturacio_"
    If Month(fecha) = 12 Then
        facTabData = facTabData & Year(fecha) + 1 & "-01_Data"
        facTabIva = facTabIva & Year(fecha) + 1 & "-01_Iva"
    Else
        facTabData = facTabData & Year(fecha) & "-" & Right("0" & Month(fecha) + 1, 2) & "_Data"
        facTabIva = facTabIva & Year(fecha) & "-" & Right("0" & Month(fecha) + 1, 2) & "_Iva"
    End If
    
    compres = 0
    sql = "select isnull(sum(import), 0) compras "
    sql = sql & "from [Facturacio_" & Year(fecha) & "-" & Right("0" & Month(fecha), 2) & "_Data] d "
    sql = sql & "left join [Facturacio_" & Year(fecha) & "-" & Right("0" & Month(fecha), 2) & "_Iva] i on d.idFactura=i.idfactura "
    sql = sql & "where i.serie not like '%RE%' and day(d.data)=" & Day(fecha) & " and month(d.data)=" & Month(fecha) & " and year(d.data)=" & Year(fecha) & " and d.client = " & codiBot & " and d.producteNom like '***%' "
    Set rsCompras = Db.OpenResultset(sql)
    
    If (rsCompras.EOF Or rsCompras("compras") = 0) And ExisteixTaula(facTabIva) Then
        sql = "select isnull(sum(import), 0) compras "
        sql = sql & "from [" & facTabData & "] d "
        sql = sql & "left join [" & facTabIva & "] i on d.idFactura=i.idfactura "
        sql = sql & "where i.serie not like '%RE%' and day(d.data)=" & Day(fecha) & " and month(d.data)=" & Month(fecha) & " and year(d.data)=" & Year(fecha) & " and d.client = " & codiBot & " and d.producteNom like '***%' "
        Set rsCompras = Db.OpenResultset(sql)
    End If
    If Not rsCompras.EOF Then compres = rsCompras("compras")
    
errFac:
    If compres = 0 Then 'servit
        sql = "select isnull(sum(s.quantitatServida * (a.preumajor - (a.preumajor * case when a.desconte=1 then cast(c.[Desconte 1] as float)/100 when a.desconte=2 then cast(c.[Desconte 2] as float)/100 when a.desconte=3 then cast(c.[Desconte 3] as float)/100 when a.desconte=4 then cast(c.[Desconte 4] as float)/100 else 0 end ))), 0) compras "
        sql = sql & "from " & DonamTaulaServit(fecha) & " s "
        sql = sql & "left join articles a on s.codiarticle = a.codi "
        sql = sql & "left join clients c on s.client=c.codi "
        sql = sql & "where client = " & codiBot & " and a.nom Like '***%' "
        Set rsCompras = Db.OpenResultset(sql)
        If Not rsCompras.EOF Then
            compres = rsCompras("compras")
        End If
    End If
    
    calculaComprasAsterisco = compres
End Function





Function calculaComprasReal(codiBot As Double, fecha As Date) As Double
    Dim facTabData As String, facTabIva As String
    Dim compres As Double
    Dim sql As String
    Dim rsCompras As rdoResultset
    
On Error GoTo errFac
    'Compras
    'Buscamos la factura
    facTabData = "Facturacio_"
    facTabIva = "Facturacio_"
    If Month(fecha) = 12 Then
        facTabData = facTabData & Year(fecha) + 1 & "-01_Data"
        facTabIva = facTabIva & Year(fecha) + 1 & "-01_Iva"
    Else
        facTabData = facTabData & Year(fecha) & "-" & Right("0" & Month(fecha) + 1, 2) & "_Data"
        facTabIva = facTabIva & Year(fecha) & "-" & Right("0" & Month(fecha) + 1, 2) & "_Iva"
    End If
    
    compres = 0
    sql = "select isnull(sum(import), 0) compras "
    sql = sql & "from [Facturacio_" & Year(fecha) & "-" & Right("0" & Month(fecha), 2) & "_Data] d "
    sql = sql & "left join [Facturacio_" & Year(fecha) & "-" & Right("0" & Month(fecha), 2) & "_Iva] i on d.idFactura=i.idfactura "
    sql = sql & "where i.serie not like '%RE%' and day(d.data)=" & Day(fecha) & " and month(d.data)=" & Month(fecha) & " and year(d.data)=" & Year(fecha) & " and d.client = " & codiBot & " "
    Set rsCompras = Db.OpenResultset(sql)
    
    If (rsCompras.EOF Or rsCompras("compras") = 0) And ExisteixTaula(facTabIva) Then
        sql = "select isnull(sum(import), 0) compras "
        sql = sql & "from [" & facTabData & "] d "
        sql = sql & "left join [" & facTabIva & "] i on d.idFactura=i.idfactura "
        sql = sql & "where i.serie not like '%RE%' and day(d.data)=" & Day(fecha) & " and month(d.data)=" & Month(fecha) & " and year(d.data)=" & Year(fecha) & " and d.client = " & codiBot & " "
        Set rsCompras = Db.OpenResultset(sql)
    End If
    If Not rsCompras.EOF Then compres = rsCompras("compras")
    
errFac:
    If compres = 0 Then 'servit
        sql = "select isnull(sum(s.quantitatServida * (a.preumajor - (a.preumajor * case when a.desconte=1 then cast(c.[Desconte 1] as float)/100 when a.desconte=2 then cast(c.[Desconte 2] as float)/100 when a.desconte=3 then cast(c.[Desconte 3] as float)/100 when a.desconte=4 then cast(c.[Desconte 4] as float)/100 else 0 end ))), 0) compras "
        sql = sql & "from " & DonamTaulaServit(fecha) & " s "
        sql = sql & "left join articles a on s.codiarticle = a.codi "
        sql = sql & "left join clients c on s.client=c.codi "
        sql = sql & "where client = " & codiBot & " "
        Set rsCompras = Db.OpenResultset(sql)
        If Not rsCompras.EOF Then
            compres = rsCompras("compras")
        End If
    End If
    
    calculaComprasReal = compres
End Function

Function calculaComprasFamilia(codiBot As Double, fecha As Date, familia As String) As Double
    Dim facTabData As String, facTabIva As String
    Dim compres As Double
    Dim sql As String
    Dim rsCompras As rdoResultset
    
On Error GoTo errFac
    'Compras
    'Buscamos la factura
    facTabData = "Facturacio_"
    facTabIva = "Facturacio_"
    If Month(fecha) = 12 Then
        facTabData = facTabData & Year(fecha) + 1 & "-01_Data"
        facTabIva = facTabIva & Year(fecha) + 1 & "-01_Iva"
    Else
        facTabData = facTabData & Year(fecha) & "-" & Right("0" & Month(fecha) + 1, 2) & "_Data"
        facTabIva = facTabIva & Year(fecha) & "-" & Right("0" & Month(fecha) + 1, 2) & "_Iva"
    End If
    
    compres = 0
    sql = "select isnull(sum(import), 0) compras "
    sql = sql & "from [Facturacio_" & Year(fecha) & "-" & Right("0" & Month(fecha), 2) & "_Data] d "
    sql = sql & "left join [Facturacio_" & Year(fecha) & "-" & Right("0" & Month(fecha), 2) & "_Iva] i on d.idFactura=i.idfactura "
    sql = sql & "left join articles a on d.producte=a.codi "
    sql = sql & "left join families f3 on a.familia = f3.nom "
    sql = sql & "left join families f2 on f3.pare = f2.nom "
    sql = sql & "left join families f1 on f2.pare = f1.nom "
    sql = sql & "where i.serie not like '%RE%' and day(d.data)=" & Day(fecha) & " and month(d.data)=" & Month(fecha) & " and year(d.data)=" & Year(fecha) & " and d.client = " & codiBot & " and (f1.nom like '%" & familia & "%' or f2.nom like '%" & familia & "%' or f3.nom like '%" & familia & "%') "
    Set rsCompras = Db.OpenResultset(sql)
    
    If (rsCompras.EOF Or rsCompras("compras") = 0) And ExisteixTaula(facTabIva) Then
        sql = "select isnull(sum(import), 0) compras "
        sql = sql & "from [" & facTabData & "] d "
        sql = sql & "left join [" & facTabIva & "] i on d.idFactura=i.idfactura "
        sql = sql & "left join articles a on d.producte=a.codi "
        sql = sql & "left join families f3 on a.familia = f3.nom "
        sql = sql & "left join families f2 on f3.pare = f2.nom "
        sql = sql & "left join families f1 on f2.pare = f1.nom "
        sql = sql & "where i.serie not like '%RE%' and day(d.data)=" & Day(fecha) & " and month(d.data)=" & Month(fecha) & " and year(d.data)=" & Year(fecha) & " and d.client = " & codiBot & " and (f1.nom like '%" & familia & "%' or f2.nom like '%" & familia & "%' or f3.nom like '%" & familia & "%') "
        Set rsCompras = Db.OpenResultset(sql)
    End If
    If Not rsCompras.EOF Then compres = rsCompras("compras")
    
errFac:
    If compres = 0 Then 'servit
        sql = "select isnull(sum(s.quantitatServida * (a.preumajor - (a.preumajor * case when a.desconte=1 then cast(c.[Desconte 1] as float)/100 when a.desconte=2 then cast(c.[Desconte 2] as float)/100 when a.desconte=3 then cast(c.[Desconte 3] as float)/100 when a.desconte=4 then cast(c.[Desconte 4] as float)/100 else 0 end ))), 0) compras "
        sql = sql & "from " & DonamTaulaServit(fecha) & " s "
        sql = sql & "left join articles a on s.codiarticle = a.codi "
        sql = sql & "left join families f3 on a.familia = f3.nom "
        sql = sql & "left join families f2 on f3.pare = f2.nom "
        sql = sql & "left join families f1 on f2.pare = f1.nom "
        sql = sql & "left join clients c on s.client=c.codi "
        sql = sql & "where client = " & codiBot & " and (f1.nom like '%" & familia & "%' or f2.nom like '%" & familia & "%' or f3.nom like '%" & familia & "%') "
        Set rsCompras = Db.OpenResultset(sql)
        If Not rsCompras.EOF Then
            compres = rsCompras("compras")
        End If
    End If
    
    calculaComprasFamilia = compres
End Function


Function calculaComprasTienda(codiBot As Double, fecha As Date) As Double
    Dim compres As Double
    Dim rsCompres As rdoResultset
    
    compres = 0
    Set rsCompres = Db.OpenResultset("select isnull(sum(import), 0) import from [" & NomTaulaMovi(fecha) & "] where botiga=" & codiBot & " and day(data)=" & Day(fecha) & " and tipus_moviment='O' and motiu not like '%Targeta%' and motiu not like '%Entrega%' and motiu not like '%TkRs%' and motiu not like '%Sortida de Canvi%'")
    If Not rsCompres.EOF Then
        compres = Abs(rsCompres("import"))
    End If
    
    calculaComprasTienda = compres
End Function

Function calculaDevoluciones(codiBot As Double, fecha As Date) As Double
    Dim rsDevDia As rdoResultset, sql As String, devDia As Double
            
    devDia = 0

    sql = "select isnull(sum(a.preu*s.quantitatTornada), 0) I "
    sql = sql & "from " & DonamTaulaServit(fecha) & " s "
    sql = sql & "left join articles a on s.codiarticle=a.codi "
    sql = sql & "left join clients c on s.client = c.codi "
    sql = sql & "Where s.client = " & codiBot & " And s.quantitatTornada > 0"
    Set rsDevDia = Db.OpenResultset(sql)
    
    If Not rsDevDia.EOF Then devDia = rsDevDia("I")
            
    calculaDevoluciones = devDia
            
End Function

Function calculaDevolucionesAcumulado(codiBot As Double, fechaIni As Date, fechaFin As Date) As Double
    Dim devAc As Double
    Dim f As Date
    Dim tServits As String
    Dim rsDevAc As rdoResultset, sql As String
    
    devAc = 0
            
    tServits = ""
    For f = fechaIni To fechaFin
        If tServits = "" Then
            tServits = "select * from " & DonamTaulaServit(f)
        Else
            tServits = tServits & " union all select * from " & DonamTaulaServit(f)
        End If
    Next
    
    sql = "select isnull(sum(a.preu*s.quantitatTornada), 0) I "
    sql = sql & "from (" & tServits & ") s "
    sql = sql & "left join articles a on s.codiarticle=a.codi "
    sql = sql & "left join clients c on s.client = c.codi "
    sql = sql & "Where s.client = " & codiBot & " And s.quantitatTornada > 0"
    Set rsDevAc = Db.OpenResultset(sql)
    If Not rsDevAc.EOF Then devAc = rsDevAc("I")
           
    calculaDevolucionesAcumulado = devAc
End Function


Function calculaHorasFichajeReal(fecha As Date, botiga As Double, usuario As String) As Double
    Dim rsReal As rdoResultset
    Dim horasReales As Double
    Dim entra As Date, sale As Date
    Dim haEntrado As Boolean, haSalido As Boolean
    
    horasReales = 0
    Set rsReal = Db.OpenResultset("select * from cdpdadesfichador where lloc='" & botiga & "' and day(tmst)=" & Day(fecha) & " and month(tmst)=" & Month(fecha) & " and year(tmst)=" & Year(fecha) & " and usuari=" & usuario & " order by usuari, accio")
    While Not rsReal.EOF
        If rsReal("accio") = 1 Then
            entra = rsReal("tmst")
            haEntrado = True
        End If
        
        If rsReal("accio") = 2 And haEntrado Then
            sale = rsReal("tmst")
            horasReales = horasReales + DateDiff("n", entra, sale)
            haSalido = True
        End If
        rsReal.MoveNext
    Wend

    'Si ha entrado pero no ha salido (y no es hoy) suponemos que ha trabajado hasta las 00:00 del dia siguiente
    If (haEntrado And Not haSalido) And DateSerial(Year(fecha), Month(fecha), Day(fecha)) < DateSerial(Year(Now()), Month(Now()), Day(Now())) Then
        horasReales = DateDiff("n", entra, DateAdd("d", 1, DateSerial(Year(fecha), Month(fecha), Day(fecha))))
    End If
    
    calculaHorasFichajeReal = horasReales / 60
End Function

Function calculaHorasReales(codiBot As Double, fecha As Date) As Double
    Dim sql As String
    Dim rsHoras As rdoResultset
    Dim horas As Double
    
    horas = 0
    sql = "select isnull(sum(case when p.idTurno like '%Extra%' then left(p.idturno, charindex('_', p.idturno)-1) else isnull(datediff(minute, t.horaInicio, t.horaFin) / 60.0, 0.0) end), 0.0) horas "
    sql = sql & "from " & taulaCdpPlanificacion(fecha) & " p "
    sql = sql & "left join cdpTurnos t on p.idTurno = t.idTurno "
    sql = sql & "where p.botiga=" & codiBot & " and day(p.fecha)=" & Day(fecha) & " and p.activo=1 and p.idturno not like '%Coordinacion%' and p.idturno not like '%aprendiz%' and p.idEmpleado is not null"
    sql = sql & " and (t.idturno is not null or p.idTurno like '%_Extra')"
    Set rsHoras = Db.OpenResultset(sql)
    If Not rsHoras.EOF Then horas = rsHoras("horas")

    calculaHorasReales = horas
End Function
Function calculaHorasPanadero(codiBot As Double, fecha As Date) As Double
    Dim sql As String
    Dim rsHoras As rdoResultset
    Dim horas As Double
    
    horas = 0
    sql = "select isnull(sum(case when p.idTurno like '%Extra%' then left(p.idturno, charindex('_', p.idturno)-1) else isnull(datediff(minute, t.horaInicio, t.horaFin) / 60.0, 0.0) end), 0.0) horas "
    sql = sql & "from " & taulaCdpPlanificacion(fecha) & " p "
    sql = sql & "left join cdpTurnos t on p.idTurno = t.idTurno "
    sql = sql & "where p.botiga=" & codiBot & " and day(p.fecha)=" & Day(fecha) & " and p.activo=1 and p.idturno not like '%Coordinacion%' and p.idturno not like '%aprendiz%' and p.idEmpleado is not null and t.tipoEmpleado like '%FORNER%' "
    Set rsHoras = Db.OpenResultset(sql)
    If Not rsHoras.EOF Then horas = rsHoras("horas")

    calculaHorasPanadero = horas
End Function

Function calculaHorasResto(codiBot As Double, fecha As Date) As Double
    Dim sql As String
    Dim rsHoras As rdoResultset
    Dim horas As Double
    
    horas = 0
    sql = "select isnull(sum(case when p.idTurno like '%Extra%' then left(p.idturno, charindex('_', p.idturno)-1) else isnull(datediff(minute, t.horaInicio, t.horaFin) / 60.0, 0.0) end), 0.0) horas "
    sql = sql & "from " & taulaCdpPlanificacion(fecha) & " p "
    sql = sql & "left join cdpTurnos t on p.idTurno = t.idTurno "
    sql = sql & "where p.botiga=" & codiBot & " and day(p.fecha)=" & Day(fecha) & " and p.activo=1 and p.idturno not like '%Coordinacion%' and p.idturno not like '%aprendiz%' and p.idEmpleado is not null and t.tipoEmpleado not like '%FORNER%' "
    Set rsHoras = Db.OpenResultset(sql)
    If Not rsHoras.EOF Then horas = rsHoras("horas")

    calculaHorasResto = horas
End Function


Function calculaHorasAprendiz(codiBot As Double, fecha As Date) As Double
    Dim sql As String
    Dim rsHoras As rdoResultset
    Dim horas As Double
    
    horas = 0
    sql = "select p.idTurno horas "
    sql = sql & "from " & taulaCdpPlanificacion(fecha) & " p "
    sql = sql & "where p.botiga=" & codiBot & " and day(p.fecha)=" & Day(fecha) & " and p.activo=1 and p.idturno like '%aprendiz%' and p.idEmpleado is not null"
    Set rsHoras = Db.OpenResultset(sql)
    While Not rsHoras.EOF
        horas = horas + CInt(Split(rsHoras("horas"), "_")(0))
        rsHoras.MoveNext
    Wend
    
    calculaHorasAprendiz = horas
End Function

Function calculaHorasCoordinacion(codiBot As Double, fecha As Date) As Double
    Dim sql As String
    Dim rsHoras As rdoResultset
    Dim horas As Double
    
    horas = 0
    sql = "select p.idTurno horas "
    sql = sql & "from " & taulaCdpPlanificacion(fecha) & " p "
    sql = sql & "where p.botiga=" & codiBot & " and day(p.fecha)=" & Day(fecha) & " and p.activo=1 and p.idturno like '%Coordinacion' and p.idEmpleado is not null"
    Set rsHoras = Db.OpenResultset(sql)
    While Not rsHoras.EOF
        horas = horas + CInt(Split(rsHoras("horas"), "_")(0))
        rsHoras.MoveNext
    Wend
    
    calculaHorasCoordinacion = horas
End Function


Function calculaHorasSinTurno(codiBot As Double, fecha As Date) As Double
    Dim sql As String
    Dim rsHoras As rdoResultset, rsFichaje As rdoResultset
    Dim horas As Double
    Dim entra As Date, sale As Date
    
    horas = 0

    sql = "select * "
    sql = sql & "from " & taulaCdpPlanificacion(fecha) & " p "
    sql = sql & "where p.botiga=" & codiBot & " and day(p.fecha)=" & Day(fecha) & " and p.activo=1 and p.idturno is null "
    sql = sql & "order by p.fecha"
    Set rsHoras = Db.OpenResultset(sql)
    While Not rsHoras.EOF
        entra = Now()
        sale = entra
        Set rsFichaje = Db.OpenResultset("select * from cdpDadesfichador where lloc='" & codiBot & "' and usuari= '" & rsHoras("idEmpleado") & "' and tmst >= '" & Format(rsHoras("fecha"), "dd/mm/yyyy hh:nn:ss") & "' and day(tmst)=" & Day(fecha) & " order by tmst")
        If Not rsFichaje.EOF Then
            If rsFichaje("accio") = 1 Then
                entra = rsFichaje("tmst")
                rsFichaje.MoveNext
                If Not rsFichaje.EOF Then
                    If rsFichaje("accio") = 2 Then
                        sale = rsFichaje("tmst")
                    End If
                End If
            End If
        End If
        
        If DateDiff("h", entra, sale) >= 1 Then horas = horas + DateDiff("h", entra, sale)
        
        rsHoras.MoveNext
    Wend
    
    calculaHorasSinTurno = horas
End Function



Function calculaNumeroClientes(codiBot As Double, fecha As Date) As Integer
    Dim rs As rdoResultset
    
    Set rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients from [" & NomTaulaVentas(fecha) & "] where botiga =  " & codiBot & " And day(data) = " & Day(fecha) & " ")
    calculaNumeroClientes = rs("Clients")

End Function

Sub calculaPrevisions(codiBot As Double, fecha As Date, previsions() As Double)
    Dim rsPrevisio As rdoResultset
    Dim pMatiB As Double, pTardaB As Double

    pMatiB = 0
    pTardaB = 0

    Set rsPrevisio = Db.OpenResultset("select * from [" & NomTaulaMovi(fecha) & "] where day(data)=" & Day(fecha) & " and botiga=" & codiBot & " and Tipus_moviment in ('MATI','TARDA','MATI_C','TARDA_C') order by Tipus_moviment")
    While Not rsPrevisio.EOF
        If rsPrevisio("Tipus_moviment") = "MATI" Then 'matí
            pMatiB = rsPrevisio("import")
        ElseIf rsPrevisio("Tipus_moviment") = "MATI_C" Then 'mati correcció
            pMatiB = rsPrevisio("import")
        ElseIf rsPrevisio("Tipus_moviment") = "TARDA" Then 'tarda
            pTardaB = rsPrevisio("import")
        ElseIf rsPrevisio("Tipus_moviment") = "TARDA_C" Then 'tarda correció
            pTardaB = rsPrevisio("import")
        End If
        
        rsPrevisio.MoveNext
    Wend
    rsPrevisio.Close
    
    previsions(0) = pMatiB
    previsions(1) = pTardaB

End Sub

Function calculaVendesFamilia(codiBot As Double, lunes As Date, domingo As Date, familia As String) As Double
    Dim sql As String, rsVendes As rdoResultset, taulesVenut As String
    
    If Month(lunes) = Month(domingo) Then
        taulesVenut = "(select * from [" & NomTaulaVentas(lunes) & "] union all select * from [" & NomTaulaAlbarans(lunes) & "])"
    Else
        taulesVenut = "(select * from [" & NomTaulaVentas(lunes) & "] union all select * from [" & NomTaulaVentas(domingo) & "] union all select * from [" & NomTaulaAlbarans(lunes) & "] union all select * from [" & NomTaulaAlbarans(domingo) & "])"
    End If
    
    sql = "select isnull(sum(import), 0) import "
    sql = sql & "from " & taulesVenut & " v "
    sql = sql & "left join articles a on v.plu=a.codi "
    sql = sql & "left join families f3 on a.familia = f3.nom "
    sql = sql & "left join families f2 on f3.pare = f2.nom "
    sql = sql & "left join families f1 on f2.pare = f1.nom "
    sql = sql & "where botiga=" & codiBot & " and data between convert(datetime, '" & lunes & "') and convert(datetime, '" & domingo & " 23:59:59')  and (f1.nom like '%" & familia & "%' or f2.nom like '%" & familia & "%' or f3.nom like '%" & familia & "%')"
    Set rsVendes = Db.OpenResultset(sql)
    
    If Not rsVendes.EOF Then
        calculaVendesFamilia = rsVendes("import")
    Else
        calculaVendesFamilia = 0
    End If
    
End Function
Function calculaVendesInterEmpreses(codiBot As Double, lunes As Date, domingo As Date) As Double
    Dim sql As String, rsVendes As rdoResultset, taulesAlbarans As String
    
    If Month(lunes) = Month(domingo) Then
        taulesAlbarans = "[" & NomTaulaAlbarans(lunes) & "]"
    Else
        taulesAlbarans = "(select * from [" & NomTaulaAlbarans(lunes) & "] union all select * from [" & NomTaulaAlbarans(domingo) & "])"
    End If
    
    sql = "select isnull(sum(import), 0) import "
    sql = sql & "from " & taulesAlbarans & " a "
    sql = sql & "left join clients c on a.otros = c.codi "
    sql = sql & "where botiga=" & codiBot & " and data between convert(datetime, '" & lunes & "') and convert(datetime, '" & domingo & " 23:59:59') and c.nif in (select valor from constantsEmpresa where camp like '%CampNif%' and isnull(valor, '')<>'') "
    Set rsVendes = Db.OpenResultset(sql)
    
    If Not rsVendes.EOF Then
        calculaVendesInterEmpreses = rsVendes("import")
    Else
        calculaVendesInterEmpreses = 0
    End If
    
End Function
Function calculaCompresInterEmpreses(codiBot As Double, lunes As Date, domingo As Date) As Double
    Dim sql As String, rsCompres As rdoResultset, taulesAlbarans As String
    
    If Month(lunes) = Month(domingo) Then
        taulesAlbarans = "[" & NomTaulaAlbarans(lunes) & "]"
    Else
        taulesAlbarans = "(select * from [" & NomTaulaAlbarans(lunes) & "] union all select * from [" & NomTaulaAlbarans(domingo) & "])"
    End If
    
    sql = "select isnull(sum(import), 0) import "
    sql = sql & "from " & taulesAlbarans & " a "
    sql = sql & "left join clients c on a.botiga = c.codi "
    sql = sql & "where a.otros=" & codiBot & " and data between convert(datetime, '" & lunes & "') and convert(datetime, '" & domingo & " 23:59:59') and c.nif in (select valor from constantsEmpresa where camp like '%CampNif%' and isnull(valor, '')<>'') "
    Set rsCompres = Db.OpenResultset(sql)
    
    If Not rsCompres.EOF Then
        calculaCompresInterEmpreses = rsCompres("import")
    Else
        calculaCompresInterEmpreses = 0
    End If
    
End Function

Function cuadrantePlanificacionTurnos3(dia As Date, Optional botiguesList As String) As String
    Dim co As String
    Dim lunes As Date, dSemana As Integer
    Dim rsQ As rdoResultset, sql As String
    Dim sqlFechas As String, sqlSelect As String, sqlPivot As String
    Dim totalBotiga As Double
    Dim color As String
    Dim totalL As Double, totalM As Double, totalX As Double, totalJ As Double, totalV As Double, totalS As Double, totalD As Double
            
On Error GoTo err

    dSemana = Weekday(dia, 2)

    lunes = DateAdd("d", -(dSemana - 1), dia)

    sqlSelect = "p1.nom, isnull(p1.[" & Format(lunes, "dd/mm/yyyy") & "]-p2.[" & Format(lunes, "dd/mm/yyyy") & "], 0) [" & Format(lunes, "dd/mm/yyyy") & "],"
    sqlSelect = sqlSelect & "isnull(p1.[" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "] - p2.[" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "],"
    sqlSelect = sqlSelect & "isnull(p1.[" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "] - p2.[" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "],"
    sqlSelect = sqlSelect & "isnull(p1.[" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "] - p2.[" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "],"
    sqlSelect = sqlSelect & "isnull(p1.[" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "] - p2.[" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "],"
    sqlSelect = sqlSelect & "isnull(p1.[" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "] - p2.[" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "],"
    sqlSelect = sqlSelect & "isnull(p1.[" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "] - p2.[" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "]"
                
                
    sqlFechas = "isnull([" & Format(lunes, "dd/mm/yyyy") & "], 0) [" & Format(lunes, "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "]"
    
    sqlPivot = "[" & Format(lunes, "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "]"
    
    sql = "select " & sqlSelect & " from "
    sql = sql & "( "
    sql = sql & "select nom, codi, " & sqlFechas & " "
    sql = sql & "From ( "
    sql = sql & "select botiga, convert(nvarchar, fecha, 103) fecha, c.nom, c.codi, sum(case when p.idturno like '%Extra' then left(p.idturno, charindex('_', p.idturno)-1) when p.idturno like '%Coordinacion' or p.idturno like '%Aprendiz' then 0.0 else datediff(minute, t.horaInicio, t.horafin) / 60.00 end) horas "
    sql = sql & "from " & taulaCdpPlanificacion(lunes) & " p "
    sql = sql & "left join cdpturnos t on p.idturno=t.idTurno "
    sql = sql & "left join clients c on p.botiga=c.codi "
    sql = sql & "Where P.activo = 1 And P.idEmpleado Is Not Null and botiga in (" & botiguesList & ") "
    sql = sql & "group by botiga, convert(nvarchar, p.fecha, 103), c.nom, c.codi ) DataTable "
    sql = sql & "PIVOT ( sum(horas) for fecha in (" & sqlPivot & ")) PivotTableReal "
    sql = sql & ") p1 "
    sql = sql & "Left Join "
    sql = sql & "( "
    sql = sql & "select botiga, " & sqlFechas & " "
    sql = sql & "From ( "
    sql = sql & "select p.botiga, convert(nvarchar, fecha, 103) fecha, sum(datediff(minute, t.horaInicio, t.horafin) / 60.00) horas "
    sql = sql & "from " & taulaCdpPlanificacion(lunes) & " p "
    sql = sql & "left join cdpturnos t on p.idturno=t.idTurno "
    sql = sql & "Where P.activo = 1 And t.idturno Is Not Null and botiga in (" & botiguesList & ") "
    sql = sql & "group by p.botiga, convert(nvarchar, p.fecha, 103) ) DataTable "
    sql = sql & "PIVOT (sum(horas) for fecha in (" & sqlPivot & ")) PivotTablePactado "
    sql = sql & ") p2 "
    sql = sql & "on P1.codi=P2.botiga "
    sql = sql & "order by p1.nom"

    Set rsQ = Db.OpenResultset(sql)
            
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'>"
    co = co & "<td><b>Botiga</b></td>"
    co = co & "<td><b>" & Format(lunes, "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>Total</b></td>"
    co = co & "</tr>"
    
    While Not rsQ.EOF
       
        co = co & "<Tr>"
        co = co & "<td><b>" & UCase(rsQ("nom")) & "</b></td>"
        If rsQ(Format(lunes, "dd/mm/yyyy")) <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
        co = co & "<td bgcolor='" & color & "' align='right'>" & FormatNumber(rsQ(Format(lunes, "dd/mm/yyyy")), 2) & "</td>"
        totalL = totalL + CDbl(rsQ(Format(lunes, "dd/mm/yyyy")))
        
        If rsQ(Format(DateAdd("d", 1, lunes), "dd/mm/yyyy")) <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
        co = co & "<td bgcolor='" & color & "' align='right'>" & FormatNumber(rsQ(Format(DateAdd("d", 1, lunes), "dd/mm/yyyy")), 2) & "</td>"
        totalM = totalM + CDbl(rsQ(Format(DateAdd("d", 1, lunes), "dd/mm/yyyy")))
        
        If rsQ(Format(DateAdd("d", 2, lunes), "dd/mm/yyyy")) <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
        co = co & "<td bgcolor='" & color & "' align='right'>" & FormatNumber(rsQ(Format(DateAdd("d", 2, lunes), "dd/mm/yyyy")), 2) & "</td>"
        totalX = totalX + CDbl(rsQ(Format(DateAdd("d", 2, lunes), "dd/mm/yyyy")))
        
        If rsQ(Format(DateAdd("d", 3, lunes), "dd/mm/yyyy")) <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
        co = co & "<td bgcolor='" & color & "' align='right'>" & FormatNumber(rsQ(Format(DateAdd("d", 3, lunes), "dd/mm/yyyy")), 2) & "</td>"
        totalJ = totalJ + CDbl(rsQ(Format(DateAdd("d", 3, lunes), "dd/mm/yyyy")))
        
        If rsQ(Format(DateAdd("d", 4, lunes), "dd/mm/yyyy")) <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
        co = co & "<td bgcolor='" & color & "' align='right'>" & FormatNumber(rsQ(Format(DateAdd("d", 4, lunes), "dd/mm/yyyy")), 2) & "</td>"
        totalV = totalV + CDbl(rsQ(Format(DateAdd("d", 4, lunes), "dd/mm/yyyy")))

        If rsQ(Format(DateAdd("d", 5, lunes), "dd/mm/yyyy")) <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
        co = co & "<td bgcolor='" & color & "' align='right'>" & FormatNumber(rsQ(Format(DateAdd("d", 5, lunes), "dd/mm/yyyy")), 2) & "</td>"
        totalS = totalS + CDbl(rsQ(Format(DateAdd("d", 5, lunes), "dd/mm/yyyy")))
        
        If rsQ(Format(DateAdd("d", 6, lunes), "dd/mm/yyyy")) <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
        co = co & "<td bgcolor='" & color & "' align='right'>" & FormatNumber(rsQ(Format(DateAdd("d", 6, lunes), "dd/mm/yyyy")), 2) & "</td>"
        totalD = totalD + CDbl(rsQ(Format(DateAdd("d", 6, lunes), "dd/mm/yyyy")))
        
        totalBotiga = CDbl(rsQ(Format(lunes, "dd/mm/yyyy"))) + CDbl(rsQ(Format(DateAdd("d", 1, lunes), "dd/mm/yyyy"))) + CDbl(rsQ(Format(DateAdd("d", 2, lunes), "dd/mm/yyyy"))) + CDbl(rsQ(Format(DateAdd("d", 3, lunes), "dd/mm/yyyy"))) + CDbl(rsQ(Format(DateAdd("d", 4, lunes), "dd/mm/yyyy"))) + CDbl(rsQ(Format(DateAdd("d", 5, lunes), "dd/mm/yyyy"))) + CDbl(rsQ(Format(DateAdd("d", 6, lunes), "dd/mm/yyyy")))
        If totalBotiga <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
        co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalBotiga, 2) & "</b></td>"

        co = co & "</Tr>"
        
        rsQ.MoveNext
    Wend

    co = co & "<tr><td><b>TOTAL</b></td>"
    If totalL <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalL, 2) & "</b></td>"
    If totalM <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalM, 2) & "</b></td>"
    If totalX <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalX, 2) & "</b></td>"
    If totalJ <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalJ, 2) & "</b></td>"
    If totalV <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalV, 2) & "</b></td>"
    If totalS <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalS, 2) & "</b></td>"
    If totalD <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalD, 2) & "</b></td>"
    If totalL + totalM + totalX + totalJ + totalV + totalS + totalD <> 0 Then color = "#FFA8A8" Else color = "#A6FFA6"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalL + totalM + totalX + totalJ + totalV + totalS + totalD, 2) & "</b></td>"
    co = co & "</tr>"
    co = co & "</table>"
    
    cuadrantePlanificacionTurnos3 = co
    
    Exit Function
    
err:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR cuadrantePlanificacionTurnos3 [" & Format(Now(), "dd/mm/yy hh:nn") & "]", sql & err.Description, "", ""
End Function

Function esFestivo(fecha As Date, codiEmp As String) As Boolean
    Dim rsCalendario As rdoResultset
    
    Set rsCalendario = Db.OpenResultset("select * from cdpCalendariLaboral_" & Year(fecha) & " where idEmpleado='" & codiEmp & "' and month(fecha)=" & Month(fecha) & " and day(fecha)=" & Day(fecha))
    If Not rsCalendario.EOF Then
        If rsCalendario("estado") = "FESTIU" Then
            esFestivo = True
        Else
            esFestivo = False
        End If
    Else
        esFestivo = False
    End If

End Function

Function getNotasIncidencia(codiBot As Double, fechaIni As Date, fechaFin As Date) As String
    Dim strInc As String
    Dim sql As String
    Dim fechaActual As Date
    Dim rsCaixes As rdoResultset, rsVentas As rdoResultset
    Dim vMati As Double, vTarda As Double
    Dim dataInici As Date, dataFi As Date
        
On Error GoTo nor:
        
    strInc = ""
    For fechaActual = fechaIni To fechaFin
        vMati = 0
        vTarda = 0
            
        sql = "select distinct data, tipus_moviment "
        sql = sql & "from [" & NomTaulaMovi(fechaActual) & "] where "
        sql = sql & "botiga='" & codiBot & "' and day(data)=" & Day(fechaActual) & " and (tipus_moviment='Wi' or tipus_moviment='W') order by data"
        Set rsCaixes = Db.OpenResultset(sql)
        
        If rsCaixes.EOF Then
            strInc = strInc & "*No hi ha caixa " & Format(fechaActual, "dd/mm/yyyy") & "<BR>"
            Set rsVentas = Db.OpenResultset("select isnull(sum(import), 0) I, isnull(sum(import/(1+(t.iva/100))), 0) sinIva, count(distinct Num_Tick ) Clients from [" & NomTaulaVentas(fechaActual) & "] v left join articles a on v.plu=a.codi left join TipusIva2012 t on a.TipoIva = t.tipus where botiga=" & codiBot & " and day(data) = " & Day(fechaActual))
            vMati = rsVentas("I")
        End If
        
        While Not rsCaixes.EOF
            If rsCaixes("tipus_moviment") = "Wi" Then
                dataInici = rsCaixes("data")
                rsCaixes.MoveNext
                
                If Not rsCaixes.EOF Then
                    If rsCaixes("tipus_moviment") = "W" Then
                        dataFi = rsCaixes("data")
                        
                        Set rsVentas = Db.OpenResultset("select isnull(sum(import), 0) I, isnull(sum(import/(1+(t.iva/100))), 0) sinIva, count(distinct Num_Tick ) Clients from [" & NomTaulaVentas(dataInici) & "] v left join articles a on v.plu=a.codi left join TipusIva2012 t on a.TipoIva = t.tipus where botiga=" & codiBot & " and data between '" & dataInici & "' and '" & dataFi & "'")
                        If Not rsVentas.EOF Then
                            If Not IsNull(rsVentas("I")) Then
                                If DatePart("h", dataInici) < 13 Then 'MATÍ
                                    vMati = vMati + rsVentas("I")
                                Else 'TARDA
                                    vTarda = vTarda + rsVentas("I")
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                rsCaixes.MoveNext
            End If
        Wend
        
        sql = "select sum(import) import "
        sql = sql & "from [" & NomTaulaMovi(fechaActual) & "] where "
        sql = sql & "botiga='" & codiBot & "' and day(data)=" & Day(fechaActual) & " and tipus_moviment='Z'"
        Set rsCaixes = Db.OpenResultset(sql)
        If Not rsCaixes.EOF Then
            If Abs((vMati + vTarda) - rsCaixes("import")) > 1 Then
                strInc = strInc & "*Desquadre vendes - caixes " & Format(fechaActual, "dd/mm/yyyy") & "<BR>"
            End If
        End If
        
        If vMati = 0 Then
            strInc = strInc & "*Falten vendes MATI dia " & Format(fechaActual, "dd/mm/yyyy") & "<BR>"
        End If
        'CORONAVIRUS, ... ALGUNAS TIENDAS NO ABREN POR LA TARDE
        'If vTarda = 0 Then
        '    strInc = strInc & "*Falten vendes TARDA dia " & Format(fechaActual, "dd/mm/yyyy") & "<BR>"
        'End If
        
        'If vMati = 0 Or vTarda = 0 Then
        '    ExecutaComandaSql "insert into missatgesAEnviar values('reenviadia', '" & codiBot & ":" & Day(fechaActual) & "-" & Month(fechaActual) & "-" & Year(fechaActual) & "')"
        'End If
    Next
    
    getNotasIncidencia = strInc
    
    Exit Function
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR getNotasIncidencia  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", sql & "  " & err.Description, "", ""
    
    
End Function

Sub SecreActualizaPrevisiones(fileName As String, emailDe As String, empresa As String)
    Dim intFile As Integer
    Dim HTMLText As String
    Dim lineHTML As String
    
    InformaMiss "SecreActualizaPrevisiones ", True
    
On Error GoTo nor

    HTMLText = ""
    
    'Open file
    intFile = FreeFile
    Open fileName For Input As intFile

    'Load XML into string strXML
    While Not EOF(intFile)
        Line Input #intFile, lineHTML
        'lineXML = Replace(lineXML, "ï»¿", "")
        HTMLText = HTMLText & lineHTML
    Wend
    Close intFile
    
    'MyKill fileName
    
    actualizaPrevisiones empresa, emailDe, HTMLText
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreActualizaPrevisiones [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "LO PIDE: [" & emailDe & "]<BR>" & err.Description, "", ""
End Sub

Sub SecreEmailResumenHoras(Optional subj As String, Optional email As String)
    Dim co As String
    Dim Semana As Integer, semanaAux As Integer, lunes As Date, f As Date
    Dim dni As String, rsDep As rdoResultset
    Dim sql As String, rsHoras As rdoResultset, botiga As String
    Dim sqlFechas As String, sqlPivot As String
    Dim rsContrato As rdoResultset, hContrato As Double, hTrabajadas As Double
    Dim vendesArr(3) As Double, clientsArr(2) As Integer, previsionsArr(2) As Double
    Dim horas(7) As Double, D As Integer, totalHoras As Double
    Dim previsioSemana As Double, vendesSemana As Double
    Dim factor(7) As Double
    Dim totalObjetivos As Double
    Dim plusFestivo As Double, importePlus As Double
        
    InformaMiss "SecreEmailResumenHoras", True

On Error GoTo ErrData

    plusFestivo = 16.05
    importePlus = 0
    
    Semana = 1
    If InStr(UCase(subj), "SEMANA") Then
        Semana = CInt(Split(UCase(subj), "SEMANA")(1))
    Else
        Semana = DatePart("ww", Now(), vbMonday)
    End If
    
    semanaAux = 0
    lunes = CDate("01/01/" & Year(Now()))

    If Semana > 0 Then
        While Semana <> semanaAux
            lunes = DateAdd("d", 1, lunes)
            semanaAux = DatePart("ww", lunes, vbMonday)
        Wend
    Else
        GoTo ErrData
    End If
    
    GoTo OkData
    
ErrData:
    lunes = DateAdd("d", -1, Now())
        
OkData:

On Error GoTo nor

    dni = Split(subj, " ")(2)
    Set rsDep = Db.OpenResultset("select d.codi, d.nom from dependentes d left join dependentesExtes de on d.codi=de.id and de.nom='DNI' where de.valor = '" & dni & "'")
    If Not rsDep.EOF Then
        co = "<H3>" & UCase(rsDep("nom")) & " SEMANA " & Semana & "</H3>"
        
        sqlFechas = "isnull([" & Format(lunes, "dd/mm/yyyy") & "], 0) [" & Format(lunes, "dd/mm/yyyy") & "],"
        sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "],"
        sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "],"
        sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "],"
        sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "],"
        sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "],"
        sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "]"
                   
        sqlPivot = "[" & Format(lunes, "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "]"
        
        sql = "select nom, codi, " & sqlFechas & " "
        sql = sql & "From "
        sql = sql & "(select c.codi, c.nom, convert(nvarchar, p.fecha, 103) fecha, case when isnull(t.horaInicio, '-')='-' then case when p.idturno like '%Aprendiz' or p.idturno like '%Coordinacion' or p.idturno like '%Extra' then left(p.idturno, charindex('_', p.idturno)-1) else '0' end else datediff(minute, t.horaInicio, t.horaFin )/60.0  end horas "
        sql = sql & "from " & taulaCdpPlanificacion(lunes) & " p "
        sql = sql & "left join cdpTurnos t on p.idturno=t.idTurno "
        sql = sql & "left join clients c on p.botiga=c.codi "
        sql = sql & "where idempleado='" & rsDep("codi") & "' and p.activo=1) dataTable "
        sql = sql & "PIVOT ( sum(horas) "
        sql = sql & "for fecha in (" & sqlPivot & ")) PivotTable "
        sql = sql & "order by nom"
        Set rsHoras = Db.OpenResultset(sql)
        
        hTrabajadas = 0
        totalObjetivos = 0
        
        While Not rsHoras.EOF

            co = co & "<TABLE BORDER=""1"" CELLPADDING=""0"" CELLSPACING=""0"">"
            co = co & "<TR><TD><B>Semana " & Semana & "</B></TD><TD align=""center""><B>Lunes</B></TD><TD align=""center""><B>Martes</B></TD><TD align=""center""><B>Miércoles</B></TD><TD align=""center""><B>Jueves</B></TD><TD align=""center""><B>Viernes</B></TD><TD align=""center""><B>Sábado</B></TD><TD align=""center""><B>Domingo</B></TD></TR>"
            co = co & "<TR><TD>&nbsp;</TD><TD align=""center""><B>" & Format(lunes, "dd/mm/yyyy") & "</B></TD><TD align=""center""><B>" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "</B></TD><TD align=""center""><B>" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "</B></TD><TD align=""center""><B>" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "</B></TD><TD align=""center""><B>" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "</B></TD><TD align=""center""><B>" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "</B></TD><TD align=""center""><B>" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "</B></TD></TR>"

            For D = 0 To 6
                factor(D) = 0
                horas(D) = 0
            Next
        
            previsioSemana = 0
            vendesSemana = 0
            
            co = co & "<TR><TD><B>" & UCase(rsHoras("nom")) & "</B></TD>"
            co = co & "<TD align=""right"">" & FormatNumber(rsHoras(Format(lunes, "dd/mm/yyyy")), 2) & " h</TD>"
            co = co & "<TD align=""right"">" & FormatNumber(rsHoras(Format(DateAdd("d", 1, lunes), "dd/mm/yyyy")), 2) & " h</TD>"
            co = co & "<TD align=""right"">" & FormatNumber(rsHoras(Format(DateAdd("d", 2, lunes), "dd/mm/yyyy")), 2) & " h</TD>"
            co = co & "<TD align=""right"">" & FormatNumber(rsHoras(Format(DateAdd("d", 3, lunes), "dd/mm/yyyy")), 2) & " h</TD>"
            co = co & "<TD align=""right"">" & FormatNumber(rsHoras(Format(DateAdd("d", 4, lunes), "dd/mm/yyyy")), 2) & " h</TD>"
            co = co & "<TD align=""right"">" & FormatNumber(rsHoras(Format(DateAdd("d", 5, lunes), "dd/mm/yyyy")), 2) & " h</TD>"
            co = co & "<TD align=""right"">" & FormatNumber(rsHoras(Format(DateAdd("d", 6, lunes), "dd/mm/yyyy")), 2) & " h</TD>"
            co = co & "</TR>"
            
            'Total horas semana todas las tiendas
            hTrabajadas = hTrabajadas + rsHoras(Format(lunes, "dd/mm/yyyy")) + rsHoras(Format(DateAdd("d", 1, lunes), "dd/mm/yyyy")) + rsHoras(Format(DateAdd("d", 2, lunes), "dd/mm/yyyy")) + rsHoras(Format(DateAdd("d", 3, lunes), "dd/mm/yyyy")) + rsHoras(Format(DateAdd("d", 4, lunes), "dd/mm/yyyy")) + rsHoras(Format(DateAdd("d", 5, lunes), "dd/mm/yyyy")) + rsHoras(Format(DateAdd("d", 6, lunes), "dd/mm/yyyy"))
            
            'Previsiones, ventas y horas trabajadas por dia/tienda
            For D = 0 To 6
                calculaVendesClients rsHoras("codi"), DateAdd("d", D, lunes), vendesArr, clientsArr 'Ventas
                calculaPrevisions rsHoras("codi"), DateAdd("d", D, lunes), previsionsArr 'Previsiones
                
                horas(D) = horas(D) + (rsHoras(Format(DateAdd("d", D, lunes), "dd/mm/yyyy"))) 'Horas dia/tienda
                If vendesArr(0) + vendesArr(1) > previsionsArr(0) + previsionsArr(1) Then 'Factor de objetivos por dia/tienda
                    factor(D) = 0.5
                End If
                
                previsioSemana = previsioSemana + previsionsArr(0) + previsionsArr(1) 'Previsiones semana/tienda
                vendesSemana = vendesSemana + vendesArr(0) + vendesArr(1) 'Ventas semana/tienda
                                
                'Plus festivo/domingo
                If rsHoras(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")) > 0 Then
                    If D = 6 Then
                        importePlus = importePlus + plusFestivo 'DOMINGO
                    Else
                        If esFestivo(DateAdd("d", D, lunes), rsDep("codi")) Then
                            importePlus = importePlus + plusFestivo 'FESTIVO
                        End If
                    End If
                End If
            Next
            
            'co = co & "<TR><TD COLSPAN=""8"">&nbsp;</TD></TR>"
            If previsioSemana < vendesSemana Then 'Si al acabar la semana se superan las previsiones se aplica un 0.75 a toda la semana
                For D = 0 To 6
                    factor(D) = 0.75
                Next
            End If
            
            co = co & "<TR><TD><B>Objetivos</TD>"
            For D = 0 To 6
                co = co & "<TD align=""right"">" & horas(D) * factor(D) & "&euro;</TD>"
                totalObjetivos = totalObjetivos + (horas(D) * factor(D))
            Next
            co = co & "</TR>"
            
            co = co & "</TABLE><BR>"
            
            rsHoras.MoveNext
        Wend

        
        co = co & "<BR>"
                
        sql = "select 40*(en.porjornada/100) jornadaHoras "
        sql = sql & "from silema_ts.sage.dbo.personas p "
        sql = sql & "left join silema_ts.sage.dbo.EmpleadoNomina en on p.dni=en.dni "
        sql = sql & "Where p.Dni='" & dni & "' and en.fechabaja Is Null And eN.CodigoEmpresa Is Not Null"
        hContrato = 0
        Set rsContrato = Db.OpenResultset(sql)
        If Not rsContrato.EOF Then hContrato = rsContrato("jornadaHoras")
        
        co = co & "<TABLE BORDER=""1"" CELLPADDING=""5"" CELLSPACING=""0"">"
        co = co & "<TR><TD><B>Horas Trabajadas semana " & Semana & "</B></TD><TD align=""right"">" & hTrabajadas & " h</TD></TR>"
        co = co & "<TR><TD><B>Horas contrato</B></TD><TD align=""right"">" & hContrato & " h</TD></TR>"
        co = co & "<TR><TD><B>Horas de más</B></TD><TD align=""right"">" & hTrabajadas - hContrato & " h</TD></TR>"
        co = co & "<TR><TD><B>Objetivos</B></TD><TD>" & FormatNumber(totalObjetivos, 2) & " &euro;</TD></TR>"
        co = co & "<TR><TD><B>Plus domingo y festivo</B></TD><TD>" & FormatNumber(importePlus, 2) & " &euro;</TD></TR>"
        co = co & "<TR><TD><B>Horas festivos</B></TD><TD>&nbsp;</TD></TR>"
        co = co & "</TABLE>"
    Else
        co = "DNI INCORRECTO"
    End If
    
    sf_enviarMail "secrehit@hit.cat", email, "Resumen horas " & dni & " [" & Format(lunes, "dd/mm/yyyy") & "]", co, "", ""
    
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreEmailResumenHoras [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "LO PIDE: [" & email & "]<BR>" & err.Description, "", ""

End Sub

Sub SecreEmailSupervisoras(Optional subj As String, Optional email As String)
    Dim rsSup As rdoResultset, rsTipus As rdoResultset, rsFranquicias As rdoResultset
    Dim fecha As Date
    Dim depId As String, depNom As String, depEMail As String
    Dim sql As String
    Dim Semana As Integer, semanaAux As Integer, lunes As Date
    Dim conCopia As Boolean
    
    InformaMiss "SecreEmailSupervisoras", True
    
On Error GoTo ErrData
'    If InStr(UCase(subj), "SEMANA") Then
'        Semana = CInt(Split(subj, " ")(3))
'        semanaAux = 0
'        lunes = CDate("01/01/" & Year(Now()))

'        If Semana > 0 Then
'            While Semana <> semanaAux
'                lunes = DateAdd("d", 1, lunes)
'                semanaAux = DatePart("ww", lunes, vbMonday)
'            Wend
    
'            fecha = DateAdd("d", 6, lunes)
'
'        End If
'    Else
'        fecha = DateAdd("d", -1, Now())
'    End If
    
    
    
   If InStr(UCase(subj), "SEMANA") Then
        Semana = CInt(Split(subj, " ")(3))
        semanaAux = 0
        If Semana = 1 Then
            fecha = CDate("26/12/" & Year(Now()) - 1) 'La primera semana del año como mucho podría ser del lunes 26 al domingo 1
            While Weekday(fecha, 2) <> 1
                fecha = DateAdd("d", 1, fecha)
            Wend
            fecha = DateAdd("d", 6, fecha)
        Else
            lunes = CDate("01/01/" & Year(Now()))
            If Semana > 0 Then
                While Semana <> semanaAux
                    lunes = DateAdd("d", 1, lunes)
                    semanaAux = DatePart("ww", lunes, vbMonday)
                Wend
        
                fecha = DateAdd("d", 6, lunes)
            Else
                fecha = DateAdd("d", -1, Now())
            End If
        End If
    Else
        fecha = DateAdd("d", -1, Now())
    End If
    
    GoTo OkData
    
ErrData:
    fecha = DateAdd("d", -1, Now())
        
OkData:
    
On Error GoTo nor

    conCopia = False
    If email <> "" Then 'ALGUIEN A PEDIDO EL INFORME
        Set rsSup = Db.OpenResultset("select * from dependentes d left join dependentesExtes d2 on d.codi=d2.id and d2.nom='EMAIL' where d2.valor like '%" & email & "%' and d.codi in (select distinct valor from constantsclient where variable = 'SupervisoraCodi' and valor<>'')")
        If rsSup.EOF Then 'NO es supervisora, comprobamos si es GERENTE
            sql = "select * "
            sql = sql & "from dependentes d "
            sql = sql & "left join dependentesExtes d1 on d.codi=d1.id and d1.nom='TIPUSTREBALLADOR' "
            sql = sql & "left join dependentesExtes d2 on d.codi=d2.id and d2.nom='EMAIL' "
            sql = sql & "where d2.valor like '%" & email & "%' and d1.valor in ('GERENT', 'GERENT_2') "

            Set rsTipus = Db.OpenResultset(sql)
            If Not rsTipus.EOF Then 'Si es GERENTE le pasamos informe de todas las supervisoras
                Set rsSup = Db.OpenResultset("select d.*, isnull(de.valor, '') eMail from dependentes d left join dependentesextes de on d.codi=de.id and de.nom='EMAIL' where d.codi in (select distinct valor from constantsclient where variable = 'SupervisoraCodi' and valor<>'')")
            Else 'SI NO ES NI GERENTE NI SUPERVISORA NO SE ENVÍA
                sf_enviarMail "secrehit@hit.cat", email, "Informe Supervisoras [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "NO TENS PERMÍS PER REBRE AQUESTA INFORMACIÓ", "", ""
                Exit Sub
            End If
        End If
    Else 'SI NO HAY EMAIL, ES EL ENVÍO AUTOMÁTICO A TODAS LAS SUPERVISORAS
        If Weekday(Now(), 2) = 1 Then conCopia = True
        Set rsSup = Db.OpenResultset("select d.*, isnull(de.valor, '') eMail from dependentes d left join dependentesextes de on d.codi=de.id and de.nom='EMAIL' where d.codi in (select distinct valor from constantsclient where variable = 'SupervisoraCodi' and valor<>'')")
    End If
    
    While Not rsSup.EOF
        depId = rsSup("codi")
        depNom = rsSup("Nom")
        If email <> "" Then
            depEMail = email
        Else
            depEMail = rsSup("EMail")
        End If
        
        If depEMail <> "" Then 'Sin no hay email ya no seguimos
            'TIENDAS PROPIAS
            SecreInformeSupervisora fecha, depId, depNom, depEMail, False, conCopia
            
            'FRANQUICIAS
            Set rsFranquicias = Db.OpenResultset("select codi, upper(nom) nom From clients where codi in (select codi from constantsclient where variable='SupervisoraCodi' and valor='" & depId & "') and codi in (select codi from constantsclient where variable='Franquicia' and valor='Franquicia')")
            If Not rsFranquicias.EOF Then SecreInformeSupervisora fecha, depId, depNom, depEMail, True, conCopia
        End If
        rsSup.MoveNext
    Wend
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreEmailSupervisoras  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", err.Description, "", ""
End Sub


Sub SecreInformePersonal(subj As String, emailDe As String, empresa As String)
    Dim co As String
    Dim rs As rdoResultset, rsHoras As rdoResultset
    Dim sql As String
    Dim lunes As Date, D As Integer, Semana As Integer, semanaAux As Integer
    Dim totalHoras As Double, horasDia As Double, estado As String
    Dim diaActual As Date, diaSql As Date
    
    InformaMiss "SecreInformePersonal ", True
    
 On Error GoTo ErrData
    Semana = -1

    If InStr(UCase(subj), "SEMANA") Then
        Semana = CInt(Split(subj, " ")(3))
        semanaAux = 0
        lunes = CDate("01/01/" & Year(Now()))

        If Semana > 0 Then
            While Semana <> semanaAux
                lunes = DateAdd("d", 1, lunes)
                semanaAux = DatePart("ww", lunes, vbMonday)
            Wend
        Else
            lunes = DateAdd("d", -((Weekday(Now(), 2)) - 1), Now())
        End If
    Else
        lunes = DateAdd("d", -((Weekday(Now(), 2)) - 1), Now())
    End If
    
GoTo OkData
    
ErrData:

    lunes = DateAdd("d", -((Weekday(Now(), 2)) - 1), Now())
    'lunes = Now()
    'While DatePart("w", lunes) <> vbMonday
    '    lunes = DateAdd("d", -1, lunes)
    'Wend
        
OkData:
    
On Error GoTo nor
     
    co = co & "<TABLE BORDER='1'>"
    co = co & "<TR>"
    co = co & "<TD><B>Empresa</B></TD><TD><B>Categoría</B></TD><TD><B>Nombre HIT</B></TD><TD><B>Nombre SAGE</B></TD><TD><B>" & utf8("Antigüedad") & "</B></TD><TD><B>Horas contratadas</B></TD>"
    For D = 0 To 6
        co = co & "<TD><B>" & Format(DateAdd("d", D, lunes), "dd/mm/yyyy") & "</B></TD>"
    Next
    co = co & "<TD><B>Total</B></TD>"
    co = co & "<TD><B>Plus</B></TD>"
    co = co & "</TR>"
    
    sql = "select distinct isnull(dHit2.codi, 9999) idHit, isnull(e.empresa, '') empresa, isnull(ep.categoria, '') categoria, isnull(dHit2.nom, '') nomHit, isnull(p.NombreEmpleado, '') + ' ' + isnull(p.PrimerApellidoEmpleado, '') + ' ' + isnull(p.SegundoApellidoEmpleado, '') nomSAGE, en.fechaAntiguedad, 40*(en.porjornada/100) jornadaHoras "
    sql = sql & "from silema_ts.sage.dbo.personas p "
    sql = sql & "left join silema_ts.sage.dbo.EmpleadoNomina en on p.dni=en.dni "
    sql = sql & "left join silema_ts.sage.dbo.empresas e on e.codigoempresa=en.CodigoEmpresa "
    sql = sql & "left join (select CodigoEmpresa, IdEmpleado, Categoria, max(fechaInicioCategoria) FechaInicioCategoria from silema_ts.sage.dbo.EmpleadoNominaCategorias group by codigoEmpresa, IdEmpleado, Categoria) ep on en.CodigoEmpresa=ep.codigoempresa and en.idempleado=ep.idEmpleado "
    sql = sql & "left join dependentesExtes dHit on p.dni collate SQL_Latin1_General_CP1_CI_AS = dHit.valor "
    sql = sql & "left join dependentes dHit2 on dHit.id=dHit2.CODI "
    sql = sql & "Where eN.fechabaja Is Null And eN.CodigoEmpresa Is Not Null "
    sql = sql & "order by isnull(e.empresa, ''), isnull(ep.categoria, ''),  isnull(p.NombreEmpleado, '') + ' ' + isnull(p.PrimerApellidoEmpleado, '') + ' ' + isnull(p.SegundoApellidoEmpleado, '')"
    Set rs = Db.OpenResultset(sql)
    
    While Not rs.EOF
        co = co & "<TR><TD>" & utf8(rs("empresa")) & "</TD><TD>" & utf8(rs("categoria")) & "</TD><TD>" & utf8(rs("nomHit")) & "</TD><TD>" & utf8(rs("nomSAGE")) & "</TD><TD>" & rs("fechaAntiguedad") & "</TD><TD>" & FormatNumber(rs("jornadaHoras"), 2)
        
        'Sql = "select convert(datetime, concat(day(p.fecha), '/', month(p.fecha),  '/', year(p.fecha)), 103) fecha, case when isnull(t.horaInicio, '-')='-' then case when p.idturno like '%Aprendiz' or p.idturno like '%Coordinacion' or p.idturno like '%Extra' then left(p.idturno, charindex('_', p.idturno)-1) end else datediff(minute, t.horaInicio, t.horaFin )/60  end horas "
        'Sql = Sql & "from " & taulaCdpPlanificacion(lunes) & " p "
        'Sql = Sql & "left join cdpturnos t on p.idturno = t.idturno "
        'Sql = Sql & "left join dependentes d on p.idEmpleado = d.codi "
        'Sql = Sql & "Where P.idEmpleado = " & Rs("idHit") & " And P.activo = 1 "
        'Sql = Sql & "order by convert(datetime, concat(day(p.fecha), '/', month(p.fecha),  '/', year(p.fecha)), 103)"
        sql = "select c.fecha, c.estado, case when isnull(t.horaInicio, '-')='-' then case when p.idturno like '%Aprendiz' or p.idturno like '%Coordinacion' or p.idturno like '%Extra' then left(p.idturno, charindex('_', p.idturno)-1) else '0' end else datediff(minute, t.horaInicio, t.horaFin )/60.0  end horas "
        sql = sql & "from cdpCalendariLaboral_" & Year(lunes) & " c "
        sql = sql & "left join " & taulaCdpPlanificacion(lunes) & " p on c.idempleado=p.idempleado and c.fecha = convert(datetime, concat(day(p.fecha), '/', month(p.fecha), '/', year(p.fecha)), 103) "
        sql = sql & "left join cdpturnos t on p.idturno = t.idturno "
        sql = sql & "left join dependentes d on p.idEmpleado = d.codi "
        sql = sql & "Where c.idEmpleado = " & rs("idHit") & " and c.fecha between convert(datetime, '" & Day(lunes) & "/" & Month(lunes) & "/" & Year(lunes) & "', 103) and convert(datetime, '" & Day(DateAdd("d", 6, lunes)) & "/" & Month(DateAdd("d", 6, lunes)) & "/" & Year(DateAdd("d", 6, lunes)) & "', 103) "
        sql = sql & "order by c.fecha"
        
        Set rsHoras = Db.OpenResultset(sql)
        totalHoras = 0
        For D = 0 To 6
            diaActual = DateAdd("d", D, lunes)
            InformaMiss "SecreInformePersonal " & diaActual & " Empleado: " & rs("idHit"), True
            horasDia = 0
            estado = ""
            If Not rsHoras.EOF Then
                diaSql = rsHoras("fecha")
                While Day(diaSql) = Day(diaActual) And Month(diaSql) = Month(diaActual) And Year(diaSql) = Year(diaActual)
                    If IsNumeric(rsHoras("horas")) Then
                        horasDia = horasDia + rsHoras("horas")
                    End If
                    estado = rsHoras("Estado")
                    If estado = "LABORABLE" Then estado = ""
                    
                    rsHoras.MoveNext
                    If Not rsHoras.EOF Then
                        diaSql = rsHoras("fecha")
                    Else
                        diaSql = CDate("01/01/2000")
                    End If
                Wend
                co = co & "<TD>" & horasDia & " " & estado & "</TD>"

                totalHoras = totalHoras + horasDia
            Else
                co = co & "<TD>-</TD>"
            End If
        Next
        
        co = co & "<TD>" & totalHoras & "</TD>"
        co = co & "<TD>" & FormatNumber(totalHoras - rs("jornadaHoras"), 2) & "</TD>"
        co = co & "</TR>"
         
        rs.MoveNext
    Wend
    
    co = co & "</TABLE>"
    
    sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
        
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "ERROR: SecreInformePersonal " & err.Description, "", ""
    
End Sub

Sub SecreInformeProductos(subj As String, emailDe As String, empresa As String)
    Dim Semana As Integer, semanaAux As Integer, lunes As Date, domingo As Date, fecha As Date, rs As rdoResultset, rsVendes As rdoResultset, sql As String, iD As String, rsA As ADODB.Recordset
    Dim nomBotiga As String, codiDep As String
    Dim co As String, supervisora As String, esGerente As Boolean
    
    InformaMiss "SecreInformeProductos", True

    ExecutaComandaSql "SET DATEFIRST 7"
                        
    Semana = -1
On Error GoTo ErrData
    If InStr(UCase(subj), "SEMANA") Then
        Semana = CInt(Split(subj, " ")(3))
        semanaAux = 0
        lunes = CDate("01/01/" & Year(Now()))
        If Semana > 0 Then
            While Semana <> semanaAux
                lunes = DateAdd("d", 1, lunes)
                semanaAux = DatePart("ww", lunes, vbMonday, vbFirstFullWeek)
            Wend
            
            domingo = DateAdd("d", 6, lunes)
            fecha = lunes
        End If
    Else
        fecha = CDate(Split(subj, " ")(2))
        If Year(fecha) > Year(Now()) Then GoTo ErrData
    End If
GoTo OkData
    
ErrData:

    fecha = Now()
    
OkData:

On Error GoTo nor
    'emailDe = "apujol@silemabcn.com"
    'emailDe = "cescuder@silemabcn.com"
    'emailDe = "lgarcia@silemabcn.com"
     
    ExecutaComandaSql "Select * from [" & NomTaulaLikes(lunes) & "] "
    
     esGerente = False
     
     Set rs = Db.OpenResultset("select * from dependentesextes where nom='EMAIL' and upper(valor) like '%' + upper('" & emailDe & "') + '%' order by len(valor)")
     If Not rs.EOF Then
         codiDep = rs("id")
         Set rs = Db.OpenResultset("select * from constantsClient where variable='SupervisoraCodi' and valor='" & codiDep & "'")
         If rs.EOF Then 'NO ES SUPERVISORA
             Set rs = Db.OpenResultset("select * from dependentesextes where nom='TIPUSTREBALLADOR' and id='" & codiDep & "'")
             If rs("Valor") = "GERENT" Or rs("Valor") = "GERENT_2" Then
                 sql = "select c.codi, c.nom, isnull(d.nom, ' Franquicia') Supervisora "
                 sql = sql & "from paramshw p "
                 sql = sql & "left join clients c on p.valor1=c.codi "
                 sql = sql & "left join constantsClient cc on c.codi=cc.codi and cc.variable='SupervisoraCodi' "
                 sql = sql & "left join dependentes d on cc.valor = d.codi "
                 sql = sql & "where isnull(c.nom, '') <> '' "
                 sql = sql & "order by isnull(d.nom, ' Franquicia') , c.nom "
                 Set rs = Db.OpenResultset(sql)
             Else
                 Exit Sub
             End If
             esGerente = True
         Else 'SUPERVISORA
             Set rs = Db.OpenResultset("Select c.Codi, c.nom, isnull(d.nom, ' Franquicia') Supervisora from ConstantsClient cc left join clients c on cc.codi=c.codi left join dependentes d on cc.valor = d.codi where variable = 'SupervisoraCodi' and valor = '" & codiDep & "' order by c.nom")
             esGerente = False
         End If
     End If
     
     co = ""
     
     'Top 10 tiendas propias
     If esGerente Then
        If Semana > 0 Then
            co = co & "<table cellpadding='0' cellspacing='3' border='0'>"
            co = co & "<tr>"
            co = co & "<td><h3>TOP TEN DE TIENDAS PROPIAS SEMANA " & Semana & "</h3></td>"
            co = co & "<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>"
            co = co & "<td><h3>TOP TEN DE TIENDAS PROPIAS SEMANA " & Semana - 1 & "</h3></td>"
            co = co & "</tr>"
        Else
            co = co & "<h3>TOP TEN DE TIENDAS PROPIAS</h3>"
        End If
    
        'Ventas
        If Semana > 0 Then
            sql = "select top 10 a.nom, sum(import) Import, sum(quantitat) Quantitat "
            If Month(lunes) <> Month(domingo) Then
                sql = sql & "from (select * from [" & NomTaulaVentas(lunes) & "] union all select * from [" & NomTaulaVentas(domingo) & "]) v  "
            Else
                sql = sql & "from [" & NomTaulaVentas(lunes) & "] v "
            End If
            sql = sql & "left join articles a on v.plu=a.codi "
            sql = sql & "where botiga in (select codi from clients where nif in (select valor from constantsempresa where camp like '%nif%' and valor is not null) and codi in (select valor1 from paramshw)) "
            sql = sql & "and data between '" & lunes & " 00:00' and '" & domingo & " 23:59' "
            sql = sql & "group by a.nom "
            sql = sql & "order by import desc"
            
            co = co & "<tr><td>"
        Else
            sql = "select top 10 a.nom, sum(import) Import, sum(quantitat) Quantitat "
            sql = sql & "from [" & NomTaulaVentas(fecha) & "] v "
            sql = sql & "left join articles a on v.plu=a.codi "
            sql = sql & "where botiga in (select codi from clients where nif in (select valor from constantsempresa where camp like '%nif%' and valor is not null) and codi in (select valor1 from paramshw)) "
            sql = sql & "and day(data)=" & Day(fecha) & " "
            sql = sql & "group by a.nom "
            sql = sql & "order by import desc"
        End If
        
        co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
        co = co & "<Tr bgColor='#DADADA'><Td><b>Producte</b></Td><Td><b>Import</b></Td><Td><b>Quantitat</b></Td></Tr>"
        
        Set rsVendes = Db.OpenResultset(sql)
        While Not rsVendes.EOF
             co = co & "<tr><Td><b>" & rsVendes("nom") & "</b></Td><Td align=""right"">" & FormatNumber(rsVendes("Import"), 2) & "&euro;</Td><Td align=""right"">" & FormatNumber(rsVendes("Quantitat"), 2) & "</Td></Tr>"
            rsVendes.MoveNext
        Wend
     
        co = co & "</table>"
        
        'Semana anterior
        If Semana > 0 Then
            co = co & "</td><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>"
            
            sql = "select top 10 a.nom, sum(import) Import, sum(quantitat) Quantitat "
            If Month(DateAdd("d", -7, lunes)) <> Month(DateAdd("d", -7, domingo)) Then
                sql = sql & "from (select * from [" & NomTaulaVentas(DateAdd("d", -7, lunes)) & "] union all select * from [" & NomTaulaVentas(DateAdd("d", -7, domingo)) & "]) v  "
            Else
                sql = sql & "from [" & NomTaulaVentas(DateAdd("d", -7, lunes)) & "] v "
            End If
            sql = sql & "left join articles a on v.plu=a.codi "
            sql = sql & "where botiga in (select codi from clients where nif in (select valor from constantsempresa where camp like '%nif%' and valor is not null) and codi in (select valor1 from paramshw)) "
            sql = sql & "and data between '" & DateAdd("d", -7, lunes) & " 00:00' and '" & DateAdd("d", -7, domingo) & " 23:59' "
            sql = sql & "group by a.nom "
            sql = sql & "order by import desc"
            
            co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
            co = co & "<Tr bgColor='#DADADA'><Td><b>Producte</b></Td><Td><b>Import</b></Td><Td><b>Quantitat</b></Td></Tr>"
            
            Set rsVendes = Db.OpenResultset(sql)
            While Not rsVendes.EOF
                co = co & "<tr><Td><b>" & rsVendes("nom") & "</b></Td><Td align=""right"">" & FormatNumber(rsVendes("Import"), 2) & "&euro;</Td><Td align=""right"">" & FormatNumber(rsVendes("Quantitat"), 2) & "</Td></Tr>"
               rsVendes.MoveNext
            Wend
        
            co = co & "</table>"
            
            co = co & "</td></tr>"
            co = co & "</table>"
        End If
        
        co = co & "<br><br>"
     
     End If
     
     
     supervisora = ""
     If Semana > 0 Then
         co = co & "<h3>Informe Productos semana " & Semana & " (" & lunes & " - " & domingo & ")</h3>"
     Else
         co = co & "<h3>Informe Productos dia " & fecha & "</h3>"
     End If
     
     While Not rs.EOF
        If supervisora <> rs("supervisora") Then
             If supervisora <> "" Then
                co = co & "</table><br>"
             End If
             co = co & "<h4>" & rs("supervisora") & "</h4>"
             co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
             co = co & "<Tr bgColor='#DADADA'><Td><b>Botiga</b></Td><Td><b>Producte</b></Td><Td><b>Import</b></Td><Td><b>Quantitat</b></Td></Tr>"
             supervisora = rs("supervisora")
        End If
         
        'Nom botiga
        nomBotiga = BotigaCodiNom(CDbl(rs("Codi")))
        co = co & "<tr bgColor='#EAEAEA'><td colspan=""4""><b>" & UCase(nomBotiga) & "</b></td></tr>"
    
        'Ventas
        If Semana > 0 Then
            sql = "select top 10 a.nom, sum(import) Import, sum(quantitat) Quantitat "
            If Month(lunes) <> Month(domingo) Then
                sql = sql & "from (select * from [" & NomTaulaVentas(lunes) & "] union all select * from [" & NomTaulaVentas(domingo) & "]) v  "
            Else
                sql = sql & "from [" & NomTaulaVentas(lunes) & "] v "
            End If
            sql = sql & "left join articles a on v.plu=a.codi "
            sql = sql & "where botiga='" & rs("Codi") & "' and data between '" & lunes & " 00:00' and '" & domingo & " 23:59' "
            sql = sql & "group by a.nom "
            sql = sql & "order by import desc"
        Else
            sql = "select top 10 a.nom, sum(import) Import, sum(quantitat) Quantitat "
            sql = sql & "from [" & NomTaulaVentas(fecha) & "] v "
            sql = sql & "left join articles a on v.plu=a.codi "
            sql = sql & "where botiga='" & rs("Codi") & "' and day(data)=" & Day(fecha) & " "
            sql = sql & "group by a.nom "
            sql = sql & "order by import desc"
        End If
        
        Set rsVendes = Db.OpenResultset(sql)
        While Not rsVendes.EOF
             co = co & "<tr><Td>&nbsp;</Td><Td><b>" & rsVendes("nom") & "</b></Td><Td align=""right"">" & FormatNumber(rsVendes("Import"), 2) & "&euro;</Td><Td align=""right"">" & FormatNumber(rsVendes("Quantitat"), 2) & "</Td></Tr>"
            rsVendes.MoveNext
        Wend
         
         rs.MoveNext
     Wend
   
     co = co & "</table><br>"
     co = co & "<br>"
     
     sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
     sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
     
nor:

End Sub

Sub SecreInformeReposicion(subj As String, emailDe As String, empresa As String)
    Dim rsBots As rdoResultset, rsProv As rdoResultset, rsMP As rdoResultset, rsInventos As rdoResultset
    Dim co As String
    
    InformaMiss "SecreInformeReposicion", True

On Error GoTo nor
     
     
    'PROVEEDORES DE REPOSICIÓN
    Set rsProv = Db.OpenResultset("select distinct id, nombre from ccproveedores where id in (select valor from ccnombrevalor where left(nombre,13) = 'P_REPOSICION_')")
    While Not rsProv.EOF
        co = co & "<h3>" & rsProv("nombre") & "</h3>"
        co = co & "<table cellpadding='0' cellspacing='3' border='1'>"
        co = co & "<tr>"
        co = co & "<td nowrap><b>Article/Botiga</b></td>"
        'TIENDAS ----------------------------------------------------------
        Set rsBots = Db.OpenResultset("select c.codi, upper(c.nom) nom from paramshw h left join clients c on h.valor1=c.codi where c.nom is not null order by c.nom")
        While Not rsBots.EOF
            co = co & "<td nowrap><b>" & rsBots("nom") & "</b></td>"
            rsBots.MoveNext
        Wend
        co = co & "</tr>"
        
        Set rsMP = Db.OpenResultset("select * from ccmateriasprimas where id in (select distinct id from ccnombrevalor Pr where Pr.valor = '" & rsProv("id") & "' and left(Pr.nombre,13) = 'P_REPOSICION_' )")
        While Not rsMP.EOF
            co = co & "<tr>"
            co = co & "<td nowrap><b>" & rsMP("nombre") & "</b></td>"
            rsBots.MoveFirst
            While Not rsBots.EOF
                Set rsInventos = Db.OpenResultset("select top 1 fecha, redondeoCliente Inventos from [" & NomTaulaPedidosUltimosDatos(Now(), rsProv("id")) & "] where cliente=" & rsBots("codi") & " and materia='" & rsMP("id") & "' order by fecha desc")
                If Not rsInventos.EOF Then
                    If rsInventos("Inventos") > 20 Then
                        co = co & "<td bgcolor='#ff1111'>" & rsInventos("Inventos") & " <br> " & Format(rsInventos("fecha"), "dd/mm/yy") & "</td>"
                    Else
                        co = co & "<td>" & rsInventos("Inventos") & " <br> " & Format(rsInventos("fecha"), "dd/mm/yy") & "</td>"
                    End If
                Else
                    Set rsInventos = Db.OpenResultset("select top 1 fecha, redondeoCliente Inventos from [" & NomTaulaPedidosUltimosDatos(DateAdd("m", -1, Now()), rsProv("id")) & "] where cliente=" & rsBots("codi") & " and materia='" & rsMP("id") & "' order by fecha desc")
                    If Not rsInventos.EOF Then
                        If rsInventos("Inventos") > 20 Then
                            co = co & "<td bgcolor='#ff1111'>" & rsInventos("Inventos") & " <br> " & Format(rsInventos("fecha"), "dd/mm/yy") & "</td>"
                        Else
                            co = co & "<td>" & rsInventos("Inventos") & " <br> " & Format(rsInventos("fecha"), "dd/mm/yy") & "</td>"
                        End If
                    Else
                        co = co & "<td>&nbsp;</td>"
                    End If
                End If
                rsBots.MoveNext
            Wend
            co = co & "</tr>"
            rsMP.MoveNext
        Wend
        
        co = co & "</table>"
        co = co & "<br>"
        rsProv.MoveNext
    Wend
     
    sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    
    Exit Sub
     
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreInformeReposicion [" & Format(Now(), "dd/mm/yy hh:nn") & "]", err.Description, "", ""
End Sub


Sub SecreInformeBocadillos(subj As String, emailDe As String, empresa As String)
    Dim co As String, sql As String
    Dim rsVendes As rdoResultset, rsBocadillos As rdoResultset, rsServit As rdoResultset
    Dim hoy As Date, semAnt As Date
    Dim botiga As String, BotigaNom As String
    Dim venut As Double, servit As Double, venutHoy As Double, faltan As Double, LatasDe As Integer, frescura As Integer
    Dim rsBotiga As rdoResultset, rsLloc As rdoResultset
    Dim ambient As String, algun As Integer, article As String
    
    InformaMiss "SecreInformeBocadillos", True
    
    botiga = ""
    hoy = Now()
    semAnt = DateAdd("d", -7, hoy)
    
On Error GoTo nor
    
    If InStr(subj, " ") Then
        Set rsBotiga = Db.OpenResultset("select * from clients where nom = '" & Split(subj, " ")(1) & "'")
        If Not rsBotiga.EOF Then
            botiga = rsBotiga("codi")
            BotigaNom = rsBotiga("nom")
        Else 'HA PUESTO MAL EL NOMBRE DE LA TIENDA
            GoTo emailInfo
        End If
    Else
        Set rsLloc = Db.OpenResultset("select isnull(lloc, '') lloc from cdpdadesfichador where usuari=(select top 1 id from dependentesextes where valor='" & emailDe & "') and accio=1 and day(tmst)=" & Day(hoy) & " and month(tmst)=" & Month(hoy) & " and year(tmst)=" & Year(hoy))
        If Not rsLloc.EOF Then
            If rsLloc("lloc") <> "" Then
                Set rsBotiga = Db.OpenResultset("select * from clients where codi = '" & rsLloc("lloc") & "'")
                If Not rsBotiga.EOF Then
                    botiga = rsBotiga("codi")
                    BotigaNom = rsBotiga("nom")
                Else 'HA FICHADO, PERO NO ES DEPENDIENTA
                    GoTo emailInfo
                End If
            Else 'HA FICHADO, PERO NO ES DEPENDIENTA
                GoTo emailInfo
            End If
        Else 'NO HA FICHADO
            GoTo emailInfo
        End If
    End If
    
    
    ambient = ""
    algun = 0
    co = ""
    
    sql = "select a.codi, a.nom, case when isnull(p.valor, '')= '' then '0' else p.valor end latas, ambient, case when isnull(p2.valor, '')= '' then '120' else p2.valor end fres, a.familia "
    sql = sql & "from teclatstpv t "
    sql = sql & "left join articles a on t.article=a.codi "
    sql = sql & "left join ArticlesPropietats p on p.CodiArticle=a.codi and p.Variable ='UnitatsCoccio' "
    sql = sql & "left join ArticlesPropietats p2 on p2.CodiArticle=a.codi and p2.Variable ='MinutosFrescura' "
    sql = sql & "where llicencia=" & botiga & " and a.codi is not null and data = (select top 1 data from teclatstpv where llicencia=" & botiga & " order by data desc) "
    sql = sql & "order by ambient, a.nom"
        
    Set rsBocadillos = Db.OpenResultset(sql)
    While Not rsBocadillos.EOF
    
        If ambient <> rsBocadillos("ambient") Then
            If algun = 1 Then
                co = co & "</TABLE>"
                
                sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
                'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", Co, "", ""
                'sf_enviarMail "secrehit@hit.cat", "jordi@hit.cat", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", Co, "", ""
                'sf_enviarMail "secrehit@hit.cat", "jaTena@SilemaBcn.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", Co, "", ""
            End If
            
            co = "<H3>" & UCase(rsBocadillos("ambient")) & BotigaNom & " ( " & Format(Now(), "dd/mm/yyyy") & ")</H3>"
            co = co & "<TABLE cellpadding='0' cellspacing='3' border='1'>"
            co = co & "<TR>"
            co = co & "<TD><B>Article</TD>"
            'Co = Co & "<TD align='center'><B>Vendido<br>HOY</B></TD>"
            'Co = Co & "<TD align='center'><B>Vendido " & Format(DateAdd("h", 3, semAnt), "dd/mm/yyyy") & "<br>de " & Format(semAnt, "hh:nn:ss") & " a " & Format(DateAdd("h", 3, semAnt), "hh:nn:ss") & "</B></TD>"
            'Co = Co & "<TD align='center'><B>Servido<br>HOY</B></TD>"
            co = co & "<TD align='center'><B>Faltan</B></TD>"
            co = co & "</TR>"
            
            algun = 0
        End If
        
        article = rsBocadillos("codi")
        ambient = rsBocadillos("ambient")
        LatasDe = rsBocadillos("latas")
        'frescura = rsBocadillos("frescura")
        'If Not IsNumeric(rsBocadillos("fres")) Then
            frescura = 120
        'Else
        '    frescura = rsBocadillos("fres")
        'End If
    
        If LatasDe > 0 Then
                
            'Vendido hoy
            'Producto $Article
            sql = "select isnull(sum(quantitat), 0) qtat from [" & NomTaulaVentas(hoy) & "] v where v.botiga=" & botiga & " and day(data)= " & Day(hoy) & " and v.plu=" & rsBocadillos("codi")
            Set rsVendes = Db.OpenResultset(sql)
            venutHoy = rsVendes("qtat")
            'Co = Co & "<TD align='center'>" & IIf(venutHoy = 0, " ", venutHoy) & "</TD>"
            
            'Productos equivalentes a $Article
            sql = "select isnull(sum(quantitat*isnull(UnitatsEquivalencia, 1)), 0) qtat "
            sql = sql & "from [" & NomTaulaVentas(hoy) & "] v "
            sql = sql & "left join EquivalenciaProductes ep on v.plu = ep.prodVenut "
            sql = sql & "where v.botiga=" & botiga & " and day(data)=" & Day(hoy) & " and ep.prodServit=" & rsBocadillos("codi")
            Set rsVendes = Db.OpenResultset(sql)
            venutHoy = venutHoy + rsVendes("qtat")
                                
            'Semana pasada
            sql = "select isnull(sum(quantitat), 0) qtat from [" & NomTaulaVentas(semAnt) & "] v where v.botiga=" & botiga & " and convert(datetime, data, 103) between '" & Format(semAnt, "dd/mm/yyyy hh:nn:ss") & "' and '" & Format(DateAdd("h", 3, semAnt), "dd/mm/yyyy hh:nn:ss") & "' and v.plu=" & rsBocadillos("codi")
            Set rsVendes = Db.OpenResultset(sql)
            venut = rsVendes("qtat")
            'Co = Co & "<TD align='center'>" & IIf(venut = 0, " ", venut) & "</TD>"
            
            'Productos equivalentes a $Article
            sql = "select isnull(sum(quantitat*isnull(UnitatsEquivalencia, 1)), 0) qtat "
            sql = sql & "from [" & NomTaulaVentas(semAnt) & "] v "
            sql = sql & "left join EquivalenciaProductes ep on v.plu = ep.prodVenut "
            sql = sql & "where v.botiga=" & botiga & " and "
            sql = sql & "convert(datetime, data, 103) between '" & Format(semAnt, "dd/mm/yyyy hh:nn:ss") & "' and '" & Format(DateAdd("h", 3, semAnt), "dd/mm/yyyy hh:nn:ss") & "' and ep.prodServit=" & rsBocadillos("codi")
            Set rsVendes = Db.OpenResultset(sql)
            venut = venut + rsVendes("qtat")
            
            'Servit
            sql = "select isnull(sum(quantitatServida), 0) qtat from " & DonamTaulaServit(hoy) & " where client=" & botiga & " and codiArticle=" & rsBocadillos("codi")
            Set rsServit = Db.OpenResultset(sql)
            servit = rsServit("qtat")
            'Co = Co & "<TD align='center'>" & IIf(servit = 0, " ", servit) & "</TD>"
            
            'Equivalentes
            sql = "select isnull(sum(quantitatServida*isnull(UnitatsEquivalencia, 1)), 0) qtat "
            sql = sql & "from " & DonamTaulaServit(hoy) & " s "
            sql = sql & "left join EquivalenciaProductes ep on s.CodiArticle = ep.prodVenut "
            sql = sql & "where client=" & botiga & " and ep.prodServit=" & rsBocadillos("codi")
            Set rsServit = Db.OpenResultset(sql)
            servit = servit + rsServit("qtat")
                        
                        
                                                       
            faltan = venutHoy + venut - servit
            If faltan < 0 Then faltan = 0
            If servit = 0 And faltan = 0 Then faltan = LatasDe
            
            If faltan > 0 Then
                If Int(faltan / LatasDe) = faltan / LatasDe Then faltan = Int(faltan / LatasDe) * LatasDe
            Else
                faltan = (Int(faltan / LatasDe) + 1) + LatasDe
            End If
            
            If faltan > 0 Then
                co = co & "<TR>"
                co = co & "<TD><B>" & rsBocadillos("nom") & "</B></TD>"
            
                co = co & "<TD align='center'>" & faltan & "</TD>"
                
                algun = 1
                
                'Actualitza servit
                ExecutaComandaSql "insert into " & DonamTaulaServit(hoy) & " (id,[timestamp],quistamp,Client,CodiArticle,PluUtilitzat,Viatge,Equip,QuantitatDemanada, QuantitatTornada,QuantitatServida,Hora,TipusComanda,Comentari,ComentariPer,Atribut) values (newid(), getdate(), 'SECRE', " & botiga & ", " & rsBocadillos("codi") & ", " & rsBocadillos("codi") & ", 'Auto', 'Auto', 0, 0, " & Int(faltan) & ", 13, 2, 'Reposicion', '', 0, '', '', '')"
        
                co = co & "</TR>"
            End If
                        
        End If
        
        rsBocadillos.MoveNext
    Wend
    
    If algun = 1 Then
        co = co & "</TABLE>"
                
        sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
        'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", Co, "", ""
        'sf_enviarMail "secrehit@hit.cat", "jordi@hit.cat", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", Co, "", ""
        'sf_enviarMail "secrehit@hit.cat", "jaTena@SilemaBcn.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", Co, "", ""
    End If
    
    Exit Sub
    
emailInfo:
    sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "EL FORMATO DEL INFORME ES: <BR>Bocadillos <I>Tienda</I><BR>Por ejemplo: Bocadillos T--073", "", ""
    'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "EL FORMATO DEL INFORME ES: <BR>Bocadillos <I>Tienda</I><BR>Por ejemplo: Bocadillos T--073", "", ""
    
    Exit Sub
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "ERROR: SecreInformeBocadillos " & err.Description, "", ""
End Sub


Sub ForzarTraspasoCajas(botiga As String, fechaIni As Date, fechaFin As Date)
    Dim sql As String
    Dim rsCajas As rdoResultset, rsVenut As rdoResultset
    Dim rsAbre As rdoResultset, fAbre As Date, fCierra As Date
    Dim Z As Double
    
    InformaMiss "ForzarTraspasoCajas", True
    
On Error GoTo nor
    
    Set rsCajas = Db.OpenResultset("select * from [" & NomTaulaMovi(fechaIni) & "] where tipus_moviment='Z' and botiga=" & botiga & " and data between '" & Format(fechaIni, "dd/mm/yyyy") & "' and '" & Format(DateAdd("d", 1, fechaFin), "dd/mm/yyyy") & "' order by data")
    While Not rsCajas.EOF
        fCierra = rsCajas("data")
        Z = rsCajas("import")
        
        If Z > 0 Then
            Set rsAbre = Db.OpenResultset("select top 1 data as dataInici from [" & NomTaulaMovi(fechaIni) & "] where tipus_moviment='Wi' and data < '" & fCierra & "' and botiga=" & botiga & " order by data desc")
            If Not rsAbre.EOF Then
                fAbre = rsAbre("dataInici")
    
                Set rsVenut = Db.OpenResultset("select sum(import) import, min(num_tick) MinTick, max(num_tick) MaxTick from [" & NomTaulaVentas(fechaIni) & "] where botiga=" & botiga & " and data between '" & fAbre & "' and '" & fCierra & "'")
                If Not rsVenut.EOF Then
                    'If Abs(rsVenut("import") - Z) < 0.01 Then
                        ExecutaComandaSql "insert into feinesafer values (newid(), 'SincroMURANOCaixaOnLine', 0, '[" & botiga & "]', '[" & fAbre & "]', '[" & fCierra & "]', '[" & rsVenut("MinTick") & "," & rsVenut("MaxTick") & "]', '[" & Z & "]', getdate())"
                    'End If
                End If
            End If
        End If
        rsCajas.MoveNext
    Wend
    
    Exit Sub
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "[" & Format(Now(), "dd/mm/yy hh:nn") & "]", "ERROR: ForzarTraspasoCajas " & err.Description, "", ""
End Sub


Sub SecreInformeSupervisora(fecha As Date, depId As String, depNom As String, depEMail As String, franquicias As Boolean, conCopia As Boolean)
    Dim rsBotigues As rdoResultset, botiguesList As String
    Dim codiBot As Double, nomBot As String
    Dim co As String, CoResum As String, CoTotal As String, CoMargeBrut As String, CoMarge As String
    Dim lunes As Date, lunesAnt As Date, lunesAnyoAnt As Date, diaAnyoAnt As Date, diaComparacion As Date, f As Date, fechaActual As Date
    Dim emailList() As String, e As Integer
    Dim color As String, colorMrg As String, sql As String
    Dim notasIncidencia As String
    
    'VENDES/PREVISIONS
    Dim rsVendes As rdoResultset, vendes As Boolean
    Dim vMatiB As Double, vTardaB As Double, vTotalB As Double, vAcumB As Double 'VENDES BOTIGA
    Dim pMatiB As Double, pTardaB As Double, pTotalB As Double, pAcumB As Double 'PREVISIONS BOTIGA
    Dim vendesSinIva As Double, vendesSinIVASINDiada As Double, vendesArr(3) As Double, previsionsArr(2) As Double
    Dim objMrgBruto As Double
    
    'CLIENTS
    Dim clients As Double, clientsAnt As Double, clientsAc As Double, clientsAcAnt As Double, clientsArr(2) As Integer
    
    'TIQUET MIG
    Dim tiquetMig As Double 'OBJETIVO
    Dim TMig As Double, TMigAc As Double
    
    'COMPRES
    Dim pctCompresObj As Double
    Dim compres As Double, pctCompres As Double
    Dim compresSINDiada As Double
    
    'DEVOLUCIONS
    Dim devDia As Double, devAc As Double, diffDias As Integer
    Dim rsDevDia As rdoResultset, rsDevAc As rdoResultset
    
    'HORES
    Dim entra As Date, sale As Date
    Dim horas As Double, horasAc As Double, minutos As Double, totalHoras As Double
    Dim Entrat As Boolean
    Dim rsHores As rdoResultset
    Dim horasAcumulado As Double
    Dim horasPlan As Double
    Dim horasReales As Double
    Dim horasPanadero As Double
    Dim horasResto As Double
    Dim gastosPersonal As Double
    Dim objEurosHora As Double
    
    'TOTALS BOTIGA
    Dim vMatiB_TB As Double, vTardaB_TB As Double, vTotalB_TB As Double, pTotalB_TB As Double, vAcumB_TB As Double, pAcumB_TB As Double
    Dim clients_TB As Double, clientsAnt_TB As Double, clientsAc_TB As Double, clientsAcAnt_TB As Double
    Dim compres_TB As Double, compresSINDiada_TB As Double, vendesSinIVA_TB As Double, vendesSinIVASinDiada_TB As Double
    Dim devDia_TB As Double, devAc_TB As Double
    Dim horas_TB As Double, horasPlan_TB As Double, totalHoras_TB As Double, horasAc_TB As Double, horasPlanAc_TB As Double
    Dim pMatiB_TB As Double, pTardaB_TB As Double
    Dim gastosPersonal_TB As Double
    
    'TOTALS RESUM
    Dim vMatiB_TR As Double, vTardaB_TR As Double, vTotalB_TR As Double, pTotalB_TR As Double, vAcumB_TR As Double, pAcumB_TR As Double
    Dim clients_TR As Double, clientsAnt_TR As Double, clientsAc_TR As Double, clientsAcAnt_TR As Double
    Dim compres_TR As Double, compresSINDiada_TR As Double, vendesSinIVA_TR As Double, vendesSinIVASinDiada_TR As Double, vendesSinIvaAc As Double, vendesSinIvaSinDiadaAc As Double, compresSinDiadaAc As Double
    Dim devDia_TR As Double, devAc_TR As Double
    Dim horas_TR As Double, horasPlan_TR As Double, totalHoras_TR As Double, horasAc_TR As Double, horasPlanAc_TR As Double, horasRealesAc As Double
    Dim pMatiB_TR As Double, pTardaB_TR As Double
    Dim gastosPersonal_TR As Double, gastosPersonalAc As Double

    InformaMiss "SecreInformeSupervisora", True

On Error GoTo nor

    lunes = fecha
    While DatePart("w", lunes) <> vbMonday
        lunes = DateAdd("d", -1, lunes)
    Wend

    co = ""
    CoMargeBrut = ""
    CoMarge = ""
    CoResum = ""
    CoTotal = ""
    If franquicias Then CoTotal = CoTotal & "<H2>FRANQUICIAS</H2>"
    CoTotal = CoTotal & "<h2>" & depNom & "</h2>"
    
    CoResum = CoResum & "<Table cellpadding='0' cellspacing='3' border='1'>"
    CoResum = CoResum & "<Tr bgColor='#DADADA'><Td align='left' colspan='15'><b>" & Format(fecha, "dd/mm/yyyy") & "</b></Td></Tr>"
    
    CoResum = CoResum & "<Tr bgColor='#FAFAFA'><Td rowspan='2'><b>Botiga</b></Td><Td align='center' colspan='4'><b>Vendes</b></Td><Td align='center' colspan='2'><b>Diferència clients</b></Td><Td align='center' colspan='2'><b>Tiquet mig</b></Td><Td align='center' colspan='2'><b>Compres</b></Td><Td colspan='2' align='center'><b>Devolucions</b></Td>"
    CoResum = CoResum & "<Td colspan='2' align='center'><b>&euro;/Hora</b></Td>"
    CoResum = CoResum & "</Tr>"
    
    CoResum = CoResum & "<Tr bgColor='#FAFAFA'>"
    CoResum = CoResum & "<Td align='center'><b>Mati</b></td><td align='center'><b>Tarda</b></Td><td align='center'><b>Dif. dia</b></Td><td align='center'><b>Dif. acumulat</b></td>"
    CoResum = CoResum & "<Td align='center'><b>Sem. " & DatePart("ww", DateAdd("d", -7, fecha)) & "<br>Sem. " & DatePart("ww", fecha) & "</b></td><td align='center'><b>Acumulat</b></Td></td>"
    CoResum = CoResum & "<Td align='center'><b>Dia</b></td><td align='center'><b>Acumulat</b></Td>"
    CoResum = CoResum & "<Td align='center'><b>% Diari</b></td><td align='center'><b>% Acumulat</b></td>"
    CoResum = CoResum & "<Td align='center'><b>Dia</b></td><td align='center'><b>Acumulat</b></Td>"
    'CoResum = CoResum & "<Td align='center'><b>Dia</b></td><td align='center'><b>%</b></Td><Td align='center'><b>Dif. Acumulat</b></td><td align='center'><b>%</b></Td>"
    CoResum = CoResum & "<Td align='center'><b>Diari</b></td><td align='center'><b>Acumulat</b></td>"
    CoResum = CoResum & "</Tr>"
    
    CoMargeBrut = CoMargeBrut & "<Table cellpadding='0' cellspacing='3' border='1'>"
    CoMargeBrut = CoMargeBrut & "<Tr bgColor='#DADADA'><Td align='left' colspan='11'><b>" & Format(fecha, "dd/mm/yyyy") & "</b></Td></Tr>"
    
    CoMargeBrut = CoMargeBrut & "<Tr bgColor='#FAFAFA'><Td rowspan='2'><b>Botiga</b></Td><Td align='center' colspan='2'><b>Vendes</b></Td><Td align='center' colspan='2'><b>Diferència clients</b></Td><Td align='center' colspan='2'><b>Compres</b></Td><Td colspan='2' align='center'><b>&euro;/Hora</b></Td><Td colspan='2' align='center'><b>Marge</b></Td></Tr>"
    
    
    CoMargeBrut = CoMargeBrut & "<Tr bgColor='#FAFAFA'>"
    CoMargeBrut = CoMargeBrut & "<td align='center'><b>Dif. dia</b></Td><td align='center'><b>Dif. acumulat</b></td>"
    CoMargeBrut = CoMargeBrut & "<Td align='center'><b>Sem. " & DatePart("ww", DateAdd("d", -7, fecha)) & "<br>Sem. " & DatePart("ww", fecha) & "</b></td><td align='center'><b>Acumulat</b></Td></td>"
    CoMargeBrut = CoMargeBrut & "<Td align='center'><b>% Diari</b></td><td align='center'><b>% Acumulat</b></td>"
    CoMargeBrut = CoMargeBrut & "<Td align='center'><b>Diari</b></td><td align='center'><b>Acumulat</b></td>"
    CoMargeBrut = CoMargeBrut & "<Td align='center'><b>Diari</b></td><td align='center'><b>Acumulat</b></td>"
    CoMargeBrut = CoMargeBrut & "</Tr>"
    
    sql = "select codi, upper(nom) nom "
    sql = sql & "from clients "
    sql = sql & "where codi in (select codi from constantsclient where variable='SupervisoraCodi' and valor='" & depId & "') "
    If franquicias Then
        sql = sql & "and codi in (select codi from constantsclient where variable='Franquicia' and valor='Franquicia') "
    Else 'NO FRANQUICIAS
        sql = sql & "and codi not in (select codi from constantsclient where variable='Franquicia' and valor='Franquicia') "
    End If
    sql = sql & "order by nom"
    
    Set rsBotigues = Db.OpenResultset(sql)
    botiguesList = ""
    While Not rsBotigues.EOF
        codiBot = rsBotigues("codi")
        nomBot = rsBotigues("nom")
        
        notasIncidencia = getNotasIncidencia(codiBot, lunes, fecha)

        compres_TB = 0: vendesSinIVA_TB = 0: compresSINDiada_TB = 0: vendesSinIVASinDiada_TB = 0: gastosPersonal_TB = 0
        devDia_TB = 0: devAc_TB = 0
        horas_TB = 0: totalHoras_TB = 0: horasAc_TB = 0
        pMatiB_TB = 0: pTardaB_TB = 0
    
        co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
        
        CoResum = CoResum & "<Tr>"
        CoMargeBrut = CoMargeBrut & "<Tr>"
        If notasIncidencia <> "" Then
            CoResum = CoResum & "<td bgcolor='#FF0000'><b>" & nomBot & "</b></td>"
            CoMargeBrut = CoMargeBrut & "<td bgcolor='#FF0000'><b>" & nomBot & "</b></td>"
        Else
            CoResum = CoResum & "<td><b>" & nomBot & "</b></td>"
            CoMargeBrut = CoMargeBrut & "<td><b>" & nomBot & "</b></td>"
        End If
        
'If nomBot = "T--126" Then
'nomBot = nomBot
'End If
        pctCompresObj = getObjetivoCompras(codiBot)
        objEurosHora = getObjetivoEurosHora(codiBot)
        objMrgBruto = getObjetivoMargenBruto(codiBot)
    
        co = co & "<Tr bgColor='#DADADA'>"
        co = co & "<Td align='left' colspan='15'><b>" & nomBot & "</b></Td></Tr>"
        co = co & "<Tr bgColor='#FAFAFA'><Td rowspan='2'><b>Dia</b></Td><Td align='center' colspan='4'><b>Vendes</b></Td><Td align='center' colspan='2'><b>Diferència clients</b></Td><Td align='center' colspan='2'><b>Tiquet mig</b></Td><Td align='center' colspan='2'><b>Compres</b></Td><Td colspan='2' align='center'><b>Devolucions</b></Td>"
        co = co & "<Td colspan='2' align='center'><b>&euro;/Hora</b><br>(Objetivo " & objEurosHora & ")</Td>"
        co = co & "</Tr>"
        
        co = co & "<Tr bgColor='#FAFAFA'>"
        Set rsVendes = Db.OpenResultset("select * from [" & NomTaulaVentas(DateAdd("yyyy", -1, fecha)) & "] where botiga = " & codiBot)
        If Not rsVendes.EOF Then
            vendes = True
            lunesAnt = lunesAnyoAnt

            co = co & "<Td align='center'><b>Mati</b></td><td align='center'><b>Tarda</b></Td><td align='center'><b>Dif. dia</b></Td><td align='center'><b>Dif. acumulat</b></td>"
            'Co = Co & "<Td align='center'><b>" & Year(fecha) - 1 & "<br>" & Year(fecha) & "</b></td><td align='center'><b>Acumulat</b></Td></td>"
        Else
            vendes = False
            lunesAnt = DateAdd("d", -7, lunes)
            
            co = co & "<Td align='center'><b>Mati</b></td><td align='center'><b>Tarda</b></Td><td align='center'><b>Dif. dia</b></Td><td align='center'><b>Dif. acumulat</b></td>"
        End If
        co = co & "<Td align='center'><b>" & Right("00" & Day(DateAdd("d", -7, fecha)), 2) & "/" & Right("00" & Month(DateAdd("d", -7, fecha)), 2) & "/" & Year(DateAdd("d", -7, fecha)) & "<br>" & Right("00" & Day(fecha), 2) & "/" & Right("00" & Month(fecha), 2) & "/" & Year(fecha) & "</b></td><td align='center'><b>Acumulat</b></Td></td>"
    
        co = co & "<Td align='center'><b>Dia</b></td><td align='center'><b>Acumulat</b></Td>"
        co = co & "<Td align='center'><b>% Dia (" & pctCompresObj & "%)</b></td><td align='center'><b>% Acumulat (" & pctCompresObj & "%)</b></td>"
        co = co & "<Td align='center'><b>Dia</b></td><td align='center'><b>Acumulat</b></Td>"
        'Co = Co & "<Td align='center'><b>Dia</b></td><td align='center'><b>%</b></Td><Td align='center'><b>Dif. Acumulat</b></td><td align='center'><b>%</b></Td>"
        co = co & "<Td align='center'><b>Dia</b></td><td align='center'><b>Acumulat</b></Td>"
        co = co & "</Tr>"
        
        'Acumulados y totales VENTAS/PREVISIONES
        vMatiB_TB = 0: vTardaB_TB = 0: vTotalB_TB = 0: pTotalB_TB = 0: vAcumB_TB = 0: pAcumB_TB = 0 'TOTALES POR TIENDA
        vAcumB = 0 'Ventas acumulado por tienda
        pAcumB = 0 'Previsiones acumulado por tienda
        'vendesSinIVA_TB = 0: vendesSinIVASinDiada_TB = 0
        
        'Acumulados y totales CLIENTES
        clientsAc = 0
        clientsAcAnt = 0
        clients_TB = 0: clientsAnt_TB = 0: clientsAc_TB = 0: clientsAcAnt_TB = 0 'TOTALES POR TIENDA
        
        'Total compras
        compres_TB = 0
        
        'Devoluciones
        devAc = 0
        devDia_TB = 0: devAc_TB = 0
        
        'Horas
        horasPlan_TB = 0: horas_TB = 0: horasPlanAc_TB = 0: horasAc_TB = 0
        horasAcumulado = 0
        
        CoMarge = ""
        
        For fechaActual = lunes To fecha
            diaAnyoAnt = DateAdd("yyyy", -1, fechaActual)
            While DatePart("w", fechaActual) <> DatePart("w", diaAnyoAnt)
                diaAnyoAnt = DateAdd("d", 1, diaAnyoAnt)
            Wend
            
            lunesAnyoAnt = diaAnyoAnt
            While DatePart("w", lunesAnyoAnt) <> vbMonday
                lunesAnyoAnt = DateAdd("d", -1, lunesAnyoAnt)
            Wend
        
            If vendes Then
                diaComparacion = DateAdd("yyyy", -1, fechaActual)
                While DatePart("w", fechaActual) <> DatePart("w", diaComparacion)
                    diaComparacion = DateAdd("d", 1, diaComparacion)
                Wend
            Else
                diaComparacion = DateAdd("d", -7, fechaActual)
            End If
            
            co = co & "<Tr>"
            co = co & "<td><b>" & Format(fechaActual, "dd/mm/yyyy") & "</b></td>"
            
'VENDES/PREVISIONS ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            calculaVendesClients codiBot, fechaActual, vendesArr, clientsArr
            calculaPrevisions codiBot, fechaActual, previsionsArr
            
            vMatiB = vendesArr(0)
            vTardaB = vendesArr(1)
            vTotalB = vendesArr(0) + vendesArr(1)
            
            pMatiB = previsionsArr(0)
            pTardaB = previsionsArr(1)
            pTotalB = previsionsArr(0) + previsionsArr(1)
        
            vAcumB = vAcumB + vTotalB
            pAcumB = pAcumB + pTotalB
            
            vendesSinIvaAc = vendesSinIvaAc + vendesArr(2) 'Acumulado ventas sin IVA de todos los días y todas las tiendas
                        
            'MATI
            color = "#A6FFA6"
            If vMatiB < pMatiB Then color = "#FFA8A8"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(vMatiB, 2) & " &euro;</Td>"
            If fecha = fechaActual Then
                CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(vMatiB, 2) & " &euro;</Td>"
                vMatiB_TR = vMatiB_TR + vMatiB
                pMatiB_TR = pMatiB_TR + pMatiB
            End If
            vMatiB_TB = vMatiB_TB + vMatiB
            pMatiB_TB = pMatiB_TB + pMatiB
            
            'TARDA
            color = "#A6FFA6"
            If vTardaB < pTardaB Then color = "#FFA8A8"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(vTardaB, 2) & " &euro;</Td>"
            If fecha = fechaActual Then
                CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(vTardaB, 2) & " &euro;</Td>"
                vTardaB_TR = vTardaB_TR + vTardaB
                pTardaB_TR = pTardaB_TR + pTardaB
            End If
            vTardaB_TB = vTardaB_TB + vTardaB
            pTardaB_TB = pTardaB_TB + pTardaB
            
            'DIF TOTAL VENDES vs PREVISIONS
            color = "#A6FFA6"
            If vTotalB < pTotalB Then color = "#FFA8A8"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs(vTotalB - pTotalB), 2) & " &euro;</Td>"
            If fecha = fechaActual Then
                CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs(vTotalB - pTotalB), 2) & " &euro;</Td>"
                CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs(vTotalB - pTotalB), 2) & " &euro;</Td>"
                vTotalB_TR = vTotalB_TR + vTotalB
                pTotalB_TR = pTotalB_TR + pTotalB
            End If
            vTotalB_TB = vTotalB_TB + vTotalB
            pTotalB_TB = pTotalB_TB + pTotalB
            
            'ACUMULAT DIF TOTAL VENDES vs PREVISIONS
            color = "#A6FFA6"
            If vAcumB < pAcumB Then color = "#FFA8A8"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs(vAcumB - pAcumB), 2) & " &euro;</Td>"
            If fecha = fechaActual Then
                CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs(vAcumB - pAcumB), 2) & " &euro;</Td>"
                CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs(vAcumB - pAcumB), 2) & " &euro;</Td>"
                vAcumB_TR = vAcumB_TR + vAcumB
                pAcumB_TR = pAcumB_TR + pAcumB
            End If
            vAcumB_TB = vAcumB 'El Acumulado total por tienda es el de la última fecha
            pAcumB_TB = pAcumB
            
'~VENDES/PREVISIONS ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'CLIENTS -------------------------------------------------------------------------------------------------------------
            clients = clientsArr(0) + clientsArr(1) 'LO TENGO CALCULADO DE LAS VENTAS (calculaVendesClients)
                        
            'Clients año anterior
            'clientsAnt = calculaNumeroClientes(codiBot, diaAnyoAnt)
            'If clientsAnt = 0 Then clientsAnt = calculaNumeroClientes(codiBot, DateAdd("d", -7, fechaActual))
            'SIEMPRE SEMANA ANTERIOR
            clientsAnt = calculaNumeroClientes(codiBot, DateAdd("d", -7, fechaActual))
            
            color = "#A6FFA6"
            If clients < clientsAnt Then color = "#FFA8A8"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs(clients - clientsAnt), 0) & "</Td>"
            If fecha = fechaActual Then
                CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs(clients - clientsAnt), 0) & "</Td>"
                CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs(clients - clientsAnt), 0) & "</Td>"
                clients_TR = clients_TR + clients
                clientsAnt_TR = clientsAnt_TR + clientsAnt
            End If
            clients_TB = clients_TB + clients
            clientsAnt_TB = clientsAnt_TB + clientsAnt
            
            'Acumulat setmanal
            clientsAc = clientsAc + clients
            clientsAcAnt = clientsAcAnt + clientsAnt
            
            color = "#A6FFA6"
            If (clientsAc - clientsAcAnt) < 0 Then color = "#FFA8A8"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs((clientsAc - clientsAcAnt)), 0) & "</Td>"
            If fecha = fechaActual Then
                CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs((clientsAc - clientsAcAnt)), 0) & "</Td>"
                CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(Abs((clientsAc - clientsAcAnt)), 0) & "</Td>"
                clientsAc_TR = clientsAc_TR + clientsAc
                clientsAcAnt_TR = clientsAcAnt_TR + clientsAcAnt
            End If
            clientsAc_TB = clientsAc 'El Acumulado total por tienda es el de la útima fecha
            clientsAcAnt_TB = clientsAcAnt
'~CLIENTS -------------------------------------------------------------------------------------------------------------
    
'TIQUET MIG -------------------------------------------------------------------------------------------------------------
            tiquetMig = getObjetivoTiquetMig(codiBot, fechaActual)
            
            TMig = 0
            If clients > 0 Then TMig = vTotalB / clients
            color = "#A6FFA6"
            If TMig < tiquetMig Then color = "#FFA8A8"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(TMig, 2) & "</Td>"
            If fecha = fechaActual Then CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(TMig, 2) & "</Td>"
    
            TMigAc = 0
            If clientsAc > 0 Then TMigAc = vAcumB / clientsAc
            If TMigAc < tiquetMig Then color = "#FFA8A8"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(TMigAc, 2) & "</Td>"
            If fecha = fechaActual Then CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(TMigAc, 2) & "</Td>"
        
'~TIQUET MIG -------------------------------------------------------------------------------------------------------------

'COMPRES -------------------------------------------------------------------------------------------------------------

            compres = calculaCompras(codiBot, fechaActual)
            compresSINDiada = calculaCompras(codiBot, fechaActual, "Diada")
            vendesSinIva = vendesArr(2)
            vendesSinIVASINDiada = calculaVendesSINDiada(codiBot, fechaActual)
            horasPanadero = calculaHorasPanadero(codiBot, fechaActual)
            horasResto = calculaHorasResto(codiBot, fechaActual)
            gastosPersonal = (horasPanadero * 13.9) + (horasResto * 11.8)
            
            vendesSinIvaSinDiadaAc = vendesSinIvaSinDiadaAc + vendesSinIVASINDiada
            gastosPersonalAc = gastosPersonalAc + gastosPersonal
            compresSinDiadaAc = compresSinDiadaAc + compresSINDiada
            
            color = "#A6FFA6"
            'Co = Co & "<Td align='right'>" & FormatNumber(compres, 2) & " &euro;</Td>"
            If vendesSinIVASINDiada > 0 Then
                If (compresSINDiada / vendesSinIVASINDiada) * 100 > pctCompresObj Then color = "#FFA8A8"
                co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((compresSINDiada / vendesSinIVASINDiada) * 100, 2) & " %</Td>"
            Else
                If compresSINDiada = 0 Then
                    co = co & "<Td align='right'>0.00 %</Td>"
                Else
                    co = co & "<Td align='right'>100.00 %</Td>"
                End If
            End If
            
            color = "#A6FFA6"
            If fecha = fechaActual Then
                'CoResum = CoResum & "<Td align='right'>" & FormatNumber(compres, 2) & " &euro;</Td>"
                
                If vendesSinIVASINDiada > 0 Then
                    If (compresSINDiada / vendesSinIVASINDiada) * 100 > pctCompresObj Then color = "#FFA8A8"
                    CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((compresSINDiada / vendesSinIVASINDiada) * 100, 2) & " %</Td>"
                    CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((compresSINDiada / vendesSinIVASINDiada) * 100, 2) & " %</Td>"
                    
                    color = "#A6FFA6"
                    If ((vendesSinIVASINDiada - compresSINDiada - gastosPersonal) * 100 / vendesSinIVASINDiada) < objMrgBruto Then color = "#FFA8A8"
                    CoMarge = "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(((vendesSinIVASINDiada - compresSINDiada - gastosPersonal) * 100 / vendesSinIVASINDiada), 2) & " %</Td>"
                Else
                    If compresSINDiada = 0 Then
                        CoResum = CoResum & "<Td align='right'>0.00 %</Td>"
                        CoMargeBrut = CoMargeBrut & "<Td align='right'>0.00 %</Td>"
                        CoMarge = "<Td align='right'>0.00 %</Td>"
                    Else
                        CoResum = CoResum & "<Td align='right'>100.00 %</Td>"
                        CoMargeBrut = CoMargeBrut & "<Td align='right'>100.00 %</Td>"
                        CoMarge = "<Td align='right'>100.00 %</Td>"
                    End If
                End If
            
                compres_TR = compres_TR + compres
                compresSINDiada_TR = compresSINDiada_TR + compresSINDiada
                vendesSinIVA_TR = vendesSinIVA_TR + vendesSinIva
                vendesSinIVASinDiada_TR = vendesSinIVASinDiada_TR + vendesSinIVASINDiada
                gastosPersonal_TR = gastosPersonal_TR + gastosPersonal
            End If

            vendesSinIVA_TB = vendesSinIVA_TB + vendesSinIva
            vendesSinIVASinDiada_TB = vendesSinIVASinDiada_TB + vendesSinIVASINDiada
            compres_TB = compres_TB + compres
            compresSINDiada_TB = compresSINDiada_TB + compresSINDiada
            gastosPersonal_TB = gastosPersonal_TB + gastosPersonal
            
            'vendesSinIvaAc = vendesSinIvaAc + vendesSinIva
            
            '% sin Diada ACUMULADO
            color = "#A6FFA6"
            If vendesSinIVASinDiada_TB > 0 Then
                If (compresSINDiada_TB / vendesSinIVASinDiada_TB) * 100 > pctCompresObj Then color = "#FFA8A8"
                co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((compresSINDiada_TB / vendesSinIVASinDiada_TB) * 100, 2) & " %</Td>"
                If fecha = fechaActual Then
                    CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((compresSINDiada_TB / vendesSinIVASinDiada_TB) * 100, 2) & " %</Td>"
                    CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((compresSINDiada_TB / vendesSinIVASinDiada_TB) * 100, 2) & " %</Td>"
                    color = "#A6FFA6"
                    If ((vendesSinIVASinDiada_TB - compresSINDiada_TB - gastosPersonal_TB) * 100 / vendesSinIVASinDiada_TB) < objMrgBruto Then color = "#FFA8A8"
                    CoMarge = CoMarge & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(((vendesSinIVASinDiada_TB - compresSINDiada_TB - gastosPersonal_TB) * 100 / vendesSinIVASinDiada_TB), 2) & " %</Td>"
                End If
            Else
                If compresSINDiada_TB = 0 Then
                    co = co & "<Td align='right'>0.00 %</Td>"
                    If fecha = fechaActual Then
                        CoResum = CoResum & "<Td align='right'>0.00 %</Td>"
                        CoMargeBrut = CoMargeBrut & "<Td align='right'>0.00 %</Td>"
                        CoMarge = CoMarge & "<Td align='right'>0.00 %</Td>"
                    End If
                Else
                    co = co & "<Td align='right'>100.00 %</Td>"
                    If fecha = fechaActual Then
                        CoResum = CoResum & "<Td align='right'>100.00 %</Td>"
                        CoMargeBrut = CoMargeBrut & "<Td align='right'>100.00 %</Td>"
                        CoMarge = CoMarge & "<Td align='right'>100.00 %</Td>"
                    End If
                End If
            End If

                        
'~COMPRES -------------------------------------------------------------------------------------------------------------

'DEVOLUCIONS -------------------------------------------------------------------------------------------------------------

            devDia = calculaDevoluciones(codiBot, fechaActual)
            
            color = "#A6FFA6"
            If devDia > 160 Or devDia < 80 Then
                color = "#FFA8A8"
            End If
        
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(devDia, 2) & " &euro;</Td>"
            If fecha = fechaActual Then
                CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(devDia, 2) & " &euro;</Td>"
                devDia_TR = devDia_TR + devDia
            End If
            devDia_TB = devDia_TB + devDia
            
            devAc = devAc + devDia
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(devAc, 2) & " &euro;</Td>"
            If fecha = fechaActual Then
                CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(devAc, 2) & " &euro;</Td>"
                devAc_TR = devAc_TR + devAc
            End If
            devAc_TB = devAc 'El Acumulado total por tienda es el de la útima fecha

'~DEVOLUCIONS -------------------------------------------------------------------------------------------------------------

'HORES --------------------------------------------------------------------------------------------------------------------

            horasPlan = getObjetivoHoras(codiBot, fechaActual) 'HORAS PROGRAMADAS EN EL CUADRANTE
            horasReales = calculaHorasReales(codiBot, fechaActual) 'HORAS QUE SE HA HECHO REALMENTE (SIN CONTAR APRENDIZ NI COORDINACIÓN)

            horasRealesAc = horasRealesAc + horasReales
            
            'color = "#A6FFA6"
            'If horasReales <> horasPlan Then
            '    color = "#FFA8A8"
            'End If
            
            'Co = Co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(horasReales - horasPlan, 2) & "</Td>"
            If horasReales > 0 Then
                colorMrg = "#A6FFA6"
                'If (vendesSinIva / horasReales) < 38 Or (vendesSinIva / horasReales) > 45 Then colorMrg = "#FFA8A8"
                If (vendesSinIva / horasReales) < objEurosHora Then colorMrg = "#FFA8A8"
                If (vendesSinIva / horasReales) >= objEurosHora Then colorMrg = "#A6FFA6"
                If (vendesSinIva / horasReales) > 50 Then colorMrg = "#FFFFCA"
                co = co & "<Td bgcolor='" & colorMrg & "' align='right'>" & FormatNumber(vendesSinIva / horasReales, 2) & " &euro;</Td>"
            Else
                co = co & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
            End If

            
            If fecha = fechaActual Then
                'CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(horasReales - horasPlan, 2) & "</Td>"
                If horasReales > 0 Then
                    colorMrg = "#A6FFA6"
                    'If (vendesSinIva / horasReales) < 38 Or (vendesSinIva / horasReales) > 45 Then colorMrg = "#FFA8A8"
                    If (vendesSinIva / horasReales) < objEurosHora Then colorMrg = "#FFA8A8"
                    If (vendesSinIva / horasReales) >= objEurosHora Then colorMrg = "#A6FFA6"
                    If (vendesSinIva / horasReales) > 50 Then colorMrg = "#FFFFCA"

                    CoResum = CoResum & "<Td bgcolor='" & colorMrg & "' align='right'>" & FormatNumber(vendesSinIva / horasReales, 2) & " &euro;</Td>"
                    CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & colorMrg & "' align='right'>" & FormatNumber(vendesSinIva / horasReales, 2) & " &euro;</Td>"
                Else
                    CoResum = CoResum & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
                    CoMargeBrut = CoMargeBrut & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
                End If
                horas_TR = horas_TR + horasReales
                horasPlan_TR = horasPlan_TR + horasPlan
            End If
            horas_TB = horas_TB + horasReales
            horasPlan_TB = horasPlan_TB + horasPlan
            
            'If vendesSinIVA > 0 Then
            '    Co = Co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((horasReales * 12 / vendesSinIVA) * 100, 2) & " %</Td>"
                'If fecha = fechaActual Then CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((horasReales * 12 / vendesSinIVA) * 100, 2) & " %</Td>"
            'Else
            '    Co = Co & "<Td bgcolor='" & color & "' align='right'>0.00 %</Td>"
                'If fecha = fechaActual Then CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>0.00 %</Td>"
            'End If
                                
            'ACUMULADO
            horasAc_TB = horasAc_TB + horasReales
            horasPlanAc_TB = horasPlanAc_TB + horasPlan
            
            horasAcumulado = horasAcumulado + (horasReales - horasPlan)
            
            'color = "#A6FFA6"
            'If horasAcumulado <> 0 Then
            '    color = "#FFA8A8"
            'End If
            
            'Co = Co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(horasAcumulado, 2) & "</Td>"
            
            
            If horasAc_TB > 0 Then
                colorMrg = "#A6FFA6"
                'If (vendesSinIVA_TB / horasAc_TB) < 38 Or (vendesSinIVA_TB / horasAc_TB) > 45 Then colorMrg = "#FFA8A8"
                If (vendesSinIVA_TB / horasAc_TB) < objEurosHora Then colorMrg = "#FFA8A8"
                If (vendesSinIVA_TB / horasAc_TB) >= objEurosHora Then colorMrg = "#A6FFA6"
                If (vendesSinIVA_TB / horasAc_TB) > 50 Then colorMrg = "#FFFFCA"
                
                co = co & "<Td bgcolor='" & colorMrg & "' align='right'>" & FormatNumber(vendesSinIVA_TB / horasAc_TB, 2) & " &euro;</Td>"
            Else
                co = co & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
            End If
            
            
            If fecha = fechaActual Then
                'CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(horasAcumulado, 2) & "</Td>"
                If horasAc_TB > 0 Then
                    colorMrg = "#A6FFA6"
                    'If (vendesSinIVA_TB / horasAc_TB) < 38 Or (vendesSinIVA_TB / horasAc_TB) > 45 Then colorMrg = "#FFA8A8"
                    If (vendesSinIVA_TB / horasAc_TB) < objEurosHora Then colorMrg = "#FFA8A8"
                    If (vendesSinIVA_TB / horasAc_TB) >= objEurosHora Then colorMrg = "#A6FFA6"
                    If (vendesSinIVA_TB / horasAc_TB) > 50 Then colorMrg = "#FFFFCA"
                                        
                    CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & colorMrg & "' align='right'>" & FormatNumber(vendesSinIVA_TB / horasAc_TB, 2) & " &euro;</Td>"
                    CoResum = CoResum & "<Td bgcolor='" & colorMrg & "' align='right'>" & FormatNumber(vendesSinIVA_TB / horasAc_TB, 2) & " &euro;</Td>"
                Else
                    CoMargeBrut = CoMargeBrut & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
                    CoResum = CoResum & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
                End If
                
                horasAc_TR = horasAc_TR + horasReales
                horasPlanAc_TR = horasPlanAc_TR + horasPlan

            End If
        
            'If vendesSinIVA_TB > 0 Then 'Acumulado
            '    Co = Co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((horasAc_TB * 12 / vendesSinIVA_TB) * 100, 2) & " %</Td>"
                'If fecha = fechaActual Then CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((horasAc_TB * 12 / vendesSinIVA_TB) * 100, 2) & " %</Td>"
            'Else
            '    Co = Co & "<Td bgcolor='" & color & "' align='right'>0.00 %</Td>"
                'If fecha = fechaActual Then CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'>0.00 %</Td>"
            'End If
            
'~HORES ----------------------------------------------------------------------------------------------------------------------------------------------------
            CoMargeBrut = CoMargeBrut & CoMarge
            co = co & "</Tr>"
        Next
        
        CoResum = CoResum & "</Tr>"
        CoMargeBrut = CoMargeBrut & "</Tr>"
        
'TOTALES POR TIENDA -----------------------------------------------------------------------------------------------------------------------------------------
        co = co & "<TR><TD><B>TOTAL</B></TD>"
        
        'VENDES
        color = "#A6FFA6"
        If pMatiB_TB > vMatiB_TB Then color = "#FFA8A8"
        co = co & "<td bgcolor='" & color & "' align='right'><B>" & FormatNumber(vMatiB_TB, 2) & " &euro;</B></td>"
        color = "#A6FFA6"
        If pTardaB_TB > vTardaB_TB Then color = "#FFA8A8"
        co = co & "<td bgcolor='" & color & "' align='right'><B>" & FormatNumber(vTardaB_TB, 2) & " &euro;</B></td>"
        color = "#A6FFA6"
        If pTotalB_TB > vTotalB_TB Then color = "#FFA8A8"
        co = co & "<td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(vTotalB_TB - pTotalB_TB), 2) & " &euro;</B></td>"
        color = "#A6FFA6"
        If pAcumB_TB > vAcumB_TB Then color = "#FFA8A8"
        co = co & "<td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(vAcumB_TB - pAcumB_TB), 2) & " &euro;</B></td>"
        
        'CLIENTS
        color = "#A6FFA6"
        If clientsAnt_TB > clients_TB Then color = "#FFA8A8"
        co = co & "<Td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(clients_TB - clientsAnt_TB), 0) & "</B></Td>"
        color = "#A6FFA6"
        If clientsAcAnt_TB > clientsAc_TB Then color = "#FFA8A8"
        co = co & "<Td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(clientsAc_TB - clientsAcAnt_TB), 0) & "</B></Td>"
        
        'TIQUET MIG
        TMig = 0
        If clients_TB > 0 Then TMig = vTotalB_TB / clients_TB
        co = co & "<Td align='right'><B>" & FormatNumber(TMig, 2) & "</B></Td>"
        TMigAc = 0
        If clientsAc_TB > 0 Then TMigAc = vAcumB_TB / clientsAc_TB
        co = co & "<Td align='right'><B>" & FormatNumber(TMigAc, 2) & "</B></Td>"
        
        'COMPRES
        pctCompres = 0
        If vendesSinIVASinDiada_TB > 0 Then
           pctCompres = (compresSINDiada_TB / vendesSinIVASinDiada_TB) * 100
        Else
            If compresSINDiada_TB > 0 Then pctCompres = 100
        End If
        'Co = Co & "<Td align='right'><B>" & FormatNumber(compres_TB, 2) & " &euro;</B></Td>"
        If vendesSinIVASinDiada_TB > 0 Then
            'Está repetido porque el acumulado de la semana, de la tienda, es el mismo dato que el acumulado
            co = co & "<Td align='right'><B>" & FormatNumber((compresSINDiada_TB / vendesSinIVASinDiada_TB) * 100, 2) & " %</B></Td>"
            co = co & "<Td align='right'><B>" & FormatNumber((compresSINDiada_TB / vendesSinIVASinDiada_TB) * 100, 2) & " %</B></Td>"
        Else
            co = co & "<Td align='right'><B>0.00 %</B></Td>"
            co = co & "<Td align='right'><B>0.00 %</B></Td>"
        End If
        
        'DEVOLUCIONS
        co = co & "<Td align='right'><B>" & FormatNumber(devDia_TB, 2) & " &euro;</B></Td>"
        co = co & "<Td align='right'><B>" & FormatNumber(devAc_TB, 2) & " &euro;</B></Td>"
        
        'HORES
        'Co = Co & "<Td align='right'><B>" & FormatNumber(horas_TB - horasPlan_TB, 2) & "</B></Td>"
        'If vendesSinIVA_TB > 0 Then
        '    Co = Co & "<Td align='right'><B>" & FormatNumber(((horas_TB * 12) / vendesSinIVA_TB) * 100, 2) & " %</B></Td>" 'El 12 es coste por hora
        'Else
        '    If horas_TB = 0 Then
        '        Co = Co & "<Td align='right'><B>0.00 %</B></Td>"
        '    Else
        '        Co = Co & "<Td align='right'><B>100.00 %</B></Td>"
        '    End If
        'End If
        
        'Co = Co & "<Td align='right'><B>" & FormatNumber(horasAc_TB - horasPlanAc_TB, 2) & "</B></Td>"
        'If vendesSinIVA_TB > 0 Then
        '    Co = Co & "<Td align='right'><B>" & FormatNumber(((horasAc_TB * 12) / vendesSinIVA_TB) * 100, 2) & " %</B></Td>" 'El 12 es coste por hora
        'Else
        '    If horasAc_TB = 0 Then
        '        Co = Co & "<Td align='right'><B>0.00 %</B></Td>"
        '    Else
        '        Co = Co & "<Td align='right'><B>100.00 %</B></Td>"
        '    End If
        'End If
          
        If horas_TB > 0 Then
             colorMrg = "#A6FFA6"
             If (vendesSinIVA_TB / horas_TB) < 38 Or (vendesSinIVA_TB / horas_TB) > 45 Then colorMrg = "#FFA8A8"
             co = co & "<Td bgcolor='" & colorMrg & "' align='right'>" & FormatNumber(vendesSinIVA_TB / horas_TB, 2) & " &euro;</Td>"
        Else
             co = co & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
        End If
        
        If horasAc_TB > 0 Then
            colorMrg = "#A6FFA6"
            If (vendesSinIVA_TB / horasAc_TB) < 38 Or (vendesSinIVA_TB / horasAc_TB) > 45 Then colorMrg = "#FFA8A8"
            co = co & "<Td bgcolor='" & colorMrg & "' align='right'>" & FormatNumber(vendesSinIVA_TB / horasAc_TB, 2) & " &euro;</Td>"
        Else
            co = co & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
        End If
          
          
        co = co & "</TR>"
'~TOTALES POR TIENDA -----------------------------------------------------------------------------------------------------------------------------------------
        co = co & "</Table>"
        co = co & notasIncidencia
        co = co & "<br>"
        
        If botiguesList <> "" Then botiguesList = botiguesList & ","
        botiguesList = botiguesList & codiBot
        
        rsBotigues.MoveNext
    Wend
    rsBotigues.Close
    
'TOTALES MARGE BRUT -----------------------------------------------------------------------------------------------------------------------------------------
    CoMargeBrut = CoMargeBrut & "<Tr><td><b>TOTAL</b></td>"
    
    'VENDES
    color = "#A6FFA6"
    If pTotalB_TR > vTotalB_TR Then color = "#FFA8A8"
    CoMargeBrut = CoMargeBrut & "<td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(vTotalB_TR - pTotalB_TR), 2) & " &euro;</B></td>"
    color = "#A6FFA6"
    If pAcumB_TR > vAcumB_TR Then color = "#FFA8A8"
    CoMargeBrut = CoMargeBrut & "<td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(vAcumB_TR - pAcumB_TR), 2) & " &euro;</B></td>"

    'CLIENTS
    color = "#A6FFA6"
    If clientsAnt_TR > clients_TR Then color = "#FFA8A8"
    CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(clients_TR - clientsAnt_TR), 0) & "</B></Td>"
    color = "#A6FFA6"
    If clientsAcAnt_TR > clientsAc_TR Then color = "#FFA8A8"
    CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(clientsAc_TR - clientsAcAnt_TR), 0) & "</B></Td>"

    'COMPRES
    pctCompres = 0
    If vendesSinIVASinDiada_TR > 0 Then
       pctCompres = (compresSINDiada_TR / vendesSinIVASinDiada_TR) * 100
    Else
        If compresSINDiada_TR > 0 Then pctCompres = 100
    End If
    CoMargeBrut = CoMargeBrut & "<Td align='right'><B>" & FormatNumber(pctCompres, 2) & " %</B></Td>"
    
    pctCompres = 0
    If vendesSinIvaSinDiadaAc > 0 Then
       pctCompres = (compresSinDiadaAc / vendesSinIvaSinDiadaAc) * 100
    Else
        If compresSinDiadaAc > 0 Then pctCompres = 100
    End If
    CoMargeBrut = CoMargeBrut & "<Td align='right'><B>" & FormatNumber(pctCompres, 2) & " %</B></Td>"

    'HORES
    If horas_TR > 0 Then
         colorMrg = "#A6FFA6"
         If (vendesSinIVA_TR / horas_TR) < 38 Or (vendesSinIVA_TR / horas_TR) > 45 Then colorMrg = "#FFA8A8"
         CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & colorMrg & "' align='right'><B>" & FormatNumber(vendesSinIVA_TR / horas_TR, 2) & " &euro;</B></Td>"
    Else
         CoMargeBrut = CoMargeBrut & "<Td  bgcolor='#FFA8A8' align='right'><B>0 &euro;</B></Td>"
    End If
     
    If horasRealesAc > 0 Then
        colorMrg = "#A6FFA6"
        If (vendesSinIvaAc / horasRealesAc) < 38 Or (vendesSinIvaAc / horasRealesAc) > 45 Then colorMrg = "#FFA8A8"
        CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & colorMrg & "' align='right'><B>" & FormatNumber(vendesSinIvaAc / horasRealesAc, 2) & " &euro;</B></Td>"
    Else
        CoMargeBrut = CoMargeBrut & "<Td  bgcolor='#FFA8A8' align='right'><B>0 &euro;</B></Td>"
    End If

    'MARGE
    If vendesSinIVASinDiada_TR > 0 Then
        CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & color & "' align='right'><B>" & FormatNumber(((vendesSinIVASinDiada_TR - compresSINDiada_TR - gastosPersonal_TR) * 100 / vendesSinIVASinDiada_TR), 2) & " %</B></Td>"
    Else
        If compresSINDiada_TR = 0 Then
            CoMargeBrut = CoMargeBrut & "<Td align='right'><B>0.00 %</B></Td>"
        Else
            CoMargeBrut = CoMargeBrut & "<Td align='right'><B>100.00 %</B></Td>"
        End If
    End If
    
    If vendesSinIvaSinDiadaAc > 0 Then
        CoMargeBrut = CoMargeBrut & "<Td bgcolor='" & color & "' align='right'><B>" & FormatNumber(((vendesSinIvaSinDiadaAc - compresSinDiadaAc - gastosPersonalAc) * 100 / vendesSinIvaSinDiadaAc), 2) & " %</B></Td>"
    Else
        If compresSinDiadaAc = 0 Then
            CoMargeBrut = CoMargeBrut & "<Td align='right'><B>0.00 %</B></Td>"
        Else
            CoMargeBrut = CoMargeBrut & "<Td align='right'><B>100.00 %</B></Td>"
        End If
    End If
                
    CoMargeBrut = CoMargeBrut & "</Tr>"
    CoMargeBrut = CoMargeBrut & "</Table>"
    CoMargeBrut = CoMargeBrut & "<br>"
'~TOTALES MARGE BRUT -----------------------------------------------------------------------------------------------------------------------------------------
    
'TOTALES RESUMEN -----------------------------------------------------------------------------------------------------------------------------------------
    CoResum = CoResum & "<Tr><td><b>TOTAL</b></td>"
    'VENDES
    color = "#A6FFA6"
    If pMatiB_TR > vMatiB_TR Then color = "#FFA8A8"
    CoResum = CoResum & "<td bgcolor='" & color & "' align='right'><B>" & FormatNumber(vMatiB_TR, 2) & " &euro;</B></td>"
    color = "#A6FFA6"
    If pTardaB_TR > vTardaB_TR Then color = "#FFA8A8"
    CoResum = CoResum & "<td bgcolor='" & color & "' align='right'><B>" & FormatNumber(vTardaB_TR, 2) & " &euro;</B></td>"
    color = "#A6FFA6"
    If pTotalB_TR > vTotalB_TR Then color = "#FFA8A8"
    CoResum = CoResum & "<td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(vTotalB_TR - pTotalB_TR), 2) & " &euro;</B></td>"
    color = "#A6FFA6"
    If pAcumB_TR > vAcumB_TR Then color = "#FFA8A8"
    CoResum = CoResum & "<td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(vAcumB_TR - pAcumB_TR), 2) & " &euro;</B></td>"
    
    'CLIENTS
    color = "#A6FFA6"
    If clientsAnt_TR > clients_TR Then color = "#FFA8A8"
    CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(clients_TR - clientsAnt_TR), 0) & "</B></Td>"
    color = "#A6FFA6"
    If clientsAcAnt_TR > clientsAc_TR Then color = "#FFA8A8"
    CoResum = CoResum & "<Td bgcolor='" & color & "' align='right'><B>" & FormatNumber(Abs(clientsAc_TR - clientsAcAnt_TR), 0) & "</B></Td>"
    
    'TIQUET MIG
    TMig = 0
    If clients_TR > 0 Then TMig = vTotalB_TR / clients_TR
    CoResum = CoResum & "<Td align='right'><B>" & FormatNumber(TMig, 2) & "</B></Td>"
    TMigAc = 0
    If clientsAc_TR > 0 Then TMigAc = vAcumB_TR / clientsAc_TR
    CoResum = CoResum & "<Td align='right'><B>" & FormatNumber(TMigAc, 2) & "</B></Td>"
    
    'COMPRES
    pctCompres = 0
    If vendesSinIVASinDiada_TR > 0 Then
       pctCompres = (compresSINDiada_TR / vendesSinIVASinDiada_TR) * 100
    Else
        If compresSINDiada_TR > 0 Then pctCompres = 100
    End If
    CoResum = CoResum & "<Td align='right'><B>" & FormatNumber(pctCompres, 2) & " %</B></Td>"
    
    pctCompres = 0
    If vendesSinIvaSinDiadaAc > 0 Then
       pctCompres = (compresSinDiadaAc / vendesSinIvaSinDiadaAc) * 100
    Else
        If compresSinDiadaAc > 0 Then pctCompres = 100
    End If
    CoResum = CoResum & "<Td align='right'><B>" & FormatNumber(pctCompres, 2) & " %</B></Td>"
    
    'DEVOLUCIONS
    CoResum = CoResum & "<Td align='right'><B>" & FormatNumber(devDia_TR, 2) & " &euro;</B></Td>"
    CoResum = CoResum & "<Td align='right'><B>" & FormatNumber(devAc_TR, 2) & " &euro;</B></Td>"
    
    'HORES
    'color = "#A6FFA6"
    'If horas_TR <> horasPlan_TR Then color = "#FFA8A8"
    'CoResum = CoResum & "<Td align='right' bgcolor='" & color & "' ><B>" & FormatNumber(horas_TR - horasPlan_TR, 2) & "</B></Td>"
    
    'If vendesSinIVA_TR > 0 Then
    '    CoResum = CoResum & "<Td align='right'><B>" & FormatNumber((horas_TR * 12 / vendesSinIVA_TR) * 100, 2) & " %</B></Td>" '12=coste por hora
    'Else
    '    If horas_TR = "0" Then
    '        CoResum = CoResum & "<Td align='right'><B>0.00 %</B></Td>"
    '    Else
    '        CoResum = CoResum & "<Td align='right'><B>100.00 %</B></Td>"
    '    End If
    'End If
    
    'color = "#A6FFA6"
    'If horasAc_TR <> horasPlanAc_TR Then color = "#FFA8A8"
    'CoResum = CoResum & "<Td align='right' bgcolor='" & color & "' ><B>" & FormatNumber(horasAc_TR - horasPlanAc_TR, 2) & "</B></Td>"
    
    'If vendesSinIVA_TR > 0 Then
    '    CoResum = CoResum & "<Td align='right'><B>" & FormatNumber((horasAc_TR * 12 / vendesSinIVA_TR) * 100, 2) & " %</B></Td>"
    'Else
    '    If horasAc_TR = 0 Then
    '        CoResum = CoResum & "<Td align='right'><B>0.00 %</B></Td>"
    '    Else
    '        CoResum = CoResum & "<Td align='right'><B>100.00 %</B></Td>"
    '    End If
    'End If
    
    If horas_TR > 0 Then
         colorMrg = "#A6FFA6"
         If (vendesSinIVA_TR / horas_TR) < 38 Or (vendesSinIVA_TR / horas_TR) > 45 Then colorMrg = "#FFA8A8"
         CoResum = CoResum & "<Td bgcolor='" & colorMrg & "' align='right'><B>" & FormatNumber(vendesSinIVA_TR / horas_TR, 2) & " &euro;</B></Td>"
    Else
         CoResum = CoResum & "<Td  bgcolor='#FFA8A8' align='right'><B>0 &euro;</B></Td>"
    End If
     
    If horasRealesAc > 0 Then
        colorMrg = "#A6FFA6"
        If (vendesSinIvaAc / horasRealesAc) < 38 Or (vendesSinIvaAc / horasRealesAc) > 45 Then colorMrg = "#FFA8A8"
        CoResum = CoResum & "<Td bgcolor='" & colorMrg & "' align='right'><B>" & FormatNumber(vendesSinIvaAc / horasRealesAc, 2) & " &euro;</B></Td>"
    Else
        CoResum = CoResum & "<Td  bgcolor='#FFA8A8' align='right'><B>0 &euro;</B></Td>"
    End If
    
    CoResum = CoResum & "</Tr>"
    CoResum = CoResum & "</Table>"
    CoResum = CoResum & "<br>"
    
'~TOTALES RESUMEN -----------------------------------------------------------------------------------------------------------------------------------------
    If botiguesList <> "" Then
        CoTotal = CoTotal & "<h3>Marge brut</h3>" & CoMargeBrut & "<h3>Diferència hores (Pactat/Real)</h3>" & cuadrantePlanificacionTurnos3(fecha, botiguesList) & "<h3>Resum</h3>" & CoResum & "<h3>Detall per botiga</h3>" & co
    
        emailList = Split(depEMail, ";")
        For e = 0 To UBound(emailList)
            sf_enviarMail "secrehit@hit.cat", Trim(emailList(e)), "Informe ventas supervisora " & depNom & " [" & Format(Now(), "dd/mm/yy hh:nn") & "]", CoTotal, "", ""
        Next
    End If
    
    If conCopia Then sf_enviarMail "secrehit@hit.cat", "atena@silemabcn.com", "Informe ventas supervisora " & depNom & " [" & Format(Now(), "dd/mm/yy hh:nn") & "]", CoTotal, "", ""
    'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "Informe ventas supervisora " & depNom & " [" & Format(Now(), "dd/mm/yy hh:nn") & "]", CoTotal, "", ""
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreInformeSupervisora " & depNom & " [" & Format(Now(), "dd/mm/yy hh:nn") & "]", sql & CoTotal & err.Description, "", ""
End Sub




Sub SecreInformeVentas(subj As String, emailDe As String, empresa As String)
    Dim Semana As Integer, semanaAux As Integer, lunes As Date, domingo As Date, fecha As Date, rs As rdoResultset, sql As String, iD As String, rsA As ADODB.Recordset
    Dim D_B(13) As String, co As String, supervisora As String, totalDia As Double, total7 As Double, strHoras As String
    Dim Totales(16) As Double, t As Integer
    Dim codiDep As String
    Dim color As String
                        
    InformaMiss "Informe Ventas " & subj & " " & emailDe & " " & empresa

    Semana = -1
On Error GoTo ErrData
    If InStr(UCase(subj), "SEMANA") Then
        Semana = CInt(Split(subj, " ")(3))
        semanaAux = 0
        lunes = CDate("01/01/" & Year(Now()))

        If Semana > 0 Then
            While Semana <> semanaAux
                lunes = DateAdd("d", 1, lunes)
                semanaAux = DatePart("ww", lunes, vbMonday, vbFirstFullWeek)
            Wend
            
            domingo = DateAdd("d", 6, lunes)
            If domingo >= Now() Then domingo = DateAdd("d", -1, CDate(Day(Now()) & "/" & Month(Now()) & "/" & Year(Now())))
            fecha = lunes
        End If
    Else
        fecha = CDate(Split(subj, " ")(2))
        If Year(fecha) > Year(Now()) Then GoTo ErrData
    End If
    
    GoTo OkData
    
ErrData:
    fecha = Now()
    
OkData:

On Error GoTo nor

    'BUSCAMOS EL USUARIO QUE ESTÁ PIDIENDO EN INFORME
    Set rs = Db.OpenResultset("select * from dependentesextes where nom='EMAIL' and upper(valor) like '%' + upper('" & emailDe & "') + '%' order by len(valor) desc")
    If Not rs.EOF Then
        codiDep = rs("id")
        Set rs = Db.OpenResultset("select * from constantsClient where variable='SupervisoraCodi' and valor='" & codiDep & "'")
        If rs.EOF Then 'NO ES SUPERVISORA
            Set rs = Db.OpenResultset("select * from dependentesextes where nom='TIPUSTREBALLADOR' and id='" & codiDep & "'")
            If rs("Valor") = "GERENT" Or rs("Valor") = "GERENT_2" Then 'SI ES GERENTE LE DAMOS INFO DE TODAS LAS TIENDAS
                sql = "select c.codi, c.nom, isnull(d.nom, ' Franquicia') Supervisora "
                sql = sql & "from paramshw p "
                sql = sql & "left join clients c on p.valor1=c.codi "
                sql = sql & "left join constantsClient cc on c.codi=cc.codi and cc.variable='SupervisoraCodi' "
                sql = sql & "left join dependentes d on cc.valor = d.codi "
                sql = sql & "where isnull(c.nom, '') <> '' "
                sql = sql & "order by isnull(d.nom, ' Franquicia') , c.nom "
                Set rs = Db.OpenResultset(sql)
            Else
               'Mirar si es franquicia
               Set rs = Db.OpenResultset("select * from constantsClient where variable='userFranquicia' and valor='" & codiDep & "'")
               If Not rs.EOF Then 'SI ES FRANQUICIA LE DAMOS INFO DE LAS TIENDAS QUE SUPERVISA
                   Set rs = Db.OpenResultset("Select c.Codi, c.nom, isnull(d.nom, ' Franquicia') Supervisora from ConstantsClient cc left join clients c on cc.codi=c.codi left join dependentes d on cc.valor = d.codi where variable = 'userFranquicia' and valor = '" & codiDep & "' order by c.nom")
               Else
                   Exit Sub
               End If
            End If
        Else 'SUPERVISORA
             Set rs = Db.OpenResultset("Select c.Codi, c.nom, isnull(d.nom, ' Franquicia') Supervisora from ConstantsClient cc left join clients c on cc.codi=c.codi left join dependentes d on cc.valor = d.codi where variable = 'SupervisoraCodi' and valor = '" & codiDep & "' and c.codi is not null order by c.nom")
         End If
     End If
     
     co = ""
     'CoE = ""
     supervisora = ""
     
     If Semana > 0 Then
         co = co & "<h3>Informe semana " & Semana & " (" & lunes & " - " & domingo & ")</h3>"
         'CoE = CoE & "<h3>Informe semana " & Semana & " (" & lunes & " - " & domingo & ")</h3>"
     Else
         co = co & "<h3>Informe dia " & fecha & "</h3>"
         'CoE = CoE & "<h3>Informe dia " & fecha & "</h3>"
     End If
     While Not rs.EOF
         If supervisora <> rs("supervisora") Then
             If supervisora <> "" Then
                'TOTALES
                co = co & "<tr><td><b>Total</b></td>"
                'CoE = CoE & "<tr><td><b>Total</b></td>"
                co = co & "<td align='right'><b>" & FormatNumber(Totales(0), 2) & " &euro;</b></td>"
                'CoE = CoE & "<td align='right'><b>" & FormatNumber(Totales(0), 2) & " &euro;</b></td>"
                co = co & "<td align='right'><b>" & FormatNumber(Totales(1), 2) & " &euro;</b></td>"
                'CoE = CoE & "<td align='right'><b>" & FormatNumber(Totales(1), 2) & " &euro;</b></td>"
                co = co & "<td align='right'><b>" & FormatNumber(Totales(2), 2) & " &euro;</b></td>"
                'CoE = CoE & "<td align='right'><b>" & FormatNumber(Totales(2), 2) & " &euro;</b></td>"
                co = co & "<td align='right'><b>" & FormatNumber(Totales(3), 0) & "</b></td>"
                'CoE = CoE & "<td align='right'><b>" & FormatNumber(Totales(3), 0) & "</b></td>"
                If Totales(3) > 0 Then
                    co = co & "<td align='right'><b>" & FormatNumber(Totales(2) / Totales(3), 2) & " &euro;</b></td>"
                    'CoE = CoE & "<td align='right'><b>" & FormatNumber(Totales(2) / Totales(3), 2) & " &euro;</b></td>"
                Else
                    co = co & "<td align='right'><b>0.00 &euro;</b></td>"
                    'CoE = CoE & "<td align='right'><b>0.00 &euro;</b></td>"
                End If
                co = co & "<td align='right'><b>" & FormatNumber(Totales(5), 2) & " &euro;</b></td>"
                co = co & "<td align='right'><b>" & FormatNumber(Totales(6), 2) & " &euro;</b></td>"
                co = co & "<td align='right'><b>" & FormatNumber(Totales(7), 2) & " &euro;</b></td>"
                co = co & "<td align='right'><b>" & FormatNumber(Totales(8), 0) & "</b></td>"
                If Totales(8) > 0 Then
                    co = co & "<td align='right'><b>" & FormatNumber(Totales(7) / Totales(8), 2) & " &euro;</b></td>"
                Else
                    co = co & "<td align='right'><b>0.00 &euro;</b></td>"
                End If

                co = co & "<td><b>&nbsp;</b></td>" '10
                co = co & "<td><b>" & FormatNumber(Totales(11), 2) & " &euro;</b></td>"
                If Totales(12) > 0 Then
                    co = co & "<td align='right'><b>" & FormatNumber((Totales(11) / Totales(12)) * 100, 2) & " %</b></td>"
                Else
                    If D_B(11) = "0" Then
                        co = co & "<Td align='right'><b>0.00 %</b></Td>"
                    Else
                        co = co & "<Td align='right'><b>100.00 %</b></Td>"
                    End If
                End If
                co = co & "<td align='right'><b>" & FormatNumber(Totales(13), 2) & " &euro;</b></td>"
                If Totales(2) > 0 Then
                    co = co & "<td align='right'><b>" & FormatNumber((Totales(13) / Totales(2)) * 100, 2) & " %</b></td>"
                Else
                    If Totales(2) = "0" Then
                        co = co & "<Td align='right'><b>0.00 %</b></Td>"
                    Else
                        co = co & "<Td align='right'>100.00 %</Td>"
                    End If
                End If
                co = co & "<td align='right'><b>" & FormatNumber(Totales(15), 2) & "</b></td>"
                If Totales(12) > 0 Then
                    co = co & "<Td align='right'><b>" & FormatNumber((Totales(15) * 12 / Totales(12)) * 100, 2) & " %</b></Td>"
                Else
                    If Totales(15) = "0" Then
                        co = co & "<Td align='right'><b>0.00 %</b></Td>"
                    Else
                        co = co & "<Td align='right'><b>100.00 %</b></Td>"
                    End If
                End If
                co = co & "</tr>"
                '~TOTALES
                co = co & "</table><br>" & strHoras
                strHoras = ""
             End If
             co = co & "<h4>" & rs("supervisora") & "</h4>"
             co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
             co = co & "<Tr bgColor='#DADADA'><Td rowspan='2'><b>Botiga</b></Td><Td align='center' colspan='3'><b>Vendes</b></Td><Td rowspan='2'><b>Clients</b></Td><Td rowspan='2'><b>Tiquet Mig</b></Td><Td colspan='3' align='center'><b>Vendes " & Year(DateAdd("yyyy", -1, fecha)) & "</b></Td><td rowspan='2' align='center'><b>Clients <br>" & Year(DateAdd("yyyy", -1, fecha)) & "</b></td><td rowspan='2' align='center'><b>Tiquet mig <br>" & Year(DateAdd("yyyy", -1, fecha)) & "</b></td><Td rowspan='2'><b>Inc</b></Td><Td colspan='2'><b>Compres</b></Td><Td colspan='2'><b>Devolucions</b></Td><Td colspan='2'><b>Horas</b></Td></Tr>"
             co = co & "<Tr bgColor='#DADADA'><Td><b>Mati</b></td><td><b>Tarda</b></Td><td><b>Total</b></Td><Td><b>Mati</b></td><td><b>Tarda</b></Td><td><b>Total</b></Td><td><b>Total</b></Td><td><b>%</b></Td><td><b>Total</b></Td><td><b>%</b></Td><td><b>Total</b></Td><td><b>%</b></Td></Tr>"
             supervisora = rs("supervisora")
             For t = 0 To 16
                Totales(t) = 0
             Next
         End If
         
         totalDia = 0
         total7 = 0
           
         CarregaReportBotiga CDbl(rs("Codi")), D_B, fecha, strHoras, Semana, empresa
         
         'Nom botiga
         co = co & "<Tr><Td><b>" & D_B(0) & "</b></Td>"
         
         'Vendes
         If InStr(D_B(1), "|") Then
             co = co & "<Td align='right'>" & Split(D_B(1), "|")(1) & " &euro;</td>"
             totalDia = Split(D_B(1), "|")(1)
             Totales(0) = Totales(0) + Split(D_B(1), "|")(1)
             If UBound(Split(D_B(1), "|")) > 1 Then
                 co = co & "<td align='right'>" & Split(D_B(1), "|")(2) & " &euro;</Td>"
                 totalDia = totalDia + Split(D_B(1), "|")(2)
                 Totales(1) = Totales(1) + Split(D_B(1), "|")(2)
             Else
                 co = co & "<td>&nbsp;</td>"
             End If
         Else
             co = co & "<Td align='right'>" & D_B(1) & " &euro;</Td><td>&nbsp;</td>"
             If D_B(1) <> "" Then totalDia = D_B(1)
         End If
         co = co & "<Td align='right'>" & FormatNumber(totalDia, 2) & " &euro;</Td>"
         Totales(2) = Totales(2) + totalDia
         
         'Clients
         co = co & "<Td align='right'>" & D_B(3) & "</Td>"
         Totales(3) = Totales(3) + D_B(3)
         
         'Tiquet mig
         If D_B(3) > 0 Then
             co = co & "<Td align='right'>" & FormatNumber(totalDia / D_B(3), 2) & " &euro;</Td>"
         Else
             co = co & "<Td align='right'>" & FormatNumber(0, 2) & " &euro;</Td>"
         End If
         Totales(4) = Totales(4) + 0 'TIENE SENTIDO TOTAL DE TIQUET MIG???
         
         'Co = Co & "<Td>" & D_B(4) & "</Td>"
         'Vendes any anterior
         If InStr(D_B(7), "|") Then
             co = co & "<Td align='right'>" & Split(D_B(7), "|")(1) & " &euro;</td>"
             total7 = Split(D_B(7), "|")(1)
             Totales(5) = Totales(5) + Split(D_B(7), "|")(1)
             If UBound(Split(D_B(7), "|")) > 1 Then
                 co = co & "<td align='right'>" & Split(D_B(7), "|")(2) & " &euro;</Td>"
                 total7 = total7 + Split(D_B(7), "|")(2)
                 Totales(6) = Totales(6) + Split(D_B(7), "|")(2)
             Else
                 co = co & "<td>&nbsp;</td>"
             End If
         Else
             co = co & "<Td align='right'>" & D_B(7) & " &euro;</Td><td>&nbsp;</td>"
             If D_B(7) <> "" Then total7 = D_B(7)
         End If
         co = co & "<Td align='right'>" & FormatNumber(total7, 2) & " &euro;</Td>"
         Totales(7) = Totales(7) + total7
         
        'Clients año anterior
         co = co & "<Td align='right'>" & D_B(10) & "</Td>"
         Totales(8) = Totales(8) + D_B(10)
         
         'Tiquet mig
         If D_B(10) > 0 Then
             co = co & "<Td align='right'>" & FormatNumber(total7 / D_B(10), 2) & " &euro;</Td>"
         Else
             co = co & "<Td align='right'>" & FormatNumber(0, 2) & " &euro;</Td>"
         End If
         Totales(9) = Totales(9) + 0 'TIENE SENTIDO TOTAL DE TIQUET MIG???
         
         '%inc
         co = co & "<Td>"
         co = co & "<font color='" & D_B(5) & "'>"
         co = co & "%</font></Td>"
         Totales(10) = Totales(10) + 0 'TIENE SENTIDO TOTAL ???
         
         'Compras / ventas sin IVA
         co = co & "<Td align='right'>" & FormatNumber(D_B(6), 2) & " &euro;</Td>"
         Totales(11) = Totales(11) + D_B(6)
         If D_B(9) > 0 Then
             co = co & "<Td align='right'>" & FormatNumber((D_B(6) / D_B(9)) * 100, 2) & " %</Td>"
         Else
             If D_B(6) = "0" Then
                 co = co & "<Td align='right'>0.00 %</Td>"
             Else
                 co = co & "<Td align='right'>100.00 %</Td>"
             End If
         End If
         Totales(12) = Totales(12) + D_B(9)
         
         'Devolucions
         
         
         color = "#A6FFA6"
         If Semana > 0 Then
             If D_B(2) > (160 * 7) Or D_B(2) < (85 * 7) Then
                 color = "#FFA8A8"
             End If
         Else
             If D_B(2) > 160 Or D_B(2) < 85 Then
                 color = "#FFA8A8"
             End If
         End If
         
         co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(D_B(2), 2) & " &euro;</Td>"
         If totalDia > 0 Then
             'If (D_B(2) / totalDia) * 100 > 10 Or (D_B(2) / totalDia) * 100 < 5 Then
             co = co & "<Td align='right'>" & FormatNumber((D_B(2) / totalDia) * 100, 2) & " %</Td>"
         Else
             If D_B(2) = "0" Then
                 co = co & "<Td align='right'>0.00 %</Td>"
             Else
                 co = co & "<Td align='right'>100.00 %</Td>"
             End If
         End If
         Totales(13) = Totales(13) + D_B(2)
         
         'Horas
         co = co & "<Td align='right'>" & FormatNumber(D_B(8), 2) & "</Td>"
         If D_B(9) > 0 Then
             co = co & "<Td align='right'>" & FormatNumber((D_B(8) * 12 / D_B(9)) * 100, 2) & " %</Td>"
         Else
             If D_B(8) = "0" Then
                 co = co & "<Td align='right'>0.00 %</Td>"
             Else
                 co = co & "<Td align='right'>100.00 %</Td>"
             End If
         End If
         Totales(15) = Totales(15) + D_B(8)
         
         co = co & "</Tr>"
         
         rs.MoveNext
    Wend
    
    'TOTALES
    co = co & "<tr><td><b>Total</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(Totales(0), 2) & " &euro;</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(Totales(1), 2) & " &euro;</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(Totales(2), 2) & " &euro;</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(Totales(3), 0) & "</b></td>"
    If Totales(3) > 0 Then
        co = co & "<td align='right'><b>" & FormatNumber(Totales(2) / Totales(3), 2) & " &euro;</b></td>"
    Else
        co = co & "<td align='right'><b>0.00 &euro;</b></td>"
    End If
    co = co & "<td align='right'><b>" & FormatNumber(Totales(5), 2) & " &euro;</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(Totales(6), 2) & " &euro;</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(Totales(7), 2) & " &euro;</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(Totales(8), 0) & "</b></td>"
    If Totales(8) > 0 Then
        co = co & "<td align='right'><b>" & FormatNumber(Totales(7) / Totales(8), 2) & " &euro;</b></td>"
    Else
        co = co & "<td align='right'><b>0.00 &euro;</b></td>"
    End If

    co = co & "<td><b>&nbsp;</b></td>" '10
    co = co & "<td><b>" & FormatNumber(Totales(11), 2) & " &euro;</b></td>"
    If Totales(12) > 0 Then
        co = co & "<td align='right'><b>" & FormatNumber((Totales(11) / Totales(12)) * 100, 2) & " %</b></td>"
    Else
        If D_B(11) = "0" Then
            co = co & "<Td align='right'><b>0.00 %</b></Td>"
        Else
            co = co & "<Td align='right'><b>100.00 %</b></Td>"
        End If
    End If
    co = co & "<td align='right'><b>" & FormatNumber(Totales(13), 2) & " &euro;</b></td>"
    If Totales(2) > 0 Then
        co = co & "<td align='right'><b>" & FormatNumber((Totales(13) / Totales(2)) * 100, 2) & " %</b></td>"
    Else
        If Totales(2) = "0" Then
            co = co & "<Td align='right'><b>0.00 %</b></Td>"
        Else
            co = co & "<Td align='right'>100.00 %</Td>"
        End If
    End If
    co = co & "<td align='right'><b>" & FormatNumber(Totales(15), 2) & "</b></td>"
    If Totales(12) > 0 Then
        co = co & "<Td align='right'><b>" & FormatNumber((Totales(15) * 12 / Totales(12)) * 100, 2) & " %</b></Td>"
    Else
        If Totales(15) = "0" Then
            co = co & "<Td align='right'><b>0.00 %</b></Td>"
        Else
            co = co & "<Td align='right'><b>100.00 %</b></Td>"
        End If
    End If
    co = co & "</tr>"
    '~TOTALES
    
    On Error Resume Next
     
    co = co & "</table><br>"
    co = co & "<br>"
    co = co & strHoras
    'CoE = CoE & "</table><br>" & strHoras
     
    sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreInformeVentas  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co & err.Description, "", ""
End Sub


Sub SecreResumenVentas(subj As String, emailDe As String)
    Dim codiDep As String
    Dim Semana As Integer, semanaAux As Integer, lunes As Date, domingo As Date, fecha As Date
    Dim rs As rdoResultset, sql As String
    Dim co As String
    Dim nomBot As String, codiBot As Double
    Dim vendesArr(3) As Double, clientsArr(2) As Integer, previsionsArr(2) As Double
    Dim vendesIva As Double, vendesNOIva As Double, vendesDiada As Double, previsions As Double, vendesInterEmp As Double
    Dim horasPactadas As Double, horasReales As Double, horasAprendiz As Double, horasCoord As Double, horasSinTurno As Double
    Dim compresReal As Double, compresDiada As Double, compresAsterisc As Double, compresBotiga As Double, compresInterEmp As Double
    Dim incidencia As Boolean
                        
    InformaMiss "Informe Resumen Ventas " & subj & " " & emailDe
    
    Semana = -1
On Error GoTo ErrData
    If InStr(UCase(subj), "SEMANA") Then
        Semana = CInt(Split(subj, " ")(3))
        semanaAux = 0
        If Semana = 1 Then
            lunes = CDate("26/12/" & Year(Now()) - 1) 'La primera semana del año como mucho podría ser del lunes 26 al domingo 1
            While Weekday(lunes, 2) <> 1
                lunes = DateAdd("d", 1, lunes)
            Wend
        Else
            lunes = CDate("01/01/" & Year(Now()))
            If Semana > 0 Then
                While Semana <> semanaAux
                    lunes = DateAdd("d", 1, lunes)
                    semanaAux = DatePart("ww", lunes, vbMonday)
                Wend
            Else
                lunes = DateAdd("d", -((Weekday(Now(), 2)) - 1), Now())
            End If
            
        End If
    Else
        lunes = DateAdd("d", -((Weekday(Now(), 2)) - 1), Now())
    End If
    
    domingo = DateAdd("d", 6, lunes)
    If domingo >= Now() Then domingo = DateAdd("d", -1, CDate(Day(Now()) & "/" & Month(Now()) & "/" & Year(Now())))
    
    GoTo OkData
    
ErrData:
    lunes = DateAdd("d", -((Weekday(Now(), 2)) - 1), Now())
    domingo = DateAdd("d", 6, lunes)
    If domingo >= Now() Then domingo = DateAdd("d", -1, CDate(Day(Now()) & "/" & Month(Now()) & "/" & Year(Now())))
    
OkData:

On Error GoTo nor

    'BUSCAMOS EL USUARIO QUE ESTÁ PIDIENDO EN INFORME
    Set rs = Db.OpenResultset("select * from dependentesextes where nom='EMAIL' and upper(valor) like '%' + upper('" & emailDe & "') + '%' order by len(valor) desc")
    If Not rs.EOF Then
       codiDep = rs("id")
       Set rs = Db.OpenResultset("select * from constantsClient where variable='SupervisoraCodi' and valor='" & codiDep & "'")
       If rs.EOF Then 'NO ES SUPERVISORA
           Set rs = Db.OpenResultset("select * from dependentesextes where nom='TIPUSTREBALLADOR' and id='" & codiDep & "'")
           If rs("Valor") = "GERENT" Or rs("Valor") = "GERENT_2" Then 'SI ES GERENTE LE DAMOS INFO DE TODAS LAS TIENDAS MENOS LAS FRANQUICIAS
               sql = "select c.codi, c.nom "
               sql = sql & "from paramshw p "
               sql = sql & "left join clients c on p.valor1=c.codi "
               sql = sql & "where isnull(c.nom, '') <> '' and c.codi not in (select codi from constantsClient where variable='Franquicia' and valor='Franquicia') "
               sql = sql & "order by c.nom"
               Set rs = Db.OpenResultset(sql)
           Else
               sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "NO TENS PERMÍS PER REBRE AQUESTA INFORMACIÓ", "", ""
               Exit Sub
           End If
       Else 'SUPERVISORA
            Set rs = Db.OpenResultset("Select c.Codi, c.nom, isnull(d.nom, ' Franquicia') Supervisora from ConstantsClient cc left join clients c on cc.codi=c.codi left join dependentes d on cc.valor = d.codi where variable = 'SupervisoraCodi' and valor = '" & codiDep & "' and c.codi is not null order by c.nom")
        End If
    End If
    
    co = ""
    co = co & "<h3>Informe semana (" & lunes & " - " & domingo & ")</h3>"

    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'><Td><b>Botiga</b></Td><Td><b>Vendes<br>amb IVA</b></Td><Td><b>Vendes<br>sense IVA</b></Td><Td><b>Vendes<br>Diada</b></Td><Td><b>Vendes<br>Inter-Empreses</b></Td><Td><b>Previsió</b></Td><Td><b>Compres<br>real</b></Td><Td><b>Compres<br>diada</b></Td><Td><b>Compres *</b></Td><Td><b>Compres<br>botiga</b></Td><Td><b>Compres<br>Inter-Empreses</b></Td><Td><b>Hores<br>pactades</b></Td><Td><b>Hores<br>real</b></Td><Td><b>Hores<br>Coordinació+Aprenent</b></Td><Td><b>Hores<br>No assignades</b></Td></Tr>"

    While Not rs.EOF
        vendesIva = 0: vendesNOIva = 0: vendesDiada = 0: previsions = 0: vendesInterEmp = 0
        incidencia = False
                
        codiBot = rs("codi")
        nomBot = rs("nom")
        'VENDES
        For fecha = lunes To domingo
            calculaVendesClients codiBot, fecha, vendesArr, clientsArr
            calculaPrevisions codiBot, fecha, previsionsArr
            
            vendesIva = vendesIva + vendesArr(0) + vendesArr(1)
            vendesNOIva = vendesNOIva + vendesArr(2)
            previsions = previsions + previsionsArr(0) + previsionsArr(1)
            If vendesArr(0) = 0 Or vendesArr(1) = 0 Then incidencia = True
        Next
        
        co = co & "<tr>"
        If incidencia Then
            co = co & "<td bgcolor='#FF0000'><b>" & UCase(nomBot) & "</b></td>"
        Else
            co = co & "<td><b>" & UCase(nomBot) & "</b></td>"
        End If
        
        co = co & "<td align='right'>" & FormatNumber(vendesIva, 2) & " &euro;</td>"
        co = co & "<td align='right'>" & FormatNumber(vendesNOIva, 2) & " &euro;</td>"
        
        vendesDiada = calculaVendesFamilia(codiBot, lunes, domingo, "Diada")
        co = co & "<td align='right'>" & FormatNumber(vendesDiada) & " &euro;</td>"
        
        vendesInterEmp = calculaVendesInterEmpreses(codiBot, lunes, domingo)
        co = co & "<td align='right'>" & FormatNumber(vendesInterEmp, 2) & " &euro;</td>"
        
        co = co & "<td align='right'>" & FormatNumber(previsions, 2) & " &euro;</td>"
        
        'COMPRES
        compresReal = 0: compresDiada = 0: compresAsterisc = 0: compresBotiga = 0: compresInterEmp = 0
        For fecha = lunes To domingo
            compresReal = compresReal + calculaComprasReal(codiBot, fecha)
            compresDiada = compresDiada + calculaComprasFamilia(codiBot, fecha, "Diada")
            compresAsterisc = compresAsterisc + calculaComprasAsterisco(codiBot, fecha)
            compresBotiga = compresBotiga + calculaComprasTienda(codiBot, fecha)
        Next
        
        co = co & "<td align='right'>" & FormatNumber(compresReal, 2) & " &euro;</td>"
        co = co & "<td align='right'>" & FormatNumber(compresDiada, 2) & " &euro;</td>"
        co = co & "<td align='right'>" & FormatNumber(compresAsterisc, 2) & " &euro;</td>"
        co = co & "<td align='right'>" & FormatNumber(compresBotiga, 2) & " &euro;</td>"
        
        compresInterEmp = calculaCompresInterEmpreses(codiBot, lunes, domingo)
        co = co & "<td align='right'>" & FormatNumber(compresInterEmp, 2) & " &euro;</td>"
        
        'HORES
        horasPactadas = 0: horasReales = 0: horasAprendiz = 0: horasCoord = 0: horasSinTurno = 0
        For fecha = lunes To domingo
            horasPactadas = horasPactadas + getObjetivoHoras(codiBot, fecha)
            horasReales = horasReales + calculaHorasReales(codiBot, fecha)
            horasAprendiz = horasAprendiz + calculaHorasAprendiz(codiBot, fecha)
            horasCoord = horasCoord + calculaHorasCoordinacion(codiBot, fecha)
            horasSinTurno = horasSinTurno + calculaHorasSinTurno(codiBot, fecha)
        Next
        co = co & "<td align='right'>" & FormatNumber(horasPactadas, 2) & "</td>"
        co = co & "<td align='right'>" & FormatNumber(horasReales, 2) & "</td>"
        co = co & "<td align='right'>" & FormatNumber((horasCoord + horasAprendiz), 2) & "</td>"
        co = co & "<td align='right'>" & FormatNumber((horasSinTurno), 2) & "</td>"
        
        
        
        co = co & "</tr>"
        
        rs.MoveNext
    Wend
    
    co = co & "</table><br>"
    sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreResumenVentas  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co & err.Description, "", ""
End Sub


Function getObjetivoHoras(codiBot As Double, fecha As Date) As Double
    Dim totalHoras As Double
    Dim rsParams As rdoResultset
    Dim sql As String
    
    totalHoras = 0
    sql = "select isnull(sum(datediff(minute, horaInicio, horafin) / 60.00), 0) TotalHoras "
    sql = sql & "from " & taulaCdpPlanificacion(fecha) & " p "
    sql = sql & "left join cdpTurnos t on p.idturno = t.idTurno "
    sql = sql & "where t.idTurno is not null and p.botiga=" & codiBot & " and convert(nvarchar, p.fecha, 103)='" & Right("00" & Day(fecha), 2) & "/" & Right("00" & Month(fecha), 2) & "/" & Year(fecha) & "' and p.activo=1"
    Set rsParams = Db.OpenResultset(sql)
    If Not rsParams.EOF Then
        If rsParams("TotalHoras") > 0 Then
            totalHoras = rsParams("TotalHoras")
        End If
    End If

    getObjetivoHoras = totalHoras
End Function


Function CarregaReportBotiga(codiBotiga As Double, D() As String, dia As Date, strHoras As String, Semana As Integer, empresa As String)
    Dim totalDiari As Double, totalDiari7 As Double, rsCaixes As rdoResultset, rs As rdoResultset, Rs2 As rdoResultset
    Dim sql As String
    Dim dataInici As Date, dataFi As Date
    Dim lunes As Date, domingo As Date, semanaAux As Integer
    Dim LunAnyoAnt As Date, DomAnyoAnt As Date
    Dim mati As Double, tarda As Double
    Dim cMati As Double, cTarda As Double
    Dim nDias As Integer
    Dim diaAnyoAnt As Date
    
 
On Error GoTo nor:
        
    'ExecutaComandaSql "SET DATEFIRST 7"
    
    diaAnyoAnt = DateAdd("yyyy", -1, dia)
    While DatePart("w", dia) <> DatePart("w", diaAnyoAnt)
        diaAnyoAnt = DateAdd("d", 1, diaAnyoAnt)
    Wend
    
    'Nom botiga
    D(0) = BotigaCodiNom(codiBotiga)
    InformaMiss "Informe Ventas " & dia & " " & D(0)
    
    'Ventas
    totalDiari = 0
    D(1) = 0
    D(9) = 0
    
    semanaAux = 0
    'If Semana < DatePart("ww", Now(), vbMonday, vbFirstFullWeek) Then
    '    lunes = CDate("01/01/" & Year(Now()) + 1)
    '    LunAnyoAnt = CDate("01/01/" & Year(Now()))
    'Else
    '    lunes = CDate("01/01/" & Year(Now()))
    '    LunAnyoAnt = CDate("01/01/" & Year(Now()) - 1)
    'End If

    lunes = CDate("01/01/" & Year(Now()))
    LunAnyoAnt = CDate("01/01/" & Year(Now()) - 1)
    If Semana > 0 Then
        While Semana <> semanaAux
            lunes = DateAdd("d", 1, lunes)
            semanaAux = DatePart("ww", lunes, vbMonday, vbFirstFullWeek)
        Wend
        
        domingo = DateAdd("d", 6, lunes)
        If domingo >= Now() Then domingo = DateAdd("d", -1, CDate(Day(Now()) & "/" & Month(Now()) & "/" & Year(Now())))
        nDias = DateDiff("d", lunes, domingo)
        
        semanaAux = 0
        While Semana <> semanaAux
            LunAnyoAnt = DateAdd("d", 1, LunAnyoAnt)
            semanaAux = DatePart("ww", LunAnyoAnt, vbMonday, vbFirstFullWeek)
        Wend
        
        DomAnyoAnt = DateAdd("d", nDias, LunAnyoAnt)
        
        
        sql = "select distinct data, tipus_moviment "
        If Month(lunes) <> Month(domingo) Then
            sql = sql & "from (select * from [" & NomTaulaMovi(lunes) & "] union all select * from [" & NomTaulaMovi(domingo) & "]) m where "
        Else
            sql = sql & "from [" & NomTaulaMovi(lunes) & "] where "
        End If
        sql = sql & "botiga='" & codiBotiga & "' and data between '" & lunes & " 00:00' and '" & domingo & " 23:59' and (tipus_moviment='Wi' or tipus_moviment='W') order by data"
    Else
        sql = "select distinct data, tipus_moviment "
        sql = sql & "from [" & NomTaulaMovi(dia) & "] where "
        sql = sql & "botiga='" & codiBotiga & "' and day(data)=" & Day(dia) & " and (tipus_moviment='Wi' or tipus_moviment='W') order by data"
    End If
    Set rsCaixes = Db.OpenResultset(sql)
    
    If rsCaixes.EOF Then
        Set rs = Db.OpenResultset("Select isnull(sum(import), 0) I from [" & NomTaulaVentas(dia) & "] where botiga =  " & codiBotiga & " And day(data) = " & Day(dia) & " ")
        D(1) = FormatNumber(rs("I"), 2)
        totalDiari = totalDiari + rs("I")
    End If
    
    mati = 0
    tarda = 0
    While Not rsCaixes.EOF
        If rsCaixes("tipus_moviment") = "Wi" Then
            dataInici = rsCaixes("data")
            rsCaixes.MoveNext
            
            If Not rsCaixes.EOF Then
                If rsCaixes("tipus_moviment") = "W" Then
                    dataFi = rsCaixes("data")
                    
                    Set Rs2 = Db.OpenResultset("select isnull(sum(import), 0) I, isnull(sum(import/(1+(t.iva/100))), 0) sinIva, count(distinct Num_Tick ) Clients from [" & NomTaulaVentas(dataInici) & "] v left join articles a on v.plu=a.codi left join TipusIva2012 t on a.TipoIva = t.tipus where botiga=" & codiBotiga & " and data between '" & dataInici & "' and '" & dataFi & "'")
                    If Not Rs2.EOF Then
                        If Not IsNull(Rs2("I")) Then
                            If DatePart("h", dataInici) < 13 Then 'MATÍ
                                mati = mati + Rs2("I")
                                cMati = cMati + Rs2("Clients")
                            Else 'TARDA
                                tarda = tarda + Rs2("I")
                                cTarda = cTarda + Rs2("Clients")
                            End If
                            D(9) = D(9) + Rs2("sinIva")
                            totalDiari = totalDiari + Rs2("I")
                        End If
                    End If
                End If
            End If
        Else
            rsCaixes.MoveNext
        End If
    Wend
    D(1) = "|" & FormatNumber(mati, 2) & "|" & FormatNumber(tarda, 2)
    If cMati = 0 Then
        D(11) = "|0.00|"
    Else
        D(11) = "|" & FormatNumber(mati / cMati, 2) & "|"
    End If
    If cTarda = 0 Then
        D(11) = D(11) & "0.00"
    Else
        D(11) = D(11) & FormatNumber(tarda / cTarda, 2) 'Tiquet mig matí/tarda
    End If
    
    'Devolucions
    Dim tServits As String
    Dim f As Date
    
    If Semana > 0 Then
        'If Month(lunes) <> Month(domingo) Then
        '    Set rs2 = Db.OpenResultset("Select isnull(sum(Import), 0) I from (select * from [" & NomTaulaDevol(lunes) & "] union all select * from [" & NomTaulaDevol(domingo) & "]) d where botiga =  " & codiBotiga & " and data between '" & lunes & " 00:00' and '" & domingo & " 23:59' ")
        'Else
        '    Set rs2 = Db.OpenResultset("Select isnull(sum(Import), 0) I from [" & NomTaulaDevol(lunes) & "] where botiga =  " & codiBotiga & "  and data between '" & lunes & " 00:00' and '" & domingo & " 23:59'")
        'End If
        tServits = ""
        For f = lunes To domingo
            If tServits = "" Then
                tServits = "select * from " & DonamTaulaServit(f)
            Else
                tServits = tServits & " union all select * from " & DonamTaulaServit(f)
            End If
        Next
        If tServits = "" Then tServits = "select * from " & DonamTaulaServit(f)
        
        sql = "select isnull(sum(a.preu*s.quantitatTornada), 0) I "
        sql = sql & "from (" & tServits & ") s "
        sql = sql & "left join articles a on s.codiarticle=a.codi "
        sql = sql & "left join clients c on s.client = c.codi "
        'sql = sql & "left join TarifesEspecials te on te.codi = c.[Desconte 5] "
        sql = sql & "Where s.client = " & codiBotiga & " And s.quantitatTornada > 0"
        Set Rs2 = Db.OpenResultset(sql)
    Else
        'Set rs2 = Db.OpenResultset("Select isnull(sum(Import), 0) I from [" & NomTaulaDevol(Dia) & "] where botiga =  " & codiBotiga & " And day(data) = " & Day(Dia) & " ")
        sql = "select isnull(sum(a.preu*s.quantitatTornada), 0) I "
        sql = sql & "from " & DonamTaulaServit(dia) & " s "
        sql = sql & "left join articles a on s.codiarticle=a.codi "
        sql = sql & "left join clients c on s.client = c.codi "
        'sql = sql & "left join TarifesEspecials te on te.codi = c.[Desconte 5] "
        sql = sql & "Where s.client = " & codiBotiga & " And s.quantitatTornada > 0"
        Set Rs2 = Db.OpenResultset(sql)
    End If
    
    D(2) = "0"
    If Not Rs2.EOF Then
        If Not IsNull(Rs2("I")) Then
            D(2) = Rs2("I")
        End If
    End If
    
    'Clientes
    If Semana > 0 Then
        If Month(lunes) <> Month(domingo) Then
            Set rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients from (select * from [" & NomTaulaVentas(lunes) & "] union all select * from [" & NomTaulaVentas(domingo) & "]) v where botiga =  " & codiBotiga & " And data between '" & lunes & " 00:00' and '" & domingo & " 23:59' ")
        Else
            Set rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients from [" & NomTaulaVentas(lunes) & "] where botiga =  " & codiBotiga & " And data between '" & lunes & " 00:00' and '" & domingo & " 23:59' ")
        End If
    Else
        Set rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients from [" & NomTaulaVentas(dia) & "] where botiga =  " & codiBotiga & " And day(data) = " & Day(dia) & " ")
    End If
    D(3) = rs("Clients")
    
    'Clientes año anterior
    If Semana > 0 Then
        If Month(LunAnyoAnt) <> Month(DomAnyoAnt) Then
            Set rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients from (select * from [" & NomTaulaVentas(LunAnyoAnt) & "] union all select * from [" & NomTaulaVentas(DomAnyoAnt) & "]) v where botiga =  " & codiBotiga & " And data between '" & LunAnyoAnt & " 00:00' and '" & DomAnyoAnt & " 23:59' ")
        Else
            Set rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients from [" & NomTaulaVentas(LunAnyoAnt) & "] where botiga =  " & codiBotiga & " And data between '" & LunAnyoAnt & " 00:00' and '" & DomAnyoAnt & " 23:59' ")
        End If
    Else
        Set rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients from [" & NomTaulaVentas(diaAnyoAnt) & "] where botiga =  " & codiBotiga & " And day(data) = " & Day(diaAnyoAnt) & " ")
    End If
    If rs("clients") = 0 Then  'Si NO hay datos del año pasado buscamos en la semana pasada
        If Semana > 0 Then
            If Month(DateAdd("d", -7, lunes)) <> Month(DateAdd("d", -7, domingo)) Then
                Set rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients from (select * from [" & NomTaulaVentas(DateAdd("d", -7, lunes)) & "] union all select * from [" & NomTaulaVentas(DateAdd("d", -7, domingo)) & "]) v where botiga =  " & codiBotiga & " And data between '" & DateAdd("d", -7, lunes) & " 00:00' and '" & DateAdd("d", -7, domingo) & " 23:59' ")
            Else
                Set rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients from [" & NomTaulaVentas(DateAdd("d", -7, lunes)) & "] where botiga =  " & codiBotiga & " And data between '" & DateAdd("d", -7, lunes) & " 00:00' and '" & DateAdd("d", -7, domingo) & " 23:59' ")
            End If
        Else
            Set rs = Db.OpenResultset("Select count(distinct Num_Tick ) Clients from [" & NomTaulaVentas(DateAdd("d", -7, dia)) & "] where botiga =  " & codiBotiga & " And day(data) = " & Day(DateAdd("d", -7, dia)) & " ")
        End If
        D(10) = rs("Clients")
        D(12) = "SEMANA"
    Else
        D(10) = rs("Clients")
        D(12) = "ANY"
    End If
    
    'Primera y última venta
    If Semana > 0 Then
        If Month(lunes) <> Month(domingo) Then
            Set rs = Db.OpenResultset("Select isnull(min(data), getdate()) hi ,isnull(Max(data), getdate()) hf from (select * from [" & NomTaulaVentas(lunes) & "] union all select * from [" & NomTaulaVentas(domingo) & "]) v where botiga =  " & codiBotiga & " And data between '" & lunes & " 00:00' and '" & domingo & " 23:59' ")
        Else
            Set rs = Db.OpenResultset("Select isnull(min(data), getdate()) hi ,isnull(Max(data), getdate()) hf from [" & NomTaulaVentas(lunes) & "] where botiga =  " & codiBotiga & " And data between '" & lunes & " 00:00' and '" & domingo & " 23:59' ")
        End If
    Else
        Set rs = Db.OpenResultset("Select isnull(min(data), getdate()) hi ,isnull(Max(data), getdate()) hf from [" & NomTaulaVentas(dia) & "] where botiga =  " & codiBotiga & " And day(data) = " & Day(dia) & " ")
    End If
    D(4) = Format(rs("Hi"), "hh:nn") & "/" & Format(rs("Hf"), "hh:nn")
    
    'Ventas 1 año atras
    totalDiari7 = 0
    D(7) = 0
    
    If Semana > 0 Then
        sql = "select distinct data, tipus_moviment "
        If Month(LunAnyoAnt) <> Month(DomAnyoAnt) Then
            sql = sql & "from (select * from [" & NomTaulaMovi(LunAnyoAnt) & "] union all select * from [" & NomTaulaMovi(DomAnyoAnt) & "]) m where "
        Else
            sql = sql & "from [" & NomTaulaMovi(LunAnyoAnt) & "] where "
        End If
        sql = sql & "botiga='" & codiBotiga & "' and data between '" & LunAnyoAnt & " 00:00' and '" & DomAnyoAnt & " 23:59' and (tipus_moviment='Wi' or tipus_moviment='W') order by data"
    Else
        sql = "select distinct data, tipus_moviment "
        sql = sql & "from [" & NomTaulaMovi(diaAnyoAnt) & "] where "
        sql = sql & "botiga='" & codiBotiga & "' and day(data)=" & Day(diaAnyoAnt) & " and (tipus_moviment='Wi' or tipus_moviment='W') order by data"
    End If
    Set rsCaixes = Db.OpenResultset(sql)
    
    If rsCaixes.EOF Then
        Set rs = Db.OpenResultset("Select isnull(sum(import), 0) I from [" & NomTaulaMovi(diaAnyoAnt) & "] where botiga =  " & codiBotiga & " And day(data) = " & Day(diaAnyoAnt) & " ")
        D(7) = FormatNumber(rs("I"), 2)
        totalDiari7 = totalDiari7 + rs("I")
    End If
    
    mati = 0
    tarda = 0
    While Not rsCaixes.EOF
        If rsCaixes("tipus_moviment") = "Wi" Then
            dataInici = rsCaixes("data")
            rsCaixes.MoveNext
            
            If Not rsCaixes.EOF Then
                If rsCaixes("tipus_moviment") = "W" Then
                    dataFi = rsCaixes("data")
                    
                    Set Rs2 = Db.OpenResultset("select isnull(sum(import), 0) I from [" & NomTaulaVentas(dataInici) & "] where botiga=" & codiBotiga & " and data between '" & dataInici & "' and '" & dataFi & "'")
                    If Not Rs2.EOF Then
                        If Not IsNull(Rs2("I")) Then
                            'd(7) = d(7) & "|" & FormatNumber(rs2("I"), 2)
                             If DatePart("h", dataInici) < 13 Then 'MATÍ
                                mati = mati + Rs2("I")
                            Else 'TARDA
                                tarda = tarda + Rs2("I")
                            End If

                            totalDiari7 = totalDiari7 + Rs2("I")
                        End If
                    End If
                End If
            End If
        Else
            rsCaixes.MoveNext
        End If
    Wend
    D(7) = "|" & FormatNumber(mati, 2) & "|" & FormatNumber(tarda, 2)
        
    'Verde incremento ventas - rojo descenso ventas
    If totalDiari >= totalDiari7 Then
        D(5) = "green"
    Else
        D(5) = "Red"
    End If
    
On Error GoTo errFac
    'Compras
    'Buscamos la factura
    Dim facTabData As String, facTabIva As String
    facTabData = "Facturacio_"
    facTabIva = "Facturacio_"
    If Month(dia) = 12 Then
        facTabData = facTabData & Year(dia) + 1 & "-01_Data"
        facTabIva = facTabIva & Year(dia) + 1 & "-01_Iva"
    Else
        facTabData = facTabData & Year(dia) & "-" & Right("0" & Month(dia) + 1, 2) & "_Data"
        facTabIva = facTabIva & Year(dia) & "-" & Right("0" & Month(dia) + 1, 2) & "_Iva"
    End If
    
    D(6) = 0
    If Semana > 0 Then
        sql = "select isnull(sum(import), 0) compras "
        sql = sql & "from [Facturacio_" & Year(dia) & "-" & Right("0" & Month(dia), 2) & "_Iva] i "
        sql = sql & "left join [Facturacio_" & Year(dia) & "-" & Right("0" & Month(dia), 2) & "_Data] d on i.idFactura=d.idFactura "
        sql = sql & "where i.serie not like '%RE%' and d.client = " & codiBotiga & " and d.producteNom not like '***%'  and i.DataInici = '" & lunes & "' and i.DataFi = '" & domingo & "' "
        Set Rs2 = Db.OpenResultset(sql)
        
        If (Rs2.EOF Or Rs2("compras") = 0) And ExisteixTaula(facTabIva) Then
            sql = "select isnull(sum(import), 0) compras "
            sql = sql & "from [" & facTabIva & "] i "
            sql = sql & "left join [" & facTabData & "] d on i.idFactura=d.idFactura "
            sql = sql & "where i.serie not like '%RE%' and d.client = " & codiBotiga & " and d.producteNom not like '***%'  and i.DataInici = '" & lunes & "' and i.DataFi = '" & domingo & "' "
            Set Rs2 = Db.OpenResultset(sql)
        End If
        If Not Rs2.EOF Then D(6) = Rs2("compras")
    Else
        sql = "select isnull(sum(import), 0) compras "
        sql = sql & "from [Facturacio_" & Year(dia) & "-" & Right("0" & Month(dia), 2) & "_Data] d "
        sql = sql & "left join [Facturacio_" & Year(dia) & "-" & Right("0" & Month(dia), 2) & "_Iva] i on d.idFactura=i.idfactura "
        sql = sql & "where i.serie not like '%RE%' and day(d.data)=" & Day(dia) & " and month(d.data)=" & Month(dia) & " and year(d.data)=" & Year(dia) & " and d.client = " & codiBotiga & " and d.producteNom not like '***%'"
        Set Rs2 = Db.OpenResultset(sql)
        
        If (Rs2.EOF Or Rs2("compras") = 0) And ExisteixTaula(facTabIva) Then
            sql = "select isnull(sum(import), 0) compras "
            sql = sql & "from [" & facTabData & "] d "
            sql = sql & "left join [" & facTabIva & "] i on d.idFactura=i.idfactura "
            sql = sql & "where i.serie not like '%RE%' and day(d.data)=" & Day(dia) & " and month(d.data)=" & Month(dia) & " and year(d.data)=" & Year(dia) & " and d.client = " & codiBotiga & " and d.producteNom not like '***%'"
            Set Rs2 = Db.OpenResultset(sql)
        End If
        If Not Rs2.EOF Then D(6) = Rs2("compras")
    End If
    
errFac:
    If D(6) = 0 Then 'servit
          sql = "select isnull(sum(s.quantitatServida * (a.preumajor - (a.preumajor * case when a.desconte=1 then cast(c.[Desconte 1] as float)/100 when a.desconte=2 then cast(c.[Desconte 2] as float)/100 when a.desconte=3 then cast(c.[Desconte 3] as float)/100 when a.desconte=4 then cast(c.[Desconte 4] as float)/100 else 0 end ))), 0) compras "
          If Semana > 0 Then
              sql = sql & "from (select * from " & DonamTaulaServit2(lunes, empresa) & " union all select * from " & DonamTaulaServit2(DateAdd("d", 1, lunes), empresa) & " union all select * from " & DonamTaulaServit2(DateAdd("d", 2, lunes), empresa) & " union all select * from " & DonamTaulaServit2(DateAdd("d", 3, lunes), empresa) & " union all select * from " & DonamTaulaServit2(DateAdd("d", 4, lunes), empresa) & " union all select * from " & DonamTaulaServit2(DateAdd("d", 5, lunes), empresa) & " union all select * from " & DonamTaulaServit2(DateAdd("d", 6, lunes), empresa) & ") s "
          Else
              sql = sql & "from " & DonamTaulaServit2(dia, empresa) & " s "
          End If
          sql = sql & "left join articles a on s.codiarticle = a.codi "
          sql = sql & "left join clients c on s.client=c.codi "
          sql = sql & "where client = " & codiBotiga & " and a.nom not like '***%'"
          Set Rs2 = Db.OpenResultset(sql)
          If Not Rs2.EOF Then
              D(6) = Rs2("compras")
          End If
     End If

    
On Error GoTo nor:
    'Horas
    Dim entra As Date, sale As Date
    Dim horas As Double, minutos As Double, horasPlan As Double
    Dim Entrat As Boolean, plegat As Boolean
    Dim dependenta As String, dependentaNom As String
    Dim rsHBase As rdoResultset, hBase As Double
    Dim turno As String, horaInicio As String, horaFin As String
    
    horas = 0
    horasPlan = 0
    minutos = 0
    Entrat = False
    
    strHoras = strHoras & "<br><b>" & UCase(D(0)) & " " & Format(dia, "dd/mm/yyyy") & "</b>:<br>"
    
    sql = "select h.accio, d.codi, d.nom, h.tmst, isnull(t.nombre, isnull(p.idTurno, 'NO ASIGNADO')) turno, ISNULL(t.horaInicio, '') horaInicio, isnull(t.horaFin, '') horaFin "
    sql = sql & "from cdpdadesfichador h "
    sql = sql & "left join " & taulaCdpPlanificacion(dia) & " p on h.tmst = p.fecha and h.lloc=p.botiga and h.usuari=p.idEmpleado "
    sql = sql & "left join cdpTurnos t on p.idTurno = t.idTurno "
    sql = sql & "left join dependentes d on h.usuari=d.codi "
    If Semana > 0 Then
        sql = sql & "where (lloc=" & codiBotiga & " or comentari like '%" & D(0) & "%') and tmst between '" & lunes & " 00:00' and '" & domingo & " 23:59' and d.codi is not null order by usuari, tmst"
    Else
        sql = sql & "where (lloc=" & codiBotiga & " or comentari like '%" & D(0) & "%') and day(tmst) = " & Day(dia) & " and month(tmst)=" & Month(dia) & " and year(tmst)= " & Year(dia) & " and d.codi is not null order by usuari, tmst"
    End If
    Set Rs2 = Db.OpenResultset(sql)
        
    If Not Rs2.EOF Then
        dependenta = Rs2("codi")
        dependentaNom = Rs2("nom")
    End If
    
    While Not Rs2.EOF
        If dependenta <> Rs2("codi") Then
            strHoras = strHoras & dependentaNom & "  "
            If minutos > 0 Then
                strHoras = strHoras & "(" & Int(minutos / 60) & ":" & Right("0" & (minutos Mod 60), 2) & " h)"
                strHoras = strHoras & "<B> " & turno & " (" & horaInicio & " a " & horaFin & ")</B>"
            Else
                strHoras = strHoras & "FICHAJE INCORRECTO"
            End If
            
            If Semana > 0 Then
                hBase = 0
                Set rsHBase = Db.OpenResultset("select valor from dependentesextes where nom like 'hBase' and id=" & dependenta)
                If Not rsHBase.EOF Then
                    If rsHBase("valor") <> "" And IsNumeric(rsHBase("valor")) Then hBase = rsHBase("valor")
                End If
                strHoras = strHoras & "  (Horas contratadas: " & hBase
                If hBase * 60 <> minutos Then
                    strHoras = strHoras & "   <font color='red'>" & Int((minutos - (hBase * 60)) / 60) & ":" & Right("0" & ((minutos - (hBase * 60)) Mod 60), 2) & " h </font>"
                End If
                strHoras = strHoras & ")"
            End If
            strHoras = strHoras & "<br>"
            
            dependenta = Rs2("codi")
            dependentaNom = Rs2("nom")
            minutos = 0
            Entrat = False
            turno = ""
            horaInicio = ""
            horaFin = ""
        End If
        
        If Rs2("Accio") = 1 Then
            If turno <> "" Then
                turno = turno & " + " & Rs2("turno")
            Else
                turno = Rs2("turno")
            End If
            If Rs2("horaInicio") <> "" And Rs2("horaFin") <> "" Then
                horaInicio = Rs2("horaInicio")
                horaFin = Rs2("horaFin")
                horasPlan = horasPlan + DateDiff("n", horaInicio, horaFin) / 60
            ElseIf InStr(Rs2("turno"), "_Extra") Then
                horasPlan = horasPlan + Split(Rs2("turno"), "_")(0)
            End If
            
            entra = Rs2("tmst")
            Entrat = True
        ElseIf Rs2("Accio") = "2" Then
            sale = Rs2("tmst")
            If Entrat Then
                If DateDiff("n", entra, sale) < 900 Then 'más de 15 horas es error
                    minutos = minutos + DateDiff("n", entra, sale)
                    horas = horas + DateDiff("n", entra, sale) / 60
                End If
                'strHoras = strHoras & Int(minutos / 60) & ":" & Right("0" & (minutos Mod 60), 2) & " h "
                
                Entrat = False
            End If
        End If
                
        Rs2.MoveNext
    Wend
    If dependentaNom <> "" Then
        strHoras = strHoras & dependentaNom & " "
        If minutos > 0 Then
            strHoras = strHoras & "(" & Int(minutos / 60) & ":" & Right("0" & (minutos Mod 60), 2) & " h)"
            strHoras = strHoras & "<B>  " & turno & " (" & horaInicio & " a " & horaFin & ")</B>"
        Else
            strHoras = strHoras & "FICHAJE INCORRECTO"
        End If
        If Semana > 0 Then
            hBase = 0
            Set rsHBase = Db.OpenResultset("select valor from dependentesextes where nom like 'hBase' and id=" & dependenta)
            If Not rsHBase.EOF Then
                If rsHBase("valor") <> "" And IsNumeric(rsHBase("valor")) Then hBase = rsHBase("valor")
            End If
            strHoras = strHoras & "  (Horas contratadas: " & hBase
            If hBase * 60 <> minutos Then
                strHoras = strHoras & "   <font color='red'>" & Int((minutos - (hBase * 60)) / 60) & ":" & Right("0" & ((minutos - (hBase * 60)) Mod 60), 2) & " h </font>"
            End If
            strHoras = strHoras & ")"
        End If
        strHoras = strHoras & "<br>"
        
    End If
    
    D(8) = horasPlan
    
    Exit Function
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR CarregaReportBotiga  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", sql & "  " & err.Description, "", ""
End Function


Sub calculaVendesClients(codiBot As Double, fecha As Date, vendes() As Double, clients() As Integer)
    Dim sql As String
    Dim rsCaixes As rdoResultset, Rs2 As rdoResultset, Rs3 As rdoResultset, RsAlbarans As rdoResultset
    Dim vMati As Double, vTarda As Double, vSinIVA As Double
    Dim dataInici As Date, dataFi As Date
    Dim cMati As Integer, cTarda As Integer
    
On Error GoTo nor:
        
    vMati = 0
    vTarda = 0
    cMati = 0
    cTarda = 0
    vSinIVA = 0
        
    sql = "select distinct data, tipus_moviment "
    sql = sql & "from [" & NomTaulaMovi(fecha) & "] where "
    sql = sql & "botiga='" & codiBot & "' and day(data)=" & Day(fecha) & " and (tipus_moviment='Wi' or tipus_moviment='W') order by data"
    Set rsCaixes = Db.OpenResultset(sql)
    
    If rsCaixes.EOF Then
        Set Rs2 = Db.OpenResultset("select isnull(sum(import), 0) I, isnull(sum(import/(1+(t.iva/100))), 0) sinIva, count(distinct Num_Tick ) Clients from (select * from [" & NomTaulaVentas(fecha) & "] union all select * from [" & NomTaulaAlbarans(fecha) & "]) v left join (select codi,tipoiva from articles union all select codi,tipoiva from articles_zombis) a on v.plu=a.codi left join TipusIva2012 t on a.TipoIva = t.tipus where botiga=" & codiBot & " and day(data) = " & Day(fecha))
        vMati = Rs2("I")
        vSinIVA = Rs2("sinIva")
        cMati = Rs2("clients")
    End If
    
    While Not rsCaixes.EOF
        If rsCaixes("tipus_moviment") = "Wi" Then
            dataInici = rsCaixes("data")
            rsCaixes.MoveNext
            
            If Not rsCaixes.EOF Then
                If rsCaixes("tipus_moviment") = "W" Then
                    dataFi = rsCaixes("data")
                    
                    Set Rs2 = Db.OpenResultset("select isnull(sum(import), 0) I, isnull(sum(import/(1+(t.iva/100))), 0) sinIva, count(distinct Num_Tick ) Clients from (select * from [" & NomTaulaVentas(fecha) & "] union all select * from [" & NomTaulaAlbarans(fecha) & "]) v left join (select codi,tipoiva from articles union all select codi,tipoiva from articles_zombis) a on v.plu=a.codi left join TipusIva2012 t on a.TipoIva = t.tipus where botiga=" & codiBot & " and data between '" & dataInici & "' and '" & dataFi & "'")
                    If Not Rs2.EOF Then
                        If Not IsNull(Rs2("I")) Then
                            If DatePart("h", dataInici) < 13 Then 'MATÍ
                                Set Rs3 = Db.OpenResultset("select import I from [" & NomTaulaMovi(fecha) & "] where botiga='" & codiBot & "' and tipus_moviment='Z' and data = '" & dataFi & "'")
                                If Not Rs3.EOF Then
                                    vMati = vMati + Rs3("I")
                                Else
                                    vMati = vMati + Rs2("I")
                                End If
                                
                                Set RsAlbarans = Db.OpenResultset("select isnull(sum(import), 0) I from [" & NomTaulaAlbarans(fecha) & "] v where botiga=" & codiBot & " and data between '" & dataInici & "' and '" & dataFi & "'")
                                If Not RsAlbarans.EOF Then
                                    vMati = vMati + RsAlbarans("I")
                                End If
                                
                                cMati = cMati + Rs2("Clients")
                            Else 'TARDA
                                Set Rs3 = Db.OpenResultset("select import I from [" & NomTaulaMovi(fecha) & "] where botiga='" & codiBot & "' and tipus_moviment='Z' and data = '" & dataFi & "'")
                                If Not Rs3.EOF Then
                                    vTarda = vTarda + Rs3("I")
                                Else
                                    vTarda = vTarda + Rs2("I")
                                End If
                                
                                Set RsAlbarans = Db.OpenResultset("select isnull(sum(import), 0) I from [" & NomTaulaAlbarans(fecha) & "] v where botiga=" & codiBot & " and data between '" & dataInici & "' and '" & dataFi & "'")
                                If Not RsAlbarans.EOF Then
                                    vTarda = vTarda + RsAlbarans("I")
                                End If
                                
                                cTarda = cTarda + Rs2("Clients")
                            End If
                            vSinIVA = vSinIVA + Rs2("sinIva")
                        End If
                    End If
                End If
            End If
        Else
            rsCaixes.MoveNext
        End If
    Wend
    
    'NO HAY CAJA, PERO PUEDE SER QUE HAYAN VENTAS
    'REPARTO LAS VENTAS EN MAÑANA DE LAS 5H A 14H Y TARDE DE 14H A 23H
    'CORONAVIRUS, .... ALGUNAS TIENDAS NO ABREN POR LAS TARDES
    'If vMati = 0 or vTarda = 0 Then
    If vMati = 0 And vTarda = 0 Then
        'Mati
        Set Rs2 = Db.OpenResultset("select isnull(sum(import), 0) I, isnull(sum(import/(1+(t.iva/100))), 0) sinIva, count(distinct Num_Tick ) Clients from (select * from [" & NomTaulaVentas(fecha) & "] union all select * from [" & NomTaulaAlbarans(fecha) & "]) v left join (select codi,tipoiva from articles union all select codi,tipoiva from articles_zombis) a on v.plu=a.codi left join TipusIva2012 t on a.TipoIva = t.tipus where botiga=" & codiBot & " and data between '" & Format(fecha, "dd/mm/yyyy") & " 05:00:00' and '" & Format(fecha, "dd/mm/yyyy") & " 13:59:59'")
        If Not Rs2.EOF Then
            vMati = Rs2("I")
            cMati = Rs2("Clients")
            vSinIVA = Rs2("sinIva")
        End If
        
        'Tarda
        Set Rs2 = Db.OpenResultset("select isnull(sum(import), 0) I, isnull(sum(import/(1+(t.iva/100))), 0) sinIva, count(distinct Num_Tick ) Clients from (select * from [" & NomTaulaVentas(fecha) & "] union all select * from [" & NomTaulaAlbarans(fecha) & "]) v left join (select codi,tipoiva from articles union all select codi,tipoiva from articles_zombis) a on v.plu=a.codi left join TipusIva2012 t on a.TipoIva = t.tipus where botiga=" & codiBot & " and data between '" & Format(fecha, "dd/mm/yyyy") & " 14:00:00' and '" & Format(fecha, "dd/mm/yyyy") & " 23:00:00'")
        If Not Rs2.EOF Then
            vTarda = Rs2("I")
            cTarda = Rs2("Clients")
            vSinIVA = vSinIVA + Rs2("sinIva")
        End If
    End If
    
    vendes(0) = vMati
    vendes(1) = vTarda
    vendes(2) = vSinIVA
    
    clients(0) = cMati
    clients(1) = cTarda
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR calculaVendes  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", sql & "  " & err.Description, "", ""
End Sub


Function calculaVendesSINDiada(codiBot As Double, fecha As Date) As Double
    Dim sql As String
    Dim rsVenut As rdoResultset
    Dim vendesSinIVASINDiada As Double
    
On Error GoTo nor:
        
    vendesSinIVASINDiada = 0
    
    sql = "select isnull(sum(import/(1+(t.iva/100))), 0) sinIva "
    sql = sql & "from (select * from [" & NomTaulaVentas(fecha) & "] union all select * from [" & NomTaulaAlbarans(fecha) & "]) v "
    sql = sql & "left join articles a on v.plu=a.codi "
    sql = sql & "left join families f3 on a.familia = f3.nom "
    sql = sql & "left join families f2 on f3.pare = f2.nom "
    sql = sql & "left join families f1 on f2.pare = f1.nom "
    sql = sql & "left join TipusIva2012 t on a.TipoIva = t.tipus "
    sql = sql & "where botiga=" & codiBot & " and day(data) = " & Day(fecha) & " and f1.nom not like '%diada%' and f2.nom not like '%diada%' and f3.nom not like '%diada%'"
    Set rsVenut = Db.OpenResultset(sql)
    
    vendesSinIVASINDiada = rsVenut("sinIva")
    
    calculaVendesSINDiada = vendesSinIVASINDiada
    
    Exit Function
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR calculaVendes  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", sql & "  " & err.Description, "", ""
End Function



Sub SecreInformeCoordinadora(codiBot As Double, nomBot As String, fecha As Date, depId As String, depNom As String, depEMail As String, Optional todas As Boolean, Optional conCopia As Boolean)
    Dim co As String
    Dim lunes As Date, f As Date, D As Integer
    Dim tiquetMig As Double, pctCompres As Double, totalHoras As Double
    
    Dim color As String, sql As String
    Dim rsVendes As rdoResultset
    Dim diaAnyoAnt As Date
    Dim clientesAnt As Integer
    Dim clientesTotal As Integer
    Dim clientesAntTotal As Integer
    
    Dim vendesArr(3) As Double, clientsArr(2) As Integer, previsionsArr(2) As Double
    Dim totalVendesMati As Double, totalVendesTarda As Double, totalVendes As Double, totalVendesDia As Double
    Dim totalClientsMati As Integer, totalClientsTarda As Integer, totalClientsDia As Integer
    
    Dim devDia As Double, devTotal As Double
    
    Dim horasPlan As Double
    Dim horasReales As Double
    Dim horasAcumulado As Double
    Dim objEurosHora As Double
    Dim pctPersonal As Double
    Dim horasPersonal As Double
    
    Dim totalPrevisionsDia As Double, difVendesAcumulat As Double, totalPrevisions As Double
    Dim compres As Double, compresSINDiada As Double
    Dim vendesSinIva As Double, vendesSinIVASINDiada As Double
    Dim totalVendesSinIVA As Double, totalVendesSinIVASINDiada As Double, totalCompres As Double, totalCompresSINDiada As Double
    
    Dim rsCoord As rdoResultset
    
    Dim rsCdA As rdoResultset, codigoAccion As String
    
    InformaMiss "SecreInformeCoordinadora", True
    
On Error GoTo nor

    pctCompres = getObjetivoCompras(codiBot)
    objEurosHora = getObjetivoEurosHora(codiBot)
    pctPersonal = getObjetivoPersonal(codiBot)

    lunes = fecha
    While DatePart("w", lunes) <> vbMonday
        lunes = DateAdd("d", -1, lunes)
    Wend
        
    Set rsCdA = Db.OpenResultset("select newid() id")
    codigoAccion = rsCdA("id")
    
    ExecutaComandaSql "Insert into  " & taulaCodigosDeAccion() & "  (IdCodigo, TipoCodigo, TmStmp, Param1, Param2) values ('" & codigoAccion & "', 'CUADRANTE', getdate(), '" & codiBot & "', '" & depEMail & "')"
    
    co = "<font size=1>CODIGO_ACCION:[" & codigoAccion & "]</font><br>"
    
    co = co & "<h3>BOTIGA " & nomBot & "</h3>"
        
    co = co & "<h4>" & depNom & "</h4>"
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'><Td rowspan='2'><b>Data</b></Td><Td align='center' colspan='2'><b>Vendes</b></Td><Td align='center' colspan='3'><b>Tiquet Mig</b></Td><Td align='center' colspan='3'><b>Clients</b></Td><Td align='center' colspan='2'><b>Compres</b></Td><Td align='center' colspan='2'><b>Devolucions</b></Td><Td colspan='2' align='center'><b>&euro;/Hora</b><br>(Objetivo " & objEurosHora & ")</Td></Tr>"
    'Set rsVendes = Db.OpenResultset("select * from [" & NomTaulaVentas(DateAdd("yyyy", -1, fecha)) & "] where botiga = " & codiBot)
    'If Not rsVendes.EOF Then
    '    Co = Co & "<Tr bgColor='#DADADA'><Td><b>Dif. dia</b></td><td><b>Dif. Acum.</b></Td><Td><b>Mati</b></td><td><b>Tarda</b></Td><td><b>Total</b></Td><td align='center'><b>" & Year(fecha) & "</b></td><td align='center'><b>" & Year(DateAdd("yyyy", -1, fecha)) & "</b></td><td><b>Dif.</b></td><td><b>Total</b></Td><td><b>% (" & pctCompres & "%)</b></Td><td><b>Total</b></Td><td><b>%</b></Td><Td align='center'><b>Diari</b></td><td align='center'><b>Acumulat</b></td></Tr>"
    'Else
        co = co & "<Tr bgColor='#DADADA'><Td><b>Dif. dia</b></td><td><b>Dif. Acum.</b></Td><Td><b>Mati</b></td><td><b>Tarda</b></Td><td><b>Total</b></Td><td align='center'><b>Set. actual</b></td><td align='center'><b>Set. anterior</b></td><td><b>Dif.</b></td><td><b>Total</b></Td><td><b>% (" & pctCompres & "%)</b></Td><td><b>Total</b></Td><td><b>%</b></Td><Td align='center'><b>Diari</b></td><td align='center'><b>Acumulat</b></td></Tr>"
    'End If
    
    D = 0
    For f = lunes To fecha
        diaAnyoAnt = DateAdd("yyyy", -1, f)
        While DatePart("w", f) <> DatePart("w", diaAnyoAnt)
            diaAnyoAnt = DateAdd("d", 1, diaAnyoAnt)
        Wend
        
        co = co & "<Tr>"
        
'VENDES ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        calculaVendesClients codiBot, f, vendesArr, clientsArr
        calculaPrevisions codiBot, f, previsionsArr
        
        vendesSinIva = vendesArr(2)
        If vendesArr(0) = 0 Or vendesArr(1) = 0 Then
            co = co & "<Td bgcolor='#FF0000'><b>" & Format(f, "dd/mm/yyyy") & "</b></Td>"
        Else
            co = co & "<Td><b>" & Format(f, "dd/mm/yyyy") & "</b></Td>"
        End If
        totalVendesDia = vendesArr(0) + vendesArr(1)
        totalClientsDia = clientsArr(0) + clientsArr(1)
        totalPrevisionsDia = previsionsArr(0) + previsionsArr(1)

        color = "#A6FFA6"
        If totalVendesDia < totalPrevisionsDia Then color = "#FFA8A8"
        co = co & "<Td align='right' bgcolor='" & color & "'>" & FormatNumber(Abs(totalPrevisionsDia - totalVendesDia), 2) & " &euro;</Td>"
        
        difVendesAcumulat = difVendesAcumulat + (totalVendesDia - totalPrevisionsDia)
        If difVendesAcumulat < 0 Then color = "#FFA8A8"
        co = co & "<Td align='right' bgcolor='" & color & "'>" & FormatNumber(Abs(difVendesAcumulat), 2) & " &euro;</Td>"
        
        totalPrevisions = totalPrevisions + totalPrevisionsDia
        
'TIQUET MEDIO -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        tiquetMig = getObjetivoTiquetMig(codiBot, f)
        
        'MATI
        If clientsArr(0) > 0 Then
            co = co & "<Td align='right'>" & FormatNumber(vendesArr(0) / clientsArr(0), 2) & " &euro;</Td>"
        Else
            co = co & "<Td align='right'>0.00 &euro;</Td>"
        End If
            
        'TARDA
        If clientsArr(1) > 0 Then
            co = co & "<Td align='right'>" & FormatNumber(vendesArr(1) / clientsArr(1), 2) & " &euro;</Td>"
        Else
            co = co & "<Td align='right'>0.00 &euro;</Td>"
        End If
            
        'TOTAL
        
        color = "#A6FFA6"
        If totalClientsDia > 0 Then
            If totalVendesDia / totalClientsDia < tiquetMig Then color = "#FFA8A8"
            co = co & "<Td align='right' bgcolor='" & color & "'>" & FormatNumber(totalVendesDia / totalClientsDia, 2) & " &euro;</Td>"
        Else
            co = co & "<Td  bgcolor='" & color & "' align='right'>0.00 &euro;</Td>"
        End If
            
        'TOTALES
        totalVendesMati = totalVendesMati + vendesArr(0)
        totalClientsMati = totalClientsMati + clientsArr(0)
        
        totalVendesTarda = totalVendesTarda + vendesArr(1)
        totalClientsTarda = totalClientsTarda + clientsArr(1)
        
        totalVendes = totalVendes + totalVendesDia
        
        clientesTotal = clientesTotal + totalClientsDia
              
'~TIQUET MEDIO -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'CLIENTES -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        co = co & "<Td align='right'>" & totalClientsDia & "</Td>"
         
        'Clients año anterior
        'clientesAnt = calculaNumeroClientes(codiBot, diaAnyoAnt)
        'If clientesAnt = 0 Then clientesAnt = calculaNumeroClientes(codiBot, DateAdd("d", -7, f))
        clientesAnt = calculaNumeroClientes(codiBot, DateAdd("d", -7, f))
        co = co & "<Td align='right'>" & clientesAnt & "</Td>"
         
        'Dif clients
        color = "#A6FFA6"
        If totalClientsDia - clientesAnt < 0 Then color = "#FFA8A8"
        co = co & "<Td bgcolor='" & color & "' align='right'>" & Abs(totalClientsDia - clientesAnt) & "</Td>"
        
        'TOTALES
        clientesAntTotal = clientesAntTotal + clientesAnt

'~CLIENTES -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         
'COMPRAS -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        compres = calculaCompras(codiBot, f)
        compresSINDiada = calculaCompras(codiBot, f, "Diada")

        vendesSinIVASINDiada = calculaVendesSINDiada(codiBot, f)
        
        color = "#A6FFA6"
        co = co & "<Td align='right'>" & FormatNumber(compres, 2) & " &euro;</Td>"
        
        If vendesSinIVASINDiada > 0 Then
            If (compresSINDiada / vendesSinIVASINDiada) * 100 > pctCompres Then color = "#FFA8A8"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber((compresSINDiada / vendesSinIVASINDiada) * 100, 2) & " %</Td>"
        Else
            If compresSINDiada = "0" Then
                co = co & "<Td align='right'>0.00 %</Td>"
            Else
                co = co & "<Td align='right'>100.00 %</Td>"
            End If
        End If
        totalCompres = totalCompres + compres
        totalCompresSINDiada = totalCompresSINDiada + compresSINDiada
        totalVendesSinIVA = totalVendesSinIVA + vendesSinIva
        totalVendesSinIVASINDiada = totalVendesSinIVASINDiada + vendesSinIVASINDiada
'~COMPRAS -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
         
'DEVOLUCIONES -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        devDia = calculaDevoluciones(codiBot, f)
        
        color = "#A6FFA6"
        If devDia > 160 Or devDia < 80 Then
            color = "#FFA8A8"
        End If
    
        co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(devDia, 2) & " &euro;</Td>"
        
        '% SOBRE VENTAS
        If totalVendesDia > 0 Then
            co = co & "<Td align='right'>" & FormatNumber((devDia / totalVendesDia) * 100, 2) & " %</Td>"
        Else
            If devDia = "0" Then
                co = co & "<Td align='right'>0.00 %</Td>"
            Else
                co = co & "<Td align='right'>100.00 %</Td>"
            End If
        End If
        devTotal = devTotal + devDia
'~DEVOLUCIONES -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'HORAS -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        'horasPlan = getObjetivoHoras(codiBot, f) 'HORAS PROGRAMADAS EN EL CUADRANTE
        horasReales = calculaHorasReales(codiBot, f) 'HORAS QUE SE HA HECHO REALMENTE (SIN CONTAR APRENDIZ NI COORDINACIÓN)
        
        'color = "#A6FFA6"
        'If horasReales <> horasPlan Then
        '    color = "#FFA8A8"
        'End If
        'Co = Co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(horasReales - horasPlan, 2) & "</Td>"
        
        'ACUMULADO
        'horasAcumulado = horasAcumulado + (horasReales - horasPlan)
        'color = "#A6FFA6"
        'If horasAcumulado <> 0 Then
        '    color = "#FFA8A8"
        'End If
        'Co = Co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(horasAcumulado, 2) & "</Td>"
        
        If horasReales > 0 Then
            color = "#A6FFA6"
            If (vendesSinIva / horasReales) < objEurosHora Then color = "#FFA8A8"
            If (vendesSinIva / horasReales) >= objEurosHora Then color = "#A6FFA6"
            If (vendesSinIva / horasReales) > 50 Then color = "#FFFFCA"

            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(vendesSinIva / horasReales, 2) & " &euro;</Td>"
        Else
            co = co & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
        End If
            
        'ACUMULADO
        horasAcumulado = horasAcumulado + horasReales
        If horasAcumulado > 0 Then
            color = "#A6FFA6"
            If (totalVendesSinIVA / horasAcumulado) < 38 Or (totalVendesSinIVA / horasAcumulado) > 45 Then color = "#FFA8A8"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(totalVendesSinIVA / horasAcumulado, 2) & " &euro;</Td>"
        Else
            co = co & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
        End If
        
'~HORAS -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        co = co & "</Tr>"
        D = D + 1
    Next
    
'TOTALES -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    If D > 0 Then
        co = co & "<tr>"
        co = co & "<td><b>Total</b></td>"
        
        'VENDES
        color = "#A6FFA6"
        If totalVendes < totalPrevisions Then color = "#FFA8A8"
        co = co & "<Td align='right' bgcolor='" & color & "'>" & FormatNumber(Abs(totalPrevisions - totalVendes), 2) & " &euro;</Td>"
        
        If difVendesAcumulat < 0 Then color = "#FFA8A8"
        co = co & "<Td align='right' bgcolor='" & color & "'>" & FormatNumber(Abs(difVendesAcumulat), 2) & " &euro;</Td>"
        
        'TIQUET MEDIO
        If totalClientsMati > 0 Then
            co = co & "<td align='right'><b>" & FormatNumber(totalVendesMati / totalClientsMati, 2) & " &euro;</b></td>"
        Else
            co = co & "<td align='right'><b>0.00 &euro;</b></td>"
        End If
        If totalClientsTarda > 0 Then
            co = co & "<td align='right'><b>" & FormatNumber(totalVendesTarda / totalClientsTarda, 2) & " &euro;</b></td>"
        Else
            co = co & "<td align='right'><b>0.00 &euro;</b></td>"
        End If
            
        color = "#A6FFA6"
        If clientesTotal > 0 Then
            If totalVendes / clientesTotal < tiquetMig Then color = "#FFA8A8"
            co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalVendes / clientesTotal, 2) & " &euro;</b></td>"
        Else
            color = "#FFA8A8"
            co = co & "<td bgcolor='" & color & "' align='right'><b>0.00 &euro;</b></td>"
        End If
        
        'CLIENTES
        co = co & "<td align='right'><b>" & clientesTotal & "</b></td><td align='right'><b>" & clientesAntTotal & "</b></td>"
        color = "#A6FFA6"
        If clientesTotal - clientesAntTotal < 0 Then color = "#FFA8A8"
        co = co & "<td bgcolor='" & color & "' align='right'><b>" & Abs(clientesTotal - clientesAntTotal) & "</b></td>"
        
        'COMPRAS
        co = co & "<td align='right'><b>" & FormatNumber(totalCompres, 2) & " &euro;</b></td>"
        color = "#A6FFA6"
        If totalVendesSinIVASINDiada > 0 Then
            If (totalCompresSINDiada / totalVendesSinIVASINDiada) * 100 > pctCompres Then color = "#FFA8A8"
            co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber((totalCompresSINDiada / totalVendesSinIVASINDiada) * 100, 2) & " %</b></td>"
        Else
            If totalCompresSINDiada = "0" Then
                co = co & "<Td align='right'><b>0.00 %</b></Td>"
            Else
                co = co & "<Td align='right'><b>100.00 %</b></Td>"
            End If
        End If
        
        'DEVOLUCIONES
        co = co & "<td align='right'><b>" & FormatNumber(devTotal, 2) & " &euro;</b></td>"
        If totalVendes > 0 Then
            co = co & "<td align='right'><b>" & FormatNumber((devTotal / totalVendes) * 100, 2) & " %</b></td>"
        Else
            If totalVendes = "0" Then
                co = co & "<Td align='right'><b>0.00 %</b></Td>"
            Else
                co = co & "<Td align='right'>100.00 %</Td>"
            End If
        End If
        
        'TOTAL HORAS: ÚLTIMO ACUMULADO
'        color = "#A6FFA6"
'        If horasAcumulado <> 0 Then color = "#FFA8A8"

'        Co = Co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(horasAcumulado, 2) & "</b></td>"
'        Co = Co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(horasAcumulado, 2) & "</b></td>"
        
        If horasAcumulado > 0 Then
            color = "#A6FFA6"

            If (totalVendesSinIVA / horasAcumulado) < objEurosHora Then color = "#FFA8A8"
            If (totalVendesSinIVA / horasAcumulado) >= objEurosHora Then color = "#A6FFA6"
            If (totalVendesSinIVA / horasAcumulado) > 50 Then color = "#FFFFCA"
            
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(totalVendesSinIVA / horasAcumulado, 2) & " &euro;</Td>"
            co = co & "<Td bgcolor='" & color & "' align='right'>" & FormatNumber(totalVendesSinIVA / horasAcumulado, 2) & " &euro;</Td>"
        Else
            co = co & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
            co = co & "<Td  bgcolor='#FFA8A8' align='right'>0 &euro;</Td>"
        End If
        
        co = co & "</tr>"
    End If

    co = co & "</table><br>"
    co = co & "<br>"
    'Co = Co & strHoras
    
    'Cuadrante V.2
    co = co & "<br><H3>Hores Pactades vs Reals</H3>"
    co = co & cuadrantePlanificacionTurnos2(fecha, codiBot, nomBot)
     
    'Previsiones semana siguiente
    co = co & "<br><H3>Previsió de vendes</H3>"
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    
    co = co & "<tr bgColor='#DADADA'><TD><B>Data</B></TD><TD><B>Mati</B></TD><TD><B>Tarda</B></TD><TD><B>Objectiu<br>Tiquet mig</B></TD><TD><B>Objectiu<br>Hores Personal</B></TD></TR>"
    For f = fecha To DateAdd("d", 6, fecha)
        calculaPrevisions codiBot, f, previsionsArr
        tiquetMig = getObjetivoTiquetMig(codiBot, f)
        
        horasPersonal = (previsionsArr(0) + previsionsArr(1)) / 1.1  'Le quitamos el 10% de IVA
        horasPersonal = horasPersonal * (pctPersonal / 100) 'Pct objetivo personal
        horasPersonal = horasPersonal / 12 'Precio hora
        
        co = co & "<tr>"
        co = co & "<td><b>" & Format(f, "dd/mm/yyyy") & "</b></td>"
        co = co & "<td align='right'>" & FormatNumber(previsionsArr(0), 2) & " &euro;</td>"
        co = co & "<td align='right'>" & FormatNumber(previsionsArr(1), 2) & " &euro;</td>"
        co = co & "<td align='right'>" & FormatNumber(tiquetMig, 2) & " &euro;</td>"
        co = co & "<td align='right'>" & FormatNumber(horasPersonal, 0) & " h</td>"
        co = co & "</tr>"
    Next

    co = co & "</table>"
    
    On Error GoTo sigue
    
    
    'Cuadrante turnos
    co = co & "<br><H3>Quadrant de torns</H3>"
    co = co & cuadrantePlanificacionTurnosEditable(codiBot)
    
sigue:
     
    sf_enviarMail "secrehit@hit.cat", depEMail, "Informe coordinadora " & nomBot & " [" & Format(fecha, "dd/mm/yyyy") & "]", co, "", ""
    
    If todas Then
        sql = "select distinct usuari, d2.valor email "
        sql = sql & "from cdpdadesfichador cdp "
        sql = sql & "left join dependentesextes d1 on cdp.usuari=d1.id and d1.nom='TIPUSTREBALLADOR' "
        sql = sql & "left join dependentesextes d2 on cdp.usuari=d2.id and d2.nom='EMAIL' "
        sql = sql & "where d1.valor='RESPONSABLE' and lloc='" & codiBot & "' and tmst between '" & DateAdd("d", -4, fecha) & "' and '" & fecha & "' and accio=1 "

        Set rsCoord = Db.OpenResultset(sql)
        While Not rsCoord.EOF
            sf_enviarMail "secrehit@hit.cat", rsCoord("email"), "Informe coordinadora " & nomBot & " [" & Format(fecha, "dd/mm/yyyy") & "]", co, "", ""
            rsCoord.MoveNext
        Wend
    End If

    'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "Informe coordinadora tienda " & nomBot & " [" & Format(fecha, "dd/mm/yyyy") & "]", Co, "", ""
    If conCopia Then sf_enviarMail "secrehit@hit.cat", "atena@silemabcn.com", "Informe coordinadora tienda " & nomBot & " [" & Format(fecha, "dd/mm/yyyy") & "]", co, "", ""
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreInformeCoordinadora [" & Format(fecha, "dd/mm/yyyy") & "]", "CODI BOTIGA [" & codiBot & "]  NOM BOTIGA [" & nomBot & "] FECHA [" & fecha & "]  CODI DEPENDENTA [" & depId & "]  NOM DEPENDENTA [" & depNom & "] EMAIL [" & depEMail & "]" & co & err.Description, "", ""
End Sub


Function cuadrantePlanificacionTurnos(dia As Date, botiga As Double, fichajes As Boolean) As String
    Dim co As String
    Dim fecha As Date, lunes As Date, dSemana As Integer
    Dim rsQ As rdoResultset, sql As String
    Dim color As String
            
    dSemana = Weekday(dia, 2)
    If fichajes Then
        If dSemana = 1 Then 'Si hoy es lunes, enviamos semana anterior completa
            lunes = DateAdd("d", -7, dia)
        Else
            lunes = DateAdd("d", -(dSemana - 1), dia)
        End If
    Else
        lunes = DateAdd("d", -(dSemana - 1), dia)
    End If
            
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'>"
    co = co & "<td>" & Format(lunes, "dd/mm/yyyy") & "</td>"
    co = co & "<td>" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "</td>"
    co = co & "<td>" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "</td>"
    co = co & "<td>" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "</td>"
    co = co & "<td>" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "</td>"
    co = co & "<td>" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "</td>"
    co = co & "<td>" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "</td>"
    co = co & "</tr>"
         
    co = co & "<tr>"
    fecha = CDate("01/01/1900")
    If fichajes Then
        sql = "select p.*, isnull(d.nom, '-') as dependenta, isnull(t.nombre, isnull(p.idTurno, 'NO CONFIGURADO')) nombre, isnull(t.horaInicio, '-') horaInicio, isnull(t.horaFin, '-') horaFin, isnull(t.color, '#ff0000') color  "
        sql = sql & "from " & taulaCdpPlanificacion(lunes) & " p "
        sql = sql & "left join cdpturnos t on p.idturno = t.idturno "
        sql = sql & "left join dependentes d on p.idEmpleado = d.codi "
        sql = sql & "Where p.botiga = " & botiga & " And p.activo = 1 "
        sql = sql & "order by convert(datetime, concat(day(fecha), '/', month(fecha),  '/', year(fecha)), 103), t.horaInicio"
    Else
        sql = "select p.*, isnull(t.nombre, isnull(p.idTurno, 'NO CONFIGURADO')) nombre, isnull(t.horaInicio, '-') horaInicio, isnull(t.horaFin, '-') horaFin, isnull(t.color, '#ff0000') color  "
        sql = sql & "from " & taulaCdpPlanificacion(lunes) & " p "
        sql = sql & "left join cdpturnos t on p.idturno = t.idturno "
        sql = sql & "Where p.botiga = " & botiga & " And p.activo = 1 and t.idTurno is not null "
        sql = sql & "order by convert(datetime, concat(day(fecha), '/', month(fecha),  '/', year(fecha)), 103), t.horaInicio"
    End If
    
    Set rsQ = Db.OpenResultset(sql)
    While Not rsQ.EOF
        If Day(fecha) <> Day(rsQ("fecha")) Or Month(fecha) <> Month(rsQ("fecha")) Or Year(fecha) <> Year(rsQ("fecha")) Then
            If Year(fecha) > 1900 Then
                co = co & "</table></td>"
            End If
            co = co & "<td valign='top'><Table cellpadding='0' cellspacing='3' border='1'>"
            fecha = rsQ("fecha")
        End If
        color = rsQ("color")
        If fichajes Then
            If rsQ("dependenta") = "-" Then color = "#ff0000"
        End If
       
        co = co & "<tr><td bgcolor='" & color & "'><b>" & rsQ("nombre")
        If rsQ("horaInicio") <> "-" And rsQ("horafin") <> "-" Then co = co & "<br>De " & rsQ("horaInicio") & " a " & rsQ("horaFin")
        If fichajes Then co = co & "<br>" & rsQ("dependenta") & "</b>"
        co = co & "</td></tr>"
        rsQ.MoveNext
    Wend
    
    co = co & "</table></td>"
    co = co & "</tr></table>"
    
    cuadrantePlanificacionTurnos = co
End Function

Function cuadrantePlanificacionTurnosEditable(botiga As Double) As String
    Dim co As String
    Dim fecha As Date
    Dim rsQ As rdoResultset, sql As String
    Dim f As Integer, diaSemana As String
    Dim i As Integer, t As Integer
    Dim tipoEmpleado As String
            
    fecha = DateAdd("d", 1, Now)
    
    co = co & "<H5>D: Dependenta  F: Forner</H5>"
   
    co = co & "<Table cellpadding=""0"" cellspacing=""3"" border=""1"" name=""Cuadrante"">"
    co = co & "<tr bgColor=""#DADADA"">"
    For f = 0 To 6
        diaSemana = Format(DateAdd("d", f, fecha), "dddd")
        co = co & "<td><b>" & UCase(Left(diaSemana, 1)) + Right(diaSemana, Len(diaSemana) - 1) & "<BR>" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "</b></td>"
    Next
    co = co & "</tr>"
         
    co = co & "<tr>"
    
    For f = 0 To 6
        i = 1
        
        co = co & "<td valign='top'>"
        
        co = co & "<Table cellpadding=""0"" cellspacing=""3"" border=""0"" name=""FECHA_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & """>"
        co = co & "<TR><TD>&nbsp;</TD><TD><B>Entra</B></TD><TD><B>Surt</B></TD></TR>"
        
        'DEPENDENTES
        sql = "select p.*, isnull(t.horaInicio, '-') horaInicio, isnull(t.horaFin, '-') horaFin, t.tipoEmpleado "
        sql = sql & "from " & taulaCdpPlanificacion(DateAdd("d", f, fecha)) & " p "
        sql = sql & "left join cdpturnos t on p.idturno = t.idturno "
        sql = sql & "Where p.botiga = " & botiga & " And p.activo = 1 And Day(fecha) = " & Day(DateAdd("d", f, fecha)) & " And t.idTurno Is Not Null and t.tipoEmpleado like '%DEPENDENTA%' "
        sql = sql & "order by t.horaInicio"
        Set rsQ = Db.OpenResultset(sql)
        While Not rsQ.EOF
            tipoEmpleado = "D"
            co = co & "<TR>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DAFCF8"" width=""20"" height=""4"" name=""TipoEmpleado_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & i & """>" & tipoEmpleado & "</TD>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DAFCF8"" width=""30"" height=""4"" name=""Entrada_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & i & """>" & rsQ("horaInicio") & "</TD>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DAFCF8"" width=""30"" height=""4"" name=""Salida_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & i & """>" & rsQ("horaFin") & "</TD>"
            co = co & "</TR>"
            
            i = i + 1
            rsQ.MoveNext
        Wend
        
        For t = i - 1 To 8
            co = co & "<TR>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DAFCF8"" width=""20"" height=""4"" name=""TipoEmpleado_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & t & """>D</TD>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DAFCF8"" width=""30"" height=""4"" name=""Entrada_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & t & """></TD>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DAFCF8"" width=""30"" height=""4"" name=""Salida_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & t & """></TD>"
            co = co & "</TR>"
        Next
        
        'FORNERS
        i = 9
        sql = "select p.*, isnull(t.horaInicio, '-') horaInicio, isnull(t.horaFin, '-') horaFin, t.tipoEmpleado "
        sql = sql & "from " & taulaCdpPlanificacion(DateAdd("d", f, fecha)) & " p "
        sql = sql & "left join cdpturnos t on p.idturno = t.idturno "
        sql = sql & "Where p.botiga = " & botiga & " And p.activo = 1 And Day(fecha) = " & Day(DateAdd("d", f, fecha)) & " And t.idTurno Is Not Null and t.tipoEmpleado like '%FORNER%' "
        sql = sql & "order by t.horaInicio"
        Set rsQ = Db.OpenResultset(sql)
        While Not rsQ.EOF
            tipoEmpleado = "F"
            co = co & "<TR>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DEFCDA"" width=""20"" height=""4"" name=""TipoEmpleado_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & i & """>" & tipoEmpleado & "</TD>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DEFCDA"" width=""30"" height=""4"" name=""Entrada_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & i & """>" & rsQ("horaInicio") & "</TD>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DEFCDA"" width=""30"" height=""4"" name=""Salida_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & i & """>" & rsQ("horaFin") & "</TD>"
            co = co & "</TR>"
            
            i = i + 1
            rsQ.MoveNext
        Wend
        
        For t = i - 1 To 9
            co = co & "<TR>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DEFCDA"" width=""20"" height=""4"" name=""TipoEmpleado_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & t & """>F</TD>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DEFCDA"" width=""30"" height=""4"" name=""Entrada_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & t & """></TD>"
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DEFCDA"" width=""30"" height=""4"" name=""Salida_" & Format(DateAdd("d", f, fecha), "dd/mm/yyyy") & "_" & t & """></TD>"
            co = co & "</TR>"
        Next
        
        co = co & "</table></td>"
    Next
    
    co = co & "</tr></table>"
    cuadrantePlanificacionTurnosEditable = co
End Function


Function cuadrantePlanificacionTurnos2(dia As Date, botiga As Double, nomBotiga As String) As String
    Dim co As String
    Dim fecha As Date, lunes As Date, dSemana As Integer
    Dim rsQ As rdoResultset, rsQ_A As rdoResultset, rsQ_C As rdoResultset, sql As String, sqlAprendiz As String, sqlCoord As String, rsReal As rdoResultset
    Dim sqlFechas As String, sqlPivot As String
    Dim totalL As Double, totalM As Double, totalX As Double, totalJ As Double, totalV As Double, totalS As Double, totalD As Double, totalContrato As Double
    Dim pactadoL As Double, pactadoM As Double, pactadoX As Double, pactadoJ As Double, pactadoV As Double, pactadoS As Double, pactadoD As Double, totalPactado As Double
    Dim rsDep As rdoResultset, hContrato As Double
    Dim color As String
    Dim entra As Date, sale As Date, horasReales As Double
    Dim rsValidado As rdoResultset, f As Integer, validado As Boolean
    Dim D As Integer, colorAprendiz As String, colorCoordinacion As String
            
    colorAprendiz = "#8000FF"
    colorCoordinacion = "#00A400"
    
    dSemana = Weekday(dia, 2)

    lunes = DateAdd("d", -(dSemana - 1), dia)
            
    sqlFechas = "isnull([" & Format(lunes, "dd/mm/yyyy") & "], 0) [" & Format(lunes, "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "]"
    
    sqlPivot = "[" & Format(lunes, "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "]"
    sql = "select nom, codi, " & sqlFechas & " "
    sql = sql & "From "
    sql = sql & "( "
    sql = sql & "select convert(nvarchar, fecha, 103) fecha, d.nom, d.codi, "
    sql = sql & "sum(case when p.idturno like '%Extra' then left(p.idturno, charindex('_', p.idturno)-1) "
    sql = sql & "else datediff(minute, t.horaInicio, t.horafin) / 60.00 end) horas "
    sql = sql & "from " & taulaCdpPlanificacion(lunes) & " p "
    sql = sql & "left join cdpturnos t on p.idturno=t.idTurno "
    sql = sql & "left join dependentes d on p.idEmpleado=d.codi "
    sql = sql & "Where P.activo = 1 And P.botiga = " & botiga & " And D.Codi Is Not Null "
    sql = sql & "group by convert(nvarchar, p.fecha, 103), d.nom, d.codi "
    sql = sql & ") DataTable "
    sql = sql & "PIVOT ( sum(horas) "
    sql = sql & "for fecha in (" & sqlPivot & ")) PivotTable "
    sql = sql & "order by nom"
    Set rsQ = Db.OpenResultset(sql)
            
    sqlAprendiz = "select nom, codi, " & sqlFechas & " "
    sqlAprendiz = sqlAprendiz & "From "
    sqlAprendiz = sqlAprendiz & "( "
    sqlAprendiz = sqlAprendiz & "select convert(nvarchar, fecha, 103) fecha, d.nom, d.codi, "
    sqlAprendiz = sqlAprendiz & "sum(case when p.idturno like '%Aprendiz' then left(p.idturno, charindex('_', p.idturno)-1) "
    sqlAprendiz = sqlAprendiz & "else 0 end) horas "
    sqlAprendiz = sqlAprendiz & "from " & taulaCdpPlanificacion(lunes) & " p "
    sqlAprendiz = sqlAprendiz & "left join cdpturnos t on p.idturno=t.idTurno "
    sqlAprendiz = sqlAprendiz & "left join dependentes d on p.idEmpleado=d.codi "
    sqlAprendiz = sqlAprendiz & "Where P.activo = 1 And P.botiga = " & botiga & " And D.Codi Is Not Null "
    sqlAprendiz = sqlAprendiz & "group by convert(nvarchar, p.fecha, 103), d.nom, d.codi "
    sqlAprendiz = sqlAprendiz & ") DataTable "
    sqlAprendiz = sqlAprendiz & "PIVOT ( sum(horas) "
    sqlAprendiz = sqlAprendiz & "for fecha in (" & sqlPivot & ")) PivotTable "
    sqlAprendiz = sqlAprendiz & "order by nom"
    Set rsQ_A = Db.OpenResultset(sqlAprendiz)
            
            
    sqlCoord = "select nom, codi, " & sqlFechas & " "
    sqlCoord = sqlCoord & "From "
    sqlCoord = sqlCoord & "( "
    sqlCoord = sqlCoord & "select convert(nvarchar, fecha, 103) fecha, d.nom, d.codi, "
    sqlCoord = sqlCoord & "sum(case when p.idturno like '%Coordinacion' then cast(left(p.idturno, charindex('_', p.idturno)-1) as float) "
    sqlCoord = sqlCoord & "else 0 end) horas "
    sqlCoord = sqlCoord & "from " & taulaCdpPlanificacion(lunes) & " p "
    sqlCoord = sqlCoord & "left join cdpturnos t on p.idturno=t.idTurno "
    sqlCoord = sqlCoord & "left join dependentes d on p.idEmpleado=d.codi "
    sqlCoord = sqlCoord & "Where P.activo = 1 And P.botiga = " & botiga & " And D.Codi Is Not Null "
    sqlCoord = sqlCoord & "group by convert(nvarchar, p.fecha, 103), d.nom, d.codi "
    sqlCoord = sqlCoord & ") DataTable "
    sqlCoord = sqlCoord & "PIVOT ( sum(horas) "
    sqlCoord = sqlCoord & "for fecha in (" & sqlPivot & ")) PivotTable "
    sqlCoord = sqlCoord & "order by nom"
    Set rsQ_C = Db.OpenResultset(sqlCoord)
            
            
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'>"
    co = co & "<td><b>" & UCase(nomBotiga) & "</b></td>"
    co = co & "<td><b>Contrato</b></td>"
    co = co & "<td><b>" & Format(lunes, "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>Total</b></td>"
    co = co & "</tr>"
    
    While Not rsQ.EOF
    
        hContrato = 0
        Set rsDep = Db.OpenResultset("select * from dependentesextes where id=" & rsQ("codi") & " and nom like '%dni%'")
        If Not rsDep.EOF Then
            Set rsDep = Db.OpenResultset("select top 1 40*(porjornada/100) jornadaHoras from silema_ts.sage.dbo.EmpleadoNomina en where en.dni='" & rsDep("valor") & "' order by en.fechaAlta desc")
            If Not rsDep.EOF Then
                hContrato = rsDep("JornadaHoras")
            End If
        End If
        
        co = co & "<Tr>"
        co = co & "<td><b>" & rsQ("nom") & "</b></td>"
        co = co & "<td align='right'>" & FormatNumber(hContrato, 2) & "</td>"
        
        For D = 0 To 6
        
            If rsQ(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")) <> 0 Then
                co = co & "<td align='right'>" & FormatNumber(rsQ(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")), 2)
                'Comprobar si tiene Aprendiz o Coordinación
                If rsQ_A(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")) <> 0 Then
                    co = co & "<font color=""" & colorAprendiz & """>(" & FormatNumber(rsQ_A(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")), 2) & ")</font>"
                End If
                If rsQ_C(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")) <> 0 Then
                    co = co & "<font color=""" & colorCoordinacion & """>(" & FormatNumber(rsQ_C(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")), 2) & ")"
                End If
                co = co & "</td>"
            Else
                If DateAdd("d", D, lunes) <= Now() Then
                    If rsQ_C(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")) = 0 And rsQ_A(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")) = 0 Then
                        horasReales = calculaHorasFichajeReal(DateAdd("d", D, lunes), botiga, rsQ("codi"))
                        If horasReales > 0 Then
                            co = co & "<td bgcolor='#FFFFCA' align='right'>" & FormatNumber(horasReales, 2) & "</td>"
                            Select Case D
                                Case 0: totalL = totalL + FormatNumber(horasReales, 2)
                                Case 1: totalM = totalM + FormatNumber(horasReales, 2)
                                Case 2: totalX = totalX + FormatNumber(horasReales, 2)
                                Case 3: totalJ = totalJ + FormatNumber(horasReales, 2)
                                Case 4: totalV = totalV + FormatNumber(horasReales, 2)
                                Case 5: totalS = totalS + FormatNumber(horasReales, 2)
                                Case 6: totalD = totalD + FormatNumber(horasReales, 2)
                            End Select
                        Else
                            co = co & "<td>&nbsp;</td>"
                        End If
                    Else
                        co = co & "<td align='right'>"
                        If rsQ_A(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")) <> 0 Then
                            co = co & "<font color=""" & colorAprendiz & """>(" & FormatNumber(rsQ_A(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")), 2) & ")</font>"
                        End If
                        If rsQ_C(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")) <> 0 Then
                            co = co & "<font color=""" & colorCoordinacion & """>(" & FormatNumber(rsQ_C(Format(DateAdd("d", D, lunes), "dd/mm/yyyy")), 2) & ")</font>"
                        End If
                        co = co & "</td>"
                    End If
                Else
                    co = co & "<td>&nbsp;</td>"
                End If
            End If
        Next
            
        co = co & "<td align='right'><b>" & FormatNumber(CInt(rsQ(Format(lunes, "dd/mm/yyyy"))) + CInt(rsQ(Format(DateAdd("d", 1, lunes), "dd/mm/yyyy"))) + CInt(rsQ(Format(DateAdd("d", 2, lunes), "dd/mm/yyyy"))) + CInt(rsQ(Format(DateAdd("d", 3, lunes), "dd/mm/yyyy"))) + CInt(rsQ(Format(DateAdd("d", 4, lunes), "dd/mm/yyyy"))) + CInt(rsQ(Format(DateAdd("d", 5, lunes), "dd/mm/yyyy"))) + CInt(rsQ(Format(DateAdd("d", 6, lunes), "dd/mm/yyyy"))), 2) & "</b></td>"
        co = co & "</tr>"
        
        totalL = totalL + rsQ(Format(lunes, "dd/mm/yyyy"))
        totalM = totalM + rsQ(Format(DateAdd("d", 1, lunes), "dd/mm/yyyy"))
        totalX = totalX + rsQ(Format(DateAdd("d", 2, lunes), "dd/mm/yyyy"))
        totalJ = totalJ + rsQ(Format(DateAdd("d", 3, lunes), "dd/mm/yyyy"))
        totalV = totalV + rsQ(Format(DateAdd("d", 4, lunes), "dd/mm/yyyy"))
        totalS = totalS + rsQ(Format(DateAdd("d", 5, lunes), "dd/mm/yyyy"))
        totalD = totalD + rsQ(Format(DateAdd("d", 6, lunes), "dd/mm/yyyy"))
        totalContrato = totalContrato + hContrato
        
        rsQ.MoveNext
        rsQ_A.MoveNext
        rsQ_C.MoveNext
    Wend
         
    co = co & "<Tr bgColor='#DADADA'>"
    co = co & "<td><b>TOTAL</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(totalContrato, 2) & "</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(totalL, 2) & "</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(totalM, 2) & "</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(totalX, 2) & "</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(totalJ, 2) & "</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(totalV, 2) & "</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(totalS, 2) & "</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(totalD, 2) & "</b></td>"
    co = co & "<td align='right'><b>" & FormatNumber(totalL + totalM + totalX + totalJ + totalV + totalS + totalD, 2) & "</b></td>"
    co = co & "</tr>"
         
    'PACTADAS
    sql = "select " & sqlFechas & " "
    sql = sql & "From ( "
    sql = sql & "select convert(nvarchar, fecha, 103) fecha, sum(datediff(minute, t.horaInicio, t.horafin) / 60.00) horas "
    sql = sql & "from " & taulaCdpPlanificacion(lunes) & " p "
    sql = sql & "left join cdpturnos t on p.idturno=t.idTurno "
    sql = sql & "Where P.activo = 1 And P.botiga = " & botiga & " And t.idturno Is Not Null "
    sql = sql & "group by convert(nvarchar, p.fecha, 103) ) DataTable "
    sql = sql & "PIVOT (sum(horas) for fecha in (" & sqlPivot & ")) PivotTable"
    Set rsQ = Db.OpenResultset(sql)
    If Not rsQ.EOF Then
        pactadoL = rsQ(Format(lunes, "dd/mm/yyyy"))
        pactadoM = rsQ(Format(DateAdd("d", 1, lunes), "dd/mm/yyyy"))
        pactadoX = rsQ(Format(DateAdd("d", 2, lunes), "dd/mm/yyyy"))
        pactadoJ = rsQ(Format(DateAdd("d", 3, lunes), "dd/mm/yyyy"))
        pactadoV = rsQ(Format(DateAdd("d", 4, lunes), "dd/mm/yyyy"))
        pactadoS = rsQ(Format(DateAdd("d", 5, lunes), "dd/mm/yyyy"))
        pactadoD = rsQ(Format(DateAdd("d", 6, lunes), "dd/mm/yyyy"))

        totalPactado = pactadoL + pactadoM + pactadoX + pactadoJ + pactadoV + pactadoS + pactadoD
        
        co = co & "<Tr bgColor='#DADADA'>"
        co = co & "<td><b>PACTADO</b></td>"
        co = co & "<td align='right'><b>" & FormatNumber(totalPactado, 2) & "</b></td>"
        co = co & "<td align='right'><b>" & FormatNumber(pactadoL, 2) & "</b></td>"
        co = co & "<td align='right'><b>" & FormatNumber(pactadoM, 2) & "</b></td>"
        co = co & "<td align='right'><b>" & FormatNumber(pactadoX, 2) & "</b></td>"
        co = co & "<td align='right'><b>" & FormatNumber(pactadoJ, 2) & "</b></td>"
        co = co & "<td align='right'><b>" & FormatNumber(pactadoV, 2) & "</b></td>"
        co = co & "<td align='right'><b>" & FormatNumber(pactadoS, 2) & "</b></td>"
        co = co & "<td align='right'><b>" & FormatNumber(pactadoD, 2) & "</b></td>"
        co = co & "<td align='right'><b>" & FormatNumber(totalPactado, 2) & "</b></td>"
        co = co & "</tr>"
    End If
    
    co = co & "<Tr>"
    co = co & "<td bgColor='#DADADA'><b>DIFERENCIA</b></td>"
    If totalContrato - totalPactado <> 0 Then color = "#FF0000" Else color = "#00FF00"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalContrato - totalPactado, 2) & "</b></td>"
    If totalL - pactadoL <> 0 Then color = "#FF0000" Else color = "#00FF00"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalL - pactadoL, 2) & "</b></td>"
    If totalM - pactadoM <> 0 Then color = "#FF0000" Else color = "#00FF00"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalM - pactadoM, 2) & "</b></td>"
    If totalX - pactadoX <> 0 Then color = "#FF0000" Else color = "#00FF00"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalX - pactadoX, 2) & "</b></td>"
    If totalJ - pactadoJ <> 0 Then color = "#FF0000" Else color = "#00FF00"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalJ - pactadoJ, 2) & "</b></td>"
    If totalV - pactadoV <> 0 Then color = "#FF0000" Else color = "#00FF00"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalV - pactadoV, 2) & "</b></td>"
    If totalS - pactadoS <> 0 Then color = "#FF0000" Else color = "#00FF00"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalS - pactadoS, 2) & "</b></td>"
    If totalD - pactadoD <> 0 Then color = "#FF0000" Else color = "#00FF00"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber(totalD - pactadoD, 2) & "</b></td>"
    If (totalL + totalM + totalX + totalJ + totalV + totalS + totalD) - totalPactado <> 0 Then color = "#FF0000" Else color = "#00FF00"
    co = co & "<td bgcolor='" & color & "' align='right'><b>" & FormatNumber((totalL + totalM + totalX + totalJ + totalV + totalS + totalD) - totalPactado, 2) & "</b></td>"
    co = co & "</tr>"

    
    co = co & "<Tr name=""VALIDACION_TURNOS"">"
    co = co & "<td bgColor=""#DADADA""><b>VALIDADO</b></td>"
    co = co & "<td bgColor=""#DADADA""><b>&nbsp;</b></td>"
    For f = 0 To 6
        validado = True
        
        sql = "select p.idplan, isnull(v.validado, 0) validado "
        sql = sql & "from [" & taulaCdpPlanificacion(DateAdd("d", f, lunes)) & "] p "
        sql = sql & "left join [" & taulaCdpValidacionHoras(DateAdd("d", f, lunes)) & "] v on p.idplan=v.idplan "
        sql = sql & "Where p.botiga = " & botiga & " And Day(p.fecha) = " & Day(DateAdd("d", f, lunes)) & " And p.activo = 1"
        Set rsValidado = Db.OpenResultset(sql)
        
        If rsValidado.EOF Then validado = False
        While Not rsValidado.EOF
            If Not rsValidado("validado") Then validado = False
            rsValidado.MoveNext
        Wend
        
        If validado Then
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DAFCF8"" width=""20"" height=""4"" name=""Validado_" & Format(DateAdd("d", f, lunes), "dd/mm/yyyy") & "_" & botiga & """>OK</TD>"
        Else
            co = co & "<TD style=""border: 1px solid black;"" bgcolor=""#DAFCF8"" width=""20"" height=""4"" name=""Validado_" & Format(DateAdd("d", f, lunes), "dd/mm/yyyy") & "_" & botiga & """></TD>"
        End If
        
    Next
    co = co & "<td bgColor=""#DADADA""><b>&nbsp;</b></td>"
    co = co & "</tr>"
    
    
    co = co & "</table></td>"
    co = co & "</tr></table>"
    
    co = co & "<h5>Para validar escribir OK en la casilla correspondiente</h5>"
    co = co & "<table border='0'>"
    co = co & "<tr><td width='10' bgcolor='#FFFFCA'>&nbsp;</td><td>Fichajes pendientes de asignar a un turno</td></tr>"
    co = co & "<tr><td colspan=""2""><FONT color=""" & colorCoordinacion & """>Horas de coordinación</FONT></td></tr>"
    co = co & "<tr><td colspan=""2""><FONT color=""" & colorAprendiz & """>Horas de aprendiz</FONT></td></tr>"
    co = co & "</table>"
    
    cuadrantePlanificacionTurnos2 = co
End Function



Function getObjetivoCompras(codiBot As Double) As Double
    Dim pctCompresObj As Double
    Dim rsParams As rdoResultset
    
    pctCompresObj = 40
    Set rsParams = Db.OpenResultset("select * from constantsClient where variable='Pct_Compras' and codi='" & codiBot & "'")
    If Not rsParams.EOF Then
        If IsNumeric(rsParams("valor")) Then
            If pctCompresObj > 0 Then pctCompresObj = rsParams("valor")
        End If
    End If
    
    getObjetivoCompras = pctCompresObj
End Function

Function getObjetivoEurosHora(codiBot As Double) As Double
    Dim eurosHora As Double
    Dim rsParams As rdoResultset
    
    eurosHora = 0
    Set rsParams = Db.OpenResultset("select * from constantsClient where variable='EurosHora' and codi='" & codiBot & "'")
    If Not rsParams.EOF Then
        If IsNumeric(rsParams("valor")) Then
            If rsParams("valor") > 0 Then eurosHora = rsParams("valor")
        End If
    End If
    
    getObjetivoEurosHora = eurosHora
End Function



Function getObjetivoPersonal(codiBot As Double) As Double
    Dim pctPersonal As Double
    Dim rsParams As rdoResultset
    
    pctPersonal = 0
    Set rsParams = Db.OpenResultset("select * from constantsClient where variable='Pct_Personal' and codi='" & codiBot & "'")
    If Not rsParams.EOF Then
        If IsNumeric(rsParams("valor")) Then
            If rsParams("valor") > 0 Then pctPersonal = rsParams("valor")
        End If
    End If
    
    getObjetivoPersonal = pctPersonal
End Function

Function getObjetivoMargenBruto(codiBot As Double) As Double
    Dim mrgBruto As Double
    Dim rsParams As rdoResultset
    
    mrgBruto = 0
    Set rsParams = Db.OpenResultset("select * from constantsClient where variable='MrgBruto' and codi='" & codiBot & "'")
    If Not rsParams.EOF Then
        If IsNumeric(rsParams("valor")) Then
            If rsParams("valor") > 0 Then mrgBruto = rsParams("valor")
        End If
    End If
    
    getObjetivoMargenBruto = mrgBruto
End Function

Sub SecreEmailCoordinadoras(Optional subj As String, Optional email As String)
    Dim rsCoord As rdoResultset, rsBotigues As rdoResultset, rsDep As rdoResultset, rsTipus As rdoResultset
    Dim fecha As Date, sql As String
    Dim depId As String, depNom As String, depEMail As String
    Dim Semana As Integer, semanaAux As Integer, lunes As Date
    Dim conCopia As Boolean, botEnviada As String
    
    InformaMiss "SecreEmailCoordinadoras", True

On Error GoTo ErrData
    If InStr(UCase(subj), "SEMANA") Then
        Semana = CInt(Split(UCase(subj), "SEMANA")(1))
        semanaAux = 0
        lunes = CDate("01/01/" & Year(Now()))

        If Semana > 0 Then
            While Semana <> semanaAux
                lunes = DateAdd("d", 1, lunes)
                semanaAux = DatePart("ww", lunes, vbMonday)
            Wend
    
            fecha = DateAdd("d", 6, lunes)
            
        End If
    Else
        fecha = DateAdd("d", -1, Now())
    End If
    GoTo OkData
    
ErrData:
    fecha = DateAdd("d", -1, Now())
        
OkData:

On Error GoTo nor
    conCopia = False
    
    'ALGUIEN LO HA PEDIDO
    If email <> "" Then
        Set rsCoord = Db.OpenResultset("select * from dependentesextes d1 left join dependentesextes d2 on d1.id=d2.id and d2.nom='EMAIL' where d1.nom='TIPUSTREBALLADOR' and d1.valor='RESPONSABLE' and d2.valor ='" & email & "'")
        'NO LO HA PEDIDO UNA COORDINADORA
        If rsCoord.EOF Then
            'SI ES GERENTE O GERENTE_2 LE ENVIAMOS TODAS LAS TIENDAS
            sql = "select * "
            sql = sql & "from dependentes d "
            sql = sql & "left join dependentesExtes d1 on d.codi=d1.id and d1.nom='TIPUSTREBALLADOR' "
            sql = sql & "left join dependentesExtes d2 on d.codi=d2.id and d2.nom='EMAIL' "
            sql = sql & "where d2.valor like '%" & email & "%' and d1.valor in ('GERENT', 'GERENT_2') "

            Set rsTipus = Db.OpenResultset(sql)
            If Not rsTipus.EOF Then 'Si es GERENTE le pasamos informe de todas las coordinadoras
                If UBound(Split(subj, " ")) >= 2 Then 'NOMÉS LA BOTIGA QUE HAN DEMANAT
                    Set rsBotigues = Db.OpenResultset("select c.Codi, c.nom from paramsHw w left join clients c on c.Codi = w.Valor1 where nom = '" & Split(subj, " ")(2) & "'")
                    If Not rsBotigues.EOF Then
                        If UBound(Split(subj, " ")) >= 3 Then
                            If UCase(Split(subj, " ")(3)) = "TODAS" Then
                                SecreInformeCoordinadora rsBotigues("codi"), rsBotigues("nom"), fecha, depId, depNom, email, True, conCopia
                            Else
                                If IsDate(Replace(Replace(Split(subj, " ")(3), "[", ""), "]", "")) Then
                                    SecreInformeCoordinadora rsBotigues("codi"), rsBotigues("nom"), CDate(Replace(Replace(Split(subj, " ")(3), "[", ""), "]", "")), depId, depNom, email, False, conCopia
                                End If
                            End If
                        Else
                            SecreInformeCoordinadora rsBotigues("codi"), rsBotigues("nom"), fecha, depId, depNom, email, False, conCopia
                        End If
                    End If
                Else 'TOTES LES BOTIGUES
                    Set rsBotigues = Db.OpenResultset("select c.Codi, c.nom from clients c join paramshw w on c.Codi = w.Valor1 where c.nif in (select valor from constantsempresa where camp like '%CampNif%' and isnull(valor, '')<>'') and c.codi is not null")
                    While Not rsBotigues.EOF
                        SecreInformeCoordinadora rsBotigues("codi"), rsBotigues("nom"), fecha, depId, depNom, email, False, conCopia
                        rsBotigues.MoveNext
                    Wend
                End If
            Else
                sf_enviarMail "secrehit@hit.cat", email, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "NO TENS PERMÍS PER REBRE AQUESTA INFORMACIÓ", "", ""
            End If
            
            Exit Sub
            
        'ESTO AHORA NO PASARÁ NUNCA PORQUE CUANDO RESPONDEN AL INFORME DE COORDINADORA SE INTERPRETA EL CUADRANTE DE TURNOS
        'Else 'LO HA PEDIDO UNA COORDINADORA (Re: Informe coordinadora tienda t--001 [21/08/19 10:34])
        '    If InStr(1, UCase(subj), UCase("Re: Informe coordinadora tienda")) Then
        '        Set rsBotigues = Db.OpenResultset("select * from clients where nom = '" & Split(subj, " ")(4) & "'")
        '        If Not rsBotigues.EOF Then
        '            fecha = CDate(Mid(Split(subj, " ")(5), 2, 10))
        '            SecreInformeCoordinadora rsBotigues("codi"), rsBotigues("nom"), fecha, depId, depNom, email
        '        End If
        '    End If
        '    Exit Sub
        End If
        
    Else 'ENVÍO AUTOMÁTICO PROGRAMADO PARA TODAS LAS COORDINADORAS
        If Weekday(Now(), 2) = 1 Then conCopia = True
        Set rsCoord = Db.OpenResultset("select * from dependentesextes where nom='TIPUSTREBALLADOR' and valor='RESPONSABLE'")
    End If
    
    botEnviada = ""
    While Not rsCoord.EOF
        depId = rsCoord("id")
        Set rsDep = Db.OpenResultset("select d.*, isnull(de.valor, '') eMail from dependentes d left join dependentesExtes de on d.codi=de.id and de.nom='EMAIL' Where D.Codi = " & depId)
        If Not rsDep.EOF Then
            depNom = rsDep("Nom")
            depEMail = rsDep("EMail")
        End If
        If depEMail <> "" Then 'Sin no hay email ya no seguimos
            Set rsBotigues = Db.OpenResultset("select distinct isnull(lloc, '') codi, isnull(c.nom, '') nom from cdpdadesfichador cdp left join clients c on cdp.lloc=c.codi where usuari=" & depId & " and tmst between '" & DateAdd("d", -4, fecha) & "' and '" & fecha & "' and accio=1")
            While Not rsBotigues.EOF  'Para cada tienda, donde ha fichado la coordinadora, generamos un informe
                If rsBotigues("codi") <> "" Then
                    If InStr(botEnviada, "[" & rsBotigues("codi") & "]") Then 'Puede ser que se envíe el informe de la misma tienda a varias coordinadoras. En ese caso no volvemos a enviar la copia.
                        SecreInformeCoordinadora rsBotigues("codi"), rsBotigues("nom"), fecha, depId, depNom, depEMail, False, False
                    Else
                        SecreInformeCoordinadora rsBotigues("codi"), rsBotigues("nom"), fecha, depId, depNom, depEMail, False, conCopia
                    End If
                    botEnviada = botEnviada & "[" & rsBotigues("codi") & "]"
                End If
                rsBotigues.MoveNext
            Wend
        End If
        rsCoord.MoveNext
    Wend
    Exit Sub
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreEmailCoordinadoras [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "LO PIDE: [" & email & "]<BR>" & err.Description, "", ""
End Sub





Sub SecreInformeProductosCompraVenta(subj As String, emailDe As String, empresa As String)
    Dim rsArticles As rdoResultset, rsFormula As rdoResultset
    Dim co As String, sql As String
    
    InformaMiss "SecreInformeProductosCompraVenta", True
    
    On Error GoTo nor
     
    co = ""
    co = co & "<H3>INFORME PRODUCTOS</H3>"
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'><Td><b>Producte Venda</b></Td><Td><b>Producte Compra</b></Td></Tr>"

    
    sql = "select a.nom, case when isnull(ap2.valor, '') = '' then isnull(mp.nombre, '')  else isnull(mpP.nombre, '') end MP "
    sql = sql & "from articles a "
    sql = sql & "left join articlesPropietats ap1 on a.codi=ap1.codiarticle and ap1.variable='MatPri' "
    sql = sql & "left join articlesPropietats ap2 on a.codi=ap2.codiarticle and ap2.variable='MatBase' "
    'Sql = Sql & "left join articlesPropietats ap11 on a.codi=ap11.codiarticle and ap11.variable='Formula' "
    sql = sql & "left join ccMateriasPrimas mp on ap1.valor=mp.id "
    sql = sql & "left join ccMateriasPrimasBase mpB on ap2.valor=mpB.id "
    sql = sql & "left join ccMateriasPrimas mpP on mpB.predeterminada=mpP.id "
    sql = sql & "where case when isnull(ap2.valor, '') = '' then isnull(mp.nombre, '')  else isnull(mpP.nombre, '') end <> '' "
    sql = sql & "order by a.nom"
    Set rsArticles = Db.OpenResultset(sql)
    
    While Not rsArticles.EOF
        co = co & "<tr><td>" & rsArticles("nom") & "</td><Td>" & rsArticles("MP") & "</Td></tr>"
        rsArticles.MoveNext
    Wend
    rsArticles.Close
    
    sql = "select a.nom, case when isnull(ap2.valor, '') = '' then isnull(mp.nombre, '')  else isnull(mpP.nombre, '') end MP "
    sql = sql & "from articles a "
    sql = sql & "left join articlesPropietats ap1 on a.codi=ap1.codiarticle and ap1.variable='MatPri' "
    sql = sql & "left join articlesPropietats ap2 on a.codi=ap2.codiarticle and ap2.variable='MatBase' "
    'Sql = Sql & "left join articlesPropietats ap3 on a.codi=ap3.codiarticle and ap3.variable='Formula' "
    sql = sql & "left join ccMateriasPrimas mp on ap1.valor=mp.id "
    sql = sql & "left join ccMateriasPrimasBase mpB on ap2.valor=mpB.id "
    sql = sql & "left join ccMateriasPrimas mpP on mpB.predeterminada=mpP.id "
    sql = sql & "where case when isnull(ap2.valor, '') = '' then isnull(mp.nombre, '')  else isnull(mpP.nombre, '') end = '' "
    sql = sql & "order by a.nom"
    Set rsArticles = Db.OpenResultset(sql)
    
    While Not rsArticles.EOF
        co = co & "<tr><td>" & rsArticles("nom") & "</td><Td>" & rsArticles("MP") & "</Td></tr>"
        rsArticles.MoveNext
    Wend
    rsArticles.Close
    
    co = co & "</table>"
    

    sf_enviarMail "secrehit@hit.cat", emailDe, UCase(subj) & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", UCase(subj) & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""

    Exit Sub
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "ERROR SecreInformeProductosCompraVenta " & err.Description, "", """"
End Sub




Sub SecreInformeInstalacion(subj As String, emailDe As String, empresa As String)
    Dim co As String
    Dim botiga As String, botigaCodi As String
    Dim codigoAccion As String
    Dim rsBotiga As rdoResultset, rsFacturas As rdoResultset, rsCdA As rdoResultset
    Dim anyo As Integer, mes As Integer, facTabData As String
    Dim f As Date
    
    InformaMiss "SecreInformeInstalacion", True

On Error GoTo nor

    botiga = Split(subj, " ")(1)
    Set rsBotiga = Db.OpenResultset("Select * from clients where nom like '%" & botiga & "%'")
    If Not rsBotiga.EOF Then
        botigaCodi = rsBotiga("codi")
        
        Set rsCdA = Db.OpenResultset("select newid() id")
        codigoAccion = rsCdA("id")
        ExecutaComandaSql "Insert into " & taulaCodigosDeAccion() & " (IdCodigo, TipoCodigo, TmStmp, Param1) values ('" & codigoAccion & "', 'INSTALACION', getdate(), '" & botigaCodi & "')"
        
        co = "<font size=1>CODIGO_ACCION:[" & codigoAccion & "]</font><br>"
        
        co = co & "<h3>INSTALACIÓN " & botiga & "</h3>"
        co = co & "<DD><TABLE name='Instalacion' BORDER='1'>"
        
        For anyo = Year(DateAdd("y", -3, Now())) To Year(Now())
            For mes = 1 To 12
                facTabData = "[facturacio_" & anyo & "-" & Right("00" & mes, 2) & "_Data]"
                If ExisteixTaula(facTabData) Then
                    Set rsFacturas = Db.OpenResultset("select * from " & facTabData & " where client = '" & botigaCodi & "'")
                    While Not rsFacturas.EOF
                        co = co & "<TR><TD>" & rsFacturas("data") & "</TD><TD>" & rsFacturas("producteNom") & "</TD><TD>" & rsFacturas("referencia") & "</TD><TD style='border: 1px solid black;' width='30'></TD></TR>"
                        rsFacturas.MoveNext
                    Wend
                    rsFacturas.Close
                End If
            Next
        Next
        
        co = co & "</TABLE>"
    Else
        co = "<h3>NO EXISTE NINGÚN CLIENTE CON NOMBRE PARECIDO A " & botiga & "</h3>"
    End If
    rsBotiga.Close
    
    sf_enviarMail "secrehit@hit.cat", emailDe, UCase(subj) & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", UCase(subj) & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreInformeInstalacion [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "PETICIÓN [" & subj & "] DE [" & emailDe & "] EMPRESA [" & empresa & "]<br>" & co & "<br>ERROR:" & err.Description, "", """"
End Sub


Sub SolicitudAutorizacionPago(idFactura As String, dataFactura As String, emailDe As String)
    Dim co As String
    Dim rsFactura As rdoResultset
    Dim f As Date
    Dim rsCdA As rdoResultset, codigoAccion As String
    
On Error GoTo nor

    f = CDate(dataFactura)
    
    'CÓDIGO DE ACCIÓN (PARA CUANDO RESPONDAN)
    Set rsCdA = Db.OpenResultset("select newid() id")
    codigoAccion = rsCdA("id")
    ExecutaComandaSql "Insert into " & taulaCodigosDeAccion() & " (IdCodigo, TipoCodigo, TmStmp, Param1, Param2) values ('" & codigoAccion & "', 'AUTORIZACION', getdate(), '" & idFactura & "', '" & dataFactura & "')"
        
    co = "<font size=1>CODIGO_ACCION:[" & codigoAccion & "]</font><br>"
        
    Set rsFactura = Db.OpenResultset("select * from ccFacturas_" & Year(f) & "_iva where idFactura='" & idFactura & "'")
    If Not rsFactura.EOF Then
        co = co & "<h3>SOLICITUD AUTORIZACION DE PAGO DE LA FACTURA " & rsFactura("numFactura") & "</h3>"
        co = co & "<DD><TABLE name='Factura' BORDER='0'>"
        
        co = co & "<TR><TD><B>Empresa </B></TD><TD>" & rsFactura("ClientNom") & "</TD></TR>"
        co = co & "<TR><TD><B>Proveedor </B></TD><TD>" & rsFactura("EmpNom") & "</TD></TR>"
        co = co & "<TR><TD><B>Fecha </B></TD><TD>" & rsFactura("dataFactura") & "</TD></TR>"
        co = co & "<TR><TD><B>Importe </B></TD><TD>" & rsFactura("total") & " &euro;</TD></TR>"
        co = co & "<TR><TD>&nbsp;</TD><TD>&nbsp;</TD></TR>"
        
        co = co & "</TABLE>"
        
        co = co & "<DD><TABLE name='Autorizacion' BORDER='0'>"
        co = co & "<TR><TD><B>Autorizar el pago?</B></TD><TD style='border: 1px solid black;' name='FACTURA_AUTORIZADA' width='30'></TD></TR>"
        co = co & "</TABLE>"
    Else
        co = "<h3>NO EXISTE LA FACTURA</h3>"
    End If
    
    sf_enviarMail "secrehit@hit.cat", emailDe, "SOLICITUD AUTORIZACION PAGO FACTURA [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SolicitudAutorizacionPago [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "ID FACTURA [" & idFactura & "] DATA FACTURA [" & dataFactura & "] <br>" & co & "<br>ERROR:" & err.Description, "", """"
End Sub



Sub SecreEnviaEmailIncidencia(idIncidencia As String, remitente As String, emailPara As String, items As String, fechaHistorico As String)
    Dim co As String, sql As String, cuerpo As String
    Dim rs As rdoResultset, rsIncidencia As rdoResultset, rsEmpresa As rdoResultset, rsEmpresaBD As rdoResultset, rsCuerpo As rdoResultset, rsLinkOtros As rdoResultset, rsEmp As rdoResultset
    Dim rsCli As rdoResultset, rsTecnico As rdoResultset, nomNouCli As String, rsNewid As rdoResultset
    Dim empCodi As String, empNom As String, empAdresa As String, empTel As String, empCp As String, empCiutat As String, empProvincia As String
    Dim rsCdA As rdoResultset, codigoAccion As String
    Dim Param1 As String, Param2 As String, Param3 As String, Param4 As String, Param5 As String, Param6 As String
    Dim emailList() As String, e As Integer
    Dim hayLink As Boolean
    Dim enviada As Boolean
    
On Error GoTo nor

    hayLink = False
    
    'CÓDIGO DE ACCIÓN (PARA CUANDO RESPONDAN)
    Set rsCdA = Db.OpenResultset("select newid() id")
    codigoAccion = rsCdA("id")
    
    Param1 = ""
    sql = "select Db from hit.dbo.web_empreses where nom ='" & EmpresaActual & "'"
    Set rsEmpresaBD = Db.OpenResultset(sql)
    If Not rsEmpresaBD.EOF Then Param1 = rsEmpresaBD("db")      'BD empresa que envía la incidencia
    Param2 = idIncidencia 'Id incidencia modificada
    
    Param3 = ""           'Id Incidencia correspondiente a la otra empresa
    Param4 = ""           'BD empresa que recibe la incidencia
    sql = "select * from Inc_Link_Otros where id=" & idIncidencia
    Set rsLinkOtros = Db.OpenResultset(sql)
    If Not rsLinkOtros.EOF Then
        Param3 = rsLinkOtros("IdOtro")
        Param4 = rsLinkOtros("Empresa")
        hayLink = True
    End If
    
    If Param4 = "" Then  'Es una incidencia nueva o no está unida a otra empresa
        sql = "select isnull(d.valor, '') NIF from incidencias i left join dependentesExtes d on i.tecnico=d.id and d.nom='DNI' where i.id=" & idIncidencia
        Set rsIncidencia = Db.OpenResultset(sql)
        If Not rsIncidencia.EOF Then
            If rsIncidencia("NIF") <> "" Then
                'BUSCAMOS POR TODAS LAS EMPRESAS EN NIF DE LA EMPRESA DESTINO
                sql = "select * from sys.databases where name like 'Fac_%' and name not like '%bak%'"
                Set rs = Db.OpenResultset(sql)
                While Not rs.EOF And Param4 = ""
                    If ExisteixTaula(rs("name") & ".dbo.constantsEmpresa") Then
                        sql = "select * from [WEB]." & rs("name") & ".dbo.constantsEmpresa where camp like '%nif%' and upper(valor) = upper('" & rsIncidencia("NIF") & "')"
                        Set rsEmp = Db.OpenResultset(sql)
                        If Not rsEmp.EOF Then
                            Param4 = rs("name")
                        End If
                    End If
                    rs.MoveNext
                Wend
            End If
        End If
    End If
    
        
    co = "<font size=1>CODIGO_ACCION:[" & codigoAccion & "]</font><br>"
        
    sql = "select i.Timestamp, i.incidencia, i.cliente, isnull(c.Nom, isnull(Icli.nom, '')) nom, isnull(c.adresa,'') adresa, isnull(c.ciutat,'') ciutat, isnull(c.cp,'') cp, isnull(cc.valor, '') as tel, isnull(d.nom,'') responsable, isnull(d2.nom,'') creador "
    sql = sql & "from incidencias i "
    sql = sql & "left join clients c on i.cliente = cast(c.codi as nvarchar) "
    sql = sql & "LEFT JOIN Inc_Clientes Icli ON cast(Icli.Id as nvarchar)= cast(i.Cliente as nvarchar) "
    sql = sql & "left join dependentes d on i.tecnico = d.codi "
    sql = sql & "left join dependentes d2 on i.usuario = d2.codi "
    sql = sql & "left join constantsclient cc on c.codi= cc.codi and cc.variable='Tel' "
    sql = sql & "Where i.iD = " & idIncidencia
    Set rsIncidencia = Db.OpenResultset(sql, rdConcurRowVer)
    
    If Not rsIncidencia.EOF Then
        If Param4 <> "" And Not hayLink Then  'ES UNA INCIDENCIA NUEVA PARA OTRA EMPRESA
            'Cliente correspondiente en la otra empresa
            Param5 = ""
            sql = "select codi from " & Param4 & ".dbo.constantsclient where variable = 'OrdreRuta' and valor='" & rsIncidencia("cliente") & "'"
            Set rsCli = Db.OpenResultset(sql)
            If Not rsCli.EOF Then
                Param5 = rsCli("codi")
            Else
                'Si no hay lo creamos
                If Len(rsIncidencia("cliente")) < 10 Then
                    sql = "select * from clients where codi='" & rsIncidencia("cliente") & "'"
                    Set rs = Db.OpenResultset(sql)
                    If Not rs.EOF Then
                        nomNouCli = rs("nom")
                    End If
                Else
                    sql = "select * from Inc_clientes where id='" & rsIncidencia("cliente") & "'"
                    Set rs = Db.OpenResultset(sql)
                    If Not rs.EOF Then
                        nomNouCli = rs("nom")
                    End If
                End If
                
                sql = "select * from " & Param4 & ".dbo.Inc_Clientes where nom = '" & nomNouCli & "'"
                Set rsNewid = Db.OpenResultset(sql)
                If Not rsNewid.EOF Then
                    Param5 = rsNewid("id")
                Else
                    sql = "select top 1 newid() id from " & Param4 & ".dbo.Inc_clientes"
                    Set rsNewid = Db.OpenResultset(sql)
                    If Not rsNewid.EOF Then Param5 = rsNewid("id")
                    sql = "INSERT INTO " & Param4 & ".dbo.Inc_Clientes (Id, Nom, Empresa, CliEmpresa) Values ('" & Param5 & "', '" & nomNouCli & "', '" & Param1 & "', '" & rsIncidencia("cliente") & "')"
                    ExecutaComandaSql sql
                End If
            End If
            
            'Tecnico(responsable) de la otra empresa
            Param6 = ""
            sql = "select top 1 id from " & Param4 & ".dbo.dependentesextes where nom = 'equips' and valor like '%TECNICOS%'"
            Set rsTecnico = Db.OpenResultset(sql)
            If Not rsTecnico.EOF Then Param6 = rsTecnico("id")
            
            If Param6 = "" Then 'Creamos un usuario para que reciba la incidencia
                Param6 = "1"
                Set rsTecnico = Db.OpenResultset("select max(c) + 1 codi from (select max(codi) c from " & Param4 & ".dbo.dependentes union select max(codi) c from " & Param4 & ".dbo.dependentes_zombis) k")
                If Not rsTecnico.EOF Then Param6 = rsTecnico("codi")

                ExecutaComandaSql "insert into " & Param4 & ".dbo.dependentesExtes (id, nom, valor) values(" & Param6 & ", 'equips', 'TECNICOS')"
                ExecutaComandaSql "insert into " & Param4 & ".dbo.dependentes (codi, nom, memo, [Hi Editem Horaris]) values(" & Param6 & ", '" & rsIncidencia("responsable") & "', '" & rsIncidencia("responsable") & "', 1)"
            End If
        End If

        sql = "SELECT convert(varchar, TimeStamp, 120) as fecha, incidencia, usuario FROM Inc_Historico where id=" & idIncidencia & " and tipo='TEXTO' ORDER BY timestamp"
        Set rsCuerpo = Db.OpenResultset(sql, rdConcurRowVer)
        cuerpo = "<table border=""0"" cellspacing=""0"" cellpadding=""0"">"
        While Not rsCuerpo.EOF
            cuerpo = cuerpo & "<TR><TD colspan=""2""><B>" & rsCuerpo("usuario") & " (" & Format(rsCuerpo("fecha"), "dd/mm/yyyy hh:nn:ss") & ")</B></TD></TR>"
            cuerpo = cuerpo & "<TR><TD>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD>" & rsCuerpo("incidencia") & "</TD></TR>"
        
            rsCuerpo.MoveNext
        Wend
        cuerpo = cuerpo & "</table>"
    
        co = co & "<table width=""95%"" style=""width:100%;border:0;background-color:#3399ff"" border=""0"" bgcolor=""#3399ff;"" align=""center"" valign=""center"">"

        empCodi = ""
        Set rsEmpresa = Db.OpenResultset("select * from constantsempresa where camp = 'Predeterminada'")
        If Not rsEmpresa.EOF Then
            If InStr(rsEmpresa("valor"), "_") Then
                empCodi = Split(rsEmpresa("valor"), "_")(0)
            End If
        End If
        
        If empCodi = "" Then
            Set rsEmpresa = Db.OpenResultset("select * from constantsempresa where camp like 'Camp%'")
        Else
            Set rsEmpresa = Db.OpenResultset("select * from constantsempresa where camp like '" & empCodi & "_Camp%'")
        End If
        While Not rsEmpresa.EOF
            If InStr(rsEmpresa("camp"), "CampNom") Then
                empNom = rsEmpresa("valor")
            ElseIf InStr(rsEmpresa("camp"), "CampAdresa") Then
                empAdresa = rsEmpresa("valor")
            ElseIf InStr(rsEmpresa("camp"), "CampTel") Then
                empTel = rsEmpresa("valor")
            ElseIf InStr(rsEmpresa("camp"), "CampCodiPostal") Then
                empCp = rsEmpresa("valor")
            ElseIf InStr(rsEmpresa("camp"), "CampCiutat") Then
                empCiutat = rsEmpresa("valor")
            ElseIf InStr(rsEmpresa("camp"), "CampProvincia") Then
                empProvincia = rsEmpresa("valor")
            End If
            rsEmpresa.MoveNext
        Wend
        
        co = co & "<tr><td align=""center""><b><FONT SIZE=""4"">" & UCase(empNom) & "</b><br>" & empTel & " - " & empAdresa & " - " & empCp & " " & empCiutat & " " & empProvincia & "</FONT></td></tr>"
        co = co & "</table>"
        
        co = co & "<H3>INCIDENCIA: " & idIncidencia & " (" & rsIncidencia("Timestamp") & ")</H3>"
        
        'Co = Co & "<img style=""float:left;margin:0;padding:0;border:0;margin-top:20px"" hspace=""0"" vspace=""0"" src=""http://silema.hiterp.com/admin/imagenes/logo.png"" border=""0"" alt=""hitsystems"" class=""CToWUd"">"
        
        co = co & "<H4>CLIENTE</H4>"
        co = co & "<DD><p>" & rsIncidencia("nom") & " - " & rsIncidencia("adresa") & " " & rsIncidencia("ciutat") & " " & rsIncidencia("cp") & " <br> Teléfono: " & rsIncidencia("tel") & "</p></DD>"
        
        co = co & "<H4>PARA</H4>"
        co = co & "<DD><p>" & rsIncidencia("responsable") & "</p></DD>"
        
        co = co & "<H4>POR</H4>"
        co = co & "<DD><p>" & rsIncidencia("creador") & "</p></DD>"
        
        co = co & "<H4>OBSERVACIONES</H4>"
        co = co & "<DD>" & cuerpo & "</DD>"
        
        co = co & "<BR>"
        
        co = co & "<H4>ENVIAR OBSERVACIONES</H4>"
        co = co & "<Table cellpadding=""0"" cellspacing=""3"" border=""0"" width=""100%"" height=""100%"">"
        co = co & "<TR>"
        co = co & "<TD style=""border: 1px solid black;"" valign=""Top"" bgcolor=""#DAFCF8"" width=""1000"" height=""100"" name=""TD_OBSERVACIONES""></TD>"
        co = co & "</TR>"
        co = co & "</TABLE><BR>"
        
        co = co & "<H4>HARDWARE</H4>"
        co = co & "<Table cellpadding=""0"" cellspacing=""3"" border=""0"">"
        co = co & "<TR>"
        co = co & "<TD width=""40"">IN</TD><TD style=""border: 1px solid black;"" bgcolor=""#DAFCF8"" width=""500"" height=""5"" name=""TD_IN""></TD>"
        co = co & "</TR>"
        co = co & "</TABLE><BR>"
        
        co = co & "<Table cellpadding=""0"" cellspacing=""3"" border=""0"">"
        co = co & "<TR>"
        co = co & "<TD width=""40"">OUT</TD><TD style=""border: 1px solid black;"" bgcolor=""#DAFCF8"" width=""500"" height=""5"" name=""TD_OUT""></TD>"
        co = co & "</TR>"
        co = co & "</TABLE>"

    Else
        co = "<h3>NO SE HA ENCONTRADO LA INCIDENCIA " & idIncidencia & "</h3>"
    End If
    
    sql = "Insert into " & taulaCodigosDeAccion() & " (IdCodigo, TipoCodigo, TmStmp, Param1, Param2, Param3, Param4, Param5, Param6, Param7, Param8) values ('" & codigoAccion & "', 'INCIDENCIA CERBERO', getdate(), '" & Param1 & "', '" & Param2 & "', '" & Param3 & "', '" & Param4 & "', '" & Param5 & "', '" & Param6 & "', '" & items & "', '" & fechaHistorico & "')"
    ExecutaComandaSql sql
    
    enviada = False
    
    If Param3 <> "" Then 'Si hay incidencia relacionada enviamos email a la secre
        sf_enviarMail "secrehit@hit.cat", "secrehit@hit.cat", "INCIDENCIA CERBERO " & idIncidencia & " EMPRESA [" & Param1 & "]", co, "", ""
        enviada = True
    End If
    
    If emailPara <> "" Then
        emailList = Split(emailPara, ";")
        For e = 0 To UBound(emailList)
            If emailList(e) = "secrehit@hit.cat" And Not enviada Then
                sf_enviarMail "secrehit@hit.cat", "secrehit@hit.cat", "INCIDENCIA CERBERO " & idIncidencia & " EMPRESA [" & Param1 & "]", co, "", ""
            Else
                sf_enviarMail "secrehit@hit.cat", emailList(e), "INCIDENCIA " & idIncidencia & " [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
                sql = "UPDATE incidencias SET enviado = enviado + 1 WHERE id = " & idIncidencia
                ExecutaComandaSql sql
            End If
        Next
    End If
    
    
    Exit Sub
    
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR SecreEnviaEmailIncidencia [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "ID Incidencia [" & idIncidencia & "] <br>" & co & "<br>ERROR:" & err.Description & "<br>" & sql, "", """"
End Sub


Sub SecreCuadranteSemanal(subj As String, emailDe As String, empresa As String)
    Dim dSemana As Integer, lunes As Date, rs As rdoResultset, rsQ As rdoResultset, sql As String
    Dim fecha As Date
    Dim codiDep As String, supervisora As String, co As String, botiga As String
    Dim rsTotal As rdoResultset
    Dim sqlFechas As String, sqlPivot As String, sqlTotal As String
    
'empresa = "Fac_Tena"
'ExecutaComandaSql "use Fac_Tena"

    InformaMiss "SecreCuadranteSemanal ", True

    On Error GoTo nor
    
    dSemana = Weekday(Now(), 2)
    lunes = DateAdd("d", -(dSemana - 1), Now())

    'emailDe = "apujol@silemabcn.com"
    'emailDe = "cescuder@silemabcn.com"
    'emailDe = "lgarcia@silemabcn.com"
    'emailDe = "jborraz@silemabcn.com"
     
     Set rs = Db.OpenResultset("select * from dependentesextes where nom='EMAIL' and upper(valor) like '%' + upper('" & emailDe & "') + '%' order by len(valor)")
     If Not rs.EOF Then
         codiDep = rs("id")
         Set rs = Db.OpenResultset("select * from constantsClient where variable='SupervisoraCodi' and valor='" & codiDep & "'")
         If rs.EOF Then 'NO ES SUPERVISORA
             Set rs = Db.OpenResultset("select * from dependentesextes where nom='TIPUSTREBALLADOR' and id='" & codiDep & "'")
             If rs("Valor") = "GERENT" Or rs("Valor") = "GERENT_2" Then
                 sql = "select c.codi, c.nom, isnull(d.nom, ' Franquicia') Supervisora "
                 sql = sql & "from paramshw p "
                 sql = sql & "left join clients c on p.valor1=c.codi "
                 sql = sql & "left join constantsClient cc on c.codi=cc.codi and cc.variable='SupervisoraCodi' "
                 sql = sql & "left join dependentes d on cc.valor = d.codi "
                 sql = sql & "where isnull(c.nom, '') <> ''  and isnull(d.nom, ' Franquicia') <> ' Franquicia' "
                 sql = sql & "and c.codi in (select c.Codi "
                 sql = sql & "from clients c "
                 sql = sql & "join paramshw w on c.Codi = w.Valor1 "
                 sql = sql & "where c.nif in (select valor from constantsempresa where camp like '%CampNif%' and isnull(valor, '')<>'')) "
                 sql = sql & "order by isnull(d.nom, ' Franquicia') , c.nom "
                 Set rs = Db.OpenResultset(sql)
             Else
                'Mirar si es franquicia
                Set rs = Db.OpenResultset("select * from constantsClient where variable='userFranquicia' and valor='" & codiDep & "'")
                If Not rs.EOF Then
                    Set rs = Db.OpenResultset("Select c.Codi, c.nom, isnull(d.nom, ' Franquicia') Supervisora from ConstantsClient cc left join clients c on cc.codi=c.codi left join dependentes d on cc.valor = d.codi where variable = 'userFranquicia' and valor = '" & codiDep & "' order by c.nom")
                Else
                    Exit Sub
                End If
             End If
         Else 'SUPERVISORA
             Set rs = Db.OpenResultset("Select c.Codi, c.nom, isnull(d.nom, ' Franquicia') Supervisora from ConstantsClient cc left join clients c on cc.codi=c.codi left join dependentes d on cc.valor = d.codi where variable = 'SupervisoraCodi' and valor = '" & codiDep & "' and c.codi is not null  and isnull(d.nom, ' Franquicia') <> ' Franquicia' order by c.nom")
         End If
     End If
     
     co = ""
     supervisora = ""
     botiga = ""
     
     
    co = co & "<h3>Total hores pactades per botiga</h3>"
    co = co & "<Table cellpadding='0' cellspacing='3' border='1'>"
    co = co & "<Tr bgColor='#DADADA'>"
    co = co & "<td><b>Botiga</b></td>"
    co = co & "<td><b>" & Format(lunes, "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "</b></td>"
    co = co & "<td><b>Total</b></td>"
    co = co & "</tr>"
     
     'sql = "select c.nom, sum(datediff(minute, t.horaInicio, t.horaFin )/60.0) horas "
     'sql = sql & "from " & taulaCdpPlanificacion(lunes) & " p "
     'sql = sql & "left join clients c on p.botiga=c.codi "
     'sql = sql & "left join cdpturnos t on p.idturno=t.idTurno "
     'sql = sql & "Where t.idTurno Is Not Null and p.activo=1 "
     'sql = sql & "group by botiga, c.nom "
     'sql = sql & "order by c.nom"
     
    sqlFechas = "isnull([" & Format(lunes, "dd/mm/yyyy") & "], 0) [" & Format(lunes, "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "],"
    sqlFechas = sqlFechas & "isnull([" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "], 0) [" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "]"
    
    sqlTotal = "isnull([" & Format(lunes, "dd/mm/yyyy") & "], 0) + "
    sqlTotal = sqlTotal & "isnull([" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "], 0) + "
    sqlTotal = sqlTotal & "isnull([" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "], 0) + "
    sqlTotal = sqlTotal & "isnull([" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "], 0) + "
    sqlTotal = sqlTotal & "isnull([" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "], 0) + "
    sqlTotal = sqlTotal & "isnull([" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "], 0) + "
    sqlTotal = sqlTotal & "isnull([" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "], 0) as Total "
    
    
    sqlPivot = "[" & Format(lunes, "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "],[" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "]"
    sql = "select nom, " & sqlFechas & ", " & sqlTotal
    sql = sql & "From "
    sql = sql & "( "
    sql = sql & "select convert(nvarchar, fecha, 103) fecha, c.nom, sum(datediff(minute, t.horaInicio, t.horaFin )/60.0) horas "
    sql = sql & "from " & taulaCdpPlanificacion(lunes) & " p "
    sql = sql & "left join cdpturnos t on p.idturno=t.idTurno "
    sql = sql & "left join clients c on p.botiga=c.codi "
    sql = sql & "Where P.activo = 1  "
    sql = sql & "group by botiga, c.nom, convert(nvarchar, p.fecha, 103) "
    sql = sql & ") DataTable "
    sql = sql & "PIVOT ( sum(horas) "
    sql = sql & "for fecha in (" & sqlPivot & ")) PivotTable "
    sql = sql & "order by nom"
     
    Set rsTotal = Db.OpenResultset(sql)
    While Not rsTotal.EOF
        co = co & "<Tr>"
        co = co & "<td align='right'>" & UCase(rsTotal("nom")) & "</td>"
        co = co & "<td align='right'>" & FormatNumber(rsTotal(Format(lunes, "dd/mm/yyyy")), 1) & "</td>"
        co = co & "<td align='right'>" & FormatNumber(rsTotal(Format(DateAdd("d", 1, lunes), "dd/mm/yyyy")), 1) & "</td>"
        co = co & "<td align='right'>" & FormatNumber(rsTotal(Format(DateAdd("d", 2, lunes), "dd/mm/yyyy")), 1) & "</td>"
        co = co & "<td align='right'>" & FormatNumber(rsTotal(Format(DateAdd("d", 3, lunes), "dd/mm/yyyy")), 1) & "</td>"
        co = co & "<td align='right'>" & FormatNumber(rsTotal(Format(DateAdd("d", 4, lunes), "dd/mm/yyyy")), 1) & "</td>"
        co = co & "<td align='right'>" & FormatNumber(rsTotal(Format(DateAdd("d", 5, lunes), "dd/mm/yyyy")), 1) & "</td>"
        co = co & "<td align='right'>" & FormatNumber(rsTotal(Format(DateAdd("d", 6, lunes), "dd/mm/yyyy")), 1) & "</td>"
        co = co & "<td align='right'>" & FormatNumber(rsTotal("Total"), 1) & "</td>"
        co = co & "</tr>"
     
        rsTotal.MoveNext
     Wend
     co = co & "</table><br>"

     co = co & "<h3>Quadrants setmana del " & Format(lunes, "dd/mm/yyyy") & " al " & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "</h3>"
     
     While Not rs.EOF
         If supervisora <> rs("supervisora") Then
             co = co & "<h4>" & rs("supervisora") & "</h4>"
             'Co = Co & "<h5>Botiga " & rs("nom") & "</h5>"
             'Co = Co & "<Table cellpadding='0' cellspacing='3' border='1'>"
             'Co = Co & "<Tr bgColor='#DADADA'>"
             'Co = Co & "<td>" & Format(lunes, "dd/mm/yyyy") & "</td>"
             'Co = Co & "<td>" & Format(DateAdd("d", 1, lunes), "dd/mm/yyyy") & "</td>"
             'Co = Co & "<td>" & Format(DateAdd("d", 2, lunes), "dd/mm/yyyy") & "</td>"
             'Co = Co & "<td>" & Format(DateAdd("d", 3, lunes), "dd/mm/yyyy") & "</td>"
             'Co = Co & "<td>" & Format(DateAdd("d", 4, lunes), "dd/mm/yyyy") & "</td>"
             'Co = Co & "<td>" & Format(DateAdd("d", 5, lunes), "dd/mm/yyyy") & "</td>"
             'Co = Co & "<td>" & Format(DateAdd("d", 6, lunes), "dd/mm/yyyy") & "</td></Tr>"
             supervisora = rs("supervisora")
         End If
         co = co & "<b>" & UCase(rs("nom")) & "</b>"
         co = co & cuadrantePlanificacionTurnos(Now(), rs("codi"), False)
         co = co & "<br>"
         
         rs.MoveNext
     Wend
         
    sf_enviarMail "secrehit@hit.cat", emailDe, subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", co, "", ""
     
    Exit Sub
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", subj & "  [" & Format(Now(), "dd/mm/yy hh:nn") & "]", "ERROR INFORME CUADRANTE: " & err.Description, "", ""
End Sub


Sub calculaInformePrevisiones(subj As String, emailDe As String, empresa As String)
    Dim fechaIni As Date, fechaFin As Date ', fecha As Date
    Dim rs As rdoResultset, rsPrev As rdoResultset
    Dim botiga As String, BotigaNom As String
    Dim msg As String
    Dim D As Integer
    Dim tMoviments As String
    Dim Semana As Integer, semanaAux As Integer
    Dim domingo As Date, lunes As Date
    Dim totalBotiga As Double, totalDia As Double

    Semana = -1
    On Error GoTo ErrData
    If InStr(UCase(subj), "SEMANA") Then
        Semana = CInt(Split(subj, " ")(3))
        semanaAux = 0
        'If Semana < DatePart("ww", Now(), vbMonday) Then
        '    lunes = CDate("01/01/" & Year(Now()) + 1)
        'Else
            lunes = CDate("01/01/" & Year(Now()))
        'End If
        If Semana > 0 Then
            While Semana <> semanaAux
                lunes = DateAdd("d", 1, lunes)
                semanaAux = DatePart("ww", lunes, vbMonday)
            Wend
            
            domingo = DateAdd("d", 6, lunes)
            'fecha = lunes
        End If
    Else
        lunes = DateAdd("d", 14, Now())
        domingo = DateAdd("d", 7, lunes)
    End If
    GoTo OkData
    
ErrData:
        lunes = DateAdd("d", 14, Now())
        domingo = DateAdd("d", 7, lunes)
        
OkData:
    On Error GoTo nor
    
    fechaIni = lunes
    fechaFin = domingo
    
    If Month(fechaIni) <> Month(fechaFin) Then
        tMoviments = "(select * from [" & NomTaulaMovi(fechaIni) & "] union all select * from [" & NomTaulaMovi(fechaFin) & "])"
    Else
        tMoviments = "[" & NomTaulaMovi(fechaIni) & "]"
    End If
   
    msg = "<TABLE BORDER='1'>"
    msg = msg & "<TR><TD ROWSPAN='2'><B>BOTIGA</B></TD>"
    For D = 0 To 6
        msg = msg & "<TD COLSPAN='2'><B>" & Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") & "</B></TD>"
    Next
    msg = msg & "<TD ROWSPAN='2'><B>TOTAL</B></TD>"
    msg = msg & "</TR>"
    msg = msg & "<TR>"
    For D = 0 To 6
        msg = msg & "<TD><B>MATI</B></TD><TD><B>TARDA</B></TD>"
    Next
    msg = msg & "</TR>"
    
    
    Set rs = Db.OpenResultset("select c.nom botigaNom, c.Codi as botiga ,w.codi as llicencia from clients c join paramshw w on c.Codi = w.Valor1 order by c.nom")
    While Not rs.EOF
        botiga = rs("botiga")
        BotigaNom = rs("botigaNom")
        
        msg = msg & "<TR><TD><B>" & BotigaNom & "</B></TD>"
        Set rsPrev = Db.OpenResultset("select * from " & tMoviments & " m where botiga = '" & botiga & "' and data between '" & Format(fechaIni, "dd/mm/yyyy") & "' and '" & Format(fechaFin, "dd/mm/yyyy") & "' and tipus_moviment in ('MATI', 'TARDA') order by data, tipus_moviment")
        
        totalBotiga = 0
        For D = 0 To 6
            If Not rsPrev.EOF Then
                If Format(rsPrev("data"), "dd/mm/yyyy") = Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") Then
                    If rsPrev("tipus_moviment") = "MATI" Then
                        msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_MATI'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'MATI
                        totalBotiga = totalBotiga + rsPrev("Import")
                        rsPrev.MoveNext
                        If Not rsPrev.EOF Then
                            If Format(rsPrev("data"), "dd/mm/yyyy") = Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") Then
                                If rsPrev("tipus_moviment") = "TARDA" Then
                                    msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_TARDA'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'TARDA
                                    totalBotiga = totalBotiga + rsPrev("Import")
                                Else
                                    msg = msg & "<TD>-1</TD>"
                                End If
                            Else
                                msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_TARDA'>-</TD>"
                            End If
                            rsPrev.MoveNext
                        Else
                            msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") & "_TARDA'>-</TD>"
                        End If
                    Else
                        If rsPrev("tipus_moviment") = "TARDA" Then
                            msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_MATI'>-</TD><TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_TARDA'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'TARDA
                            totalBotiga = totalBotiga + rsPrev("Import")
                            rsPrev.MoveNext
                        Else
                            msg = msg & "<TD>-4</TD>"
                        End If
                    End If
                Else
                    msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_MATI'>-</TD><TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_TARDA'>-</TD>" 'NO HI HA PREVISIÓ PER AQUEST DIA
                End If
            Else
                msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") & "_MATI'>-</TD><TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") & "_TARDA'>-</TD>" 'NO HI HA PREVISIÓ PER AQUEST DIA
            End If
        Next
    
        msg = msg & "<TD align='right'><b>" & FormatNumber(totalBotiga, 2) & "<b></TD>" 'TOTAL BOTIGA
        rs.MoveNext
    Wend
    
    sf_enviarMail "secrehit@hit.cat", emailDe, "Previsiones semana " & Semana, msg, "", ""
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "Previsiones semana " & Semana, msg, "", ""
    Exit Sub
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "Previsiones semana " & Semana, msg & "<h1>ERROR: " & err.Description & "</h1>", "", ""
End Sub


Sub calculaInformePrevisiones2(subj As String, emailDe As String, empresa As String)
    Dim fechaIni As Date, fechaFin As Date ', fecha As Date
    Dim fechaIniPasado As Date, fechaFinPasado As Date
    Dim rs As rdoResultset, rsPrev As rdoResultset, rsSup As rdoResultset
    Dim botiga As String, BotigaNom As String
    Dim msg As String
    Dim D As Integer
    Dim tMoviments As String, tMovimentsPasado As String
    Dim Semana As Integer, semanaAux As Integer
    Dim domingo As Date, lunes As Date
    Dim totalBotiga As Double, totalDiaMati(7) As Double, totalDiaTarda(7) As Double, Total As Double, totalDiaMati_C(7) As Double, totalDiaTarda_C(7) As Double, Total_C As Double
    Dim totalDiaMati_R(7) As Double, totalDiaTarda_R(7) As Double, Total_R As Double
    Dim totalDiaMati_RP(7) As Double, totalDiaTarda_RP(7) As Double, Total_RP As Double
    
    Semana = -1
    On Error GoTo ErrData
    If InStr(UCase(subj), "SEMANA") Then
        Semana = CInt(Split(subj, " ")(3))
        semanaAux = 0
        
        If Semana = 1 Then
            lunes = CDate("26/12/" & Year(Now()) - 1)
            While DatePart("w", lunes) <> vbMonday
                lunes = DateAdd("d", 1, lunes)
            Wend
            
            domingo = DateAdd("d", 6, lunes)
        Else
            lunes = CDate("01/01/" & Year(Now()))
            If Semana > 0 Then
                While Semana <> semanaAux
                    lunes = DateAdd("d", 1, lunes)
                    semanaAux = DatePart("ww", lunes, vbMonday)
                Wend
                
                domingo = DateAdd("d", 6, lunes)
            End If
        End If
    Else
        lunes = DateAdd("d", 14, Now())
        domingo = DateAdd("d", 7, lunes)
    End If
    GoTo OkData
    
ErrData:
    lunes = DateAdd("d", 14, Now())
    domingo = DateAdd("d", 7, lunes)
        
OkData:
    On Error GoTo nor
    
    fechaIni = lunes
    fechaFin = domingo
    
    fechaIniPasado = DateAdd("yyyy", -1, fechaIni)
    'While DatePart("w", fechaIniPasado, vbMonday) <> vbMonday
    '    fechaIniPasado = DateAdd("d", 1, fechaIniPasado)
    'Wend
    fechaFinPasado = DateAdd("d", 6, fechaIniPasado)
    
    'ExecutaComandaSql "Insert into CodigosDeAccion (IdCodigo, TipoCodigo, TmStmp, Param1, Param2) values ('" & codigoAccion & "', 'PREVISIONES', getdate(), '[" & fechaIni & "]', '[" & fechaFin & "]')"
    
    If Month(fechaIni) <> Month(fechaFin) Then
        tMoviments = "(select * from [" & NomTaulaMovi(fechaIni) & "] union all select * from [" & NomTaulaMovi(fechaFin) & "])"
    Else
        tMoviments = "[" & NomTaulaMovi(fechaIni) & "]"
    End If
    
    If Month(fechaIniPasado) <> Month(fechaFinPasado) Then
        tMovimentsPasado = "(select * from [" & NomTaulaMovi(fechaIniPasado) & "] union all select * from [" & NomTaulaMovi(fechaFinPasado) & "])"
    Else
        tMovimentsPasado = "[" & NomTaulaMovi(fechaIniPasado) & "]"
    End If
    
    msg = ""
    Set rsSup = Db.OpenResultset("select isnull(d.nom, '') nom, isnull(d.codi, '') codi from dependentes d left join dependentesextes de on d.codi=de.id and de.nom='EMAIL' where d.codi in (select distinct valor from constantsclient where variable = 'SupervisoraCodi' and valor<>'')")
    While Not rsSup.EOF
        msg = msg & "<H3>" & rsSup("nom") & "</H3>"
        msg = msg & "<TABLE BORDER='1'>"
        msg = msg & "<TR><TD ROWSPAN='2'><B>BOTIGA</B></TD>"
        For D = 0 To 6
            msg = msg & "<TD COLSPAN='2'><B>" & Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") & "</B></TD>"
        Next
        msg = msg & "<TD ROWSPAN='2'><B>TOTAL</B></TD>"
        msg = msg & "</TR>"
        msg = msg & "<TR>"
        For D = 0 To 6
            msg = msg & "<TD><B>MATI</B></TD><TD><B>TARDA</B></TD>"
            
            totalDiaMati(D) = 0
            totalDiaTarda(D) = 0
            
            totalDiaMati_C(D) = 0
            totalDiaTarda_C(D) = 0
            
            totalDiaMati_R(D) = 0
            totalDiaTarda_R(D) = 0
            
            totalDiaMati_RP(D) = 0
            totalDiaTarda_RP(D) = 0
        Next
        msg = msg & "</TR>"
        
        Total = 0
        Set rs = Db.OpenResultset("select c.nom botigaNom, c.Codi as botiga ,w.codi as llicencia from paramshw w left join clients c  on c.Codi = w.Valor1 left join constantsClient cc on c.codi=cc.codi and cc.variable='SupervisoraCodi' where cc.valor = '" & rsSup("codi") & "' and c.codi is not null order by c.nom")
        While Not rs.EOF
            botiga = rs("botiga")
            BotigaNom = rs("botigaNom")
            
            'PREVISIÓN CALCULADA
            msg = msg & "<TR BGCOLOR='#DDDDDD'><TD><B>" & UCase(BotigaNom) & " (Proposta)</B></TD>"
            Set rsPrev = Db.OpenResultset("select * from " & tMoviments & " m where botiga = '" & botiga & "' and data between '" & Format(fechaIni, "dd/mm/yyyy") & "' and '" & Format(fechaFin, "dd/mm/yyyy") & "' and tipus_moviment in ('MATI', 'TARDA') order by data, tipus_moviment")
                        
            totalBotiga = 0
            For D = 0 To 6
                If Not rsPrev.EOF Then
                    If Format(rsPrev("data"), "dd/mm/yyyy") = Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") Then
                        If rsPrev("tipus_moviment") = "MATI" Then
                            msg = msg & "<TD width='80' align='right'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'MATI
                            totalBotiga = totalBotiga + rsPrev("Import")
                            totalDiaMati(D) = totalDiaMati(D) + rsPrev("Import")
                            rsPrev.MoveNext
                            If Not rsPrev.EOF Then
                                If Format(rsPrev("data"), "dd/mm/yyyy") = Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") Then
                                    If rsPrev("tipus_moviment") = "TARDA" Then
                                        msg = msg & "<TD width='80' align='right'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'TARDA
                                        totalBotiga = totalBotiga + rsPrev("Import")
                                        totalDiaTarda(D) = totalDiaTarda(D) + rsPrev("Import")
                                    Else
                                        msg = msg & "<TD>-1</TD>"
                                    End If
                                Else
                                    msg = msg & "<TD width='80' align='right'>-</TD>"
                                End If
                                rsPrev.MoveNext
                            Else
                                msg = msg & "<TD width='80' align='right'>-</TD>"
                            End If
                        Else
                            If rsPrev("tipus_moviment") = "TARDA" Then
                                msg = msg & "<TD width='80' align='right'>-</TD><TD width='80' align='right'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'TARDA
                                totalBotiga = totalBotiga + rsPrev("Import")
                                totalDiaTarda(D) = totalDiaTarda(D) + rsPrev("Import")
                                rsPrev.MoveNext
                            Else
                                msg = msg & "<TD>-4</TD>"
                            End If
                        End If
                    Else
                        msg = msg & "<TD width='80' align='right'>-</TD><TD width='80' align='right'>-</TD>" 'NO HI HA PREVISIÓ PER AQUEST DIA
                    End If
                Else
                    msg = msg & "<TD width='80' align='right'>-</TD><TD width='80' align='right'>-</TD>" 'NO HI HA PREVISIÓ PER AQUEST DIA
                End If
            Next
        
            msg = msg & "<TD align='right'><b>" & FormatNumber(totalBotiga, 2) & "<b></TD></TR>" 'TOTAL BOTIGA
            
            
            'REALIDAD AÑO PASADO
            msg = msg & "<TR BGCOLOR='#EEEEEE'><TD><B>" & UCase(BotigaNom) & " (Realitat " & Year(fechaIniPasado) & ")</B></TD>"
            Set rsPrev = Db.OpenResultset("select * from " & tMovimentsPasado & " m where botiga = '" & botiga & "' and data between '" & Format(fechaIniPasado, "dd/mm/yyyy") & "' and '" & Format(fechaFinPasado, "dd/mm/yyyy") & "' and tipus_moviment = 'Z' order by data")
                        
            totalBotiga = 0
            For D = 0 To 6
                If Not rsPrev.EOF Then
                    If Format(rsPrev("data"), "dd/mm/yyyy") = Format(DateAdd("d", D, fechaIniPasado), "dd/mm/yyyy") Then
                        If DatePart("h", rsPrev("data")) < 16 Then 'MATI
                            msg = msg & "<TD width='80' align='right'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'MATI
                            totalBotiga = totalBotiga + rsPrev("Import")
                            totalDiaMati_RP(D) = totalDiaMati_RP(D) + rsPrev("Import")
                            rsPrev.MoveNext
                            
                            If Not rsPrev.EOF Then
                                If Format(rsPrev("data"), "dd/mm/yyyy") = Format(DateAdd("d", D, fechaIniPasado), "dd/mm/yyyy") Then
                                    If DatePart("h", rsPrev("data")) > 16 Then 'TARDA
                                        msg = msg & "<TD width='80' align='right'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'TARDA
                                        totalBotiga = totalBotiga + rsPrev("Import")
                                        totalDiaTarda_RP(D) = totalDiaTarda_RP(D) + rsPrev("Import")
                                        rsPrev.MoveNext
                                    Else 'NO HAY TARDA
                                        msg = msg & "<TD width='80' align='right'>-</TD>"
                                    End If
                                Else 'NO HAY TARDA
                                    msg = msg & "<TD width='80' align='right'>-</TD>"
                                End If
                            Else 'NO HAY TARDA
                                msg = msg & "<TD width='80' align='right'>-</TD>"
                            End If
                        Else 'NO HAY MATI
                            msg = msg & "<TD width='80' align='right'>-</TD>"
                            If DatePart("h", rsPrev("data")) > 16 Then 'TARDA
                                msg = msg & "<TD width='80' align='right'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'TARDA
                                totalBotiga = totalBotiga + rsPrev("Import")
                                totalDiaTarda_RP(D) = totalDiaTarda_RP(D) + rsPrev("Import")
                                rsPrev.MoveNext
                            Else 'NO HAY TARDA
                                msg = msg & "<TD width='80' align='right'>-</TD>"
                            End If
                        End If
                    Else
                        msg = msg & "<TD width='80' align='right'>-</TD><TD width='80' align='right'>-</TD>"
                    End If
                Else
                    msg = msg & "<TD width='80' align='right'>-</TD><TD width='80' align='right'>-</TD>"
                End If
            Next
        
            msg = msg & "<TD align='right'><b>" & FormatNumber(totalBotiga, 2) & "<b></TD></TR>" 'TOTAL BOTIGA
            
            'REALIDAD
            msg = msg & "<TR BGCOLOR='#EEEEEE'><TD><B>" & UCase(BotigaNom) & " (Realitat)</B></TD>"
            Set rsPrev = Db.OpenResultset("select * from " & tMoviments & " m where botiga = '" & botiga & "' and data between '" & Format(fechaIni, "dd/mm/yyyy") & "' and '" & Format(fechaFin, "dd/mm/yyyy") & "' and tipus_moviment = 'Z' order by data")
                        
'If UCase(BotigaNom) = "T--022" Or UCase(BotigaNom) = "T--005" Then
'totalBotiga = 0
'End If
            totalBotiga = 0
            For D = 0 To 6
                If Not rsPrev.EOF Then
                    If Format(rsPrev("data"), "dd/mm/yyyy") = Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") Then
                        If DatePart("h", rsPrev("data")) < 16 Then 'MATI
                            msg = msg & "<TD width='80' align='right'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'MATI
                            totalBotiga = totalBotiga + rsPrev("Import")
                            totalDiaMati_R(D) = totalDiaMati_R(D) + rsPrev("Import")
                            rsPrev.MoveNext
                            
                            If Not rsPrev.EOF Then
                                If Format(rsPrev("data"), "dd/mm/yyyy") = Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") Then
                                    If DatePart("h", rsPrev("data")) > 16 Then 'TARDA
                                        msg = msg & "<TD width='80' align='right'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'TARDA
                                        totalBotiga = totalBotiga + rsPrev("Import")
                                        totalDiaTarda_R(D) = totalDiaTarda_R(D) + rsPrev("Import")
                                        rsPrev.MoveNext
                                    Else 'NO HAY TARDA
                                        msg = msg & "<TD width='80' align='right'>-</TD>"
                                    End If
                                Else 'NO HAY TARDA
                                    msg = msg & "<TD width='80' align='right'>-</TD>"
                                End If
                            Else 'NO HAY TARDA
                                msg = msg & "<TD width='80' align='right'>-</TD>"
                            End If
                        Else 'NO HAY MATI
                            msg = msg & "<TD width='80' align='right'>-</TD>"
                            If DatePart("h", rsPrev("data")) > 16 Then 'TARDA
                                msg = msg & "<TD width='80' align='right'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'TARDA
                                totalBotiga = totalBotiga + rsPrev("Import")
                                totalDiaTarda_R(D) = totalDiaTarda_R(D) + rsPrev("Import")
                                rsPrev.MoveNext
                            Else 'NO HAY TARDA
                                msg = msg & "<TD width='80' align='right'>-</TD>"
                            End If
                        End If
                    Else
                        msg = msg & "<TD width='80' align='right'>-</TD><TD width='80' align='right'>-</TD>"
                    End If
                Else
                    msg = msg & "<TD width='80' align='right'>-</TD><TD width='80' align='right'>-</TD>"
                End If
            Next
        
            msg = msg & "<TD align='right'><b>" & FormatNumber(totalBotiga, 2) & "<b></TD></TR>" 'TOTAL BOTIGA
                                
            
            'PREVISIÓN MODIFICADA
            msg = msg & "<TR><TD><B>" & UCase(BotigaNom) & " (Modificada)</B></TD>"
            Set rsPrev = Db.OpenResultset("select * from " & tMoviments & " m where botiga = '" & botiga & "' and data between '" & Format(fechaIni, "dd/mm/yyyy") & "' and '" & Format(fechaFin, "dd/mm/yyyy") & "' and tipus_moviment in ('MATI_C', 'TARDA_C') order by data, tipus_moviment")
                        
            totalBotiga = 0
            For D = 0 To 6
                If Not rsPrev.EOF Then
                    If Format(rsPrev("data"), "dd/mm/yyyy") = Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") Then
                        If rsPrev("tipus_moviment") = "MATI_C" Then
                            msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_MATI'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'MATI
                            totalBotiga = totalBotiga + rsPrev("Import")
                            totalDiaMati_C(D) = totalDiaMati_C(D) + rsPrev("Import")
                            rsPrev.MoveNext
                            If Not rsPrev.EOF Then
                                If Format(rsPrev("data"), "dd/mm/yyyy") = Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") Then
                                    If rsPrev("tipus_moviment") = "TARDA_C" Then
                                        msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_TARDA'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'TARDA
                                        totalBotiga = totalBotiga + rsPrev("Import")
                                        totalDiaTarda_C(D) = totalDiaTarda_C(D) + rsPrev("Import")
                                    Else
                                        msg = msg & "<TD>-1</TD>"
                                    End If
                                Else
                                    msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_TARDA'>-</TD>"
                                End If
                                rsPrev.MoveNext
                            Else
                                msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") & "_TARDA'>-</TD>"
                            End If
                        Else
                            If rsPrev("tipus_moviment") = "TARDA_C" Then
                                msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_MATI'>-</TD><TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_TARDA'>" & FormatNumber(rsPrev("Import"), 2) & "</TD>" 'TARDA
                                totalBotiga = totalBotiga + rsPrev("Import")
                                totalDiaTarda_C(D) = totalDiaTarda_C(D) + rsPrev("Import")
                                rsPrev.MoveNext
                            Else
                                msg = msg & "<TD>-4</TD>"
                            End If
                        End If
                    Else
                        msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_MATI'>-</TD><TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(rsPrev("data"), "dd-mm-yyyy") & "_TARDA'>-</TD>" 'NO HI HA PREVISIÓ PER AQUEST DIA
                    End If
                Else
                    msg = msg & "<TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") & "_MATI'>-</TD><TD style='border: 1px solid black;' width='80' align='right' name='P_" & botiga & "_" & Format(DateAdd("d", D, fechaIni), "dd/mm/yyyy") & "_TARDA'>-</TD>" 'NO HI HA PREVISIÓ PER AQUEST DIA
                End If
            Next
        
            msg = msg & "<TD align='right'><b>" & FormatNumber(totalBotiga, 2) & "<b></TD></TR>" 'TOTAL BOTIGA

            msg = msg & "<TR HEIGHT='1'><TD COLSPAN='16'><hr></TD></TR>"
            rs.MoveNext
        Wend
        
        msg = msg & "<TR BGCOLOR='#DDDDDD'><TD><B>TOTAL (Proposta)</B></TD>"
        For D = 0 To 6
            msg = msg & "<TD align='right'><b>" & FormatNumber(totalDiaMati(D), 2) & "<b></TD><TD align='right'><b>" & FormatNumber(totalDiaTarda(D), 2) & "<b></TD>"
            Total = Total + totalDiaMati(D) + totalDiaTarda(D)
        Next
        msg = msg & "<TD>" & FormatNumber(Total, 2) & "</TD>"
        msg = msg & "</TR>"
        
        msg = msg & "<TR BGCOLOR='#EEEEEE'><TD><B>TOTAL (Realitat (" & Year(fechaIniPasado) & ")</B></TD>"
        For D = 0 To 6
            msg = msg & "<TD align='right'><b>" & FormatNumber(totalDiaMati_RP(D), 2) & "<b></TD><TD align='right'><b>" & FormatNumber(totalDiaTarda_RP(D), 2) & "<b></TD>"
            Total_RP = Total_RP + totalDiaMati_RP(D) + totalDiaTarda_RP(D)
        Next
        msg = msg & "<TD>" & FormatNumber(Total_RP, 2) & "</TD>"
        msg = msg & "</TR>"
        
        
        msg = msg & "<TR BGCOLOR='#EEEEEE'><TD><B>TOTAL (Realitat)</B></TD>"
        For D = 0 To 6
            msg = msg & "<TD align='right'><b>" & FormatNumber(totalDiaMati_R(D), 2) & "<b></TD><TD align='right'><b>" & FormatNumber(totalDiaTarda_R(D), 2) & "<b></TD>"
            Total_R = Total_R + totalDiaMati_R(D) + totalDiaTarda_R(D)
        Next
        msg = msg & "<TD>" & FormatNumber(Total_R, 2) & "</TD>"
        msg = msg & "</TR>"
        
        
        msg = msg & "<TR><TD><B>TOTAL (Modificada)</B></TD>"
        For D = 0 To 6
            msg = msg & "<TD align='right'><b>" & FormatNumber(totalDiaMati_C(D), 2) & "<b></TD><TD align='right'><b>" & FormatNumber(totalDiaTarda_C(D), 2) & "<b></TD>"
            Total_C = Total_C + totalDiaMati_C(D) + totalDiaTarda_C(D)
        Next
        msg = msg & "<TD>" & FormatNumber(Total_C, 2) & "</TD>"
        msg = msg & "</TR>"
        
        msg = msg & "</TABLE>"
        rsSup.MoveNext
    Wend
    
    sf_enviarMail "secrehit@hit.cat", emailDe, "Previsiones semana " & Semana, msg, "", ""
    'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "Previsiones semana " & Semana, msg, "", ""
    Exit Sub
nor:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "Previsiones semana " & Semana, msg & "<h1>ERROR: " & err.Description & "</h1>", "", ""
End Sub


Function utf8(cadena As String) As String
    Dim utf8Str As String
    
    utf8Str = cadena
    
    utf8Str = Replace(utf8Str, "à", "&agrave;")
    utf8Str = Replace(utf8Str, "á", "&aacute;")
    utf8Str = Replace(utf8Str, "À", "&Agrave;")
    utf8Str = Replace(utf8Str, "Á", "&Aacute;")
    
    utf8Str = Replace(utf8Str, "è", "&egrave;")
    utf8Str = Replace(utf8Str, "é", "&eacute;")
    utf8Str = Replace(utf8Str, "È", "&Egrave;")
    utf8Str = Replace(utf8Str, "É", "&Eacute;")
    
    'utf8Str = Replace(utf8Str, "ì", "&igrave;")
    utf8Str = Replace(utf8Str, "í", "&iacute;")
    'utf8Str = Replace(utf8Str, "Ì", "&Igrave;")
    utf8Str = Replace(utf8Str, "Í", "&Iacute;")
    
    utf8Str = Replace(utf8Str, "ò", "&ograve;")
    utf8Str = Replace(utf8Str, "ó", "&oacute;")
    utf8Str = Replace(utf8Str, "Ò", "&Ograve;")
    utf8Str = Replace(utf8Str, "Ó", "&Oacute;")
    
    utf8Str = Replace(utf8Str, "ü", "&uuml;")
    'utf8Str = Replace(utf8Str, "ù", "&ugrave;")
    utf8Str = Replace(utf8Str, "ú", "&uacute;")
    utf8Str = Replace(utf8Str, "Ü", "&Uuml;")
    'utf8Str = Replace(utf8Str, "Ù", "&Ugrave;")
    utf8Str = Replace(utf8Str, "Ú", "&Uacute;")
    
        
    utf8Str = Replace(utf8Str, "ñ", "&ntilde;")
    utf8Str = Replace(utf8Str, "Ñ", "&Ntilde;")
        
    utf8Str = Replace(utf8Str, "º", "&ordm;")
    utf8Str = Replace(utf8Str, "ª", "&ordf;")
    
    utf8 = utf8Str
End Function


