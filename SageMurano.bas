Attribute VB_Name = "SageMurano"
Global EmpresaMurano As String, nDigitos As Double
Function DigitosCC(empresa As Double) As Double
    Dim rs As rdoResultset, prefijo As String
    
    DigitosCC = 9
    
    If empresa = 0 Then
        prefijo = ""
    Else
        prefijo = empresa & "_"
    End If
    
    Set rs = Db.OpenResultset("select isnull(valor, 9) nDigitos from constantsempresa where camp = '" & prefijo & "CampDSubcuenta'")
    If Not rs.EOF Then
        If IsNumeric(rs("nDigitos")) Then DigitosCC = rs(0)
    End If
End Function

Sub ExportaMURANO(Tipus As String, p1 As String, P2 As String, Optional P3 As String, Optional P4 As String, Optional P5 As String, Optional idCalcul As String)
    Dim sql As String, Descripcio As String, Botigues() As String, i, rsEmp As rdoResultset, rsNEmp As rdoResultset, nEmpresa As Double, rsNA As rdoResultset, rsCtb As rdoResultset, numAsiento As String, botiga As String
    Dim Di As Date, Df As Date, D As Integer, intTickets As String, Z As Double
    Dim idFactura As String, dataFactura As Date, numFactura As String, taulaFactura As String, rsHist As rdoResultset, dataInici As Date, dataFi As Date, idNorma43 As String
    
    If Not frmSplash.Debugant Then On Error GoTo norR

    nDigitos = 9
    dbSage = "[silema_Ts].sage"
    
    If idCalcul <> "" Then 'CALCULO EXTERNO
        Set rsEmp = Db.OpenResultset("select * from hit.dbo.CalculsEspecials where id = '" & idCalcul & "'")
        If Not rsEmp.EOF Then EmpresaActual = rsEmp("Empresa")
    End If

    Select Case Tipus
        'Case "CAIXA"
        '    botiga = Car(p1)
        '    Di = Car(P2)
        '    If botiga = "00" Then
        '        Set rsNEmp = Db.OpenResultset("Select c.codi from paramshw w join clients c on w.valor1=c.codi ")
        '        botiga = ""
        '        While Not rsNEmp.EOF
        '            botiga = botiga & rsNEmp(0) & ","
        '            rsNEmp.MoveNext
        '        Wend
        '    End If
           
        '    Botigues = Split(botiga & ",", ",")
        '    For i = 0 To UBound(Botigues) - 1
        '        botiga = Botigues(i)
        '        nEmpresa = 0
        '        sql = "select cc.Valor, c.Codi, c.Nom "
        '        sql = sql & "from constantsclient cc "
        '        sql = sql & "left join clients c on cc.Codi=c.codi "
        '        sql = sql & "where cc.variable = 'EmpresaVendes' and cc.Codi in (select valor1 from ParamsHw) and c.codi = '" & botiga & "' "
        '        sql = sql & "order by cc.Valor, c.nom"
        '        Set rsNEmp = Db.OpenResultset(sql)
        '        If Not rsNEmp.EOF Then nEmpresa = CDbl(rsNEmp("valor"))
        '        If InStr(P2, "[00") Then
        '            For D = 1 To 30
        '                Di = Car(Replace(P2, "[00", "[" & D))
        '                ExportaMURANO_CaixaBotiga nEmpresa, botiga, Di, idCalcul
        '            Next
        '        Else
        '            ExportaMURANO_CaixaBotiga nEmpresa, botiga, Di, idCalcul
        '        End If
        '    Next
        Case "CAIXA_ONLINE"
        
            botiga = Car(p1)
            Di = Car(P2)
            Df = Car(P3)
            intTickets = Car(P4)
            Z = Car(P5)
            
            nEmpresa = -1
            sql = "select cc.Valor, c.Codi, c.Nom "
            sql = sql & "from constantsclient cc "
            sql = sql & "left join clients c on cc.Codi=c.codi "
            sql = sql & "where cc.variable = 'EmpresaVendes' and cc.Codi in (select valor1 from ParamsHw) and c.codi = '" & botiga & "' "
            sql = sql & "order by cc.Valor, c.nom"
            Set rsNEmp = Db.OpenResultset(sql)
            If Not rsNEmp.EOF Then nEmpresa = CDbl(rsNEmp("valor"))
            
            If nEmpresa > -1 Then
                nDigitos = DigitosCC(nEmpresa)
                
                If UCase(EmpresaActual) <> UCase("Tena") Then
                    ExportaMURANO_CaixaBotigaOnLine_V3 nEmpresa, botiga, Di, Df, intTickets, Z 'NO SE EXPORTAN VENTAS, SOLO MOVIMIENTOS
                Else
                    If botiga = "761" Then 'probando con T--073
                        ExportaMURANO_CaixaBotigaOnLine_V3 nEmpresa, botiga, Di, Df, intTickets, Z 'NO SE EXPORTAN VENTAS, SOLO MOVIMIENTOS
                    Else
                        ExportaMURANO_CaixaBotigaOnLine nEmpresa, botiga, Di, Df, intTickets, Z
                    End If
                End If
            End If
        
        Case "VENDES"
            botiga = p1
            botiga = Car(botiga)
            
            nEmpresa = -1
            sql = "select cc.Valor, c.Codi, c.Nom "
            sql = sql & "from constantsclient cc "
            sql = sql & "left join clients c on cc.Codi=c.codi "
            sql = sql & "where cc.variable = 'EmpresaVendes' and cc.Codi in (select valor1 from ParamsHw) and c.codi = '" & botiga & "' "
            sql = sql & "order by cc.Valor, c.nom"
            Set rsNEmp = Db.OpenResultset(sql)
            If Not rsNEmp.EOF Then nEmpresa = CDbl(rsNEmp("valor"))
            
            If nEmpresa > -1 Then
                nDigitos = DigitosCC(nEmpresa)
                
                If UCase(EmpresaActual) = UCase("Concordia") Then
                    ExportaMURANO_VendesBotiga 0, botiga, "NO"
                    ExportaMURANO_VendesBotiga 1, botiga, "SI"
                Else
                    ExportaMURANO_VendesBotiga nEmpresa, botiga
                End If
            End If
        
        
        Case "FACTURA"
            idFactura = Car(p1)
            dataFactura = Car(P2)
            numFactura = Car(P3)
            taulaFactura = Car(P4)
            
            ExportaMURANO_FacturaEmesa idFactura, dataFactura, numFactura, taulaFactura
            
            'ExportaANALITICA_FacturaEmesa idFactura, dataFactura, numFactura, taulaFactura
            
        'Case "FACTURA_DEL"
        '    idFactura = Car(p1)
        '    nEmpresa = Car(P3)
        '    dataFactura = Car(P2)
        '
        '    If nEmpresa = "0" Then
        '        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampEnlaceNominas' ")
        '    Else
        '        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampEnlaceNominas' ")
        '    End If
        '    If Not rsCtb.EOF Then EmpresaMurano = Trim(Left(rsCtb("Valor"), InStr(rsCtb("Valor"), " ")))
        
        '    Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(dataFactura) & " where Param2 = '" & idFactura & "' and CodigoEmpresa='" & EmpresaMurano & "' and TipoExportacion='FACTURA'")
        '    While Not rsHist.EOF
        '        MuranoExecute "Delete from [WEB].[Sage].[dbo].[Movimientos] where Asiento = " & rsHist("Asiento") & " and CodigoEmpresa=" & EmpresaMurano
        '        rsHist.MoveNext
        '    Wend
            
        Case "FACTURA_REBUDA"
            idFactura = Car(p1)
            dataFactura = Car(P2)
            numFactura = Car(P3)
            taulaFactura = Car(P4)
            
            ExportaMURANO_FacturaRebuda idFactura, dataFactura, numFactura, taulaFactura
            
            ExportaANALITICA_FacturaRebuda idFactura, dataFactura, numFactura, taulaFactura
            
        'Case "FACTURA_REBUDA_DEL"
        '    idFactura = Car(p1)
        '    nEmpresa = Car(P3)
        '    dataFactura = Car(P2)
            
        '    If nEmpresa = "0" Then
        '        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampEnlaceNominas' ")
        '    Else
        '        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampEnlaceNominas' ")
        '    End If
        '    If Not rsCtb.EOF Then EmpresaMurano = Trim(Left(rsCtb("Valor"), InStr(rsCtb("Valor"), " ")))
       '
       '     Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(dataFactura) & " where Param2 = '" & idFactura & "' and CodigoEmpresa='" & EmpresaMurano & "' and TipoExportacion='FACTURA_REBUDA'")
       '     While Not rsHist.EOF
       '         MuranoExecute "Delete from [WEB].[Sage].[dbo].[Movimientos] where Asiento = " & rsHist("Asiento") & " and CodigoEmpresa=" & EmpresaMurano
       '         rsHist.MoveNext
       '     Wend
            
        Case "NOMINES"
            nEmpresa = Car(p1)
            dataInici = Car(P2)
            dataFi = Car(P3)
            
            ExportaMURANO_Nomines nEmpresa, dataInici, dataFi
        Case "BANCS"
            idNorma43 = Car(p1)
            dataInici = Car(P2)
        
            ExportaMURANO_Bancs idNorma43, dataInici
        Case "CALAIX"
            idNorma43 = Car(p1)
            dataInici = Car(P2)
        
            ExportaMURANO_Calaixos idNorma43, dataInici
        Case "MANUAL"
            idNorma43 = Car(p1)
            dataInici = Car(P2)
            nEmpresa = Car(P3)
        
            ExportaMURANO_Manuals nEmpresa, idNorma43, dataInici
        Case "MANUAL_DEL"
            idNorma43 = Car(p1)
            dataInici = Car(P2)
            nEmpresa = Car(P3)
        
            If nEmpresa = 0 Then
                Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampEnlaceNominas' ")
            Else
                Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampEnlaceNominas' ")
            End If
            If Not rsCtb.EOF Then EmpresaMurano = Trim(Left(rsCtb("Valor"), InStr(rsCtb("Valor"), " ")))
        
            Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(dataInici) & " where Param1 = '" & idNorma43 & "' and CodigoEmpresa='" & EmpresaMurano & "' and TipoExportacion='MANUAL'")
            While Not rsHist.EOF
                MuranoExecute "Delete from [WEB].[Sage].[dbo].[Movimientos] where Asiento = " & rsHist("Asiento") & " and CodigoEmpresa=" & EmpresaMurano
                rsHist.MoveNext
            Wend
        Case "SUBCUENTAS"
            ExportaMURANO_Subcuentas
      End Select
    'End If
  
    Exit Sub
    
norR:

Debug.Print "error"

Resume Next

End Sub


Sub ExportaMURANO_FacturaRebuda(idFactura As String, dataFactura As Date, numFactura As String, taulaFactura As String)
    Dim valorIva1 As Double, valorIva2 As Double, valorIva3 As Double, valorIva4 As Double, valorRec1 As Double, valorRec2 As Double, valorRec3 As Double, valorRec4 As Double
    Dim totalImporte As Double, Total As Double
    Dim rsCtb As rdoResultset, rsFac As rdoResultset, TeRetencio As Boolean
    Dim nEmpresa As String, D As Date
    Dim rsHist As rdoResultset, rsNA As rdoResultset, Rs3 As rdoResultset, rsSage As rdoResultset
    Dim codiProv As String, ContaVentes As String, numAsiento As String
    Dim RecConta As String, RecConta19 As String, RecConta21 As String, contraSin As String
    Dim BaseSinIva As Double, SinIva As Double
    Dim t As Integer
    Dim ref As String
    Dim i As Integer
    Dim fechaVencimiento As Date
    Dim i1 As Integer, baseIva1() As Double, iva1() As Double, rec1() As Double, Contra1() As String, baseIva1Total As Double, Iva1Total As Double, Rec1Total As Double
    Dim i2 As Integer, baseIva2() As Double, iva2() As Double, rec2() As Double, Contra2() As String, baseIva2Total As Double, Iva2Total As Double, Rec2Total As Double
    Dim i3 As Integer, baseIva3() As Double, iva3() As Double, rec3() As Double, Contra3() As String, baseIva3Total As Double, Iva3Total As Double, Rec3Total As Double
    Dim i4 As Integer, baseIva4() As Double, iva4() As Double, rec4() As Double, Contra4() As String, baseIva4Total As Double, Iva4Total As Double, Rec4Total As Double
    Dim nifEmpMurano As String, rsEmpMurano As rdoResultset
    Dim rsCActividad As rdoResultset, actividad As String
    Dim rsIntraComunitaria As rdoResultset, intraComunitaria As Boolean

    On Error GoTo noExportat
    
    Informa "MURANO Factura rebuda " & numFactura
    
    D = dataFactura
    Set rsFac = Db.OpenResultset("select (case when charindex('[Retencio',isnull(reservat,'Directa'))  = 0 then 0 else 1 end) TeRetencio, * from " & taulaFactura & " where idFactura='" & idFactura & "'")
    If Not rsFac.EOF Then
        'BUSCAMOS EMPRESA CORRESPONDIENTE EN SAGE POR NIF -----------------------------------------------------------------------------
        nEmpresa = rsFac("ClientCodi")
        
        nDigitos = DigitosCC(CDbl(nEmpresa))
        
        If nEmpresa = "0" Then
            Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampNif' ")
            Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampActividadSAGE' ")
        Else
            Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampNif' ")
            Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampActividadSAGE' ")
        End If
        
        actividad = ""
        If Not rsCActividad.EOF Then actividad = rsCActividad("valor")
        
        If Not rsCtb.EOF Then nifEmpMurano = rsCtb("valor")
        Set rsEmpMurano = Db.OpenResultset("select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'")
        If Not rsEmpMurano.EOF Then
            EmpresaMurano = rsEmpMurano("CodigoEmpresa")
        Else
            GoTo noMurano
        End If

        'NÚMERO DE ASIENTO HIT -------------------------------------------------------------------------------------------------------
        Set rsNA = Db.OpenResultset("select isnull(max(Asiento), 0) + 1 NumAsiento from " & TaulaHistoricoMURANO(D) & " where CodigoEmpresa = " & EmpresaMurano & " and year(fechaAsiento) = " & Year(D))
        If Not rsNA.EOF Then numAsiento = rsNA("NumAsiento")

        'REVISAMOS SI YA ESTABA EXPORTADA Y SI ESTÁ PENDIENTE O EL TRASPASO ES ERRÓNEO SE ELIMINA DE LA TABLA INTERMEDIA Y SE VUELVE A TRASPASAR
        'Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(D) & " where month(FechaAsiento)=" & Month(D) & " and day(FechaAsiento)=" & Day(D) & " and Param2 = '" & idFactura & "' and TipoExportacion='FACTURA_REBUDA' and CodigoEmpresa=" & EmpresaMurano)
        Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(D) & " where Param2 = '" & idFactura & "' and TipoExportacion='FACTURA_REBUDA' and CodigoEmpresa=" & EmpresaMurano)
        If Not rsHist.EOF Then numAsiento = rsHist("Asiento")
        While Not rsHist.EOF
            ExecutaComandaSql "Delete from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & rsHist("Asiento") & " and (statusTraspasadoIME=0 or statusTraspasadoIME=2) and CodigoEmpresa=" & EmpresaMurano & " and ejercicio=" & Year(D)
            rsHist.MoveNext
        Wend
        
        'SI YA ESTÁ TRASPASADA A SAGE CORRECTAMENTE NO VOLVEMOS A TRASPASAR
        Set rsSage = Db.OpenResultset("select * from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & numAsiento & " and statusTraspasadoIME=1 and CodigoEmpresa=" & EmpresaMurano & " and ejercicio=" & Year(D))
        If Not rsSage.EOF Then GoTo jaExportat
    
        TipusDeIva valorIva1, valorIva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, dataFactura
        
        codiProv = rsFac("EmpresaCodi")
        
        intraComunitaria = False
        Set rsIntraComunitaria = Db.OpenResultset("select isnull(pais, 'ES') pais from ccproveedores where id='" & codiProv & "'")
        If Not rsIntraComunitaria.EOF Then
            If rsIntraComunitaria("pais") <> "ES" And rsIntraComunitaria("pais") <> "" Then
                intraComunitaria = True
            End If
        End If
    
'BASES CON IVA --------------------------------------------------------------------------------------------------------------------
        i1 = 0: i2 = 0: i3 = 0: i4 = 0
        Set Rs3 = Db.OpenResultset("Select tipusIva, rec, isnull(referencia,'') referencia, SUM(import) import from [" & tablaFacturaProformaData(D) & "] where idfactura = '" & idFactura & "' and tipusiva<>0 group by tipusIva, rec, referencia ")
        While Not Rs3.EOF
            ref = Rs3("Referencia")
            
            If ref = "" Then
                If Rs3("TipusIva") = 4 Then
                    ref = "623000000"
                Else
                    ref = "600000000"
                End If
            End If
            
            If Len(ref) < nDigitos Then
                ref = Mid(ref, 1, 4) & Left("000000", nDigitos - Len(ref)) & Right(ref, 4)
            End If
        
            If Rs3("TipusIva") = 1 Then
                i1 = i1 + 1
                
                ReDim Preserve baseIva1(i1)
                baseIva1(i1) = Round(Rs3("Import"), 2)
                ReDim Preserve iva1(i1)
                iva1(i1) = Round(Rs3("Import") * (valorIva1 / 100), 2)
                ReDim Preserve rec1(i1)
                rec1(i1) = Round(Rs3("Import") * (Rs3("rec") / 100), 2)
                ReDim Preserve Contra1(i1)
                Contra1(i1) = ref
                If Contra1(i1) = 0 Then Contra1(i1) = "600000000"
            End If
        
            If Rs3("TipusIva") = 2 Then
                i2 = i2 + 1
                ReDim Preserve baseIva2(i2)
                baseIva2(i2) = Round(Rs3("Import"), 2)
                ReDim Preserve iva2(i2)
                iva2(i2) = Round(Rs3("Import") * (valorIva2 / 100), 2)
                ReDim Preserve rec2(i2)
                rec2(i2) = Round(Rs3("Import") * (Rs3("rec") / 100), 2)
                ReDim Preserve Contra2(i2)
                Contra2(i2) = ref
                If Contra2(i2) = 0 Then Contra2(i2) = "600000000"
            End If
                
            If Rs3("TipusIva") = 3 Then
                i3 = i3 + 1
                ReDim Preserve baseIva3(i3)
                baseIva3(i3) = Round(Rs3("Import"), 2)
                ReDim Preserve iva3(i3)
                iva3(i3) = Round(Rs3("Import") * (valorIva3 / 100), 2)
                ReDim Preserve rec3(i3)
                rec3(i3) = Round(Rs3("Import") * (Rs3("rec") / 100), 2)
                ReDim Preserve Contra3(i3)
                Contra3(i3) = ref
                If Contra3(i3) = 0 Then Contra3(i3) = "600000000"
            End If
            
            If Rs3("TipusIva") = 4 Then
                i4 = i4 + 1
                ReDim Preserve baseIva4(i4)
                baseIva4(i4) = Round(Rs3("Import"), 2)
                ReDim Preserve iva4(i4)
                iva4(i4) = Round(Rs3("Import") * (valorIva4 / 100), 2)
                ReDim Preserve rec4(i4)
                rec4(i4) = Round(Rs3("Import") * (Rs3("rec") / 100), 2)
                ReDim Preserve Contra4(i4)
                Contra4(i4) = ref
                If Contra4(i4) = 0 Then Contra4(i4) = "623000000"
            End If
            
            
            Rs3.MoveNext
        Wend
        Rs3.Close
        
        'REDONDEOS Y TOTALES
        baseIva1Total = 0
        Iva1Total = 0
        Rec1Total = 0
        If i1 > 0 Then
            For i = 1 To i1
                iva1(i) = Round(iva1(i), 3)
                Iva1Total = Iva1Total + iva1(i)
                Rec1Total = Rec1Total + rec1(i)
                
                baseIva1(i) = Round(baseIva1(i), 2)
                baseIva1Total = baseIva1Total + baseIva1(i)
            Next i
        End If
        
        baseIva2Total = 0
        Iva2Total = 0
        Rec2Total = 0
        If i2 > 0 Then
            For i = 1 To i2
                iva2(i) = Round(iva2(i), 3)
                Iva2Total = Iva2Total + iva2(i)
                Rec2Total = Rec2Total + rec2(i)
                
                baseIva2(i) = Round(baseIva2(i), 2)
                baseIva2Total = baseIva2Total + baseIva2(i)
            Next i
        End If
        
        baseIva3Total = 0
        Iva3Total = 0
        Rec3Total = 0
        If i3 > 0 Then
            For i = 1 To i3
                iva3(i) = Round(iva3(i), 3)
                Iva3Total = Iva3Total + iva3(i)
                Rec3Total = Rec3Total + rec3(i)
                
                baseIva3(i) = Round(baseIva3(i), 2)
                baseIva3Total = baseIva3Total + baseIva3(i)
            Next i
        End If
        
        baseIva4Total = 0
        Iva4Total = 0
        Rec4Total = 0
        If i4 > 0 Then
            For i = 1 To i4
                iva4(i) = Round(iva4(i), 3)
                Iva4Total = Iva4Total + iva4(i)
                Rec4Total = Rec4Total + rec4(i)
                
                baseIva4(i) = Round(baseIva4(i), 2)
                baseIva4Total = baseIva4Total + baseIva4(i)
            Next i
        End If
        
'~BASES CON IVA --------------------------------------------------------------------------------------------------------------------
        
'SENSE IVA -------------------------------------------------------------------------------------------------------------------------
        Dim teSenseIva As Boolean, impSenseIva As Double, contraSenseIva As String
        teSenseIva = False
        impSenseIva = 0
        contraSenseIva = ""
        Set Rs3 = Db.OpenResultset("Select * from [" & tablaFacturaProformaData(D) & "] where idfactura = '" & idFactura & "' and producteNom='Sin iva' ")
        If Not Rs3.EOF Then
            impSenseIva = Rs3("import")
            contraSenseIva = Rs3("referencia")
            If Len(contraSenseIva) < nDigitos Then
                contraSenseIva = Mid(contraSenseIva, 1, 4) & Left("000000", nDigitos - Len(contraSenseIva)) & Right(contraSenseIva, 4)
            End If
            
            teSenseIva = True
        End If
        Rs3.Close
'~SENSE IVA -------------------------------------------------------------------------------------------------------------------------
       
'RETENCIONS -------------------------------------------------------------------------------------------------------------------------
        Dim pctRetencion As Double, impRetencio As Double, contraRetencio As String
        TeRetencio = False
        pctRetencion = 0
        Set Rs3 = Db.OpenResultset("Select preu pctRetencion, import, referencia from [" & tablaFacturaProformaData(D) & "] where idfactura = '" & idFactura & "' and producteNom='IRPF' ")
        If Not Rs3.EOF Then
            pctRetencion = Rs3("pctRetencion")
            impRetencio = Round((baseIva1Total + baseIva2Total + baseIva3Total + baseIva4Total) * (pctRetencion / 100), 2)
            'impRetencio = rs3("import")
            contraRetencio = Rs3("referencia")
            If Len(contraRetencio) < nDigitos Then
                contraRetencio = Mid(contraRetencio, 1, 4) & Left("000000", nDigitos - Len(contraRetencio)) & Right(contraRetencio, 4)
            End If

            TeRetencio = True
        End If
        Rs3.Close
'~RETENCIONS -------------------------------------------------------------------------------------------------------------------------
        
'CUADRAR EL IMPORTE !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'DEBERÍA ESTAR CUADRADO EN LA BD
        totalImporte = Round((Iva1Total + Iva2Total + Iva3Total + Iva4Total + baseIva1Total + baseIva2Total + baseIva3Total + baseIva4Total + Rec1Total + Rec2Total + Rec3Total + Rec4Total + impSenseIva - impRetencio), 2)
        
        Total = rsFac("Total")
        
        If totalImporte <> 0 Then
            t = 0
            While Total <> totalImporte And t < 30
                If Abs(baseIva1Total) > 0 Then
                    baseIva1Total = baseIva1Total + (Total - totalImporte)
                    baseIva1(i1) = baseIva1(i1) + (Total - totalImporte)
                ElseIf Abs(baseIva2Total) > 0 Then
                    baseIva2Total = baseIva2Total + (Total - totalImporte)
                    baseIva2(i2) = baseIva2(i2) + (Total - totalImporte)
                ElseIf Abs(baseIva3Total) > 0 Then
                    baseIva3Total = baseIva3Total + (Total - totalImporte)
                    baseIva3(i3) = baseIva3(i3) + (Total - totalImporte)
                ElseIf Abs(baseIva4Total) > 0 Then
                    baseIva4Total = baseIva4Total + (Total - totalImporte)
                    baseIva4(i4) = baseIva4(i4) + (Total - totalImporte)
                End If

                totalImporte = Round((Iva1Total + Iva2Total + Iva3Total + Iva4Total + baseIva1Total + baseIva2Total + baseIva3Total + baseIva4Total + Rec1Total + Rec2Total + Rec3Total + Rec4Total - impRetencio), 2)
                t = t + 1
                DoEvents
            Wend
        Else
            GoTo noMurano
        End If
'~CUADRAR EL IMPORTE !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

        ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", '" & D & "', '" & numFactura & "', '" & idFactura & "', " & numAsiento & ", 'FACTURA_REBUDA')"
        
        Dim nifProv As String
        nifProv = rsFac("EmpNif")
        
        If nifProv = "" Then
            Set Rs3 = Db.OpenResultset("select * from ccproveedores where id='" & rsFac("EmpresaCodi") & "'")
            If Not Rs3.EOF Then
                nifProv = Rs3("nif")
            End If
            Rs3.Close
        End If
        
        Dim ccp As String
        ccp = codiContableProveedorMURANO(codiProv)
        fechaVencimiento = rsFac("dataVenciment")

        If intraComunitaria Then
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ccp, "R", "INTRA", numFactura, 0, totalImporte, Iva1Total, baseIva1Total, Iva2Total, baseIva2Total, Iva3Total, baseIva3Total, Iva4Total, baseIva4Total + impSenseIva, Rec1Total, Rec2Total, Rec3Total, pctRetencion, teSenseIva, nifProv, Left("F.num : " & numFactura & Space(40), 40), idFactura, "", "", "", "P", Format(fechaVencimiento, "dd/mm/yy")
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("477" & Right("000000000000", nDigitos - 3)) + valorIva2, "E", "INTRA", numFactura, 0, totalImporte * 0.1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifProv, Left("F.num : " & numFactura & Space(40), 40), idFactura, "", "", "", "P", Format(fechaVencimiento, "dd/mm/yy") '10% al haber de la cuenta 477
            PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, cIvaCompras(valorIva2), numFactura, totalImporte * 0.1, 0, actividad '10% al debe de la cuenta 472
        Else
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ccp, "R", "", numFactura, 0, totalImporte, Iva1Total, baseIva1Total, Iva2Total, baseIva2Total, Iva3Total, baseIva3Total, Iva4Total, baseIva4Total + impSenseIva, Rec1Total, Rec2Total, Rec3Total, pctRetencion, teSenseIva, nifProv, Left("F.num : " & numFactura & Space(40), 40), idFactura, "", "", "", "", Format(fechaVencimiento, "dd/mm/yy")
        End If

        'IVAS
        If i1 > 0 Then
            For i = 1 To i1
                PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, cIvaCompras(valorIva1), numFactura, iva1(i), 0, actividad
                PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, Contra1(i), numFactura, baseIva1(i), 0, actividad
                If Abs(rec1(i)) > 0 Then PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, cIvaCompras(valorIva1), numFactura, rec1(i), 0, actividad
            Next i
        End If
        
        If i2 > 0 Then
            For i = 1 To i2
                PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, cIvaCompras(valorIva2), numFactura, iva2(i), 0, actividad
                PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, Contra2(i), numFactura, baseIva2(i), 0, actividad
                If Abs(rec2(i)) > 0 Then PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, cIvaCompras(valorIva2), numFactura, rec2(i), 0, actividad
            Next i
        End If
        
        If i3 > 0 Then
            For i = 1 To i3
                PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, cIvaCompras(valorIva3), numFactura, iva3(i), 0, actividad
                PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, Contra3(i), numFactura, baseIva3(i), 0, actividad
                If Abs(rec3(i)) > 0 Then PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, cIvaCompras(valorIva3), numFactura, rec3(i), 0, actividad
            Next i
        End If
        
        If i4 > 0 Then
            For i = 1 To i4
                PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, cIvaCompras(valorIva4), numFactura, iva4(i), 0, actividad
                PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, Contra4(i), numFactura, baseIva4(i), 0, actividad
                If Abs(rec4(i)) > 0 Then PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, cIvaCompras(valorIva4), numFactura, rec4(i), 0, actividad
            Next i
        End If
        
        'FALTA IMPORTE SIN IVA EN EN DEBE
        If teSenseIva Then PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, contraSenseIva, numFactura, impSenseIva, 0, actividad
        
        'RETENCIÓN
        If TeRetencio Then PintaDetallIvaRebudesMURANO numAsiento, codiProv, D, contraRetencio, numFactura, 0, impRetencio, actividad
        
    End If
    
jaExportat:

    Exit Sub
    
noMurano:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR ExportaMURANO_FacturaRebuda", "[" & idFactura & "][" & dataFactura & "][" & numFactura & "][" & taulaFactura & "]", "", ""
    Exit Sub
    
noExportat:

    InsertFeineaAFer "SincroMURANOFacturaRebuda", "[" & idFactura & "]", "[" & dataFactura & "]", "[" & numFactura & "]", "[" & taulaFactura & "]"
End Sub


Sub PintaDetallIvaRebudesMURANO(numAsiento As String, codiProv As String, dataFactura As Date, nCuenta As String, numFactura As String, debe As Double, haber As Double, codigoActividad As String)
   
    debe = Round(debe, 2)
    haber = Round(haber, 2)
    
    If nCuenta = "0" Then nCuenta = "60000000"
    
    If debe <> 0 Or haber <> 0 Then
        'AsientoAddMURANO 0, numAsiento, dataFactura, nCuenta, codiContableProveedorMURANO(codiProv), Left("S/F " & numFactura & Space(25), 25), debe, haber
        AsientoAddMURANO_TS codigoActividad, 0, numAsiento, dataFactura, nCuenta, "", "", numFactura, debe, haber, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(40), 40), ""
    End If
    
End Sub

Sub ExportaANALITICA_FacturaRebuda(idFactura As String, dataFactura As Date, numFactura As String, taulaFactura As String)
    Dim rsFac As rdoResultset, rsEmpHit As rdoResultset, rsEmpSage As rdoResultset, rsSage As rdoResultset
    Dim numAsiento As String, sql As String, ordMovimiento As String, empresa As String, familia As String
    Dim Semana As Integer, anyo As Integer, empNif As String, empSage As String
    Dim nFactura As String, provCodi As String, provNom As String
    Dim rsRepartoTot As rdoResultset, rsReparto As rdoResultset
    Dim importeReparto As Double, importe As Double
    Dim tabAnalitica As String
    Dim rsProds As rdoResultset, rsMP As rdoResultset
    Dim asignado As Boolean, nProds As Integer
    
    On Error GoTo noExportat
    
    Informa "ANALITICA Factura " & numFactura & " data Factura " & dataFactura
    
    Semana = DatePart("WW", dataFactura, vbMonday, vbFirstJan1)
    anyo = Year(dataFactura)
    tabAnalitica = tablaAnaliticaSemanal(anyo, Semana)
        
    Set rsFac = Db.OpenResultset("select * from " & taulaFactura & " where idfactura='" & idFactura & "'")
    If Not rsFac.EOF Then
        empresa = rsFac("ClientCodi")
        nFactura = rsFac("numFactura")
        provCodi = rsFac("empresaCodi")
        provNom = rsFac("empNom")
        
        If empresa = "0" Then
            Set rsEmpHit = Db.OpenResultset("select * from constantsempresa where camp like 'CampNif'")
        Else
            Set rsEmpHit = Db.OpenResultset("select * from constantsempresa where camp like '" & empresa & "_CampNif'")
        End If
        If Not rsEmpHit.EOF Then empNif = rsEmpHit("valor")
    
        Set rsEmpSage = Db.OpenResultset("select * from " & dbSage & ".dbo.empresas where CifDni='" & empNif & "'")
        If Not rsEmpSage.EOF Then empSage = rsEmpSage("CodigoEmpresa")
        'empSage = "0"
        
        Set rsSage = Db.OpenResultset("select * from " & dbSage & ".dbo.Movimientos where ejercicio=" & Year(dataFactura) & " and Codigoempresa=" & empSage & " and comentario = 'F.Num : " & numFactura & "' and CargoAbono='H'")
        If Not rsSage.EOF Then
            numAsiento = rsSage("Asiento")
            'numAsiento = "0000"
           
            'POR PRODUCTO
            'sql = "select f.idFactura, f.producte MP, isnull(ap.CodiArticle, isnull(fp.codiarticle, 99999)) as article, f1.NOM familia, isnull(sum(f.import-(f.import*(f.desconte/100))), 0) Importe "
            'sql = sql & "from [" & tablaFacturaProformaData(dataFactura) & "] f "
            'sql = sql & "left join [" & tablaFacturaProforma(dataFactura) & "] fi on f.idfactura=fi.idfactura "
            'sql = sql & "left join articlespropietats ap on ap.valor=f.producte and ap.variable='MatPri' "
            'sql = sql & "Left Join "
            'sql = sql & "(select ap.codiArticle, f.kg, isnull(case when tipo ='MP' then materia else (select isnull(predeterminada, '') from ccMateriasPrimasBase where id=f.materia) end, '') mp "
            'sql = sql & "from formulaMasaDetalle f "
            'sql = sql & "left join articlesPropietats ap on cast(f.formula as nvarchar)=ap.valor and ap.variable='Formula' "
            'sql = sql & "where tipo in ('MPB', 'MP') and ap.codiarticle is not null) fp on fp.mp = f.producte "
            'sql = sql & "left join articles a on isnull(ap.CodiArticle, fp.codiarticle)=CAST(a.codi AS nvarchar) "
            'sql = sql & "left join families f3 on a.familia=f3.nom "
            'sql = sql & "left join families f2 on f2.nom=f3.pare "
            'sql = sql & "left join families f1 on f1.nom=f2.pare "
            'sql = sql & "left join ccproveedores p on fi.EmpresaCodi = p.id "
            'sql = sql & "Where f.idfactura='" & idFactura & "' "
            'sql = sql & "group by f.idFactura, f.producte, isnull(ap.CodiArticle, isnull(fp.codiarticle, 99999)), f1.NOM"
            
            sql = "Select f.producte, f.producteNom, isnull(sum(f.import-(f.import*(f.desconte/100))), 0) Importe "
            sql = sql & "From [" & tablaFacturaProformaData(dataFactura) & "] f "
            sql = sql & "Where f.idfactura='" & idFactura & "' "
            sql = sql & "Group by f.producte, f.producteNom"
            Set rsMP = Db.OpenResultset(sql)
            
            'ExecutaComandaSql "delete from " & tabAnalitica & " where Emp_HIT=" & Empresa & " and Emp_SAGE=" & empSage & " and Asiento_SAGE=" & numAsiento & " and Ejercicio= " & Year(dataFactura) & " and IdFactura = '" & idFactura & "' "
            ExecutaComandaSql "delete from " & tabAnalitica & " where Emp_HIT=" & empresa & " and Ejercicio= " & Year(dataFactura) & " and IdFactura = '" & idFactura & "' "

            importeReparto = 0
            While Not rsMP.EOF
                asignado = False
                importe = rsMP("Importe")
                
                'COMPROBAR SI TIENE ARTÍCULO DE VENTA ASOCIADO
                sql = "select isnull(a.codi, 0) article, isnull(a.nom, '') nom, isnull(f1.nom, '') familia "
                sql = sql & "from articlespropietats ap "
                sql = sql & "left join articles a on ap.CodiArticle=CAST(a.codi AS nvarchar) "
                sql = sql & "left join families f3 on a.familia=f3.nom "
                sql = sql & "left join families f2 on f2.nom=f3.pare "
                sql = sql & "left join families f1 on f1.nom=f2.pare "
                sql = sql & "where ap.variable='MatPri' and ap.valor='" & rsMP("producte") & "'"

                Set rsProds = Db.OpenResultset(sql)
                If Not rsProds.EOF Then
                    If rsProds("article") <> 0 Then
                        sql = "Insert into " & tabAnalitica & " (Emp_HIT, Emp_SAGE, Ejercicio, Asiento_SAGE, Semana, IdFactura, FechaFactura, NumFactura, ClienteProveedorCodi, ClienteProveedor, TipoFactura, TablaFactura, Producto, Familia, Importe, Msg_Importacion, MP) Values "
                        sql = sql & "(" & empresa & ", " & empSage & ", " & Year(dataFactura) & ", " & numAsiento & ", " & Semana & ", '" & idFactura & "', '" & dataFactura & "', '" & nFactura & "', '" & provCodi & "', '" & provNom & "', 'R', '" & taulaFactura & "', " & rsProds("article") & ", '" & rsProds("familia") & "', " & importe & ", 'OK', '" & rsMP("producte") & "') "
                        ExecutaComandaSql sql
                        asignado = True
                    End If
                    rsProds.Close
                End If
                
                If Not asignado Then
                    'COMPROBAR SI ESTA EN ALGUNA FÓRMULA
                    sql = "select count (*) n "
                    sql = sql & "from formulaMasaDetalle f "
                    sql = sql & "left join articlesPropietats ap on cast(f.formula as nvarchar)=ap.valor and ap.variable='Formula' "
                    sql = sql & "left join ccMateriasPrimasBase mpb on f.materia = mpb.id "
                    sql = sql & "left join ccMateriasPrimas mp on mpb.predeterminada = mp.id or f.materia=mp.id "
                    sql = sql & "where (f.materia='" & rsMP("producte") & "' or mpb.predeterminada = '" & rsMP("producte") & "') and ap.codiArticle is not null"
                    Set rsProds = Db.OpenResultset(sql)
                    If rsProds("n") > 0 Then
                        nProds = rsProds("n")
                        sql = "select ap.codiArticle article, isnull(f1.nom, '') familia, f.kg, mp.nombre "
                        sql = sql & "from formulaMasaDetalle f "
                        sql = sql & "left join articlesPropietats ap on cast(f.formula as nvarchar)=ap.valor and ap.variable='Formula' "
                        sql = sql & "left join articles a on ap.CodiArticle=CAST(a.codi AS nvarchar) "
                        sql = sql & "left join families f3 on a.familia=f3.nom "
                        sql = sql & "left join families f2 on f2.nom=f3.pare "
                        sql = sql & "left join families f1 on f1.nom=f2.pare "
                        sql = sql & "left join ccMateriasPrimasBase mpb on f.materia = mpb.id "
                        sql = sql & "left join ccMateriasPrimas mp on mpb.predeterminada = mp.id or f.materia=mp.id "
                        sql = sql & "where (f.materia='" & rsMP("producte") & "' or mpb.predeterminada = '" & rsMP("producte") & "') and ap.codiArticle is not null"
                        Set rsProds = Db.OpenResultset(sql)
                        While Not rsProds.EOF
                            sql = "Insert into " & tabAnalitica & " (Emp_HIT, Emp_SAGE, Ejercicio, Asiento_SAGE, Semana, IdFactura, FechaFactura, NumFactura, ClienteProveedorCodi, ClienteProveedor, TipoFactura, TablaFactura, Producto, Familia, Importe, Msg_Importacion, MP) Values "
                            sql = sql & "(" & empresa & ", " & empSage & ", " & Year(dataFactura) & ", " & numAsiento & ", " & Semana & ", '" & idFactura & "', '" & dataFactura & "', '" & nFactura & "', '" & provCodi & "', '" & provNom & "', 'R', '" & taulaFactura & "', " & rsProds("article") & ", '" & rsProds("familia") & "', " & Round(importe / nProds, 2) & ", 'OK', '" & rsMP("producte") & "') "
                            ExecutaComandaSql sql
                            
                            rsProds.MoveNext
                        Wend
                        asignado = True
                    End If
                    rsProds.Close
                End If
                
                If Not asignado Then
                    'LO REPARTIMOS ENTRE TODOS LOS PRODUCTOS DE LA SEMANA
                    'Set rsRepartoTot = Db.OpenResultset("select COUNT (DISTINCT PRODUCTO) nProds from " & tabAnalitica & " an where semana=" & Semana & " and tipoFactura='E' and emp_HIT=" & Empresa & " AND MSG_IMPORTACION='OK'")
                    'If Not rsRepartoTot.EOF Then
                    '    If rsRepartoTot("nProds") > 0 Then nProds = rsRepartoTot("nProds")
                    'End If
                    
                    'Set rsReparto = Db.OpenResultset("select distinct producto, familia, MP from " & tabAnalitica & " an where semana=" & Semana & " and tipoFactura='E' and emp_HIT=" & Empresa & " AND MSG_IMPORTACION='OK'")
                    'While Not rsReparto.EOF
                    '    sql = "Insert into " & tabAnalitica & " (Emp_HIT, Emp_SAGE, Ejercicio, Asiento_SAGE, Semana, IdFactura, FechaFactura, NumFactura, ClienteProveedorCodi, ClienteProveedor, TipoFactura, TablaFactura, Producto, Familia, Importe, Msg_Importacion, MP) Values "
                    '    sql = sql & "(" & Empresa & ", " & empSage & ", " & Year(dataFactura) & ", " & numAsiento & ", " & Semana & ", '" & idFactura & "', '" & dataFactura & "', '" & nFactura & "', '" & provCodi & "', '" & provNom & "', 'R', '" & taulaFactura & "', " & rsReparto("producto") & ", '" & rsReparto("familia") & "', " & Round(importe / nProds, 2) & ", 'OK', '" & rsMP("ProducteNom") & "') "
                    '    ExecutaComandaSql sql
                
                    '    rsReparto.MoveNext
                    'Wend
                    
                    sql = "Insert into " & tabAnalitica & " (Emp_HIT, Emp_SAGE, Ejercicio, Asiento_SAGE, Semana, IdFactura, FechaFactura, NumFactura, ClienteProveedorCodi, ClienteProveedor, TipoFactura, TablaFactura, Producto, Familia, Importe, Msg_Importacion, MP) Values "
                    sql = sql & "(" & empresa & ", " & empSage & ", " & Year(dataFactura) & ", " & numAsiento & ", " & Semana & ", '" & idFactura & "', '" & dataFactura & "', '" & nFactura & "', '" & provCodi & "', '" & provNom & "', 'R', '" & taulaFactura & "', 99999, 'GASTOS SIN IMPUTAR', " & Round(importe, 2) & ", 'OK', '" & rsMP("ProducteNom") & "') "
                    ExecutaComandaSql sql
                    
                End If
                
                rsMP.MoveNext
            Wend
            rsMP.Close
        Else
            sql = "Insert into " & tabAnalitica & " (Emp_HIT, Emp_SAGE, Ejercicio, Semana, IdFactura, FechaFactura, NumFactura, ClienteProveedorCodi, ClienteProveedor, TipoFactura, TablaFactura, Msg_Importacion) Values "
            sql = sql & "(" & empresa & ", " & empSage & ", " & Year(dataFactura) & ", " & Semana & ", '" & idFactura & "', '" & dataFactura & "', '" & nFactura & "', '" & provCodi & "', '" & provNom & "', 'R', '" & taulaFactura & "', 'ERROR: NO IMPORTADO EN SAGE!!!')"
            ExecutaComandaSql sql
        End If
    End If
    
    Exit Sub
    
noExportat:

 
End Sub



Sub ExportaANALITICA_FacturaEmesa(idFactura As String, dataFactura As Date, numFactura As String, taulaFactura As String)
    Dim rsFac As rdoResultset, rsHist As rdoResultset, rsSage As rdoResultset, rsProds As rdoResultset
    Dim numAsiento As String, sql As String, ordMovimiento As String
    Dim Semana As Integer, anyo As Integer, sqlData As String, empSage As String, empresa As String
    Dim rsEmpHit As rdoResultset, rsEmpSage As rdoResultset, empNif As String
    Dim nFactura As String, clientCodi As String, clientNom As String, serieFac As String
    Dim tabAnalitica As String
    
    On Error GoTo noExportat
    
    Informa "ANALITICA Factura " & numFactura

    Semana = DatePart("WW", dataFactura, vbMonday, vbFirstJan1)
    anyo = Year(dataFactura)
    tabAnalitica = tablaAnaliticaSemanal(anyo, Semana)
        
    Set rsFac = Db.OpenResultset("select * from [" & taulaFactura & "] where idfactura='" & idFactura & "'")
    If Not rsFac.EOF Then
        empresa = rsFac("EmpresaCodi")
        serieFac = rsFac("Serie")
        nFactura = rsFac("NumFactura")
        clientCodi = rsFac("ClientCodi")
        clientNom = rsFac("ClientNom")
        
        If empresa = "0" Then
            Set rsEmpHit = Db.OpenResultset("select * from constantsempresa where camp like 'CampNif'")
        Else
            Set rsEmpHit = Db.OpenResultset("select * from constantsempresa where camp like '" & empresa & "_CampNif'")
        End If
        If Not rsEmpHit.EOF Then empNif = rsEmpHit("valor")
    
        Set rsEmpSage = Db.OpenResultset("select * from " & dbSage & ".dbo.empresas where CifDni='" & empNif & "'")
        If Not rsEmpSage.EOF Then empSage = rsEmpSage("CodigoEmpresa")
            
        Set rsSage = Db.OpenResultset("select * from " & dbSage & ".dbo.Movimientos where ejercicio=" & Year(dataFactura) & " and Codigoempresa=" & empSage & " and comentario = 'F.Num : " & numFactura & "' and CargoAbono='D'")
        If Not rsSage.EOF Then
            numAsiento = rsSage("Asiento")
               
            'POR PRODUCTO
            sqlData = "select fd.idFactura, fd.producte, isnull(sum(import-(import*(fd.desconte/100))), 0) Importe , f1.Nom Familia "
            sqlData = sqlData & "from [" & NomTaulaFacturaData(dataFactura) & "] fd "
            sqlData = sqlData & "left join articles a on fd.producte=a.codi "
            sqlData = sqlData & "left join families f3 on a.familia=f3.nom "
            sqlData = sqlData & "left join families f2 on f2.nom=f3.pare "
            sqlData = sqlData & "left join families f1 on f1.nom=f2.pare "
            sqlData = sqlData & "where fd.idfactura='" & idFactura & "' "
            sqlData = sqlData & "group by fd.idFactura, fd.client, fd.producte , f1.Nom"
            Set rsProds = Db.OpenResultset(sqlData)
            
            ExecutaComandaSql "delete from " & tabAnalitica & " where Emp_HIT=" & empresa & " and Emp_SAGE=" & empSage & " and Asiento_SAGE=" & numAsiento & " and Ejercicio= " & Year(dataFactura) & " and IdFactura = '" & idFactura & "' "
    
            While Not rsProds.EOF
                sql = "Insert into " & tabAnalitica & " (Emp_HIT, Emp_SAGE, Ejercicio, Asiento_SAGE, Semana, IdFactura, FechaFactura, NumFactura, ClienteProveedorCodi, ClienteProveedor, TipoFactura, TablaFactura, Producto, Familia, Importe, Msg_Importacion) Values "
                sql = sql & "(" & empresa & ", " & empSage & ", " & Year(dataFactura) & ", " & numAsiento & ", " & Semana & ", '" & idFactura & "', '" & dataFactura & "', '" & serieFac & nFactura & "', '" & clientCodi & "', '" & clientNom & "', 'E', '" & taulaFactura & "', " & rsProds("producte") & ", '" & rsProds("familia") & "', " & rsProds("importe") & ", 'OK') "
                ExecutaComandaSql sql
                
                rsProds.MoveNext
            Wend
        Else
            sql = "Insert into " & tabAnalitica & " (Emp_HIT, Emp_SAGE, Ejercicio, Semana, IdFactura, FechaFactura, NumFactura, ClienteProveedorCodi, ClienteProveedor, TipoFactura, TablaFactura, Msg_Importacion) Values "
            sql = sql & "(" & empresa & ", " & empSage & ", " & Year(dataFactura) & ", " & Semana & ", '" & idFactura & "', '" & dataFactura & "', '" & serieFac & nFactura & "', '" & clientCodi & "', '" & clientNom & "', 'E', '" & taulaFactura & "', 'ERROR: NO IMPORTADO EN SAGE!!!')"
            ExecutaComandaSql sql
        End If
    End If
noExportat:

End Sub


Sub ExportaMURANO_FacturaEmesa(idFactura As String, dataFactura As Date, numFactura As String, taulaFactura As String)
    Dim valorIva1 As Double, valorIva2 As Double, valorIva3 As Double, valorIva4 As Double, valorRec1 As Double, valorRec2 As Double, valorRec3 As Double, valorRec4 As Double
    Dim iva1 As Double, iva2 As Double, iva3 As Double, iva4 As Double, baseIva1 As Double, baseIva2 As Double, baseIva3 As Double, baseIva4 As Double, rec1 As Double, rec2 As Double, rec3 As Double
    Dim BaseRec1  As Double, BaseRec2  As Double, BaseRec3  As Double, IvaRec1  As Double, IvaRec2  As Double, IvaRec3 As Double, totalImporte As Double
    Dim rsCtb As rdoResultset, rsFac As rdoResultset, TeRetencio As Boolean
    Dim nEmpresa As String, D As Date, serie As String, cNif As String
    Dim rsHist As rdoResultset, rsNA As rdoResultset, rsSage As rdoResultset
    Dim clientCodi As String, ContaVentes As String, numAsiento As String
    Dim Tenim1 As Boolean, Tenim2 As Boolean, Tenim3 As Boolean, Tenim4 As Boolean, Tenim5 As Boolean, Tenim6 As Boolean
    Dim TenimRec1 As Boolean, TenimRec2 As Boolean, TenimRec3 As Boolean
    Dim sqlData As String, rsData As rdoResultset
    Dim haberReten As Double, nCuenta As String
    Dim totalFamilias As Double
    Dim fechaVencimiento As Date
    Dim SqlIBEE As String, rsIBEE As rdoResultset
    Dim rsCActividad As rdoResultset, actividad As String
    
    On Error GoTo noExportat
'If Month(dataFactura) <> 1 Then GoTo noExportat
    
    Informa "MURANO Factura " & numFactura
    
    D = dataFactura
    If Left(taulaFactura, 1) <> "[" Then taulaFactura = "[" & taulaFactura & "]"
    Set rsFac = Db.OpenResultset("select (case when charindex('[Retencio',isnull(reservat,'Directa'))  = 0 then 0 else 1 end) TeRetencio, * from " & taulaFactura & " where idFactura='" & idFactura & "'")
    If Not rsFac.EOF Then
        nEmpresa = rsFac("EmpresaCodi")
        
        nDigitos = DigitosCC(CDbl(nEmpresa))
        
        If nEmpresa = "0" Then
            Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampNif' ")
            Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampActividadSAGE' ")
        Else
            Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampNif' ")
            Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampActividadSAGE' ")
        End If
        
        actividad = ""
        If Not rsCActividad.EOF Then actividad = rsCActividad("valor")
        If actividad = "0" Then actividad = ""
        
        Dim nifEmpMurano As String, rsEmpMurano As rdoResultset
        
        If Not rsCtb.EOF Then nifEmpMurano = rsCtb("valor")
        Set rsEmpMurano = Db.OpenResultset("select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'")
        If Not rsEmpMurano.EOF Then
            EmpresaMurano = rsEmpMurano("CodigoEmpresa")
        Else
            GoTo noExportat
        End If
    
        Set rsNA = Db.OpenResultset("select isnull(max(Asiento), 0) + 1 NumAsiento from " & TaulaHistoricoMURANO(D) & " where CodigoEmpresa = " & EmpresaMurano & " and year(fechaAsiento) = " & Year(D))
        If Not rsNA.EOF Then numAsiento = rsNA("NumAsiento")

        'Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(D) & " where month(FechaAsiento)=" & Month(D) & " and day(FechaAsiento)=" & Day(D) & " and Param1 = '" & numFactura & "' and Param2 = '" & idFactura & "' and CodigoEmpresa=" & EmpresaMurano & " and TipoExportacion='FACTURA'")
        Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(D) & " where month(FechaAsiento)=" & Month(D) & " and day(FechaAsiento)=" & Day(D) & " and Param2 = '" & idFactura & "' and CodigoEmpresa=" & EmpresaMurano & " and TipoExportacion='FACTURA'")
        If Not rsHist.EOF Then numAsiento = rsHist("Asiento")
        While Not rsHist.EOF
            'MuranoExecute "Delete from [WEB].[Sage].[dbo].[Movimientos] where Asiento = " & rsHist("Asiento") & " and CodigoEmpresa=" & EmpresaMurano
            ExecutaComandaSql "Delete from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & rsHist("Asiento") & " and (statusTraspasadoIME=0 or statusTraspasadoIME=2) and CodigoEmpresa=" & EmpresaMurano & " and ejercicio=" & Year(D)
            rsHist.MoveNext
        Wend
        
        Set rsSage = Db.OpenResultset("select * from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & numAsiento & " and statusTraspasadoIME=1 and CodigoEmpresa=" & EmpresaMurano & " and ejercicio=" & Year(D))
        If Not rsSage.EOF Then GoTo jaExportat  'Si ya se ha traspasado correctamente a MURANO no volvemos a traspasarla
        
        TipusDeIva valorIva1, valorIva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, dataFactura
        
        iva1 = Round(rsFac("Iva1"), 2)
        iva2 = Round(rsFac("Iva2"), 2)
        iva3 = Round(rsFac("Iva3"), 2)
        iva4 = Round(rsFac("Iva4"), 2)
        baseIva1 = Round(rsFac("baseIva1"), 2)
        baseIva2 = Round(rsFac("baseIva2"), 2)
        baseIva3 = Round(rsFac("baseIva3"), 2)
        baseIva4 = Round(rsFac("baseIva4"), 2)
        rec1 = Round(rsFac("Rec1"), 2)
        rec2 = Round(rsFac("Rec2"), 2)
        rec3 = Round(rsFac("Rec3"), 2)
        IvaRec1 = Round(rsFac("IvaRec1"), 2)
        IvaRec2 = Round(rsFac("IvaRec2"), 2)
        IvaRec3 = Round(rsFac("IvaRec3"), 2)
        BaseRec1 = Round(rsFac("BaseRec1"), 2)
        BaseRec2 = Round(rsFac("BaseRec2"), 2)
        BaseRec3 = Round(rsFac("BaseRec3"), 2)
        
    
        'IBEE
        Dim ibee As Double, ivaIbee As Double
        ibee = 0
        SqlIBEE = "select sum( "
        SqlIBEE = SqlIBEE & "case when isnumeric(SUBSTRING(referencia, charindex('IBEE', referencia)+5, charindex(']', referencia, charindex('IBEE', referencia))-len('[IBEE:')-1))=1 then "
        SqlIBEE = SqlIBEE & "SUBSTRING(referencia, charindex('IBEE', referencia)+5, charindex(']', referencia, charindex('IBEE', referencia))-len('[IBEE:')-1)*(servit-tornat) else 0 end) IBEE "
        SqlIBEE = SqlIBEE & "from [" & NomTaulaFacturaData(dataFactura) & "] where idfactura='" & idFactura & "' and referencia like '%IBEE:%' "
        Set rsIBEE = Db.OpenResultset(SqlIBEE)
        
        If Not rsIBEE.EOF Then
            If rsIBEE("IBEE") <> 0 Then
                ibee = Round(rsIBEE("ibee"), 2)
                ivaIbee = Round(ibee * 0.1, 2)
            End If
        End If
    
        totalImporte = Round((ibee + ivaIbee + iva1 + iva2 + iva3 + iva4 + baseIva1 + baseIva2 + baseIva3 + baseIva4 + rec1 + rec2 + rec3 + BaseRec1 + BaseRec2 + BaseRec3 + IvaRec1 + IvaRec2 + IvaRec3), 2)
    
        If totalImporte <> (ibee + ivaIbee + iva1 + iva2 + iva3 + iva4 + baseIva1 + baseIva2 + baseIva3 + baseIva4 + rec1 + rec2 + rec3 + BaseRec1 + BaseRec2 + BaseRec3 + IvaRec1 + IvaRec2 + IvaRec3) Then
            If baseIva1 > 0 Then
                iva1 = iva1 + (totalImporte - (ibee + ivaIbee + iva1 + iva2 + iva3 + iva4 + baseIva1 + baseIva2 + baseIva3 + baseIva4 + rec1 + rec2 + rec3 + BaseRec1 + BaseRec2 + BaseRec3 + IvaRec1 + IvaRec2 + IvaRec3))
            ElseIf baseIva2 > 0 Then
                iva2 = iva2 + (totalImporte - (ibee + ivaIbee + iva1 + iva2 + iva3 + iva4 + baseIva1 + baseIva2 + baseIva3 + baseIva4 + rec1 + rec2 + rec3 + BaseRec1 + BaseRec2 + BaseRec3 + IvaRec1 + IvaRec2 + IvaRec3))
            ElseIf baseIva3 > 0 Then
                iva3 = iva3 + (totalImporte - (ibee + ivaIbee + iva1 + iva2 + iva3 + iva4 + baseIva1 + baseIva2 + baseIva3 + baseIva4 + rec1 + rec2 + rec3 + BaseRec1 + BaseRec2 + BaseRec3 + IvaRec1 + IvaRec2 + IvaRec3))
            ElseIf ibee > 0 Then
                ivaIbee = ivaIbee + (totalImporte - (ibee + ivaIbee + iva1 + iva2 + iva3 + iva4 + baseIva1 + baseIva2 + baseIva3 + baseIva4 + rec1 + rec2 + rec3 + BaseRec1 + BaseRec2 + BaseRec3 + IvaRec1 + IvaRec2 + IvaRec3))
            ElseIf BaseRec1 > 0 Then
                IvaRec1 = IvaRec1 + (totalImporte - (ibee + ivaIbee + iva1 + iva2 + iva3 + iva4 + baseIva1 + baseIva2 + baseIva3 + baseIva4 + rec1 + rec2 + rec3 + BaseRec1 + BaseRec2 + BaseRec3 + IvaRec1 + IvaRec2 + IvaRec3))
            ElseIf BaseRec2 > 0 Then
                IvaRec2 = IvaRec2 + (totalImporte - (ibee + ivaIbee + iva1 + iva2 + iva3 + iva4 + baseIva1 + baseIva2 + baseIva3 + baseIva4 + rec1 + rec2 + rec3 + BaseRec1 + BaseRec2 + BaseRec3 + IvaRec1 + IvaRec2 + IvaRec3))
            ElseIf BaseRec3 > 0 Then
                IvaRec3 = IvaRec3 + (totalImporte - (ibee + ivaIbee + iva1 + iva2 + iva3 + iva4 + baseIva1 + baseIva2 + baseIva3 + baseIva4 + rec1 + rec2 + rec3 + BaseRec1 + BaseRec2 + BaseRec3 + IvaRec1 + IvaRec2 + IvaRec3))
            End If
        End If
    
        Dim pctRetencio As Double
        pctRetencio = 0
        TeRetencio = rsFac("TeRetencio")
        If TeRetencio Then
            pctRetencio = CDbl(Split(Split(rsFac("Reservat"), "[Retencio")(1), "]")(0))
            totalImporte = totalImporte - Round((ibee + baseIva1 + baseIva2 + baseIva3 + baseIva4) * (pctRetencio / 100), 2)
        End If
    
        clientCodi = Trim(rsFac("clientCodi"))
        If totalImporte < 0 Then
            ContaVentes = cVentaMercaderiesDevolucions(clientCodi, TeRetencio)
        Else
            ContaVentes = cVentaMercaderies(clientCodi, TeRetencio)
        End If
    
        ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", '" & D & "', " & numFactura & ", '" & idFactura & "', " & numAsiento & ", 'FACTURA')"
        'AsientoAddMURANO 0, numAsiento, D, codiContable2(clientCodi), ContaVentes, Left("F.num : " & Abs(numFactura) & Space(25), 25), totalImporte, 0
        
        fechaVencimiento = rsFac("DataVenciment")
        serie = rsFac("Serie")
        cNif = rsFac("ClientNif")
        AsientoAddMURANO_TS actividad, 0, numAsiento, D, codiContable2(clientCodi), "E", serie, numFactura, totalImporte, 0, iva1 + IvaRec1, baseIva1 + BaseRec1, ivaIbee + iva2 + IvaRec2, ibee + baseIva2 + BaseRec2, iva3 + IvaRec3, baseIva3 + BaseRec3, iva4, baseIva4, rec1, rec2, rec3, pctRetencio, False, cNif, Left("F.num : " & numFactura & Space(25), 25), idFactura, "", "", "", "", Format(fechaVencimiento, "dd/mm/yy")
 
        Tenim1 = False
        Tenim2 = False
        Tenim3 = False
        Tenim4 = False
        Tenim5 = False
        Tenim6 = False
        TenimRec1 = False
        TenimRec2 = False
        TenimRec3 = False
    
        If Abs(iva1) > 0 Then Tenim1 = True
        If Abs(iva2) > 0 Then Tenim2 = True
        If Abs(iva3) > 0 Then Tenim3 = True
    
        If (Tenim1) And Abs(baseIva1) > 0 Then Tenim4 = True
        If (Tenim2) And Abs(baseIva2) > 0 Then Tenim5 = True
        If (Tenim3) And Abs(baseIva3) > 0 Then Tenim6 = True
        'If Abs(baseIva4) > 0 Then Tenim7 = True
        
        If Abs(rec1) > 0 Then TenimRec1 = True
        If Abs(rec2) > 0 Then TenimRec2 = True
        If Abs(rec3) > 0 Then TenimRec3 = True
    
        If (TenimRec1) And Abs(BaseRec1) > 0 Then Tenim4 = True
        If (TenimRec2) And Abs(BaseRec2) > 0 Then Tenim5 = True
        If (TenimRec3) And Abs(BaseRec3) > 0 Then Tenim6 = True
        
        'IVAS 47700004,47700010,47700021
        If Tenim1 Then PintaDetallIvaMURANO actividad, numAsiento, clientCodi, dataFactura, iva1, baseIva1, numFactura, valorIva1 & ".00", TenimRec1
        If Tenim2 Then PintaDetallIvaMURANO actividad, numAsiento, clientCodi, dataFactura, iva2, baseIva2, numFactura, valorIva2 & ".00", TenimRec2
        If Tenim3 Then PintaDetallIvaMURANO actividad, numAsiento, clientCodi, dataFactura, iva3, baseIva3, numFactura, valorIva3 & ".00", TenimRec3
        'If Tenim7 Then PintaDetallIvaMURANO numAsiento, clientCodi, dataFactura, Iva4, baseIva4, numFactura, valorIva4 & ".00", TenimRec3
       
        'IVAS 47700004,47700010,47700021
        If TenimRec1 Then PintaDetallIvaMURANO actividad, numAsiento, clientCodi, dataFactura, rec1, BaseRec1, numFactura, valorIva1 & ".00", True
        If TenimRec2 Then PintaDetallIvaMURANO actividad, numAsiento, clientCodi, dataFactura, rec2, BaseRec2, numFactura, valorIva2 & ".00", True
        If TenimRec3 Then PintaDetallIvaMURANO actividad, numAsiento, clientCodi, dataFactura, rec3, BaseRec3, numFactura, valorIva3 & ".00", True
       
        'IVAS CON RECS 47700404,47700408,47700418
        If TenimRec1 Then PintaDetallIvaMURANO actividad, numAsiento, clientCodi, dataFactura, IvaRec1, 0, numFactura, "0.5"
        If TenimRec2 Then PintaDetallIvaMURANO actividad, numAsiento, clientCodi, dataFactura, IvaRec2, 0, numFactura, "1"
        If TenimRec3 Then PintaDetallIvaMURANO actividad, numAsiento, clientCodi, dataFactura, IvaRec3, 0, numFactura, "4"
       
        If ibee <> 0 Then
            AsientoAddMURANO_TS actividad, 0, numAsiento, dataFactura, "477000010", "", "", numFactura, 0, ivaIbee, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(25), 25), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, dataFactura, "475000055", "", "", numFactura, 0, ibee, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(25), 25), ""
        End If
        
        'If iva4 <> 0 Then
        '     AsientoAddMURANO_TS 0, numAsiento, dataFactura, "477000010", "", "", numFactura, 0, iva4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(25), 25), ""
        'End If
        
        'If baseIva4 <> 0 And iva4 <> 0 Then
        '     AsientoAddMURANO_TS 0, numAsiento, dataFactura, "475000055", "", "", numFactura, 0, baseIva4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(25), 25), ""
        'End If
        
        'BASES -------------------------------------------------------------------------------------------------------------------------------------
        
        '70000000
        'If Tenim4 Then PintaDetallIvaMURANO numAsiento, clientCodi, dataFactura, baseIva1 + BaseRec1, "0.00", numFactura, "0.00", False, TeRetencio
        'If Tenim5 Then PintaDetallIvaMURANO numAsiento, clientCodi, dataFactura, baseIva2 + BaseRec2, "0.00", numFactura, "0.00", False, TeRetencio
        'If Tenim6 Then PintaDetallIvaMURANO numAsiento, clientCodi, dataFactura, baseIva3 + BaseRec3, "0.00", numFactura, "0.00", False, TeRetencio
            
        'POR FAMILIAS
        'SqlData = "select ISNULL(fe1.valor, '0') CC, tipusIva, case when c.[Tipus Iva]=1 then 'IVA_INCLOS' when c.[Tipus Iva]=2 then 'AMB_RECARREC' when c.[Tipus Iva]=3 then 'SENSE_RECARREC' when c.[Tipus Iva]=4 then 'ESTRANGER' end  TE_RECARREC, i.iva PCT_IVA, i.irpf PCT_REC, sum(import) BASE_IVA "
        sqlData = "select ISNULL(fe3.valor, isnull(fe2.valor, isnull(fe1.valor, '" & Left("7000000000000", nDigitos) & "'))) CC, sum(import) BASE_FAM "
        sqlData = sqlData & "from [" & NomTaulaFacturaData(dataFactura) & "] fd "
        'SqlData = SqlData & "left join clients c on fd.client=c.codi "
        sqlData = sqlData & "left join articles a on fd.producte=a.codi "
        'SqlData = SqlData & "left join " & DonamTaulaTipusIva(dataFactura) & " i on fd.tipusiva=i.tipus "
        sqlData = sqlData & "left join families F3 on a.familia = F3.nom "
        sqlData = sqlData & "left join families F2 on F2.nom = F3.pare "
        sqlData = sqlData & "left join families F1 on F1.nom = F2.Pare "
        sqlData = sqlData & "left join familiesExtes fe1 on F1.nom=fe1.familia and fe1.variable='CUENTA_CONTABLE' "
        sqlData = sqlData & "left join familiesExtes fe2 on F2.nom=fe2.familia and fe2.variable='CUENTA_CONTABLE' "
        sqlData = sqlData & "left join familiesExtes fe3 on F3.nom=fe3.familia and fe3.variable='CUENTA_CONTABLE' "
        sqlData = sqlData & "where fd.idfactura='" & idFactura & "' "
        'SqlData = SqlData & "group by ISNULL(fe1.valor, '0'), tipusIva, i.iva, i.irpf, c.[Tipus Iva]"
        sqlData = sqlData & "group by ISNULL(fe3.valor, isnull(fe2.valor, isnull(fe1.valor, '" & Left("7000000000000", nDigitos) & "')))"
        Set rsData = Db.OpenResultset(sqlData)
        
        totalFamilias = 0
        While Not rsData.EOF
            'Bases
            nCuenta = rsData("CC")
            If Len(nCuenta) <> nDigitos Then nCuenta = Left("7000000000000", nDigitos)
            AsientoAddMURANO_TS actividad, 0, numAsiento, dataFactura, nCuenta, "", "", numFactura, 0, Round(rsData("BASE_FAM"), 2), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(25), 25), ""
            totalFamilias = totalFamilias + Round(rsData("BASE_FAM"), 2)
            
            If TeRetencio Then
                haberReten = Round(rsData("BASE_FAM") * 0.19, 2)
                nCuenta = Left(("473" & Right("000000000000", nDigitos - 3)) & Space(12), 12)
                AsientoAddMURANO_TS actividad, 0, numAsiento, dataFactura, nCuenta, "", "", numFactura, haberReten, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(25), 25), "" 'La retención va al Debe
            End If
    
            
            rsData.MoveNext
        Wend
        'Regularizar decimales
        If Len(nCuenta) <> nDigitos Then nCuenta = Left("7000000000000", nDigitos)
        totalFamilias = totalFamilias + ibee
        If totalFamilias <> (baseIva1 + BaseRec1 + baseIva2 + BaseRec2 + baseIva3 + BaseRec3 + baseIva4 + ibee) Then
            If (baseIva1 + BaseRec1 + baseIva2 + BaseRec2 + baseIva3 + BaseRec3 + ibee) - totalFamilias > 0 Then
                AsientoAddMURANO_TS actividad, 0, numAsiento, dataFactura, nCuenta, "", "", numFactura, 0, FormatNumber(Abs((ibee + baseIva1 + BaseRec1 + baseIva2 + BaseRec2 + baseIva3 + BaseRec3) - totalFamilias), 2), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(25), 25), ""
            Else
                AsientoAddMURANO_TS actividad, 0, numAsiento, dataFactura, nCuenta, "", "", numFactura, FormatNumber(Abs((ibee + baseIva1 + BaseRec1 + baseIva2 + BaseRec2 + baseIva3 + BaseRec3) - totalFamilias), 2), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(25), 25), ""
            End If
        End If
        '~POR FAMILIAS
        
    End If
    
    Exit Sub
    
noExportat:

    InsertFeineaAFer "SincroMURANOFactura", "[" & idFactura & "]", "[" & dataFactura & "]", "[" & numFactura & "]", "[" & taulaFactura & "]"
    
jaExportat:

End Sub


Sub PintaDetallIvaMURANO(codigoActividad As String, numAsiento As String, clientCodi As String, dataFactura As Date, haber As Double, BaseIva As Double, numFactura As String, tipoIva As String, Optional TeRecarreg As Boolean = False, Optional TeRetencio As Boolean = False)
    
    Dim haberPts, baseIvaPts, nCuenta As String, Quota, datos, TipoRec, haberReten
    Dim valorIva1 As Double, valorIva2 As Double, valorIva3 As Double, valorIva4 As Double, valorRec1 As Double, valorRec2 As Double, valorRec3 As Double, valorRec4 As Double
    
    TipusDeIva valorIva1, valorIva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, dataFactura
   
    
    haber = haber
    haber = Round(haber, 3)
    
    haberPts = 0 'Haber * 166.386
    haberPts = 0 'Round(haberPts, 2)
    
    BaseIva = BaseIva
    BaseIva = Round(BaseIva, 2)
    
    baseIvaPts = 0 'BaseIva * 166.386
    baseIvaPts = 0 'Round(baseIvaPts, 2)
    TipoRec = "0.00"
    
    If (tipoIva = "0.00") Then
        nCuenta = cVentaMercaderies(clientCodi, TeRetencio)
        If TeRetencio Then
            haberReten = Round(haber * 0.19, 2)
        End If
    Else
        Quota = Left(tipoIva, 2)
        Quota = Replace(Quota, ".", "")
        nCuenta = cIVA(Quota, TeRecarreg)

        If TeRecarreg Then
            If tipoIva = valorIva1 & ".00" Then TipoRec = valorRec1
            If tipoIva = valorIva2 & ".00" Then TipoRec = valorRec2
            If tipoIva = valorIva3 & ".00" Then TipoRec = valorRec3
        End If
        
        If tipoIva = "0.5" Then
            'nCuenta = ("475" & Right("000000000000", nDigitos - 3)) + valorIva1
            nCuenta = ("477" & Right("000000000000", nDigitos - 3)) + valorIva1
            tipoIva = valorIva1 & ".00"
            TipoRec = valorRec1
        End If
        
        If tipoIva = "1" Then
            'nCuenta = ("475" & Right("000000000000", nDigitos - 3)) + valorIva2
            nCuenta = ("477" & Right("000000000000", nDigitos - 3)) + valorIva2
            tipoIva = valorIva2 & ".00"
            TipoRec = valorRec2
        End If
        
        If tipoIva = "4" Then
            'nCuenta = ("475" & Right("000000000000", nDigitos - 3)) + valorIva3
            nCuenta = ("477" & Right("000000000000", nDigitos - 3)) + valorIva3
            tipoIva = valorIva3 & ".00"
            TipoRec = valorRec3
        End If
        
        nCuenta = Left(nCuenta & Space(12), 12)
        
    End If
  
    'AsientoAdd numAsiento, dataFactura, nCuenta, codiContable2(ClientCodi), numFactura, Left("Factura num: " & numFactura & Space(25), 25), 0, Haber, BaseIva, tipoIva, TipoRec, ""
    'AsientoAddMURANO 0, numAsiento, dataFactura, nCuenta, codiContable2(clientCodi), Left("F.num : " & Abs(numFactura) & Space(25), 25), 0, haber
    AsientoAddMURANO_TS codigoActividad, 0, numAsiento, dataFactura, nCuenta, "", "", numFactura, 0, haber, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(25), 25), ""

    If TeRetencio Then
        nCuenta = Left(("473" & Right("000000000000", nDigitos - 3)) & Space(12), 12)
        'AsientoAdd numAsiento, dataFactura, nCuenta, nCuenta, numFactura, Left("Factura num: " & numFactura & Space(25), 25), HaberReten, 0, baseIvaPts, tipoIva, TipoRec, ""
        'AsientoAddMURANO 0, numAsiento, dataFactura, nCuenta, nCuenta, Left("F.num : " & Abs(numFactura) & Space(25), 25), HaberReten, 0
        AsientoAddMURANO_TS codigoActividad, 0, numAsiento, dataFactura, nCuenta, "", "", numFactura, haberReten, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("F.num : " & numFactura & Space(25), 25), "" 'La retención va al Debe
    End If

End Sub



Sub AsientoAddMURANO_TS(codigoActividad As String, TipoMov, numAsiento, data As Date, Cuenta1 As String, tipoFactura, serie, numFactura As String, debe, haber, iva1, baseIva1, iva2, baseIva2, iva3, baseIva3, iva4, baseIva4, rec1, rec2, rec3, pctRetencion, importSenseIVA As Boolean, cNif As String, comentari As String, idFactura As String, Optional serieTick As String, Optional primerTick As String, Optional ultimTick As String, Optional claveOperacion As String, Optional fechaVencimiento As String)
   
    Dim sq As String, Sq2 As String

    Dim valorIva1 As Double, valorIva2 As Double, valorIva3 As Double, valorIva4 As Double, valorRec1 As Double, valorRec2 As Double, valorRec3 As Double, valorRec4 As Double
    TipusDeIva valorIva1, valorIva2, valorIva3, valorIva4, valorRec1, valorRec2, valorRec3, valorRec4, data

    comentari = Left(comentari, 40)
    
    sq = "":     Sq2 = ""
    sq = sq & "Insert into " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS (":     Sq2 = Sq2 & " Values ("
    sq = sq & "[CodigoEmpresa]":                                       Sq2 = Sq2 & EmpresaMurano
    sq = sq & ",[Asiento]":                                            Sq2 = Sq2 & ", '" & numAsiento & "' "
    sq = sq & ",[A_ASSENTAMENT]":                                      Sq2 = Sq2 & ", '" & numAsiento & "' "
    sq = sq & ",[Ejercicio]":                                          Sq2 = Sq2 & ", '" & Format(data, "yyyy") & "' "
    If debe = 0 Then
        sq = sq & ",[CargoAbono]":                                     Sq2 = Sq2 & ",'H'"
    Else
        sq = sq & ",[CargoAbono]":                                     Sq2 = Sq2 & ",'D'"
    End If
    sq = sq & ",[CodigoCuenta]":                                       Sq2 = Sq2 & ",'" & Cuenta1 & "'"
    sq = sq & ",[Contrapartida]":                                      Sq2 = Sq2 & ",''"
    sq = sq & ",[FechaAsiento]":                                       Sq2 = Sq2 & ",convert(datetime,'" & Format(data, "dd/mm/yy") & "',3)"
    sq = sq & ",[Comentario]":                                         Sq2 = Sq2 & ",'" & comentari & "'"
    If debe = 0 Then
        sq = sq & ",[ImporteAsiento]":                                 Sq2 = Sq2 & ",'" & CDbl(haber) & " '"
    Else
        sq = sq & ",[ImporteAsiento]":                                 Sq2 = Sq2 & ",'" & CDbl(debe) & " '"
    End If
    sq = sq & ",[CodigoDiario]":                                       Sq2 = Sq2 & ",0"
    If codigoActividad <> "" Then
        sq = sq & ",[CodigoActividad]":                                Sq2 = Sq2 & ",'" & codigoActividad & "'"
        sq = sq & ",[CodigoCanal]":                                    Sq2 = Sq2 & ",'" & codigoActividad & "'"
    Else
        sq = sq & ",[CodigoActividad]":                                Sq2 = Sq2 & ",''"
    End If
    If fechaVencimiento <> "" Then
        sq = sq & ",[FechaVencimiento]":                               Sq2 = Sq2 & ",convert(datetime,'" & fechaVencimiento & "',3)"
    Else
        sq = sq & ",[FechaVencimiento]":                               Sq2 = Sq2 & ",NULL"
    End If
    sq = sq & ",[NumeroPeriodo]":                                      Sq2 = Sq2 & "," & Month(data)
    sq = sq & ",[FechaGrabacion]":                                     Sq2 = Sq2 & ",getdate()"
    sq = sq & ",[TipoEntrada]":                                        Sq2 = Sq2 & ",'EN'"
    sq = sq & ",[CodigoDepartamento]":                                 Sq2 = Sq2 & ",''"
    sq = sq & ",[CodigoSeccion]":                                      Sq2 = Sq2 & ",''"
    sq = sq & ",[CodigoDivisa]":                                       Sq2 = Sq2 & ",''"
    sq = sq & ",[ImporteCambio]":                                      Sq2 = Sq2 & ",0"
    sq = sq & ",[ImporteDivisa]":                                      Sq2 = Sq2 & ",0"
    sq = sq & ",[FactorCambio]":                                       Sq2 = Sq2 & ",0"
    sq = sq & ",[CodigoProyecto]":                                     Sq2 = Sq2 & ",''"
    sq = sq & ",[LibreN1]":                                            Sq2 = Sq2 & ",0"
    sq = sq & ",[LibreN2]":                                            Sq2 = Sq2 & ",0"
    sq = sq & ",[LibreA1]":                                            Sq2 = Sq2 & ",''"
    sq = sq & ",[LibreA2]":                                            Sq2 = Sq2 & ",''"
    sq = sq & ",[IdDelegacion]":                                       Sq2 = Sq2 & ",''"
    sq = sq & ",[baseIva1]":                                           Sq2 = Sq2 & ", " & baseIva1 & " "
    If baseIva1 <> 0 Then
        sq = sq & ",[PorIva1]":                                        Sq2 = Sq2 & ", " & valorIva1 & " "
        If tipoFactura = "E" Then sq = sq & ",[CodigoTransaccion1]":   Sq2 = Sq2 & ", 1 "
    Else
        sq = sq & ",[PorIva1]":                                        Sq2 = Sq2 & ", 0 "
    End If
    sq = sq & ",[CuotaIva1]":                                          Sq2 = Sq2 & ", " & iva1 & " "
    If rec1 <> 0 Then
        sq = sq & ",[PorRecargoEquivalencia1]":                        Sq2 = Sq2 & ", " & valorRec1 & " "
    Else
        sq = sq & ",[PorRecargoEquivalencia1]":                        Sq2 = Sq2 & ", 0 "
    End If
    sq = sq & ",[RecargoEquivalencia1]":                               Sq2 = Sq2 & ", " & rec1 & " "
    sq = sq & ",[baseIva2]":                                           Sq2 = Sq2 & ", " & baseIva2 & " "
    If baseIva2 <> 0 Then
        sq = sq & ",[PorIva2]":                                        Sq2 = Sq2 & ", " & valorIva2 & " "
        If tipoFactura = "E" Then sq = sq & ",[CodigoTransaccion2]":   Sq2 = Sq2 & ", 1 "
    Else
        sq = sq & ",[PorIva2]":                                        Sq2 = Sq2 & ", 0 "
    End If
    sq = sq & ",[CuotaIva2]":                                          Sq2 = Sq2 & ", " & iva2 & " "
    If rec2 <> 0 Then
        sq = sq & ",[PorRecargoEquivalencia2]":                        Sq2 = Sq2 & ", " & valorRec2 & " "
    Else
        sq = sq & ",[PorRecargoEquivalencia2]":                        Sq2 = Sq2 & ", 0 "
    End If
    sq = sq & ",[RecargoEquivalencia2]":                               Sq2 = Sq2 & ", " & rec2 & " "
    sq = sq & ",[baseIva3]":                                           Sq2 = Sq2 & ", " & baseIva3 & " "
    If baseIva3 <> 0 Then
        sq = sq & ",[PorIva3]":                                        Sq2 = Sq2 & ", " & valorIva3 & " "
        If tipoFactura = "E" Then sq = sq & ",[CodigoTransaccion3]":   Sq2 = Sq2 & ", 1 "
    Else
        sq = sq & ",[PorIva3]":                                        Sq2 = Sq2 & ", 0 "
    End If
    sq = sq & ",[CuotaIva3]":                                          Sq2 = Sq2 & ", " & iva3 & " "
    
    'SENSE IVA / IBEE
    sq = sq & ",[baseIva4]":                                           Sq2 = Sq2 & ", " & baseIva4 & " "
    If iva4 > 0 Then 'IBEE
        sq = sq & ",[PorIva4]":                                        Sq2 = Sq2 & ", 10 "
        If tipoFactura = "E" Then sq = sq & ",[CodigoTransaccion4]":   Sq2 = Sq2 & ", 1 "
    Else
        sq = sq & ",[PorIva4]":                                        Sq2 = Sq2 & ", 0 "
    End If
    sq = sq & ",[CuotaIva4]":                                          Sq2 = Sq2 & ", " & iva4 & " "
    
    If rec3 <> 0 Then
        sq = sq & ",[PorRecargoEquivalencia3]":                        Sq2 = Sq2 & ", " & valorRec3 & " "
    Else
        sq = sq & ",[PorRecargoEquivalencia3]":                        Sq2 = Sq2 & ", 0 "
    End If
    sq = sq & ",[RecargoEquivalencia3]":                               Sq2 = Sq2 & ", " & rec3 & " "
    
    sq = sq & ",[PorRecargoEquivalencia4]":                            Sq2 = Sq2 & ", 0"
    sq = sq & ",[RecargoEquivalencia4]":                               Sq2 = Sq2 & ", 0"

    If pctRetencion <> 0 Then
        If importSenseIVA Then
            sq = sq & ",[BaseRetencion]":                               Sq2 = Sq2 & ", " & baseIva1 + baseIva2 + baseIva3 & " "
            sq = sq & ",[PorRetencion]":                                Sq2 = Sq2 & ", " & pctRetencion & " "
            sq = sq & ",[ImporteRetencion]":                            Sq2 = Sq2 & ", " & Round((baseIva1 + baseIva2 + baseIva3) * (pctRetencion / 100), 2) & " "
        Else
            sq = sq & ",[BaseRetencion]":                               Sq2 = Sq2 & ", " & baseIva1 + baseIva2 + baseIva3 + baseIva4 & " "
            sq = sq & ",[PorRetencion]":                                Sq2 = Sq2 & ", " & pctRetencion & " "
            sq = sq & ",[ImporteRetencion]":                            Sq2 = Sq2 & ", " & Round((baseIva1 + baseIva2 + baseIva3 + baseIva4) * (pctRetencion / 100), 2) & " "
        End If
    End If

    If tipoFactura = "" Or tipoFactura = "B" Then
        sq = sq & ",[Año]":                                             Sq2 = Sq2 & ", 0 "
        sq = sq & ",[Serie]":                                           Sq2 = Sq2 & ", '' "
        sq = sq & ",[Factura]":                                         Sq2 = Sq2 & ", 0 "
        sq = sq & ",[SuFacturaNo]":                                     Sq2 = Sq2 & ", '' "
        sq = sq & ",[FechaFactura]":                                    Sq2 = Sq2 & ", null "
        sq = sq & ",[ImporteFactura]":                                  Sq2 = Sq2 & ",0 "
    Else
        sq = sq & ",[Año]":                                             Sq2 = Sq2 & ", " & Format(data, "yyyy") & " "
        sq = sq & ",[Serie]":                                           Sq2 = Sq2 & ", '" & Left(serie, 10) & "' "
        If tipoFactura = "R" Then
            sq = sq & ",[Factura]":                                     Sq2 = Sq2 & ", 1 "
        Else
            sq = sq & ",[Factura]":                                     Sq2 = Sq2 & ", '" & numFactura & "' "
        End If
        sq = sq & ",[SuFacturaNo]":                                     Sq2 = Sq2 & ", '" & numFactura & "' "
        sq = sq & ",[FechaFactura]":                                    Sq2 = Sq2 & ", convert(datetime,'" & Format(data, "dd/mm/yy") & "',3) "
        If debe = 0 Then
            sq = sq & ",[ImporteFactura]":                              Sq2 = Sq2 & "," & haber & " "
        Else
            sq = sq & ",[ImporteFactura]":                              Sq2 = Sq2 & "," & debe & " "
        End If
    End If
    If tipoFactura = "B" Then
        sq = sq & ",[TipoFactura]":                                     Sq2 = Sq2 & ",'' "
    Else
        sq = sq & ",[TipoFactura]":                                     Sq2 = Sq2 & ",'" & tipoFactura & "' "
    End If
    
    If tipoFactura = "" Then
        sq = sq & ",[CifDni]":                                          Sq2 = Sq2 & ", '' "
        sq = sq & ",[CifEuropeo]":                                      Sq2 = Sq2 & ", '' "
    Else
        sq = sq & ",[CifDni]":                                          Sq2 = Sq2 & ", left('" & cNif & "', 13) "
        'If InStr(cNif, ",") Then
        'sq = sq & ",[CifEuropeo]":                                      Sq2 = Sq2 & ", left('" & cNif & "', 13) "
        'Else
            sq = sq & ",[CifEuropeo]":                                      Sq2 = Sq2 & ", '" & cNif & "' "
        'End If
    End If
    sq = sq & ",[Metalico347]":                                         Sq2 = Sq2 & ",0"
    sq = sq & ",[FechaFacturaOriginal]":                                Sq2 = Sq2 & ",convert(datetime,'" & Format(data, "dd/mm/yy") & "',3) "

    If serieTick <> "" Then
        sq = sq & ",[ClaveOperacionFactura_]":                          Sq2 = Sq2 & ", '" & claveOperacion & "' "
        sq = sq & ",[serieAgrupacion_]":                                Sq2 = Sq2 & ", '" & serieTick & "' "
        sq = sq & ",[NumeroFacturaInicial_]":                           Sq2 = Sq2 & ", '" & primerTick & "' "
        sq = sq & ",[NumeroFacturaFinal_]":                             Sq2 = Sq2 & ", '" & ultimTick & "' "
    ElseIf claveOperacion <> "" Then
        sq = sq & ",[ClaveOperacionFactura_]":                          Sq2 = Sq2 & ", '" & claveOperacion & "' "
    End If

    sq = sq & ") ":                                                     Sq2 = Sq2 & ") "
    sq = sq & Sq2
    
On Error GoTo norR

    MuranoExecute sq
    
    Dim sql As String
    If cNif <> "" Then
        If tipoFactura = "E" Then 'EL NIF ES DEL CLIENTE
            'sql = "insert into [silema_Ts].sage.dbo.A_IMPORTACIO_CLIPRO ( "
            'sql = sql & "CodigoEmpresa, ClienteOProveedor, CodigoClienteProveedor, RazonSocial, Nombre, Domicilio, codigoCuenta, CifDni, CodigoSigla, CodigoPostal,Municipio, CodigoNacion, telefono, email1, codigoBanco, CodigoAgencia, DC, CCC) "
            'sql = sql & "select " & EmpresaMurano & ", 'C', '" & Cuenta1 & "', left([nom llarg], 35), left(nom, 25), left(Adresa, 25), '" & Cuenta1 & "', left(nif,10) , '', cP , left(ciutat,25), 108, left(cc1.valor,15), "
            'sql = sql & "left(cc2.valor,100), left(isnulL(cc3.valor, '                    '), 4), substring(isnulL(cc3.valor, '                    '), 5, 4), substring(isnulL(cc3.valor, '                    '), 9, 2), "
            'sql = sql & "right(isnulL(cc3.valor, '                    '), 10) "
            'sql = sql & "from clients c "
            'sql = sql & "left join constantsClient cc1 on c.codi=cc1.codi and cc1.variable='Tel' "
            'sql = sql & "left join constantsClient cc2 on c.codi=cc2.codi and cc2.variable='eMail' "
            'sql = sql & "left join constantsClient cc3 on c.codi=cc3.codi and cc3.variable='CompteCorrent' "
            'sql = sql & "where c.nif like '%" & cNif & "%'"
            'ExecutaComandaSql sql
            
            sql = "insert into " & dbSage & ".dbo.A_IMPORTACIO_CLIPRO ( "
            sql = sql & "CodigoEmpresa, ClienteOProveedor, CodigoClienteProveedor, RazonSocial, Nombre, Domicilio, codigoCuenta, CifDni, CifEuropeo, CodigoSigla, CodigoPostal,Municipio, CodigoNacion, telefono, email1, codigoBanco, CodigoAgencia, DC, CCC) "
            sql = sql & "select " & EmpresaMurano & ", 'C', '', left(fi.clientNom , 35), left(fi.clientNom, 25), left(fi.ClientAdresa, 25), '', left(fi.clientNif,13), left(fi.clientNif,15) , '', fi.ClientcP , left(fi.clientCiutat,25), 108, left(fi.Tel,15), "
            sql = sql & "left(fi.eMail,100), left(isnulL(fr.ClientCompte, '                    '), 4), substring(isnulL(fr.ClientCompte, '                    '), 5, 4), substring(isnulL(fr.ClientCompte, '                    '), 9, 2), "
            sql = sql & "right(isnulL(fr.ClientCompte, '                    '), 10) "
            sql = sql & "from [" & NomTaulaFacturaIva(data) & "] fi "
            sql = sql & "left join [" & NomTaulaFacturaReb(data) & "] fr on fi.idFactura  collate Modern_Spanish_CI_AS = fr.idFactura collate Modern_Spanish_CI_AS "
            sql = sql & "where fi.idfactura='" & idFactura & "'"
            ExecutaComandaSql sql
            
        ElseIf tipoFactura = "R" Or tipoFactura = "B" Then 'EL NIF ES EL DEL PROVEEDOR
            sql = "insert into " & dbSage & ".dbo.A_IMPORTACIO_CLIPRO ( "
            sql = sql & "CodigoEmpresa, ClienteOProveedor, CodigoClienteProveedor, RazonSocial, Nombre, Domicilio, codigoCuenta, CifDni, CifEuropeo, CodigoSigla, CodigoPostal,Municipio, CodigoNacion, telefono, email1, codigoBanco, CodigoAgencia, DC, CCC) "
            sql = sql & "select " & EmpresaMurano & ", 'P', '" & Cuenta1 & "', left(nombre, 35), left(nombre, 25), left(direccion, 25), '" & Cuenta1 & "', left(nif,13), left(nif,15), '', cP , left(ciudad,25), 108, left(tlf1,15), left(email,100), left(isnulL(pe.valor, '                    '), 4), substring(isnulL(pe.valor, '                    '), 5, 4), substring(isnulL(pe.valor, '                    '), 9, 2), right(isnulL(pe.valor, '                    '), 10) "
            sql = sql & "from ccProveedores p "
            sql = sql & "left join ccProveedoresExtes pe on p.id=pe.id and pe.nom='NumeroCuenta' "
            sql = sql & "WHERE p.nif like '%" & cNif & "%'"
            ExecutaComandaSql sql
        End If
    End If

    Exit Sub
    
norR:
    sf_enviarMail "email@hit.cat", "ana@solucionesit365.com", "ERROR TRASPASO MURANO", sq, "", ""

End Sub


Sub MuranoExecute(st As String)
On Error GoTo emailnorrR


    Db.Execute "SET LANGUAGE us_english ; " & st
    Exit Sub
emailnorrR:

'    sf_enviarMail "email@hit.cat", "ana@solucionesit365.com", "ERROR TRASPASO MURANO", st, "", ""
    Debug.Print "ERROR !! " & st
End Sub


Function getEmpresaMURANOBotiga(botiga As Integer) As Integer
    Dim nEmpresa As Double 'Empresa HIT
    Dim sql As String
    Dim rsPC As rdoResultset, rsNEmp As rdoResultset, rsEmpMurano As rdoResultset
    Dim nifEmp As String, rsNif As rdoResultset

    dbSage = "Silema_ts.sage"
    
    nEmpresa = -1
    
    Set rsPC = Db.OpenResultset("select isnull(Valor, '') Pc from constantsempresa where camp='ProgramaContable'")
    If Not rsPC.EOF Then
        If rsPC("Pc") = "SAGE" Then
            sql = "select cc.Valor, c.Codi, c.Nom "
            sql = sql & "from constantsclient cc "
            sql = sql & "left join clients c on cc.Codi=c.codi "
            sql = sql & "where cc.variable = 'EmpresaVendes' and cc.Codi in (select valor1 from ParamsHw) and c.codi = '" & botiga & "' "
            sql = sql & "order by cc.Valor, c.nom"
            Set rsNEmp = Db.OpenResultset(sql)
            If Not rsNEmp.EOF Then nEmpresa = CDbl(rsNEmp("valor"))
            
            If nEmpresa = 0 Then
                Set rsNif = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampNif' ")
            Else
                Set rsNif = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampNif' ")
            End If
            If Not rsNif.EOF Then nifEmp = rsNif("valor")
        
            Set rsEmpMurano = Db.OpenResultset("select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmp & "'")
            If Not rsEmpMurano.EOF Then
                nEmpresa = rsEmpMurano("CodigoEmpresa")
            Else
                nEmpresa = -1
            End If
        End If
    End If
    
    getEmpresaMURANOBotiga = nEmpresa
End Function

Sub ExportaMURANO_CaixaBotigaOnLine(nEmpresa As Double, botiga As String, Di As Date, Df As Date, intTickets As String, Z As Double)
    Dim D As Date, sql As String, rsCtb As rdoResultset, rsCActividad As rdoResultset
    Dim import As Double, tipoIva, PctIva, Base, Quota, CuentaVentas
    Dim T1 As Double, T2 As Double, T3 As Double, T4 As Double, TR1 As Double, TR2 As Double, TR3 As Double, TR4 As Double
    Dim cCtble, rsCodi As rdoResultset
    Dim rsHist As rdoResultset
    Dim rsNA As rdoResultset, rsSage As rdoResultset
    Dim rsCaixes As rdoResultset, rsCCBanc As rdoResultset, ccBanc As String
    Dim CcVentas As String
    Dim iTargeta As Double, iTkRs As Double, importZ As Double, import43 As Double
    Dim numAsiento As String
    Dim Motiu As String, nifClienteContado As String
    Dim rsBotiga As rdoResultset
    Dim primerTick As String, ultimTick As String
    Dim iva1 As Double, baseIva1 As Double, iva2 As Double, baseIva2 As Double, iva3 As Double, baseIva3 As Double
    Dim msgError As String
    Dim actividad As String
    
    Dim nCierres As Integer
    Dim importZTotal As Double, importVTotal As Double
  
    Dim rsRectifica As rdoResultset, rsAnulats As rdoResultset
    
    Dim asientosCaixa As String
    asientosCaixa = "0"
    
    On Error GoTo noExportat
    
    If nEmpresa = "0" Then
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampNif' ")
        Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampActividadSAGE' ")
    Else
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampNif' ")
        Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampActividadSAGE' ")
    End If
    
    actividad = ""
    If Not rsCActividad.EOF Then actividad = rsCActividad("valor")
    
    Dim nifEmpMurano As String, rsEmpMurano As rdoResultset
    
    If Not rsCtb.EOF Then nifEmpMurano = rsCtb("valor")
    Set rsEmpMurano = Db.OpenResultset("select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'")
    If Not rsEmpMurano.EOF Then
        EmpresaMurano = rsEmpMurano("CodigoEmpresa")
    Else
        msgError = "NO HAY EMPRESA MURANO:<br>" & "select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'"
        GoTo noExportat
    End If

    nifClienteContado = "22222222J"

    D = Di
    
    primerTick = Split(intTickets, ",")(0)
    ultimTick = Split(intTickets, ",")(1)
    
    Set rsNA = Db.OpenResultset("select isnull(max(Asiento), 0) + 1 NumAsiento from " & TaulaHistoricoMURANO(D) & " where CodigoEmpresa = " & EmpresaMurano & " and year(fechaAsiento) = " & Year(D))
    If Not rsNA.EOF Then numAsiento = rsNA("NumAsiento")
    
    Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(D) & " where month(FechaAsiento)=" & Month(D) & " and day(FechaAsiento)=" & Day(D) & " and Param1 = '" & botiga & "' and Param2 = '" & primerTick & "' and Param3 = '" & ultimTick & "' and TipoExportacion='CAIXA' and CodigoEmpresa=" & EmpresaMurano & " order by asiento")
    If Not rsHist.EOF Then numAsiento = rsHist("Asiento")
    While Not rsHist.EOF
        ExecutaComandaSql "Delete from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & rsHist("Asiento") & " and (statusTraspasadoIME=0 or statusTraspasadoIME=2) and CodigoEmpresa=" & EmpresaMurano
        rsHist.MoveNext
    Wend
    
    Set rsSage = Db.OpenResultset("select * from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & numAsiento & " and statusTraspasadoIME=1 and CodigoEmpresa=" & EmpresaMurano & " and ejercicio = " & Year(D))
    If Not rsSage.EOF Then
        msgError = "YA ESTABA TRASPASADA<br> " & "select * from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & numAsiento & " and statusTraspasadoIME=1 and CodigoEmpresa=" & EmpresaMurano & " and ejercicio = " & Year(D)
        GoTo noExportat  'Si ya se ha traspasado correctamente a MURANO no volvemos a traspasarla
    End If
    
    TipusDeIva T1, T2, T3, T4, TR1, TR2, TR3, TR4, D
    
    cCtble = botiga
    Set rsCodi = Db.OpenResultset("SELECT Valor FROM " & tablaConstantsClient() & " WHERE  codi = " & botiga & " AND variable = 'CodiContable' ")
    If Not rsCodi.EOF Then If Not IsNull(rsCodi("Valor")) And (Len(rsCodi("Valor")) > 0) And IsNumeric(rsCodi("Valor")) Then cCtble = CDbl(rsCodi("Valor"))
    
    InformaMiss "MURANO Vendes Botiga: " & BotigaCodiNom(botiga) & " Dia: " & D, True
    
    importZTotal = 0
    
    'TOTAL VENTAS
    sql = "Select isnull(sum(vv.import), 0) Import "
    sql = sql & "From [" & NomTaulaVentas(D) & "] vv "
    sql = sql & "where day(vv.data) = " & Day(Di) & " and vv.botiga = '" & botiga & "' and num_tick>=" & primerTick & " and num_Tick<=" & ultimTick & " "
    Set rsCtb = Db.OpenResultset(sql)
                    
    If rsCtb.EOF Then
        InsertFeineaAFer "SincroMURANOCaixaOnLine", "[" & botiga & "]", "[" & Format(Di, "dd-mm-yyyy") & " " & Format(Di, "hh:mm:ss") & "]", "[" & Format(Df, "dd-mm-yyyy") & " " & Format(Df, "hh:mm:ss") & "]", "[" & primerTick & "," & ultimTick & "]", "[" & Z & "]"
        msgError = "QUEDA PENDIENTE. NO HAY VENTAS<br>" & sql
        GoTo noExportat
    Else
        'Comprobar si las ventas = Z, si no son iguales, faltan ventas y hay que dejar la caja pendiente
        If rsCtb("import") > Z Then 'HI HAN MÉS VENDES QUE CAIXA, COMPROVEM TIQUETS ANULATS
            Set rsRectifica = Db.OpenResultset("select num_tick, sum(import) i from [" & NomTaulaVentas(D) & "] where botiga=" & botiga & " and day(data)=" & Day(D) & " group by botiga, num_tick having sum(import)=" & Round((rsCtb("import") - Z), 2))
            While Not rsRectifica.EOF
                Set rsAnulats = Db.OpenResultset("select * from [" & NomTaulaAnulats(D) & "] where botiga=" & botiga & " and day(data)=" & Day(D) & " and num_tick=" & rsRectifica("num_tick"))
                If Not rsAnulats.EOF Then ExecutaComandaSql "delete from [" & NomTaulaVentas(D) & "] where botiga=" & botiga & " and day(data)=" & Day(D) & " and num_tick=" & rsRectifica("num_tick")
                rsRectifica.MoveNext
            Wend
            sql = "Select isnull(sum(vv.import), 0) Import "
            sql = sql & "From [" & NomTaulaVentas(D) & "] vv "
            sql = sql & "where day(vv.data) = " & Day(Di) & " and vv.botiga = '" & botiga & "' and num_tick>=" & primerTick & " and num_Tick<=" & ultimTick & " "
            Set rsCtb = Db.OpenResultset(sql)
        End If
        
        If Abs(Z - rsCtb("import")) > 5 Then
            msgError = "QUEDA PENDIENTE. FALTAN VENTAS<br>" & sql & "<BR>IMPORT Z: " & Z & " VENDES: " & rsCtb("import")
            InsertFeineaAFer "SincroMURANOCaixaOnLine", "[" & botiga & "]", "[" & Format(Di, "dd-mm-yyyy") & " " & Format(Di, "hh:mm:ss") & "]", "[" & Format(Df, "dd-mm-yyyy") & " " & Format(Df, "hh:mm:ss") & "]", "[" & primerTick & "," & ultimTick & "]", "[" & Z & "]"
            GoTo noExportat
        End If

    End If
    
    
    'BASES DE IVA
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
        AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("477" & Right("000000000000", nDigitos - 3)) + PctIva, "", "", "", 0, Quota, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Ventas " & BotigaCodiNom(botiga), ""
        
        InformaMiss "Ventas " & D, True
        DoEvents

        rsCtb.MoveNext
    Wend
    rsCtb.Close
    
    'FAMILIAS
    sql = "Select a.tipoiva, v.botiga, sum(v.import) as import,  ISNULL(fe1.valor, '" & Left("7000000000000", nDigitos) & "') CC "
    sql = sql & "from ( "
    sql = sql & "Select vv.Botiga, vv.Plu ,sum(vv.import) Import "
    sql = sql & "From [" & NomTaulaVentas(D) & "] vv "
    sql = sql & "where vv.data between CONVERT(datetime, '" & Di & "', 103) and convert(datetime,'" & Df & "', 103) and vv.botiga = '" & botiga & "' "
    sql = sql & "group by vv.botiga, vv.Plu) v "
    sql = sql & "Left Join "
    sql = sql & "(select Aa.codi, aa.TipoIva, aa.familia "
    sql = sql & "From (select Familia,codi,tipoiva from articles union select Familia,codi,tipoiva from articles_zombis) aa ) a on a.codi = v.plu "
    sql = sql & "left join families F3 on a.familia = F3.nom "
    sql = sql & "left join families F2 on F2.nom = F3.pare "
    sql = sql & "left join families F1 on F1.nom = F2.Pare "
    sql = sql & "left join familiesExtes fe1 on F1.nom=fe1.familia and fe1.variable='CUENTA_CONTABLE' "
    sql = sql & "group by botiga, tipoiva, ISNULL(fe1.valor, '" & Left("7000000000000", nDigitos) & "') "
    sql = sql & "order by botiga, tipoiva "
    Set rsCtb = Db.OpenResultset(sql)
                    
    While Not rsCtb.EOF
        import = rsCtb("Import")
        tipoIva = rsCtb("TipoIva")
        CcVentas = rsCtb("CC")
        
        InformaMiss "MURANO Vendes Botiga FAMILIA: " & BotigaCodiNom(botiga) & " Dia: " & D, True
        
        Select Case tipoIva
            Case 1: PctIva = T1
            Case 2: PctIva = T2
            Case 3: PctIva = T3
        End Select
        Base = Round(import / (1 + (PctIva / 100)), 2)
        
        AsientoAddMURANO_TS actividad, 0, numAsiento, D, CcVentas, "", "", "", 0, Base, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Ventas " & BotigaCodiNom(botiga), ""
        
        InformaMiss "Ventas " & D, True
        DoEvents

        rsCtb.MoveNext
    Wend
    rsCtb.Close
    
    AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("43" & Right("000000000000" & cCtble, nDigitos - 2)), "E", BotigaCodiNom(botiga), primerTick, import43, 0, iva1, baseIva1, iva2, baseIva2, iva3, baseIva3, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Ventas " & BotigaCodiNom(botiga), "", Left(BotigaCodiNom(botiga), 10), primerTick, ultimTick, "B"
                    
    ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
        
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

    'ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", " & numAsiento & ", 'CAIXA')"
    ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
    AsientoAddMURANO_TS actividad, 0, numAsiento, D, "43" & Right("000000000000" & cCtble, nDigitos - 2), "B", "", "", 0, import, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "COBRO EN METALICO CAJA REGISTRA " & BotigaCodiNom(botiga), ""
    AsientoAddMURANO_TS actividad, 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", import, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "COBRO EN METALICO CAJA REGISTRA " & BotigaCodiNom(botiga), ""
    
    ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
    
    asientosCaixa = asientosCaixa + "," + numAsiento
                    
    numAsiento = numAsiento + 1
    
    'Moviments ENTRADA/SORTIDA
    InformaMiss "MURANO Moviments ENTRADA/SORTIDA Botiga: " & BotigaCodiNom(botiga) & " Dia: " & D, True
    
    sql = "Select motiu, tipus_moviment, c.codi botigaCodi, sum(import) as import "
    sql = sql & "from [" & NomTaulaMovi(D) & "] v "
    sql = sql & "left join clients c on v.botiga = c.codi "
    sql = sql & "where tipus_moviment in ('O','A') and Data between convert(datetime,'" & Di & "', 103) and convert(datetime,'" & Df & "', 103) And botiga = '" & botiga & "' and Import<>0 "
    sql = sql & "Group By motiu, tipus_moviment, c.codi "
    
    Set rsCtb = Db.OpenResultset(sql)
    While Not rsCtb.EOF
        import = Format(rsCtb("Import"), "0.0#")
        Motiu = rsCtb("motiu") & " " & BotigaCodiNom(botiga)

        If rsCtb("motiu") = "Entrega Diària" Or UCase(rsCtb("motiu")) = UCase("Entrega Diaria") Or rsCtb("motiu") = "Sortida de Canvi" Then
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "5701" & Right("000000000000" & cCtble, nDigitos - 4), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("SALIDA DE CAJA A CAJA FUERTE " & BotigaCodiNom(botiga), 40), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "SALIDA DE CAJA A CAJA FUERTE " & BotigaCodiNom(botiga), ""
            
            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
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
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("43" & Right("000000000000" & cCtble, nDigitos - 2)), "B", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Cobros ticket restaurant " & BotigaCodiNom(botiga), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("44" & Right("000000000000", nDigitos - 2)), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Cobros ticket restaurant " & BotigaCodiNom(botiga), ""
            
            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
            
        ElseIf Left(rsCtb("motiu"), 9) = "Excs.TkRs" Then
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("768" & Right("000000000000", nDigitos - 3)), "B", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Exc. cobro CHQ RTE", ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("44" & Right("000000000000", nDigitos - 2)), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Exc. cobro CHQ RTE", ""
            
            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
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
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "629005" & Right("000000000000" & cCtble, nDigitos - 6), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "GASTOS " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "GASTOS " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            
            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
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
    sql = sql & "where tipus_moviment = 'J' and Data between convert(datetime,'" & Di & "', 103) and convert(datetime,'" & Df & "', 103) And botiga = '" & botiga & "' "
    sql = sql & "Order By data "
    
    Set rsCtb = Db.OpenResultset(sql)
    While Not rsCtb.EOF
        import = rsCtb("import")

        If import < 0 Then 'DESCUADRE NEGATIVO
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "659" & Right("000000000000" & cCtble, nDigitos - 3), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            
            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
        Else 'DESCUADRE POSITIVO
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "659" & Right("000000000000" & cCtble, nDigitos - 3), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""

            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
        End If
        rsCtb.MoveNext
    Wend
    rsCtb.Close
    Exit Sub
    
noExportat:
    'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR ExportaMURANO_CaixaBotigaOnLine", Err.Description & "<br>" & msgError, "", ""
End Sub


Sub ExportaMURANO_CaixaBotigaOnLine_V3(nEmpresa As Double, botiga As String, Di As Date, Df As Date, intTickets As String, Z As Double)
    Dim D As Date, sql As String, rsCtb As rdoResultset, rsCActividad As rdoResultset
    Dim import As Double, tipoIva, PctIva, Base, Quota, CuentaVentas
    Dim T1 As Double, T2 As Double, T3 As Double, T4 As Double, TR1 As Double, TR2 As Double, TR3 As Double, TR4 As Double
    Dim cCtble, rsCodi As rdoResultset
    Dim rsHist As rdoResultset
    Dim rsNA As rdoResultset, rsSage As rdoResultset
    Dim rsCaixes As rdoResultset, rsCCBanc As rdoResultset, ccBanc As String
    Dim CcVentas As String
    Dim iTargeta As Double, iTkRs As Double, importZ As Double, import43 As Double, iDeudas As Double, iCobros As Double
    Dim numAsiento As String
    Dim Motiu As String, nifClienteContado As String
    Dim rsBotiga As rdoResultset
    Dim primerTick As String, ultimTick As String
    Dim iva1 As Double, baseIva1 As Double, iva2 As Double, baseIva2 As Double, iva3 As Double, baseIva3 As Double
    Dim msgError As String
    Dim actividad As String
    Dim rsPP As rdoResultset, pp As Double
    
    Dim nCierres As Integer
    Dim importZTotal As Double, importVTotal As Double
  
    Dim rsRectifica As rdoResultset, rsAnulats As rdoResultset
    
    Dim asientosCaixa As String
    asientosCaixa = "0"
    
    On Error GoTo noExportat
    
    If nEmpresa = "0" Then
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampNif' ")
        Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampActividadSAGE' ")
    Else
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampNif' ")
        Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampActividadSAGE' ")
    End If
    
    actividad = ""
    If Not rsCActividad.EOF Then actividad = rsCActividad("valor")
    If actividad = "0" Then actividad = ""
    
    Dim nifEmpMurano As String, rsEmpMurano As rdoResultset
    
    If Not rsCtb.EOF Then nifEmpMurano = rsCtb("valor")
    Set rsEmpMurano = Db.OpenResultset("select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'")
    If Not rsEmpMurano.EOF Then
        EmpresaMurano = rsEmpMurano("CodigoEmpresa")
    Else
        msgError = "NO HAY EMPRESA MURANO:<br>" & "select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'"
        GoTo noExportat
    End If

    nifClienteContado = "22222222J"

    D = Di
    
    primerTick = Split(intTickets, ",")(0)
    ultimTick = Split(intTickets, ",")(1)
    
    Set rsNA = Db.OpenResultset("select isnull(max(Asiento), 0) + 1 NumAsiento from " & TaulaHistoricoMURANO(D) & " where CodigoEmpresa = " & EmpresaMurano & " and year(fechaAsiento) = " & Year(D))
    If Not rsNA.EOF Then numAsiento = rsNA("NumAsiento")
    
    Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(D) & " where month(FechaAsiento)=" & Month(D) & " and day(FechaAsiento)=" & Day(D) & " and Param1 = '" & botiga & "' and Param2 = '" & primerTick & "' and Param3 = '" & ultimTick & "' and TipoExportacion='CAIXA' and CodigoEmpresa=" & EmpresaMurano & " order by asiento")
    If Not rsHist.EOF Then numAsiento = rsHist("Asiento")
    While Not rsHist.EOF
        ExecutaComandaSql "Delete from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & rsHist("Asiento") & " and (statusTraspasadoIME=0 or statusTraspasadoIME=2) and CodigoEmpresa=" & EmpresaMurano
        rsHist.MoveNext
    Wend
    
    Set rsSage = Db.OpenResultset("select * from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & numAsiento & " and statusTraspasadoIME=1 and CodigoEmpresa=" & EmpresaMurano & " and ejercicio = " & Year(D))
    If Not rsSage.EOF Then
        msgError = "YA ESTABA TRASPASADA<br> " & "select * from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & numAsiento & " and statusTraspasadoIME=1 and CodigoEmpresa=" & EmpresaMurano & " and ejercicio = " & Year(D)
        GoTo noExportat  'Si ya se ha traspasado correctamente a MURANO no volvemos a traspasarla
    End If
    
    TipusDeIva T1, T2, T3, T4, TR1, TR2, TR3, TR4, D
    
    cCtble = botiga
    Set rsCodi = Db.OpenResultset("SELECT Valor FROM " & tablaConstantsClient() & " WHERE  codi = " & botiga & " AND variable = 'CodiContable' ")
    If Not rsCodi.EOF Then If Not IsNull(rsCodi("Valor")) And (Len(rsCodi("Valor")) > 0) And IsNumeric(rsCodi("Valor")) Then cCtble = CDbl(rsCodi("Valor"))
    
    'If UCase(EmpresaActual) <> UCase("Tena") Then nifClienteContado = nifClienteContado & "," & cCtble
    
    importZTotal = 0
                  
    'Metálico
    'import = total de ventas - cobros tarjeta - cobros cheques - descuadre
    InformaMiss "MURANO metàl·lic Botiga: " & BotigaCodiNom(botiga) & " Dia: " & D, True
                    
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

    If import = 0 Then GoTo noExportat 'NO HAN LLEGADO AÚN LOS DATOS
    
    'ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", " & numAsiento & ", 'CAIXA')"
    ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
    AsientoAddMURANO_TS actividad, 0, numAsiento, D, "43" & Right("000000000000" & cCtble, nDigitos - 2), "B", "", "", 0, import, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "COBRO EN METALICO CAJA REGISTRA " & BotigaCodiNom(botiga), ""
    AsientoAddMURANO_TS actividad, 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", import, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "COBRO EN METALICO CAJA REGISTRA " & BotigaCodiNom(botiga), ""
    
    ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
    
    asientosCaixa = asientosCaixa + "," + numAsiento
                    
    numAsiento = numAsiento + 1
    
    'DIFERENCIA PREVISION vs Z
    pp = 0
    Set rsPP = Db.OpenResultset("select isnull([Desconte ProntoPago], 0) pp from clients where codi=" & botiga)
    If Not rsPP.EOF Then pp = rsPP("PP")
    If pp > 0 Then
        sql = "select isnull(sum(import), 0) as import "
        sql = sql & "from [" & NomTaulaVentasAoB(D, True) & "] "
        sql = sql & "where Data between convert(datetime,'" & Di & "', 103) and convert(datetime,'" & Df & "', 103) And botiga = '" & botiga & "' "
        Set rsCtb = Db.OpenResultset(sql)
        
        If Not rsCtb.EOF Then
            If rsCtb("Import") > 0 Then
                ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
                AsientoAddMURANO_TS actividad, 0, numAsiento, D, "43" & Right("000000000000" & cCtble, nDigitos - 2), "B", "", "", importZ - rsCtb("import"), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "DIFERENCIA PREVISIÓN " & BotigaCodiNom(botiga), ""
                AsientoAddMURANO_TS actividad, 0, numAsiento, D, "70" & Right("000000000000", nDigitos - 2), "", "", "", 0, importZ - rsCtb("import"), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DIFERENCIA PREVISIÓN " & BotigaCodiNom(botiga), ""
                
                asientosCaixa = asientosCaixa + "," + numAsiento
                
                numAsiento = numAsiento + 1
            End If
        End If
    End If

    If EmpresaMurano = "62" Then 'PROBANDO PAGOS CON VISA A CUENTA PUENTE
        If Abs(iTargeta) > 0 Then
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("43" & Right("000000000000" & cCtble, nDigitos - 2)), "B", "", "", 0, Abs(iTargeta), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Cobros tarjeta " & BotigaCodiNom(botiga), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("572" & Right("99999999", nDigitos - 3)), "", "", "", Abs(iTargeta), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Cobros tarjeta " & BotigaCodiNom(botiga), ""
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
        End If
    End If
    
    'DEUDAS
    sql = "Select c.codi botigaCodi, abs(sum(import)) as import "
    sql = sql & "from [" & NomTaulaMovi(D) & "] v "
    sql = sql & "left join clients c on v.botiga = c.codi "
    sql = sql & "where Data between convert(datetime,'" & Di & "', 103) and convert(datetime,'" & Df & "', 103) And botiga = '" & botiga & "' and motiu like 'D.%' "
    sql = sql & "Group By c.codi"
    Set rsCtb = Db.OpenResultset(sql)
    If Not rsCtb.EOF Then iDeudas = rsCtb("import")
    If Abs(iDeudas) > 0 Then
        ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
        AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("446" & Right("000000000000", nDigitos - 3)), "", "", "", 0, Abs(iDeudas), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Deudas " & BotigaCodiNom(botiga), ""
        AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("43" & Right("000000000000" & cCtble, nDigitos - 2)), "B", "", "", Abs(iDeudas), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Deudas " & BotigaCodiNom(botiga), ""
        
        ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
        
        asientosCaixa = asientosCaixa + "," + numAsiento
        
        numAsiento = numAsiento + 1
    End If

    
    'COBROS DEUDAS
    sql = "Select c.codi botigaCodi, abs(sum(import)) as import "
    sql = sql & "from [" & NomTaulaMovi(D) & "] v "
    sql = sql & "left join clients c on v.botiga = c.codi "
    sql = sql & "where Data between convert(datetime,'" & Di & "', 103) and convert(datetime,'" & Df & "', 103) And botiga = '" & botiga & "' and motiu like 'P.%' "
    sql = sql & "Group By c.codi"
    Set rsCtb = Db.OpenResultset(sql)
    If Not rsCtb.EOF Then iCobros = rsCtb("import")
    If Abs(iCobros) > 0 Then
        ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
        AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("43" & Right("000000000000" & cCtble, nDigitos - 2)), "B", "", "", 0, Abs(iCobros), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Cobros Deudas " & BotigaCodiNom(botiga), ""
        AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("446" & Right("000000000000", nDigitos - 3)), "", "", "", Abs(iCobros), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Cobros Deudas " & BotigaCodiNom(botiga), ""
        
        ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
        
        asientosCaixa = asientosCaixa + "," + numAsiento
        
        numAsiento = numAsiento + 1
    End If
    
    'Moviments ENTRADA/SORTIDA
    InformaMiss "MURANO Moviments ENTRADA/SORTIDA Botiga: " & BotigaCodiNom(botiga) & " Dia: " & D, True
    
    sql = "Select motiu, tipus_moviment, c.codi botigaCodi, sum(import) as import "
    sql = sql & "from [" & NomTaulaMovi(D) & "] v "
    sql = sql & "left join clients c on v.botiga = c.codi "
    sql = sql & "where tipus_moviment in ('O','A') and Data between convert(datetime,'" & Di & "', 103) and convert(datetime,'" & Df & "', 103) And botiga = '" & botiga & "' and Import<>0 "
    sql = sql & "and motiu not like 'Pagat Targeta%' and motiu not like 'P.%' and motiu not like 'D.%' "
    sql = sql & "Group By motiu, tipus_moviment, c.codi "
    
    Set rsCtb = Db.OpenResultset(sql)
    While Not rsCtb.EOF
        import = Format(rsCtb("Import"), "0.0#")
        Motiu = rsCtb("motiu") & " " & BotigaCodiNom(botiga)

        If rsCtb("motiu") = "Entrega Diària" Or UCase(rsCtb("motiu")) = UCase("Entrega Diaria") Or rsCtb("motiu") = "Sortida de Canvi" Then
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "5701" & Right("000000000000" & cCtble, nDigitos - 4), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left("SALIDA DE CAJA A CAJA FUERTE " & BotigaCodiNom(botiga), 40), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "SALIDA DE CAJA A CAJA FUERTE " & BotigaCodiNom(botiga), ""
            
            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
            
        'ElseIf Left(rsCtb("motiu"), 13) = "Pagat Targeta" Then

        ElseIf Left(rsCtb("motiu"), 10) = "Pagat TkRs" Then
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("43" & Right("000000000000" & cCtble, nDigitos - 2)), "B", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Cobros ticket restaurant " & BotigaCodiNom(botiga), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("44" & Right("000000000000", nDigitos - 2)), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Cobros ticket restaurant " & BotigaCodiNom(botiga), ""
            
            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
            
        ElseIf Left(rsCtb("motiu"), 9) = "Excs.TkRs" Then
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("768" & Right("000000000000", nDigitos - 3)), "B", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Exc. cobro CHQ RTE", ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("44" & Right("000000000000", nDigitos - 2)), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Exc. cobro CHQ RTE", ""
            
            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
            
        ElseIf Left(rsCtb("motiu"), 6) = "Albara" Then 'ALBARANES
            InformaMiss "MURANO Moviments Botiga Albara " & D, True

        'ElseIf Left(rsCtb("motiu"), 2) = "D." Then 'DEIXA A DEURE
            'InformaMiss "MURANO Moviments Deixen deute " & D, True
            
        'ElseIf Left(rsCtb("motiu"), 2) = "P." Then 'PAGA DEUTE
        '    InformaMiss "MURANO Moviments Paguen deute " & D, True
        
        'SOTIDES DE DINERS
        Else
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "629005" & Right("000000000000" & cCtble, nDigitos - 6), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "GASTOS " & BotigaCodiNom(rsCtb("botigaCodi")) & " " & rsCtb("motiu"), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "GASTOS " & BotigaCodiNom(rsCtb("botigaCodi")) & " " & rsCtb("motiu"), ""
            
            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
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
    sql = sql & "where tipus_moviment = 'J' and Data between convert(datetime,'" & Di & "', 103) and convert(datetime,'" & Df & "', 103) And botiga = '" & botiga & "' "
    sql = sql & "Order By data "
    
    Set rsCtb = Db.OpenResultset(sql)
    While Not rsCtb.EOF
        import = rsCtb("import")

        If import < 0 Then 'DESCUADRE NEGATIVO
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "659" & Right("000000000000" & cCtble, nDigitos - 3), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            
            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
        Else 'DESCUADRE POSITIVO
            ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", 'CAIXA')"
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "57" & Right("000000000000" & cCtble, nDigitos - 2), "", "", "", Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""
            AsientoAddMURANO_TS actividad, 0, numAsiento, D, "659" & Right("000000000000" & cCtble, nDigitos - 3), "", "", "", 0, Abs(import), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "DESCUADRE " & BotigaCodiNom(rsCtb("botigaCodi")), ""

            ExportaANALITICA_CaixaBotigaOnLine numAsiento, D, botiga
            
            asientosCaixa = asientosCaixa + "," + numAsiento
            
            numAsiento = numAsiento + 1
        End If
        rsCtb.MoveNext
    Wend
    rsCtb.Close
    
    Exit Sub
    
noExportat:
    'ExecutaComandaSql "delete from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where asiento in (" & asientosCaixa & ")"
    InsertFeineaAFer "SincroMURANOCaixaOnLine", "[" & botiga & "]", "[" & Format(Di, "dd-mm-yyyy") & " " & Format(Di, "hh:mm:ss") & "]", "[" & Format(Df, "dd-mm-yyyy") & " " & Format(Df, "hh:mm:ss") & "]", "[" & primerTick & "," & ultimTick & "]", "[" & Z & "]"
    'sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR ExportaMURANO_CaixaBotigaOnLine_V3", err.Description & "<br>" & msgError, "", ""
End Sub



Sub ExportaMURANO_VendesBotiga(nEmpresa As Double, botiga As String, Optional reventa As String)
    Dim D As Date, sql As String, f As Date
    Dim rsCtb As rdoResultset
    Dim rsCActividad As rdoResultset, actividad As String
    Dim nifEmpMurano As String, rsEmpMurano As rdoResultset
    Dim nifClienteContado As String, nifCContadoEU As String
    Dim rsSiguienteTick As rdoResultset, siguienteTick As String, fUltimoTickExportado As Date
    Dim rsNA As rdoResultset, numAsiento As String
    Dim cCtble, rsCodi As rdoResultset
    Dim T1 As Double, T2 As Double, T3 As Double, T4 As Double, TR1 As Double, TR2 As Double, TR3 As Double, TR4 As Double
    Dim primerTick As Long, ultimTick As Long, fUltim As Date
    Dim import As Double, tipoIva, PctIva, Base, Quota, CuentaVentas
    Dim CcVentas As String
    Dim Motiu As String
    Dim iva1 As Double, baseIva1 As Double, iva2 As Double, baseIva2 As Double, iva3 As Double, baseIva3 As Double
    Dim msgError As String
    Dim tVenut As String, rsPP As rdoResultset, pp As Double
    
    On Error GoTo noExportat
    
    D = Now()
    
    If nEmpresa = "0" Then
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampNif' ")
        Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampActividadSAGE' ")
    Else
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampNif' ")
        Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampActividadSAGE' ")
    End If
    
    'ACTIVIDAD
    actividad = ""
    If Not rsCActividad.EOF Then actividad = rsCActividad("valor")
    If actividad = "0" Then actividad = ""
    
    'EMPRESA CORRESPONDIENTE EN SAGE
    If Not rsCtb.EOF Then nifEmpMurano = rsCtb("valor")
    Set rsEmpMurano = Db.OpenResultset("select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'")
    If Not rsEmpMurano.EOF Then
        EmpresaMurano = rsEmpMurano("CodigoEmpresa")
    Else
        msgError = "NO HAY EMPRESA MURANO:<br>" & "select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'"
        GoTo noExportat
    End If

    'IVAS
    TipusDeIva T1, T2, T3, T4, TR1, TR2, TR3, TR4, D
    
    nifClienteContado = "22222222J"
    
    cCtble = botiga
    Set rsCodi = Db.OpenResultset("SELECT Valor FROM " & tablaConstantsClient() & " WHERE  codi = " & botiga & " AND variable = 'CodiContable' ")
    If Not rsCodi.EOF Then If Not IsNull(rsCodi("Valor")) And (Len(rsCodi("Valor")) > 0) And IsNumeric(rsCodi("Valor")) Then cCtble = CDbl(rsCodi("Valor"))
    
    'EL ÚLTIMO TICKET EXPORTADO + 1 SERÁ EL PRIMER TICKET DE LA SIGUIENTE EXPORTACIÓN
    'siguienteTick = 1
    fUltimoTickExportado = CDate("01-01-" & Year(D))
    Set rsSiguienteTick = Db.OpenResultset("select *, param3 + 1 nTick from " & TaulaHistoricoMURANO(D) & " where codigoempresa=" & EmpresaMurano & " and tipoExportacion='VENTAS" & actividad & "' and param1='" & botiga & "' order by fechaAsiento desc")
    If Not rsSiguienteTick.EOF Then
        'siguienteTick = rsSiguienteTick("nTick")
        fUltimoTickExportado = rsSiguienteTick("fechaAsiento")
    Else
        Set rsSiguienteTick = Db.OpenResultset("select top 1 NumeroFacturaFinal_+1 nTick, fechaAsiento From " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where codigocuenta = '" & ("43" & Right("000000000000" & cCtble, nDigitos - 2)) & "' and serie='" & Left(BotigaCodiNom(botiga), 10) & "' and CargoAbono='D' and statusTraspasadoime in (0,1) and codigoActividad='" & actividad & "' and ejercicio=" & Year(D) & " and codigoEmpresa=" & EmpresaMurano & " order by fechaasiento desc")
        If Not rsSiguienteTick.EOF Then
            'siguienteTick = rsSiguienteTick("nTick")
            fUltimoTickExportado = rsSiguienteTick("fechaAsiento")
        End If
    End If
    
    'For f = fUltimoTickExportado To DateAdd("d", -1, D)
    For f = fUltimoTickExportado To D
    
        'TRASPASO DE VENTAS
        InformaMiss "MURANO Vendes Botiga: " & BotigaCodiNom(botiga) & " Dia: " & f, True
        
        'NÚMERO DE ASIENTO
        Set rsNA = Db.OpenResultset("select isnull(max(Asiento), 0) + 1 NumAsiento from " & TaulaHistoricoMURANO(D) & " where CodigoEmpresa = " & EmpresaMurano & " and year(fechaAsiento) = " & Year(D))
        If Not rsNA.EOF Then numAsiento = rsNA("NumAsiento")

        pp = 0
        Set rsPP = Db.OpenResultset("select isnull([Desconte ProntoPago], 0) pp from clients where codi=" & botiga)
        If Not rsPP.EOF Then
            pp = rsPP("PP")
        End If
        tVenut = NomTaulaVentasAoB(f, (pp > 0))
        
        'TOTAL VENTAS
        sql = "Select isnull(sum(vv.import), 0) Import, min(num_tick) primerTick, max(num_tick) ultimTick, max(data) fUltim "
        sql = sql & "From [" & tVenut & "] vv "
        sql = sql & "where day(vv.data) = " & Day(f) & " and vv.botiga = '" & botiga & "' and data > '" & fUltimoTickExportado & "' " 'and num_tick>=" & siguienteTick
        Set rsCtb = Db.OpenResultset(sql)
        If Not rsCtb.EOF Then
            If rsCtb("Import") > 0 Then
                primerTick = rsCtb("primerTick")
                ultimTick = rsCtb("ultimTick")
                fUltim = rsCtb("fUltim")
                
                'CUOTAS DE IVA
                sql = "Select isnull(a.tipoiva, 2) tipoiva, v.botiga, sum(v.import) as import "
                sql = sql & "from ( "
                sql = sql & "Select vv.Botiga, vv.Plu ,sum(vv.import) Import "
                sql = sql & "From [" & tVenut & "] vv "
                If reventa <> "" Then
                    sql = sql & "left join ArticlesPropietats ap on vv.plu=ap.CodiArticle and ap.variable='MatPri' "
                    sql = sql & "left join ccMateriasPrimas mp on ap.valor=mp.id "
                    sql = sql & "left join ccNombreValor nv on mp.id = nv.id and nv.nombre='RecargoEquivalencia' "
                End If
                sql = sql & "where day(vv.data) = " & Day(f) & " and vv.botiga = '" & botiga & "' and data > '" & fUltimoTickExportado & "' "  'and num_tick>=" & siguienteTick & " "
                If reventa = "SI" Then
                    sql = sql & " and isnull(ap.valor, '') <> '' and isnull(nv.valor, '') = 'on' "
                ElseIf reventa = "NO" Then
                    sql = sql & " and isnull(ap.valor, '') = '' and isnull(nv.valor, '') = '' "
                End If
                sql = sql & "group by vv.botiga, vv.Plu) v "
                sql = sql & "Left Join "
                sql = sql & "(select Aa.codi, aa.TipoIva "
                sql = sql & "From (select Familia,codi,tipoiva from articles union select Familia,codi,tipoiva from articles_zombis) aa ) a on a.codi = v.plu "
                sql = sql & "group by botiga, isnull(a.tipoiva, 2) "
                sql = sql & "order by botiga, isnull(a.tipoiva, 2) "
                Set rsCtb = Db.OpenResultset(sql)
                
                import43 = 0
                iva1 = 0: baseIva1 = 0:  iva2 = 0: baseIva2 = 0: iva3 = 0: baseIva3 = 0
                While Not rsCtb.EOF
                    import = rsCtb("Import")
                    tipoIva = rsCtb("TipoIva")
                    
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
                    AsientoAddMURANO_TS actividad, 0, numAsiento, f, ("477" & Right("000000000000", nDigitos - 3)) + PctIva, "", "", "", 0, Quota, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Ventas " & BotigaCodiNom(botiga), ""
                    
                    DoEvents
            
                    rsCtb.MoveNext
                Wend
                rsCtb.Close
                
                'BASES IVA
                sql = "Select isnull(a.tipoiva, 2) tipoiva, v.botiga, sum(v.import) as import,  ISNULL(fe1.valor, '" & Left("7000000000000", nDigitos) & "') CC "
                sql = sql & "from ( "
                sql = sql & "Select vv.Botiga, vv.Plu ,sum(vv.import) Import "
                sql = sql & "From [" & tVenut & "] vv "
                If reventa <> "" Then
                    sql = sql & "left join ArticlesPropietats ap on vv.plu=ap.CodiArticle and ap.variable='MatPri' "
                    sql = sql & "left join ccMateriasPrimas mp on ap.valor=mp.id "
                    sql = sql & "left join ccNombreValor nv on mp.id = nv.id and nv.nombre='RecargoEquivalencia' "
                End If
                sql = sql & "where day(vv.data) = " & Day(f) & " and vv.botiga = '" & botiga & "' and data > '" & fUltimoTickExportado & "' "  ' and num_tick>=" & siguienteTick & " "
                If reventa = "SI" Then
                    sql = sql & " and isnull(ap.valor, '') <> '' and isnull(nv.valor, '') = 'on' "
                ElseIf reventa = "NO" Then
                    sql = sql & " and isnull(ap.valor, '') = '' and isnull(nv.valor, '') = '' "
                End If
                sql = sql & "group by vv.botiga, vv.Plu) v "
                sql = sql & "Left Join "
                sql = sql & "(select Aa.codi, aa.TipoIva, aa.familia "
                sql = sql & "From (select Familia,codi,tipoiva from articles union select Familia,codi,tipoiva from articles_zombis) aa ) a on a.codi = v.plu "
                sql = sql & "left join families F3 on a.familia = F3.nom "
                sql = sql & "left join families F2 on F2.nom = F3.pare "
                sql = sql & "left join families F1 on F1.nom = F2.Pare "
                sql = sql & "left join familiesExtes fe1 on F1.nom=fe1.familia and fe1.variable='CUENTA_CONTABLE' "
                sql = sql & "group by botiga, isnull(a.tipoiva, 2), ISNULL(fe1.valor, '" & Left("7000000000000", nDigitos) & "') "
                sql = sql & "order by botiga, isnull(a.tipoiva, 2) "
                Set rsCtb = Db.OpenResultset(sql)
                                
                While Not rsCtb.EOF
                    import = rsCtb("Import")
                    tipoIva = rsCtb("TipoIva")
                    CcVentas = rsCtb("CC")
                    
                    Select Case tipoIva
                        Case 1: PctIva = T1
                        Case 2: PctIva = T2
                        Case 3: PctIva = T3
                    End Select
                    Base = Round(import / (1 + (PctIva / 100)), 2)
                    
                    AsientoAddMURANO_TS actividad, 0, numAsiento, f, CcVentas, "", "", "", 0, Base, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Ventas " & BotigaCodiNom(botiga), ""
                    
                    DoEvents
            
                    rsCtb.MoveNext
                Wend
                rsCtb.Close
                
                'TOTAL
                AsientoAddMURANO_TS actividad, 0, numAsiento, f, ("43" & Right("000000000000" & cCtble, nDigitos - 2)), "E", BotigaCodiNom(botiga), CStr(primerTick), import43, 0, iva1, baseIva1, iva2, baseIva2, iva3, baseIva3, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Ventas " & BotigaCodiNom(botiga), "", Left(BotigaCodiNom(botiga), 10), CStr(primerTick), CStr(ultimTick), "B"
                                
                ExportaANALITICA_CaixaBotigaOnLine numAsiento, f, botiga
                
                ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & Format(fUltim, "dd/mm/yyyy hh:nn:ss") & "', 103), " & botiga & ", '" & Format(f, "dd/mm/yyyy") & "', '" & ultimTick & "', " & numAsiento & ", 'VENTAS" & actividad & "')"
                
                'siguienteTick = ultimTick + 1
            End If
        End If
    Next
                  
    Exit Sub
    
noExportat:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR ExportaMURANO_VendesBotiga", err.Description & "<br>" & msgError, "", ""
End Sub


'NO SE USA, ERA PARA CONCORDIA, AHORA SE EXPORTAN LAS VENTAS POR ExportaMURANO_VendesBotiga
Sub ExportaMURANO_CaixaBotigaOnLine_V2(botiga As String, Di As Date, Df As Date, intTickets As String, Z As Double, reventa As Boolean)
    Dim nEmpresa As String
    Dim D As Date, sql As String, rsCtb As rdoResultset, rsCActividad As rdoResultset
    Dim import As Double, tipoIva, PctIva, Base, Quota, CuentaVentas
    Dim T1 As Double, T2 As Double, T3 As Double, T4 As Double, TR1 As Double, TR2 As Double, TR3 As Double, TR4 As Double
    Dim cCtble, rsCodi As rdoResultset
    Dim rsHist As rdoResultset
    Dim rsNA As rdoResultset, rsSage As rdoResultset
    Dim rsCaixes As rdoResultset, rsCCBanc As rdoResultset, ccBanc As String
    Dim CcVentas As String
    Dim iTargeta As Double, iTkRs As Double, importZ As Double, import43 As Double
    Dim numAsiento As String
    Dim Motiu As String, nifClienteContado As String
    Dim rsBotiga As rdoResultset
    Dim primerTick As String, ultimTick As String
    Dim iva1 As Double, baseIva1 As Double, iva2 As Double, baseIva2 As Double, iva3 As Double, baseIva3 As Double
    Dim msgError As String
    Dim actividad As String
    
    Dim nCierres As Integer
    Dim importZTotal As Double, importVTotal As Double
  
    Dim rsRectifica As rdoResultset, rsAnulats As rdoResultset
    
    Dim asientosCaixa As String
    asientosCaixa = "0"
    
    Dim tipoExportacion As String
    
    On Error GoTo noExportat
    
    If reventa Then
        nEmpresa = "1"
        tipoExportacion = "CAIXA_R"
    Else
        nEmpresa = "0"
        tipoExportacion = "CAIXA"
    End If
    
    If nEmpresa = "0" Then
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampNif' ")
        Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampActividadSAGE' ")
    Else
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampNif' ")
        Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampActividadSAGE' ")
    End If
    
    actividad = ""
    If Not rsCActividad.EOF Then actividad = rsCActividad("valor")
    
    Dim nifEmpMurano As String, rsEmpMurano As rdoResultset
    
    If Not rsCtb.EOF Then nifEmpMurano = rsCtb("valor")
    Set rsEmpMurano = Db.OpenResultset("select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'")
    If Not rsEmpMurano.EOF Then
        EmpresaMurano = rsEmpMurano("CodigoEmpresa")
    Else
        msgError = "NO HAY EMPRESA MURANO:<br>" & "select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'"
        GoTo noExportat
    End If

    nifClienteContado = "22222222J"

    D = Di
    
    primerTick = Split(intTickets, ",")(0)
    ultimTick = Split(intTickets, ",")(1)
    
    Set rsNA = Db.OpenResultset("select isnull(max(Asiento), 0) + 1 NumAsiento from " & TaulaHistoricoMURANO(D) & " where CodigoEmpresa = " & EmpresaMurano & " and year(fechaAsiento) = " & Year(D))
    If Not rsNA.EOF Then numAsiento = rsNA("NumAsiento")
    
    Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(D) & " where month(FechaAsiento)=" & Month(D) & " and day(FechaAsiento)=" & Day(D) & " and Param1 = '" & botiga & "' and Param2 = '" & primerTick & "' and Param3 = '" & ultimTick & "' and TipoExportacion='" & tipoExportacion & "' and CodigoEmpresa=" & EmpresaMurano & " order by asiento")
    If Not rsHist.EOF Then numAsiento = rsHist("Asiento")
    While Not rsHist.EOF
        ExecutaComandaSql "Delete from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & rsHist("Asiento") & " and (statusTraspasadoIME=0 or statusTraspasadoIME=2) and CodigoEmpresa=" & EmpresaMurano
        rsHist.MoveNext
    Wend
    
    Set rsSage = Db.OpenResultset("select * from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & numAsiento & " and statusTraspasadoIME=1 and CodigoEmpresa=" & EmpresaMurano & " and ejercicio = " & Year(D))
    If Not rsSage.EOF Then
        msgError = "YA ESTABA TRASPASADA<br> " & "select * from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & numAsiento & " and statusTraspasadoIME=1 and CodigoEmpresa=" & EmpresaMurano & " and ejercicio = " & Year(D)
        GoTo noExportat  'Si ya se ha traspasado correctamente a MURANO no volvemos a traspasarla
    End If
    
    TipusDeIva T1, T2, T3, T4, TR1, TR2, TR3, TR4, D
    
    cCtble = botiga
    Set rsCodi = Db.OpenResultset("SELECT Valor FROM " & tablaConstantsClient() & " WHERE  codi = " & botiga & " AND variable = 'CodiContable' ")
    If Not rsCodi.EOF Then If Not IsNull(rsCodi("Valor")) And (Len(rsCodi("Valor")) > 0) And IsNumeric(rsCodi("Valor")) Then cCtble = CDbl(rsCodi("Valor"))
    
    InformaMiss "MURANO Vendes Botiga: " & BotigaCodiNom(botiga) & " Dia: " & D, True
    
    importZTotal = 0
    
    'BASES DE IVA
    sql = "Select a.tipoiva, v.botiga, sum(v.import) as import "
    sql = sql & "from ( "
    sql = sql & "Select vv.Botiga, vv.Plu ,sum(vv.import) Import "
    sql = sql & "From [" & NomTaulaVentas(D) & "] vv "
    sql = sql & "left join ArticlesPropietats ap on vv.plu=ap.CodiArticle and ap.variable='MatPri' "
    sql = sql & "left join ccMateriasPrimas mp on ap.valor=mp.id "
    sql = sql & "left join ccNombreValor nv on mp.id = nv.id and nv.nombre='RecargoEquivalencia' "
    sql = sql & "where vv.data between CONVERT(datetime, '" & Di & "', 103) and convert(datetime,'" & Df & "', 103) and vv.botiga = '" & botiga & "' "
    If reventa Then
        sql = sql & " and isnull(ap.valor, '') <> '' and isnull(nv.valor, '') = 'on' "
    Else
        sql = sql & " and isnull(ap.valor, '') = '' and isnull(nv.valor, '') = '' "
    End If
    sql = sql & "group by vv.botiga, vv.Plu) v "
    sql = sql & "Left Join "
    sql = sql & "(select Aa.codi, aa.TipoIva "
    sql = sql & "From (select Familia,codi,tipoiva from articles union select Familia,codi,tipoiva from articles_zombis) aa ) a on a.codi = v.plu "
    sql = sql & "group by botiga, tipoiva "
    sql = sql & "order by botiga, tipoiva "
    Set rsCtb = Db.OpenResultset(sql)
    
    ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(D) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Param2, Param3, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", convert(datetime, '" & D & "', 103), " & botiga & ", '" & primerTick & "', '" & ultimTick & "', " & numAsiento & ", '" & tipoExportacion & "')"
            
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
        AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("477" & Right("000000000000", nDigitos - 3)) + PctIva, "", "", "", 0, Quota, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Ventas " & BotigaCodiNom(botiga), ""
        
        InformaMiss "Ventas " & D, True
        DoEvents

        rsCtb.MoveNext
    Wend
    rsCtb.Close
    
    'FAMILIAS
    sql = "Select a.tipoiva, v.botiga, sum(v.import) as import,  ISNULL(fe1.valor, '" & Left("7000000000000", nDigitos) & "') CC "
    sql = sql & "from ( "
    sql = sql & "Select vv.Botiga, vv.Plu ,sum(vv.import) Import "
    sql = sql & "From [" & NomTaulaVentas(D) & "] vv "
    sql = sql & "left join ArticlesPropietats ap on vv.plu=ap.CodiArticle and ap.variable='MatPri' "
    sql = sql & "left join ccMateriasPrimas mp on ap.valor=mp.id "
    sql = sql & "left join ccNombreValor nv on mp.id = nv.id and nv.nombre='RecargoEquivalencia' "
    sql = sql & "where vv.data between CONVERT(datetime, '" & Di & "', 103) and convert(datetime,'" & Df & "', 103) and vv.botiga = '" & botiga & "' "
    If reventa Then
        sql = sql & " and isnull(ap.valor, '') <> '' and isnull(nv.valor, '') = 'on' "
    Else
        sql = sql & " and isnull(ap.valor, '') = '' and isnull(nv.valor, '') = '' "
    End If
    sql = sql & "group by vv.botiga, vv.Plu) v "
    sql = sql & "Left Join "
    sql = sql & "(select Aa.codi, aa.TipoIva, aa.familia "
    sql = sql & "From (select Familia,codi,tipoiva from articles union select Familia,codi,tipoiva from articles_zombis) aa ) a on a.codi = v.plu "
    sql = sql & "left join families F3 on a.familia = F3.nom "
    sql = sql & "left join families F2 on F2.nom = F3.pare "
    sql = sql & "left join families F1 on F1.nom = F2.Pare "
    sql = sql & "left join familiesExtes fe1 on F1.nom=fe1.familia and fe1.variable='CUENTA_CONTABLE' "
    sql = sql & "group by botiga, tipoiva, ISNULL(fe1.valor, '" & Left("7000000000000", nDigitos) & "') "
    sql = sql & "order by botiga, tipoiva "
    Set rsCtb = Db.OpenResultset(sql)
                    
    While Not rsCtb.EOF
        import = rsCtb("Import")
        tipoIva = rsCtb("TipoIva")
        CcVentas = rsCtb("CC")
        
        InformaMiss "MURANO Vendes Botiga FAMILIA: " & BotigaCodiNom(botiga) & " Dia: " & D, True
        
        Select Case tipoIva
            Case 1: PctIva = T1
            Case 2: PctIva = T2
            Case 3: PctIva = T3
        End Select
        Base = Round(import / (1 + (PctIva / 100)), 2)
        
        AsientoAddMURANO_TS actividad, 0, numAsiento, D, CcVentas, "", "", "", 0, Base, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", "Ventas " & BotigaCodiNom(botiga), ""
        
        InformaMiss "Ventas " & D, True
        DoEvents

        rsCtb.MoveNext
    Wend
    rsCtb.Close
    
    AsientoAddMURANO_TS actividad, 0, numAsiento, D, ("43" & Right("000000000000" & cCtble, nDigitos - 2)), "E", BotigaCodiNom(botiga), primerTick, import43, 0, iva1, baseIva1, iva2, baseIva2, iva3, baseIva3, 0, 0, 0, 0, 0, 0, False, nifClienteContado, "Ventas " & BotigaCodiNom(botiga), "", Left(BotigaCodiNom(botiga), 10), primerTick, ultimTick, "B"
                    
    Exit Sub
    
noExportat:
    sf_enviarMail "secrehit@hit.cat", "ana@solucionesit365.com", "ERROR ExportaMURANO_CaixaBotigaOnLine_V2", err.Description & "<br>" & msgError, "", ""
End Sub



Sub ExportaMURANO_Bancs(idNorma43 As String, fecha As Date)
    Dim rsCtb As rdoResultset, rsN43 As rdoResultset, rsEmp As rdoResultset, rsNA As rdoResultset, rsHist As rdoResultset, Cuenta As String
    Dim nEmpresa As String, numAsiento As String, debe As Double, haber As Double, fechaAsiento As Date
    Dim rsNif As rdoResultset
    Dim rsSage As rdoResultset
    Dim totalDebe As Double, totalHaber As Double
    Dim calaix As String
    Dim rsCActividad As rdoResultset, actividad As String
    
    fechaAsiento = fecha
    On Error GoTo noExportat
    nEmpresa = "0"
    InformaMiss "MURANO BANCS: " & idNorma43 & " Dia: " & fecha, True

    totalDebe = 0
    totalHaber = 0
    
'If Month(fecha) <> 1 Then GoTo noExportat

    nEmpresa = ""
    Set rsN43 = Db.OpenResultset("select * from norma43 where idNorma43 ='" & idNorma43 & "'")
    If Not rsN43.EOF Then
        'nEmpresa = "0"
        Set rsEmp = Db.OpenResultset("select * from constantsempresa where camp like '%CampCuentaContable' and Valor='" & rsN43("Comu_NumCuenta") & "'")
        If Not rsEmp.EOF Then
            If InStr(1, rsEmp("camp"), "_") Then
                nEmpresa = Split(rsEmp("Camp"), "_")(0)
            Else
                nEmpresa = "0"
            End If
        End If
    Else
        If InStr(idNorma43, "|") Then
            calaix = Split(idNorma43, "|")(0)
            Set rsEmp = Db.OpenResultset("select valor from constantsclient where codi='" & calaix & "' and variable='EmpCalaix'")
            If Not rsEmp.EOF Then nEmpresa = rsEmp("valor")
        End If
    End If
    
    If nEmpresa <> "" Then
        
        If nEmpresa = "0" Then
            Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampNif' ")
            Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampActividadSAGE' ")
        Else
            Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampNif' ")
            Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampActividadSAGE' ")
        End If
        
        actividad = ""
        If Not rsCActividad.EOF Then actividad = rsCActividad("valor")
        
        
        Dim nifEmpMurano As String, rsEmpMurano As rdoResultset
        
        If Not rsCtb.EOF Then nifEmpMurano = rsCtb("valor")
        Set rsEmpMurano = Db.OpenResultset("select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'")
        If Not rsEmpMurano.EOF Then
            EmpresaMurano = rsEmpMurano("CodigoEmpresa")
        Else
            GoTo noExportat
        End If
    
    
        Set rsNA = Db.OpenResultset("select isnull(max(Asiento), 0) + 1 NumAsiento from " & TaulaHistoricoMURANO(fecha) & " where CodigoEmpresa = " & EmpresaMurano & " and year(fechaAsiento) = " & Year(fecha))
        If Not rsNA.EOF Then numAsiento = rsNA("NumAsiento")
    
        Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(fecha) & " where month(FechaAsiento)=" & Month(fecha) & " and day(FechaAsiento)=" & Day(fecha) & " and Param1 = '" & idNorma43 & "' and CodigoEmpresa=" & EmpresaMurano & " and TipoExportacion='BANCS'")
        If Not rsHist.EOF Then numAsiento = rsHist("Asiento")
        While Not rsHist.EOF
            ExecutaComandaSql "Delete from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & rsHist("Asiento") & " and (statusTraspasadoIME=0 or statusTraspasadoIME=2) and CodigoEmpresa=" & EmpresaMurano & " and ejercicio=" & Year(fecha)
            rsHist.MoveNext
        Wend
    
        Set rsSage = Db.OpenResultset("select * from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & numAsiento & " and (statusTraspasadoIME=1 or statusTraspasadoIME=2) and CodigoEmpresa=" & EmpresaMurano & " and ejercicio=" & Year(fecha))
        If Not rsSage.EOF Then GoTo noExportat  'Si ya se ha traspasado correctamente a MURANO no volvemos a traspasarla
    
        ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(fecha) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", '" & fecha & "', '" & idNorma43 & "', " & numAsiento & ", 'BANCS')"
        
        Set rsCtb = Db.OpenResultset("select IdNorma43, DEBE, HABER, SUBCUENTA, FECHA, DESCRIPCION, FacturaId, TaulaFactura, CONCEPTO, isnull(ReferenciaInterna, '') ReferenciaInterna from " & DonamNomTaulaAsientosContables(fecha) & " where idNorma43 ='" & idNorma43 & "' order by orden")
    
        Dim sql As String
        Dim fechaAct, cBotiga, cBotigaAct, CcBotiga, CcBotigaAct, nBotiga, Total, tipoIva, strDebe, strIVA, strBase, ReferenciaInterna, FacturaId
        Dim Descripcio, Texte, ImportReal, Deve, acd, Haver, Ach, nif As String
        
        While Not rsCtb.EOF
            nif = ""
            debe = rsCtb("DEBE")
            debe = Round(debe * 10000) / 10000
            haber = rsCtb("HABER")
            haber = Round(haber * 10000) / 10000
            
            If Left(rsCtb("SUBCUENTA"), 1) = "4" Then haber = rsCtb("HABER")
    
            fecha = rsCtb("FECHA")
            'fecha = rsCtb("dataValor")
            ReferenciaInterna = rsCtb("ReferenciaInterna")
            If Len(ReferenciaInterna) > 5 Then
               'Cuenta1 = codiContableProveedor(ReferenciaInterna)
               Set rsNif = Db.OpenResultset("select nif from ccproveedores where id='" & ReferenciaInterna & "'")
               If Not rsNif.EOF Then nif = rsNif("nif")
            Else
               'Cuenta1 = codiContable(ReferenciaInterna)
               Set rsNif = Db.OpenResultset("select nif from clients where codi='" & ReferenciaInterna & "'")
               If Not rsNif.EOF Then nif = rsNif("nif")
            End If
            'If ReferenciaInterna = "" Then
            '    If Cuenta1 = "" Or Cuenta1 = "0" Or Cuenta1 = 0 Then
            '        Cuenta1 = BancoNumConta(rsCtb("IdNorma43"), rsCtb("SUBCUENTA"))
            '    End If
            'End If
            If (Descripcio = "" Or Descripcio = 0) Then Descripcio = rsCtb("DESCRIPCION")
    
            'If Not rsCtb("FacturaId") = "Pendiente" Then
            '    If (FacturaId = "" Or FacturaId = 0) Then
            '        FacturaId = rsCtb("FacturaId")
            '        If FacturaId = "Nomina" Then
            '            Dim codiCentre, CntBrut
            '            codiCentre = CodiCentreTreball(ReferenciaInterna)
                        'Cuenta1 = "4650" & Right("0000000", 5 - Len(codiCentre)) + codiCentre
            '        ElseIf Len(FacturaId) > 5 And InStr(FacturaId, "|") = 0 Then
            '            Cuenta1 = codiContableFacturaProveedor(FacturaId, fecha, ReferenciaInterna)
            '        Else
            '            'Cuenta1 = codiContable(FacturaId)
            '        End If
            '    End If
            'End If
    
            If nif = "" Then
                If InStr(rsCtb("TaulaFactura"), "ccFacturas") Then
                    Set rsNif = Db.OpenResultset("select * from " & rsCtb("TaulaFactura") & " where idFactura='" & rsCtb("facturaId") & "'")
                    If Not rsNif.EOF Then
                        nif = rsNif("EmpNif")
                    End If
                End If
            End If
            
            Texte = rsCtb("CONCEPTO")
            If (Descripcio = "" Or Descripcio = 0) Then Descripcio = rsCtb("CONCEPTO")
    
            If Not rsCtb("DEBE") = "" Then Deve = Round(rsCtb("DEBE"), 2)
            If Deve = "" Then Deve = "0"
            acd = acd + CDbl(Deve)
    
            If Not rsCtb("HABER") = "" Then Haver = Round(rsCtb("HABER"), 2)
            If Haver = "" Then Haver = "0"
            Ach = Ach + CDbl(Haver)
    
    
            Cuenta = Left(rsCtb("SUBCUENTA"), 4) & "0" & Right(rsCtb("SUBCUENTA"), 4)

            AsientoAddMURANO_TS actividad, 0, numAsiento, rsCtb("FECHA"), Cuenta, "B", "", "", Deve, Haver, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nif, Left(LTrim(RTrim(rsCtb("DESCRIPCION"))), 40), ""
            
            totalDebe = totalDebe + Deve
            totalHaber = totalHaber + Haver
            
            fecha = rsCtb("FECHA")
            
            rsCtb.MoveNext
        Wend
        
        If totalDebe <> totalHaber Then
            If (totalHaber - totalDebe) < 0 Then
                AsientoAddMURANO_TS actividad, 0, numAsiento, fechaAsiento, "768000000", "B", "", "", 0, Abs(Round(totalHaber - totalDebe, 2)), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nif, "DESCUADRE BANCOS", ""
            Else
                AsientoAddMURANO_TS actividad, 0, numAsiento, fechaAsiento, "668000000", "B", "", "", Abs(Round(totalHaber - totalDebe)), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, nif, "DESCUADRE BANCOS", ""
            End If
        End If
    End If
noExportat:

End Sub


Sub ExportaMURANO_Manuals(nEmpresa As Double, idNorma43 As String, fecha As Date)
    Dim rsCtb As rdoResultset, rsN43 As rdoResultset, rsEmp As rdoResultset, rsNA As rdoResultset, rsHist As rdoResultset
    Dim numAsiento As String
    Dim debe As Double, haber As Double
    Dim fechaAsiento As Date
    Dim rsCActividad As rdoResultset, actividad As String
    
    fechaAsiento = fecha
      
    On Error GoTo noExportat
    
    InformaMiss "MURANO MANUAL: " & idNorma43 & " Dia: " & fecha, True

    'If nEmpresa = "0" Then
    '    Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampEnlaceNominas' ")
    'Else
    '    Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampEnlaceNominas' ")
    'End If
    'If Not rsCtb.EOF Then EmpresaMurano = Trim(Left(rsCtb("Valor"), InStr(rsCtb("Valor"), " ")))
    
    
    If nEmpresa = "0" Then
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampNif' ")
        Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = 'CampActividadSAGE' ")
    Else
        Set rsCtb = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampNif' ")
        Set rsCActividad = Db.OpenResultset("Select Isnull(valor,'') Valor from constantsempresa  where camp = '" & nEmpresa & "_CampActividadSAGE' ")
    End If
    
    actividad = ""
    If Not rsCActividad.EOF Then actividad = rsCActividad("valor")

    
    Dim nifEmpMurano As String, rsEmpMurano As rdoResultset
    
    If Not rsCtb.EOF Then nifEmpMurano = rsCtb("valor")
    Set rsEmpMurano = Db.OpenResultset("select CodigoEmpresa from " & dbSage & ".dbo.empresas where cifDni='" & nifEmpMurano & "'")
    If Not rsEmpMurano.EOF Then
        EmpresaMurano = rsEmpMurano("CodigoEmpresa")
    Else
        GoTo noExportat
    End If
        
    'Set rsNA = Db.OpenResultset("select isnull(max(Asiento), 0) + 1 NumAsiento from [WEB].[Sage].[dbo].[Movimientos] where CodigoEmpresa = " & EmpresaMurano & " and year(fechaAsiento) = " & Year(fecha))
    'If Not rsNA.EOF Then numAsiento = rsNA("NumAsiento")
    
    Set rsNA = Db.OpenResultset("select isnull(max(Asiento), 0) + 1 NumAsiento from " & TaulaHistoricoMURANO(fecha) & " where CodigoEmpresa = " & EmpresaMurano & " and year(fechaAsiento) = " & Year(fecha))
    If Not rsNA.EOF Then numAsiento = rsNA("NumAsiento")
    
    Set rsHist = Db.OpenResultset("select distinct(Asiento) from " & TaulaHistoricoMURANO(fecha) & " where month(FechaAsiento)=" & Month(fecha) & " and day(FechaAsiento)=" & Day(fecha) & " and Param1 = '" & idNorma43 & "' and CodigoEmpresa=" & EmpresaMurano & " and TipoExportacion='MANUAL'")
    While Not rsHist.EOF
        'MuranoExecute "Delete from [WEB].[Sage].[dbo].[Movimientos] where Asiento = " & rsHist("Asiento") & " and CodigoEmpresa=" & EmpresaMurano
        ExecutaComandaSql "Delete from " & dbSage & ".dbo.A_IMPORTACIO_ASSENTAMENTS where Asiento = " & rsHist("Asiento") & " and CodigoEmpresa=" & EmpresaMurano & " and ejercicio=" & Year(fecha)
        rsHist.MoveNext
    Wend
    
    ExecutaComandaSql "Insert into " & TaulaHistoricoMURANO(fecha) & " (FechaGrabacion, CodigoEmpresa, FechaAsiento, Param1, Asiento, TipoExportacion) values (getdate(), " & EmpresaMurano & ", '" & fecha & "', '" & idNorma43 & "', " & numAsiento & ", 'MANUAL')"
        
    Set rsCtb = Db.OpenResultset("select * from " & DonamNomTaulaAsientosContables(fecha) & " where idNorma43 ='" & idNorma43 & "' order by orden")
    While Not rsCtb.EOF
        debe = rsCtb("DEBE")
        haber = rsCtb("HABER")
        
        'AsientoAddMURANO 0, numAsiento, rsCtb("FECHA"), Left(rsCtb("SUBCUENTA"), 4) & "0" & Right(rsCtb("SUBCUENTA"), 4), "", Left(LTrim(RTrim(rsCtb("DESCRIPCION"))), 20), debe, haber
        AsientoAddMURANO_TS actividad, 0, numAsiento, rsCtb("FECHA"), Left(rsCtb("SUBCUENTA"), 4) & "0" & Right(rsCtb("SUBCUENTA"), 4), "", "", "", debe, haber, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, False, "", Left(LTrim(RTrim(rsCtb("DESCRIPCION"))), 20), ""
        
        rsCtb.MoveNext
    Wend
    
noExportat:

End Sub


