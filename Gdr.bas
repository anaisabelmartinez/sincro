Attribute VB_Name = "Gdr"
'**** Inicializamos variables y conectamos con la BBDD de empresa
Dim slo_connexio, SGS_ND, SGS_REMOTEADDR, SGC_DBUSER, SGC_DBPASSWORD, SGC_DBGDR, SGC_SERVER, SGS_CONGDR, SGS_CONUSER
'*********************************************************************************
'Creación del diccionario con los literales.
'Como vamos a necesitar en el mismo proceso varios idiomas, creamos un diccionario, con los idiomas delante de los codigos de los textos
' es- , ca- o cualquier otro en el futuro
'*********************************************************************************
Dim SGD_Literal
Dim ultimaSql

Dim slo_fs, slo_fname As Object


Sub Horario()
On Error GoTo TractaError 'ErrorHandler
sls_fecha = Now
sls_hoy = Date
'Set slo_fname = Nothing
'****Creamos un fichero de log
If SGS_ND > 0 Then
    sls_dia = Day(sls_fecha)
    sls_mes = Month(sls_fecha)
    sls_any = Year(sls_fecha)
    sls_hora = Hour(sls_fecha)
    sls_minuto = Minute(sls_fecha)
    If Len(sls_dia) = 1 Then
        sls_dia = "0" & sls_dia
    End If
    If Len(sls_mes) = 1 Then
        sls_mes = "0" & sls_mes
    End If
    If Len(sls_hora) = 1 Then
        sls_hora = "0" & sls_hora
    End If
    If Len(sls_minuto) = 1 Then
        sls_minuto = "0" & sls_minuto
    End If
    sls_fname = sls_any & sls_mes & sls_dia & "-" & sls_hora & sls_minuto & ".log"
    Loga ("Iniciamos batch horario para el  " & sls_fecha)
    Loga ("==================================================================================")
End If
'Abrimos cursor sobre la BBDD de empresas
Sql = "select * from  gdrempresas with (nolock) where regEstado='A' and codigo='" & EmpresaActual & "' "
Set rsEmp = sf_recGdr(Sql)
Do While Not rsEmp.EOF
'**** Cargamos empresa
    If SGS_ND > 0 Then
        frmSplash.NomEmpresa = "H - " & Format(Now, "hh:mm") & " - " & rsEmp("codigo")
        InformaMiss "Tratando " & rsEmp("nombre") & " a las " & Time
        Loga ("Tratando empresa " & rsEmp("nombre") & " a las " & Time)
        Loga ("==================================================================================")
    End If
    SGC_DBGDR = rsEmp("database")
    SGC_SERVER = rsEmp("dbServer")
    Set SGS_CONUSER = CreateObject("ADODB.Connection")
    SGS_CONUSER.Open "WSID=" & SGS_REMOTEADDR & ";UID=" & SGC_DBUSER & ";PWD=" & SGC_DBPASSWORD & ";Database=" & SGC_DBGDR & ";Server=" & SGC_SERVER & ";Driver={SQL Server};DSN='';"
'****   Cargamos lso residentes activos con contratos vigentes de CDIA,HRES,RASI activos
    sls_resiCad = ""
    sls_fecha1 = (FormatDateTime(DateAdd("d", 1, sls_hoy), 2) & " 00:00:00.001")   'Dia siguiente
    sls_fecha2 = (FormatDateTime(DateAdd("d", -1, sls_hoy), 2) & " 23:59:59.999")   'Dia anterior
    sls_resiCad = sf_residentes("", 0)
'****
    If SGS_ND > 0 Then
        If SGS_ND > 4 Then
            Loga ("Residentes cargados " & sls_resiCad)
        Else
            Loga ("Residentes cargados")
        End If
        Loga ("==================================================================================")
    End If
'****   Seguimos las alertas definidas en la residencia
    Sql = "select distinct idAlerta from alertasParam with (nolock) where regEstado='A'"
    Set rsAlert = sf_rec(Sql)
    'sls_resiCad=""
    'sls_horasCDIACad=""
    Do While Not rsAlert.EOF
        sls_idAlerta = rsAlert("idAlerta")
        Sql = "select id,tipo,comentario,periodicidad from gdrAlertas with (nolock) where id='" & sls_idAlerta & "' and regEstado='A' and periodicidad='H' "      'Periodicidad H para alertas horarias
        Set rsIdAlert = sf_recGdr(Sql)
        If Not rsIdAlert.EOF Then
            DoEvents
            sls_idAlertaN = rsIdAlert("tipo")
            sla_resiCad = Split(sls_resiCad, ",")
            'Pasamos funciones de alertas para cada residente contenido en la cadena sls_resiCad
            For sln_i = 0 To UBound(sla_resiCad)
                DoEvents
                If SGS_ND > 0 Then
                    Loga ("Tratando residente :" & sla_resiCad(sln_i))
                    Loga ("==================================================================================")
                End If
                Sql = "select centro from residentes with (nolock) where id=" & sla_resiCad(sln_i) & " and estado='A' and regEstado='A'"
                If SGS_ND > 9 Then
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=================================================================================")
                End If
                Set rsResi = sf_rec(Sql)
                If Not rsResi.EOF Then
                    sls_centro = rsResi("centro")
                Else
                    sls_centro = ""
                End If
                Select Case sls_idAlertaN
                    Case "AFEITADO"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de AFEITADO")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("MAXDIAS", sls_idAlertaN)
                        sls_saco = sf_alertaAfeitado(sla_resiCad(sln_i), sls_centro, sls_param1, sls_idAlerta, sls_idAlertaN)
                    Case "BANO"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de BAÑO")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("MAXDIAS", sls_idAlertaN)
                        sls_saco = sf_alertaBano(sla_resiCad(sln_i), sls_centro, sls_param1, sls_idAlerta, sls_idAlertaN)
                    Case "CAIDA"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de CAIDA")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("NCAIDAS", sls_idAlertaN)
                        sls_param2 = sf_getParamAlerta("PERIODOTIEMPO", sls_idAlertaN)
                        sls_saco = sf_alertaCaida(sla_resiCad(sln_i), sls_centro, sls_param1, sls_param2, sls_idAlerta, sls_idAlertaN)
                    Case "CAMAS"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de CAMAS")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("MAXDIAS", sls_idAlertaN)
                        sls_saco = sf_alertaCamas(sla_resiCad(sln_i), sls_centro, sls_param1, sls_idAlerta, sls_idAlertaN)
                    Case "DEPOSICION"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de DEPOSICION")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("MAXHORAS", sls_idAlertaN)
                        sls_saco = sf_alertaDeposicion(sla_resiCad(sln_i), sls_centro, sls_param1, sls_idAlerta, sls_idAlertaN)
                    Case "FCARDIACA"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de FCARDIACA")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("VMAXFCARDIACA", sls_idAlertaN)
                        sls_param2 = sf_getParamAlerta("VMINFCARDIACA", sls_idAlertaN)
                        sls_saco = sf_alertaFCardiaca(sla_resiCad(sln_i), sls_centro, sls_param1, sls_param2, sls_idAlerta, sls_idAlertaN)
                Case "GLICEMIA"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de GLICEMIA")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("VMAXGLICEMIA", sls_idAlertaN)
                        sls_param2 = sf_getParamAlerta("VMINGLICEMIA", sls_idAlertaN)
                        sls_saco = sf_alertaGlicemia(sla_resiCad(sln_i), sls_centro, sls_param1, sls_param2, sls_idAlerta, sls_idAlertaN)
                    Case "O2"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de O2")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("VMAXO2", sls_idAlertaN)
                        sls_param2 = sf_getParamAlerta("VMINO2", sls_idAlertaN)
                        sls_saco = sf_alertaO2(sla_resiCad(sln_i), sls_centro, sls_param1, sls_param2, sls_idAlerta, sls_idAlertaN)
                    Case "PESO"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de PESO")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("VKG", sls_idAlertaN)
                        sls_saco = sf_alertaPeso(sla_resiCad(sln_i), sls_centro, sls_param1, sls_idAlerta, sls_idAlertaN)
                    Case "REGRESAR"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de REGRESAR")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("MAXHORAS", sls_idAlertaN)
                        sls_saco = sf_alertaRegresar(sla_resiCad(sln_i), sls_centro, sls_param1, sls_idAlerta, sls_idAlertaN)
                    Case "TEMPERATURA"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de TEMPERATURA")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("VMAXTEMP", sls_idAlertaN)
                        sls_param2 = sf_getParamAlerta("VMINTEMP", sls_idAlertaN)
                        sls_saco = sf_alertaTemperatura(sla_resiCad(sln_i), sls_centro, sls_param1, sls_param2, sls_idAlerta, sls_idAlertaN)
                    Case "TENSION"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de TENSION")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("VMAXTENSIONMAX", sls_idAlertaN)
                        sls_param2 = sf_getParamAlerta("VMINTENSIONMAX", sls_idAlertaN)
                        sls_param3 = sf_getParamAlerta("VMAXTENSIONMIN", sls_idAlertaN)
                        sls_param4 = sf_getParamAlerta("VMINTENSIONMIN", sls_idAlertaN)
                        sls_saco = sf_alertaTension(sla_resiCad(sln_i), sls_centro, sls_param1, sls_param2, sls_param3, sls_param4, sls_idAlerta, sls_idAlertaN)
                    Case "UNAS"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de UNAS")
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("MAXDIAS", sls_idAlertaN)
                        sls_saco = sf_alertaUnas(sla_resiCad(sln_i), sls_centro, sls_param1, sls_idAlerta, sls_idAlertaN)
                End Select
            Next
        End If
'*Pasamos al siguiente registro
        rsAlert.MoveNext
        DoEvents
    Loop
'********************************************************
'*************MENSAJES AGENDA****************************
    If sf_getParam("RECORDAGENDA") = 1 Then
'*Control de envio de mensajes recordatorio
        Sql = "select a.id,isnull(a.momentoIni,'') momentoIni,isnull(a.momentoFin,'') momentoFin,a.resumen,a.detalle,a.usrAlta "
        Sql = Sql & " from agenda a with (nolock) left join agendaResidente ar with (nolock) on (a.id=ar.idAnotacion and ar.regEstado='A') left join agendaUsuario au on "
        Sql = Sql & "(ar.idAnotacion=au.idAnotacion and au.regEstado='A') where a.regEstado='A' and a.momentoIni<'" & sls_hoy & " 23:59:59.999' "
        Sql = Sql & "and a.momentoFin>convert(datetime,'" & sls_fecha2 & "',103) "
        Sql = Sql & "group by a.id,momentoIni,momentoFin,a.resumen,a.detalle,a.usrAlta "
        Set rsAgenda = sf_rec(Sql)
        If SGS_ND > 0 Then
            Loga ("Control de mensajes recordatorio")
            Loga ("==================================================================================")
        End If
'Genera contenido tabla html a partir de rsAgenda
        sls_idioma = sf_getIdioma(sls_centro)
        Do While Not rsAgenda.EOF
            sls_idAgenda = rsAgenda("id")
            sls_resumen = rsAgenda("resumen")
            sls_momentoIni = rsAgenda("momentoIni")
            sls_momentoFin = rsAgenda("momentoFin")
            sls_detalle = rsAgenda("detalle")
            sls_usrAlta = rsAgenda("usrAlta")
            Sql = "select id from agendaAviso with (nolock) where idAgenda='" & sls_idAgenda & "' and avisado='0' "
            Sql = Sql & "and fecha>=convert(datetime,'" & DateAdd("d", -1, sls_hoy) & "',103) and fecha <convert(datetime,'" & sls_hoy & " 23:59:59.999',103) and regEstado='A' "
            Set rsAviso = sf_rec(Sql)
            If Not rsAviso.EOF Then
                'Enviamos un mail interno al usuario pertinente
                sls_idAviso = rsAviso("id")
                sls_empresa = rsEmp("id")
                sls_momento = sls_hoy
                sls_asunto = SGD_Literal.Item(sls_idioma & "-recordatorioAgenda") & ": " & sls_resumen
                sls_cuerpo = ""
                If sls_momentoIni <> "" Then sls_cuerpo = sls_cuerpo & "<p>" & SGD_Literal.Item(sls_idioma & "-fechaIni") & ": " & sls_momentoIni & "</p>"
                If sls_momentoFin <> "" Then sls_cuerpo = sls_cuerpo & "<p>" & SGD_Literal.Item(sls_idioma & "-fechaFin") & ": " & sls_momentoFin & "</p>"
                Sql = "select  valor from infoEditor with (nolock) where id='" & sls_detalle & "' and regEstado='A' "
                Set rsInfo = sf_rec(Sql)
                If Not rsInfo.EOF Then
                    sls_valor = rsInfo("valor")
                    sls_valor = sf_restaura(sls_valor)
                Else
                    sls_valor = ""
                End If
                sls_cuerpo = sls_cuerpo & sls_valor
                sls_saco = sf_enviarMailInterno("GdR", sls_empresa, sls_usrAlta, sls_asunto, sls_cuerpo)
                If SGS_ND > 0 Then
                    Loga ("Enviado mensaje para usuario " & sls_usrAlta)
                    Loga ("==================================================================================")
                End If
                Sql = "select usuario from agendaUsuario with (nolock) where idAnotacion='" & sls_idAgenda & "' and regEstado='A' "
                Set rsAvisoUsu = sf_rec(Sql)
                Do While Not rsAvisoUsu.EOF
                    sls_a = rsAvisoUsu("usuario")
                    If sls_a <> "" And sls_a <> sls_usrAlta Then sls_saco = sf_enviarMailInterno("GdR", sls_empresa, sls_a, sls_asunto, sls_cuerpo)
                    If SGS_ND > 0 Then
                        Loga ("Enviado mensaje para usuario " & sls_a)
                        Loga ("==================================================================================")
                    End If
                    rsAvisoUsu.MoveNext
                Loop
                Sql = "UPDATE agendaAviso set regEstado='H',fecMod=getDate(),usrMod='GdR' where id='" & sls_idAviso & "' and regEstado='A'"
                Set rsUpd = sf_rec(Sql)
                Sql = "INSERT into agendaAviso(id,idAgenda,fecha,avisado,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                Sql = Sql & "values('" & sls_idAviso & "','" & sls_idAgenda & "','" & FormatDateTime(sls_momentoIni, 2) & "',1,'A',"
                Sql = Sql & "'GdR','GdR',getDate(),getDate())"
                Set rsIns = sf_rec(Sql)
            End If
            rsAgenda.MoveNext
        Loop
    End If
'Quizas guardar un registro de batch correctamente pasado para la fecha
    rsEmp.MoveNext
    DoEvents
Loop
If SGS_ND > 0 Then
    InformaMiss "Fin batch horario para el " & sls_fecha & " . " & Now
    Loga ("Finalizamos batch horario para el " & sls_fecha & " . " & Now)
    Loga ("==================================================================================")
End If
'ErrorHandler
TractaError:
    If Err.Number <> 0 Then
        ultimaSql = sf_limpia(ultimaSql)
        Sql = "INSERT INTO gdrErrorLog (ERROR_ID,ER_TIME,ER_NUMBER,ER_SOURCE,ER_PAGE,ER_DESC,ER_CODE,ER_LINE, ER_REMOTE_ADDR,ER_REMOTE_HOST,ER_LOCAL_ADDR,ER_SQL,ER_EMPRESA) values "
        Sql = Sql & "(NewID(),getDate(),'" & Err.Number & "','" & Err.Source & "','horario','" & Err.Description & "','" & Err.HelpContext & "',"
        Sql = Sql & "'','','','','" & ultimaSql & "','" & EmpresaActual & "')"
        Set rsE = sf_recGdr(Sql)
        InformaMiss "ERROR!!!" & Now & " : " & Err.Number & " " & Err.Description & " "
        Loga ("ERROR!!!" & Now & " : " & Err.Number & " " & Err.Description)
        Loga ("==================================================================================")
        Err.Clear
        Resume Next
    End If
    
    If Not slo_fname Is Nothing Then slo_fname.Close

End Sub

Sub Diario()
On Error GoTo TractaError 'ErrorHandler
'**** Inicializamos fecha
sls_fecha = Now
sls_hoy = Date
Set slo_fname = Nothing
'****Creamos un fichero de log
If SGS_ND > 0 Then
    sls_dia = Day(sls_fecha)
    sls_mes = Month(sls_fecha)
    sls_any = Year(sls_fecha)
    sls_hora = Hour(sls_fecha)
    sls_minuto = Minute(sls_fecha)
    If Len(sls_dia) = 1 Then
        sls_dia = "0" & sls_dia
    End If
    If Len(sls_mes) = 1 Then
        sls_mes = "0" & sls_mes
    End If
    If Len(sls_hora) = 1 Then
        sls_hora = "0" & sls_hora
    End If
    If Len(sls_minuto) = 1 Then
        sls_minuto = "0" & sls_minuto
    End If
    sls_fname = sls_any & sls_mes & sls_dia & "-" & sls_hora & sls_minuto & ".log"
    Loga ("Iniciamos batch diario para el  " & sls_fecha)
    Loga ("==================================================================================")
End If
'Abrimos cursor sobre la BBDD de empresas
Sql = "select * from  gdrempresas with (nolock) where regEstado='A' and codigo='" & EmpresaActual & "'"
Set rsEmp = sf_recGdr(Sql)
Do While Not rsEmp.EOF
'**** Cargamos empresa
    If SGS_ND > 0 Then
        frmSplash.NomEmpresa = "D - " & Format(Now, "hh:mm") & " - " & rsEmp("codigo")
        InformaMiss "Tratando " & rsEmp("nombre") & " a las " & Time
        Loga ("Tratando empresa " & rsEmp("nombre") & " a las " & Time)
        Loga ("==================================================================================")
    End If
    SGC_DBGDR = rsEmp("database")
    SGC_SERVER = rsEmp("dbServer")
    SGS_EMPID = rsEmp("id")
    Set SGS_CONUSER = CreateObject("ADODB.Connection")
    SGS_CONUSER.Open "WSID=" & SGS_REMOTEADDR & ";UID=" & SGC_DBUSER & ";PWD=" & SGC_DBPASSWORD & ";Database=" & SGC_DBGDR & ";Server=" & SGC_SERVER & ";Driver={SQL Server};DSN='';"
'****   Cargamos lso residentes activos con contratos vigentes de CDIA,HRES,RASI activos
    sls_resiCad = sf_residentes("", 0)
'****
    If SGS_ND > 0 Then
        If SGS_ND > 4 Then
            Loga ("Residentes cargados " & sls_resiCad)
        Else
            Loga ("Residentes cargados")
        End If
        Loga ("==================================================================================")
    End If
'****   Cargamos los empleados activos con contratos vigentes
    sls_usuCad = ""
    Sql = "select distinct u.id from usuarios as u with (nolock),contratosUsr as c with (nolock) where u.id=c.usuario and u.regEstado='A' and u.estado='A' "
    Sql = Sql & "   and c.estado='A' "
    Set rsUsu = sf_rec(Sql)
    Do While Not rsUsu.EOF
        sls_usuario = rsUsu("id")
        sls_usuCad = sls_usuCad & "'" & sls_usuario & "',"
        rsUsu.MoveNext
        DoEvents
    Loop
    If sls_usuCad <> "" Then
        sls_usuCad = Mid(sls_usuCad, 1, Len(sls_usuCad) - 1)
    End If
'****   Seguimos las alertas definidas en la residencia
    Sql = "select distinct idAlerta from alertasParam with (nolock) where regEstado='A'"
    Set rsAlert = sf_rec(Sql)
    Do While Not rsAlert.EOF
        sls_idAlerta = rsAlert("idAlerta")
        Sql = "select id,tipo,comentario,periodicidad from gdrAlertas with (nolock) where id='" & sls_idAlerta & "' and regEstado='A' and periodicidad='D' "     'Periodicidad H para alertas horarias
        Set rsIdAlert = sf_recGdr(Sql)
        If Not rsIdAlert.EOF Then
            sls_idAlertaN = rsIdAlert("tipo")
            sla_usuCad = Split(sls_usuCad, ",")
            sla_resiCad = Split(sls_resiCad, ",")
            'Pasamos funciones de alertas para cada usuario contenido en la cadena sls_usuCad
            For sln_i = 0 To UBound(sla_resiCad)
                Sql = "select centro from residentes with (nolock) where id=" & sla_resiCad(sln_i) & " and estado='A' and regEstado='A'"
                If SGS_ND > 9 Then
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=================================================================================")
                End If
                Set rsResi = sf_rec(Sql)
                If Not rsResi.EOF Then
                    sls_centro = rsResi("centro")
                Else
                    sls_centro = ""
                End If
                Select Case sls_idAlertaN
                    Case "CONTRESICADUCA"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de CONTRESICADUCA para residente: " & sla_resiCad(sln_i))
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("MAXDIAS", sls_idAlertaN)
                        sls_saco = sf_alertaContResiCaduca(sla_resiCad(sln_i), sls_centro, sls_param1, sls_idAlerta, sls_idAlertaN)
                End Select
            Next
            'Pasamos funciones de alertas para cada usuario contenido en la cadena sls_usuCad
            For sln_i = 0 To UBound(sla_usuCad)
                Select Case sls_idAlertaN
                    Case "DOCCADUCA"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de DOCCADUCA para usuario: " & sla_usuCad(sln_i))
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("MAXDIAS", sls_idAlertaN)
                        sls_saco = sf_alertaDocCaduca(sla_usuCad(sln_i), SGS_EMPID, sls_param1, sls_idAlerta, sls_idAlertaN)
                    Case "CONTEMPCADUCA"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de CONTEMPCADUCA para usuario: " & sla_usuCad(sln_i))
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("MAXDIAS", sls_idAlertaN)
                        sls_saco = sf_alertaContEmpCaduca(sla_usuCad(sln_i), SGS_EMPID, sls_param1, sls_idAlerta, sls_idAlertaN)
                    Case "EMPLEADOREGRESI"
                        If SGS_ND > 0 Then
                            Loga ("Tratando alerta de EMPLEADOREGRESI para usuario: " & sla_usuCad(sln_i))
                            Loga ("==================================================================================")
                        End If
                        sls_param1 = sf_getParamAlerta("MAXDIAS", sls_idAlertaN)
                        sls_saco = sf_alertaEmpleadoRegResi(sla_usuCad(sln_i), SGS_EMPID, sls_param1, sls_idAlerta, sls_idAlertaN)
                End Select
            Next
        End If
'*Pasamos al siguiente registro
        rsAlert.MoveNext
        DoEvents
    Loop
'Quizas guaradr un registro de batch correctamente pasado para la fecha
    rsEmp.MoveNext
    DoEvents
Loop
If SGS_ND > 0 Then
    InformaMiss "Fin batch diario para el " & sls_fecha
    Loga ("Finalizamos batch diario para el " & sls_fecha)
    Loga ("==================================================================================")
End If
'ErrorHandler
TractaError:
    If Err.Number <> 0 Then
        ultimaSql = sf_limpia(ultimaSql)
        Sql = "INSERT INTO gdrErrorLog (ERROR_ID,ER_TIME,ER_NUMBER,ER_SOURCE,ER_PAGE,ER_DESC,ER_CODE,ER_LINE, ER_REMOTE_ADDR,ER_REMOTE_HOST,ER_LOCAL_ADDR,ER_SQL,ER_EMPRESA) values "
        Sql = Sql & "(NewID(),getDate(),'" & Err.Number & "','" & Err.Source & "','diario','" & Err.Description & "','" & Err.HelpContext & "',"
        Sql = Sql & "'','','','','" & ultimaSql & "','" & EmpresaActual & "')"
        Set rsE = sf_recGdr(Sql)
        InformaMiss "ERROR!!!" & Now & " : " & Err.Number & " " & Err.Description
        Loga ("ERROR!!!" & Now & " : " & Err.Number & " " & Err.Description)
        Loga ("==================================================================================")
        Err.Clear
        Resume Next
    End If
    If Not slo_fname Is Nothing Then slo_fname.Close
    
End Sub



Sub Noche()
On Error GoTo TractaError 'ErrorHandler
'**** Inicializamos fecha
sls_fecha = DateAdd("d", 1, Date)
'sls_fecha = Date
Set slo_fname = Nothing
'****Creamos un fichero de log
If SGS_ND > 0 Then
    sls_dia = Day(sls_fecha)
    sls_mes = Month(sls_fecha)
    sls_any = Year(sls_fecha)
    If Len(sls_dia) = 1 Then
        sls_dia = "0" & sls_dia
    End If
    If Len(sls_mes) = 1 Then
        sls_mes = "0" & sls_mes
    End If
    sls_fname = sls_any & sls_mes & sls_dia & ".log"
    Loga ("Iniciamos batch nocturno para el dia " & sls_fecha & " a las " & Now)
    Loga ("==================================================================================")
End If
'Abrimos cursor sobre la BBDD de empresas
Sql = "select * from  gdrempresas with (nolock) where regEstado='A' and codigo='" & EmpresaActual & "' "
Set rsEmp = sf_recGdr(Sql)
Do While Not rsEmp.EOF
'**** Cargamos empresa
    If SGS_ND > 0 Then
        frmSplash.NomEmpresa = "N - " & Format(Now, "hh:mm") & " - " & rsEmp("codigo") & " - " & sls_fecha
        InformaMiss "Tratando " & rsEmp("nombre") & " a las " & Time
        Loga ("==================================================================================")
        Loga ("Tratando empresa " & rsEmp("nombre") & " a las " & Time)
        Loga ("==================================================================================")
    End If
    SGC_DBGDR = rsEmp("database")
    SGC_SERVER = rsEmp("dbServer")
    Set SGS_CONUSER = CreateObject("ADODB.Connection")
    SGS_CONUSER.Open "WSID=" & SGS_REMOTEADDR & ";UID=" & SGC_DBUSER & ";PWD=" & SGC_DBPASSWORD & ";Database=" & SGC_DBGDR & ";Server=" & SGC_SERVER & ";Driver={SQL Server};DSN='';"
'****   Seguimos los residentes
    If SGS_ND > 0 Then
        Loga ("Tratando RESIDENTES")
        Loga ("==================================================================================")
    End If
    Sql = "select id,nombre,apellido1 from residentes with (nolock) where regEstado='A' and estado='A'"
    Set rsRes = sf_rec(Sql)
    If SGS_ND > 0 Then
        Loga ("Tratando residentes activos ")
        Loga ("==================================================================================")
    End If
    Do While Not rsRes.EOF
        If SGS_ND > 0 Then
            Loga ("Tratando residente " & rsRes("nombre") & " " & rsRes("apellido1") & " " & rsRes("id"))
            Loga ("==================================================================================")
        End If
        sls_id = rsRes("id")
'* Tratamiento Registro medicación
        If SGS_ND > 4 Then
            Loga ("Tratando medicación")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_regMed(sls_fecha, sls_id)
'*  Tratamiento  Otros Registros
        If SGS_ND > 4 Then
            Loga ("Tratando Otros registros")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_regOtrosRegistros(sls_fecha, sls_id)
'*  Tratamiento Registros Contenciones
        If SGS_ND > 4 Then
            Loga ("Tratando Contenciones")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_regContenciones(sls_fecha, sls_id)
'*  Tratamiento Registros CambiosPosturales
        If SGS_ND > 4 Then
            Loga ("Tratando Cambios Posturales")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_regCPosturales(sls_fecha, sls_id)
'*  Tratamiento Registros Curas
        If SGS_ND > 4 Then
            Loga ("Tratando Curas")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_regCuras(sls_fecha, sls_id)
'*  Tratamiento Registros Pañales
        If SGS_ND > 4 Then
            Loga ("Tratando Pañales")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_regPanales(sls_fecha, sls_id)
'* Tratamiento Registro nutricion enteral
        If SGS_ND > 4 Then
            Loga ("Tratando Nutricion Enteral")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_regNEnteral(sls_fecha, sls_id)
'*  Tratamiento finalizacion contratos
        If SGS_ND > 4 Then
            Loga ("Finalizando Contratos Residentes")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_finContrato(sls_fecha, sls_id, "residente")
'*  Tratamiento aniversario
        If SGS_ND > 4 Then
            Loga ("Aniversarios Residentes")
            Loga ("==================================================================================")
        End If
        'sls_saco=sf_aniversario(sls_id)
'*Pasamos al siguiente registro
        rsRes.MoveNext
        DoEvents
    Loop
'****   Seguimos los residentes en pre-alta
    Sql = "select id,nombre,apellido1 from residentes with (nolock) where regEstado='A' and estado='P'"
    Set rsResP = sf_rec(Sql)
    If SGS_ND > 0 Then
        Loga ("Tratando residentes en pre-alta")
        Loga ("==================================================================================")
    End If
    Do While Not rsResP.EOF
        If SGS_ND > 0 Then
            Loga ("Tratando residente " & rsResP("nombre") & " " & rsResP("apellido1") & " " & rsResP("id"))
            Loga ("==================================================================================")
        End If
        sls_id = rsResP("id")
'*  Tratamiento alta contratos
        If SGS_ND > 4 Then
            Loga ("Altas Contratos Residentes")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_altaContrato(sls_fecha, sls_id, "residente")
'*Pasamos al siguiente registro
        rsResP.MoveNext
        DoEvents
    Loop
'****   Cargamos los empleados activos
    If SGS_ND > 0 Then
        Loga ("Tratando USUARIOS")
        Loga ("==================================================================================")
    End If
    Sql = "select id,nombre,apellido1,apellido2 from usuarios with (nolock) where regEstado='A' and estado='A' "
    Set rsUsu = sf_rec(Sql)
    Do While Not rsUsu.EOF
        If SGS_ND > 0 Then
            Loga ("Tratando usuario " & rsUsu("nombre") & " " & rsUsu("apellido1") & " " & rsUsu("id"))
            Loga ("==================================================================================")
        End If
        sls_id = rsUsu("id")
'*  Tratamiento finalizacion contratos
        If SGS_ND > 4 Then
            Loga ("Finalizando Contratos Usuarios")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_finContrato(sls_fecha, sls_id, "usuario")
        rsUsu.MoveNext
        DoEvents
    Loop
    '****   Seguimos los usuarios en pre-alta
    Sql = "select id,nombre,apellido1,apellido2 from usuarios with (nolock) where regEstado='A' and estado='P' "
    Set rsUsuP = sf_rec(Sql)
    Do While Not rsUsuP.EOF
        If SGS_ND > 0 Then
            Loga ("Tratando usuario " & rsUsuP("nombre") & " " & rsUsuP("apellido1") & " " & rsUsuP("id"))
            Loga ("==================================================================================")
        End If
        sls_id = rsUsuP("id")
'*  Tratamiento alta contratos
        If SGS_ND > 4 Then
            Loga ("Altas Contratos Usuarios")
            Loga ("==================================================================================")
        End If
        sls_saco = sf_altaContrato(sls_fecha, sls_id, "usuario")
'*Pasamos al siguiente registro
        rsUsuP.MoveNext
        DoEvents
    Loop
'*Tratamiento tablas de farmacia
    sls_saco = sf_farmaciaBlister(ByVal sls_fecha)
'*Pasamos a la siguiente empresa
    rsEmp.MoveNext
    DoEvents
Loop
If SGS_ND > 0 Then
    InformaMiss "Fin batch nocturno para el " & sls_fecha & " a las " & Now
    Loga ("Finalizamos batch nocturno para el dia " & sls_fecha & " a las " & Now)
    Loga ("==================================================================================")
End If
'ErrorHandler
TractaError:
    If Err.Number <> 0 Then
        ultimaSql = sf_limpia(ultimaSql)
        Sql = "INSERT INTO gdrErrorLog (ERROR_ID,ER_TIME,ER_NUMBER,ER_SOURCE,ER_PAGE,ER_DESC,ER_CODE,ER_LINE, ER_REMOTE_ADDR,ER_REMOTE_HOST,ER_LOCAL_ADDR,ER_SQL,ER_EMPRESA) values "
        Sql = Sql & "(NewID(),getDate(),'" & Err.Number & "','" & Err.Source & "','noche','" & Err.Description & "','" & Err.HelpContext & "',"
        Sql = Sql & "'','','','','" & ultimaSql & "','" & EmpresaActual & "')"
        Set rsE = sf_recGdr(Sql)
        InformaMiss "ERROR!!! " & Now & " : " & Err.Number & " - " & Err.Description
        Loga ("ERROR!!! " & Now & " : " & Err.Number & " - " & Err.Description)
        Loga ("==================================================================================")
        Err.Clear
        Resume Next
    End If
    If Not slo_fname Is Nothing Then slo_fname.Close
End Sub

Sub GdrMain(sls_proceso As String)
    'If SGS_CONGDR Is Nothing Then GdrInit 'Si no hi ha conexio
        
    ExecutaComandaSql "Delete hit.dbo.gosdetura Where NomObella = '" & sls_proceso & "' "
    ExecutaComandaSql "insert into hit.dbo.gosdetura values (newId(),'" & sls_proceso & "','GosDeTura','jordi.bosch.maso@gmail.com','Cada 10 n',getDate(),'','')"
    
    Select Case sls_proceso
        Case "noche"
            'Exit Sub
            InformaMiss "Proces vespre " & Now
            Noche
        Case "diario"
            'Exit Sub
            InformaMiss "Proces matinal " & Now
            Diario
        Case "horario"
            'Exit Sub
            InformaMiss "Proces horari " & Now
            Horario
    End Select
    
    ExecutaComandaSql "Delete hit.dbo.gosdetura Where NomObella = '" & sls_proceso & "' "
    
End Sub
Sub GdrInit()
    SGS_REMOTEADDR = "BATCH"
    SGS_ND = 1 'Nivel de debug 1: mínimo,  5:Básico, 10:Completo
    SGC_DBUSER = "GdrAdmin"
    SGC_DBGDR = "gdr"
    SGC_DBPASSWORD = "A4M719XfU792hK5"
    SGC_SERVER = "SERVERCLOUD" ' "10.1.2.16" 'Nom del servidor de base de dades
    SGC_PATH = App.Path & "c:\data\gdr\batch"
    Set SGS_CONGDR = CreateObject("ADODB.Connection")
    SGS_CONGDR.Open "WSID=" & SGS_REMOTEADDR & ";UID=" & SGC_DBUSER & ";PWD=" & SGC_DBPASSWORD & ";Database=" & SGC_DBGDR & ";Server=" & SGC_SERVER & ";Driver={SQL Server};DSN='';"
    Set SGD_Literal = CreateObject("Scripting.Dictionary")
    SGD_Literal.CompareMode = 2
    'Español
    SGD_Literal.Item("es-SupNumMax") = "Superado el número máximo de "
    SGD_Literal.Item("es-CaidasPeriodo") = " caidas establecido para un periodo dado "
    SGD_Literal.Item("es-SupMax") = "Superado el valor máximo establecido "
    SGD_Literal.Item("es-InfMin") = "Inferior al valor mínimo establecido "
    SGD_Literal.Item("es-SupVarMes") = "Superada la variación mensual máxima de "
    SGD_Literal.Item("es-Faltan") = "Faltan menos de "
    SGD_Literal.Item("es-RegPeso") = " Kg. para registros de peso."
    SGD_Literal.Item("es-RegAutSalida") = " hora/s establecido para una autorización de salida."
    SGD_Literal.Item("es-RegCorteUnas") = " día/s establecido para registros de corte de uñas."
    SGD_Literal.Item("es-RegAfeitado") = " día/s establecido para registros de afeitado."
    SGD_Literal.Item("es-RegBano") = " día/s establecido para registros de baño."
    SGD_Literal.Item("es-RegCambioCama") = " día/s establecido para registros de cambio ropa cama."
    SGD_Literal.Item("es-RegDeposi") = " hora/s establecido para registros de deposición."
    SGD_Literal.Item("es-RegFCardiaca") = " para registros de frecuencia cardiaca."
    SGD_Literal.Item("es-RegTensMax") = " para registros de tensión máxima."
    SGD_Literal.Item("es-RegTensMin") = " para registros de tensión mínima"
    SGD_Literal.Item("es-RegGlicemia") = " para registros de glicemia."
    SGD_Literal.Item("es-RegSatO2") = " para registros de saturación de Oxígeno"
    SGD_Literal.Item("es-RegTemp") = " para registros de temperatura."
    SGD_Literal.Item("es-DocCaduca") = " día/s para que caduque el documento de identidad del empleado."
    SGD_Literal.Item("es-ContEmpCaduca") = " día/s para que caduque el contrato del empleado."
    SGD_Literal.Item("es-ContResiCaduca") = " día/s para que caduque el contrato del residente."
    SGD_Literal.Item("es-fechaIni") = "Fecha inicio"
    SGD_Literal.Item("es-fechaFin") = "Fecha fin"
    SGD_Literal.Item("es-recordatorioAgenda") = "Recordatorio agenda"
    SGD_Literal.Item("es-EmpleadoRegResi") = " dia/s desde que el empleado visitó las fichas de los residentes."
    'Catalán
    SGD_Literal.Item("ca-SupNumMax") = "Superat el número màxim de "
    SGD_Literal.Item("ca-CaidasPeriodo") = " caigudes establert per un periode donat "
    SGD_Literal.Item("ca-SupMax") = "Superat el valor màxim establert "
    SGD_Literal.Item("ca-InfMin") = "Inferior al valor mínim establert "
    SGD_Literal.Item("ca-SupVarMes") = "Superada la variació mensual màxima "
    SGD_Literal.Item("ca-Faltan") = "Falten menys de "
    SGD_Literal.Item("ca-RegPeso") = " Kg. per registres de pes."
    SGD_Literal.Item("ca-RegAutSalida") = " hora/es establert per una autorització de sortida."
    SGD_Literal.Item("ca-RegCorteUnas") = " dia/es establert per registres de tall d'ungles."
    SGD_Literal.Item("ca-RegAfeitado") = " dia/es establert per registres d'afaitat."
    SGD_Literal.Item("ca-RegBano") = " dia/es establert per registres de bany."
    SGD_Literal.Item("ca-RegCambioCama") = " dia/es establert per registres de canvi roba llit."
    SGD_Literal.Item("ca-RegDeposi") = " hora/es establert per registres de deposició."
    SGD_Literal.Item("ca-RegFCardiaca") = " per registres de freqüència cardiaca."
    SGD_Literal.Item("ca-RegTensMax") = " per registres de tensió màxima."
    SGD_Literal.Item("ca-RegTensMin") = " per registres de tensió mínima"
    SGD_Literal.Item("ca-RegGlicemia") = " per registres de glicemia."
    SGD_Literal.Item("ca-RegSatO2") = " per registres de saturació d'Oxígen"
    SGD_Literal.Item("ca-RegTemp") = " per registres de temperatura."
    SGD_Literal.Item("ca-DocCaduca") = " dia/es per a que caduqui el document d@@qs@@identitat de l@@qs@@empleat."
    SGD_Literal.Item("ca-ContEmpCaduca") = " dia/es per a que caduqui el contracte de l@@qs@@empleat."
    SGD_Literal.Item("ca-ContResiCaduca") = " dia/es per a que caduqui el contracte del resident."
    SGD_Literal.Item("ca-fechaIni") = "Data inici"
    SGD_Literal.Item("ca-fechaFin") = "Data fi"
    SGD_Literal.Item("ca-recordatorioAgenda") = "Recordatori agenda"
    SGD_Literal.Item("ca-EmpleadoRegResi") = " dia/es des que l@@qs@@empleat va visitar les fitxes dels residents."
End Sub

'************************************************************************
'* Funcion para sustituir caracteres en base a split + Join
'************************************************************************
Function sf_replace(ByVal sls_entrada, ByVal sls_in, ByVal sls_out)
    If sls_entrada <> "" Then
        sf_replace = Join(Split(sls_entrada, sls_in), sls_out)
    Else
        sf_replace = sls_entrada
    End If
End Function
'
'**********************************************************************************
'* sf_actualizaStock(empresa,database,codigo,residente,fecha,cantidad,usuario)
'*Inserta registro en tabla farmaciaStock-MM-YYYY segun parametros pasados
'**********************************************************************************'
Function sf_actualizaStock(sls_empId, sls_db, sls_codigo, sls_residente, sls_fecha, sln_envase, sls_usuario)
Dim sls_mes, sls_anyo, sls_timeStamp, sls_id, sls_centro, sls_pautaId, sls_articulo, sln_stock, sls_codigo2
Dim sln_stockF, sln_unidades, sln_unidadesF, sls_semana, sls_fechaIni, sls_fechaFin, sls_id2, sls_tipo, sls_fechaDiff

'Declaramos periodo
sls_mes = Month(sls_fecha)
If Len(sls_mes) = 1 Then sls_mes = "0" & sls_mes
sls_anyo = Year(sls_fecha)
sls_semana = sf_fechaSemana(sls_fecha)
sla_semana = Split(sls_semana, ",")
sls_fechaIni = sla_semana(0)
sls_fechaFin = sla_semana(1)
'Formateamos el codigo de dos maneras, con punto separador de ultimo digito o sin
sls_codigo2 = sls_codigo
If InStr(1, sls_codigo2, ".") > 1 Then
    sls_codigo2 = Replace(sls_codigo2, ".", "")
Else
    sls_codigo2 = Mid(sls_codigo, 1, 6) & "." & Mid(sls_codigo, 7, 1)
End If
'Comprobamos si existen tablas farmaciaStock-MM-YYYY
Sql = "SELECT * FROM sys.objects WHERE name='farmaciaStock-" & sls_mes & "-" & sls_anyo & "' AND type='U'"
Set rsTbl = sf_conexionSQL(sls_empId, sls_db, Sql)
If rsTbl.EOF Then
    Sql = "CREATE TABLE [farmaciaStock-" & sls_mes & "-" & sls_anyo & "]("
    Sql = Sql & "[id]             [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[centro]         [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[codResi]        [nvarchar] (25)    COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[residente]      [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[fechaIni]       [datetime]                                        NULL,"
    Sql = Sql & "[fechaFin]       [datetime]                                        NULL,"
    Sql = Sql & "[tipo]           [nvarchar] (25)    COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[articulo]       [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NULL,"
    Sql = Sql & "[cantidad]       [float]                                           NULL,"
    Sql = Sql & "[pautaId]        [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NULL,"
    Sql = Sql & "[stock]          [float]                                           NULL,"
    Sql = Sql & "[regEstado]      [nvarchar] (25)    COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[usrAlta]        [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[usrMod]         [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[fecAlta]        [datetime]                                        NOT NULL,"
    Sql = Sql & "[fecMod]         [datetime]                                        NOT NULL,"
    Sql = Sql & "CONSTRAINT       [PK_farmaciaStock-" & sls_mes & "-" & sls_anyo & "]  PRIMARY KEY CLUSTERED "
    Sql = Sql & "([id] ASC,[regEstado] ASC,[fecAlta] ASC) WITH (PAD_INDEX=OFF, STATISTICS_NORECOMPUTE=OFF, "
    Sql = Sql & "IGNORE_DUP_KEY=OFF, ALLOW_ROW_LOCKS=ON, ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]"
    Sql = Sql & ") ON [PRIMARY]"
    Set Rs = sf_conexionSQL(sls_empId, sls_db, Sql)
    'Indice
    Sql = "CREATE NONCLUSTERED INDEX [GK_farmaciaStock-" & sls_mes & "-" & sls_anyo & "] ON [farmaciaStock-" & sls_mes & "-" & sls_anyo & "] ("
    Sql = Sql & "[residente] ASC,"
    Sql = Sql & "[articulo] ASC,"
    Sql = Sql & "[fechaIni] ASC,"
    Sql = Sql & "[fechaFin] ASC,"
    Sql = Sql & "[regEstado] ASC,"
    Sql = Sql & "[fecAlta] ASC"
    Sql = Sql & " )WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]"
    Set Rs = sf_conexionSQL(sls_empId, sls_db, Sql)
End If
'SQL
Sql = " select r.centro,r.codigo,r.id residente,ar.id pautaId,ar.articulo,a.unidades from articuloResidente ar "
Sql = Sql & "left join residentes r on (ar.residente=r.id and r.regEstado='A') "
Sql = Sql & "left join articulos a on (ar.articulo=a.id and a.regEstado='A') "
Sql = Sql & "where ar.residente in ('" & sls_residente & "') and (a.codigo like '%" & sls_codigo & "%' or a.codigo like '%" & sls_codigo2 & "%') and ar.regEstado='A' "
Sql = Sql & "and ar.fechaInicio<=convert(datetime,'" & sls_fechaFin & "',103) and (ar.fechaFin>=convert(datetime,'" & sls_fechaIni & "',103) or ar.fechaFin is null) "
Set rsMed = sf_conexionSQL(sls_empId, sls_db, Sql)
Do While Not rsMed.EOF
    Sql = "select getDate() timeStamp,newId() id"
    Set rsTime = sf_rec(Sql)
    sls_timeStamp = rsTime("timeStamp")
    sls_id = rsTime("id")
    sls_centro = rsMed("centro")
    sls_codigo = rsMed("codigo")
    sls_residente = rsMed("residente")
    sls_pautaId = rsMed("pautaId")
    sls_articulo = rsMed("articulo")
    sln_unidades = rsMed("unidades")
    sln_stock = 0
    sln_stockF = 0
    'Ultimo stock segun fecha
    Sql = " select top 1 stock from [farmaciaStock-" & sls_mes & "-" & sls_anyo & "] with (nolock) where regEstado='A' and pautaId='" & sls_pautaId & "' "
    Sql = Sql & "and fechaIni<=convert(datetime,'" & sls_fechaFin & "',103) and fechaFin>=convert(datetime,'" & sls_fechaIni & "',103) order by fechaFin desc,fecAlta desc "
    Set rsFs = sf_conexionSQL(sls_empId, sls_db, Sql)
    If Not rsFs.EOF Then
        sln_stock = rsFs("stock")
        sln_stock = sf_replace(sln_stock, ".", ",") 'Para asegurar que las operaciones sean correctas
    End If
    sln_unidadesF = CDbl(sln_envase) * CDbl(sln_unidades)
    sln_stockF = CDbl(sln_unidadesF) + CDbl(sln_stock)
    sln_stockF = sf_replace(sln_stockF, ",", ".")
    Sql = "INSERT into [farmaciaStock-" & sls_mes & "-" & sls_anyo & "] (id,centro,codResi,residente,fechaIni,fechaFin,tipo,articulo,cantidad,pautaId,stock,regEstado,"
    Sql = Sql & "usrAlta,usrMod,fecAlta,fecMod) values('" & sls_id & "','" & sls_centro & "','" & sls_codigo & "','" & sls_residente & "',"
    Sql = Sql & "convert(datetime,'" & sls_fechaIni & "',103),convert(datetime,'" & sls_fechaFin & "',103),'ENTRADA','" & sls_articulo & "'," & sln_unidadesF & ",'" & sls_pautaId & "'," & sln_stockF & ","
    Sql = Sql & "'A','" & sls_usuario & "','" & sls_usuario & "','" & sls_timeStamp & "','" & sls_timeStamp & "')"
    Set rsIns = sf_conexionSQL(sls_empId, sls_db, Sql)
    'Actualizar stock para fechas posteriores
    sls_fechaDiff = DateDiff("m", sls_fechaIni, Date)
    For sln_i = 0 To sls_fechaDiff
        'Declaramos periodo
        sls_mesDiff = Month(DateAdd("m", sln_i, sls_fechaIni))
        If Len(sls_mesDiff) = 1 Then sls_mesDiff = "0" & sls_mesDiff
        sls_anyoDiff = Year(DateAdd("m", sln_i, sls_fechaIni))
        Sql = " select id,tipo,stock from [farmaciaStock-" & sls_mesDiff & "-" & sls_anyoDiff & "] where fechaIni>=convert(datetime,'" & sls_fechaFin & "',103) and regEstado='A' "
        Sql = Sql & "and residente='" & sls_residente & "' and articulo='" & sls_articulo & "' order by fechaIni,fecAlta "
        Set rsStock = sf_conexionSQL(sls_empId, sls_db, Sql)
        sln_stockF = sf_replace(sln_stockF, ".", ",") 'Para asegurar que las operaciones sean correctas
        Do While Not rsStock.EOF
            sls_id2 = rsStock("id")
            sls_tipo = rsStock("tipo")
            If sls_tipo = "ENTRADA" Then
                sln_stockF = CDbl(sln_unidadesF) + CDbl(sln_stockF)
            ElseIf sls_tipo = "SALIDA" Then
                sln_stockF = CDbl(sln_stockF) - CDbl(sln_unidadesF)
            End If
            sln_stockG = sf_replace(sln_stockG, ",", ".")
            Sql = " update [farmaciaStock-" & sls_mesDiff & "-" & sls_anyoDiff & "] set stock='" & sln_stockF & "',usrMod='GdR',fecMod=getDate() where id='" & sls_id2 & "' and regEstado='A' "
            Set rsUpd = sf_conexionSQL(sls_empId, sls_db, Sql)
            rsStock.MoveNext
        Loop
    Next
    rsMed.MoveNext
Loop
End Function

'************************************************************************
'* Funcion sf_fechaSemana devuelve una cadena con la fecha de inicio de la semana  de la fecha pasada y la fecha fin
'* Espera una fecha en formato string con dd/mm/aaaa
'* Las devuelve separadas por ,
'************************************************************************
Function sf_fechaSemana(sls_fec)
    Dim sld_fec
    If sls_fec = "" Or Not IsDate(sls_fec) Then
        sld_fec = Date
    Else
        sld_fec = CDate(sls_fec)
    End If
    sln_ara = Weekday(sld_fec) 'devuelve 1-domingo a 7-sábado
    sln_lunes = sln_ara - 2
    If sln_lunes = -1 Then
        sln_lunes = 6
    End If
    sld_lunes = DateAdd("d", sln_lunes * (-1), sld_fec)
    sld_domingo = DateAdd("d", 6, sld_lunes)
    sf_fechaSemana = CStr(sld_lunes) & "," & CStr(sld_domingo)
End Function

'**********************************************************************************
'* sf_calcPlan(tipo planificacion, planificacion, fecha inicio,fecha calculo)
'* devuelve un array con si toca o no en (0) y la posición para el caso de repetición en 1
'**********************************************************************************
Function sf_calcPlan(ByVal sls_tp, ByVal sls_pla, ByVal sls_fIni, ByVal sls_fecha)
    Dim sln_dif 'Diferencia en dias entre fecha inicio y fecha pasada
    Dim sln_wDIni 'Dia de la semana de la fecha inicio
    Dim sln_wDHoy 'Dia de la semana de la fecha a que estamos calculando
    Dim sln_dHoy ' Dia del mes de la fecha que pasamos
    Dim sln_fm ' Ultimo dia del mes de la fecha que pasamos
    Dim sln_ara 'Temporal
    Dim slb_toca 'flag de si toca o no hoy
    Dim sln_pos ' En caso de repetición, que repetición debemos coger
    Dim sln_k 'Contador for
    slb_toca = 0
    sln_pos = 0
    sln_dif = DateDiff("d", sls_fIni, sls_fecha)
    sln_wDIni = Weekday(sls_fIni, 2)
    sln_wDHoy = Weekday(sls_fecha, 2)
    sln_dHoy = Day(sls_fecha)
    sln_fm = sf_finalMes(sls_fecha)
    If sls_tp = "DIARIA" Then
        If Not IsNumeric(sls_pla) Then
            sls_pla = 1
        End If
        sls_pla = CInt(sls_pla)
        If sls_pla = 0 Then
            sls_pla = 1
            If SGS_ND > 0 Then
                Loga ("Tratando planificacion diaria con dias sin informar. Lo ponemos a 1.")
                Loga ("=================================================================================")
            End If
        End If
        If (sln_dif Mod sls_pla) = 0 Then
            slb_toca = 1
        End If
    ElseIf sls_tp = "SEMANAL" Then
        If Mid(sls_pla, sln_wDHoy, 1) = "1" Then
            slb_toca = 1
        End If
    ElseIf sls_tp = "QUINCENAL" Then
        sln_ara = ((sln_dif + sln_wDIni - 1) Mod 14) + 1
        If Mid(sls_pla, sln_ara, 1) = "1" Then
            slb_toca = 1
        End If
    ElseIf sls_tp = "MENSUALFECHA" Then
'**** Si el mes tiene ese número de días lo devolvemos sin mas. Pero si es el último dia del mes miramos ese y los posteriores hasta el final
        sln_ara = 0
        If sln_dHoy < sln_fm Then 'Si es antes del último dia de mes miramos ese dia
            sln_ara = CInt(Mid(sls_pla, sln_dHoy, 1))
        Else ' Si es el último dia de mes, miramos ese dia y todos hasta el 31
            For sln_k = sln_fm To 31
                sln_ara = sln_ara + CInt(Mid(sls_pla, sln_k, 1))
            Next
        End If
        If sln_ara > 0 Then
            slb_toca = 1
        End If
    ElseIf sls_tp = "MENSUALDIASEMANA" Then
'**** Hacemos una interpretación especial de la última semana
        sln_ara = 0
        sln_k = (Int(sln_dHoy / 7) * 7) + sln_wDHoy ' Posición a buscar en la cadena de la planificación
        sln_ara = Mid(sls_pla, sln_k, 1)
        If sln_k >= 22 And sln_k <= 28 And sln_ara = "0" And ((sln_dHoy + 7) > sln_fm) Then
            sln_ara = Mid(sls_pla, sln_k + 7, 1)
        End If
        If sln_ara = "1" Then
            slb_toca = 1
        End If
    ElseIf Mid(sls_tp, 1, 10) = "REPETICION" Then
        sln_ara = CInt(Mid(sls_tp, 11, 1))
        If sln_ara = 0 Then
            sln_ara = 1
            If SGS_ND > 0 Then
                Loga ("Tratando planificacion repeticion sin dias. No debe ocurrir. Lo ponemos a 1.")
                Loga ("=================================================================================")
            End If
        End If
        sln_pos = (sln_dif) Mod sln_ara
        slb_toca = 1
    End If
    sf_calcPlan = CStr(slb_toca) & CStr(sln_pos)
End Function
'************************************************************************************************
'* Funciones para crear los registros de detalle y el de resumen
'************************************************************************************************
'* Tratamiento Registro medicación
'************************************************************************************************
Function sf_regMed(ByVal sls_fecha, ByVal sls_id)
'Pauta en -> articuloResidente
'Registros en -> registroMedicacion
    Dim sls_tp 'Tipo planificacion
    Dim sls_pla 'Planificacion
    Dim sls_fIni ' Fecha inicio
    Dim sls_cadPautas ' Para guardar una cadena con las pautas tratadas
    Dim slb_toca 'flag de si toca o no hoy
    Dim sln_pos ' En caso de repetición, que repetición debemos coger
    sln_tomasEx = sf_getParam("NUMTOMASEXTRA") 'Nº de tomas extra
    If sln_tomasEx = "" Then sln_tomasEx = 0
    Dim sla_poso(13) 'Array para las posologias
    Dim sla_ara ' Array para las repeticiones
    Dim sls_ara 'Para cadenas temporales
    Dim sla_hora(13)
    For sln_i = 1 To UBound(sla_hora)
        sla_hora(sln_i) = sf_getParam("MEDLIMITETOMA" & sln_i)
    Next
    sls_cadPautas = ""
    'Borramos datos de controlRegMed, conservando solo los del mes actual i el anterior
    sls_fechaAnt = DateAdd("m", -1, sls_fecha)
    Sql = "delete from controlRegMed where momento<'01/" & Month(sls_fechaAnt) & "/" & Year(sls_fechaAnt) & "' "
    Set Rs = sf_rec(Sql)
    'Sql pautas
    Sql = "select  id,tipoPlanifica,planificacion,fechaInicio,articulo,siPrecisa,tipo,"
    Sql = Sql & "isnull(posologia1,0) posologia1,isnull(posologia2,0) posologia2,isnull(posologia3,0) posologia3,isnull(posologia4,0) posologia4,"
    Sql = Sql & "isnull(posologia5,0) posologia5,isnull(posologia6,0) posologia6,isnull(posologia7,0) posologia7,isnull(posologia8,0) posologia8,"
    Sql = Sql & "isnull(posologia9,0) posologia9,isnull(posologia10,0) posologia10,isnull(posologia11,0) posologia11,isnull(posologia12,0) posologia12 "
    Sql = Sql & "from articuloResidente with (nolock) where residente='" & sls_id & "' and regEstado='A' and fechaInicio<=convert(datetime,'" & sls_fecha & "',103) "
    Sql = Sql & "and (fechaFin>=convert(datetime,'" & sls_fecha & " 23:59:59.999',103) or fechaFin is null) "
    If SGS_ND > 9 Then
        Loga ("Hacemos la selección de la tabla articuloResidente")
        Loga ("Ejecutando: " & Sql)
        Loga ("=================================================================================")
    End If
    Set rsAR = sf_rec(Sql)
    Do While Not rsAR.EOF
        sls_tp = rsAR("tipoPlanifica")
        sls_pla = rsAR("planificacion")
        sls_fIni = rsAR("fechaInicio")
        sls_arId = rsAR("id")
        sls_arArId = rsAR("articulo")
        sls_arSP = rsAR("siPrecisa")
        sls_arTipo = rsAR("tipo")
        If sls_arSP Then
            sls_arSP = 1
        Else
            sls_arSP = 0
        End If
        sls_ara = sf_calcPlan(sls_tp, sls_pla, sls_fIni, sls_fecha)
        sln_pos = CInt(Mid(sls_ara, 2))
        sls_ara = Mid(sls_ara, 1, 1)
        If sls_ara = "1" Then
            slb_toca = True
        Else
            slb_toca = False
        End If
'****Si el medicamento toca hoy
        If slb_toca Then
'***** Revisamos los posibles registros
            For sln_k = 1 To CInt(4 + sln_tomasEx)
                sla_poso(sln_k) = rsAR("posologia" & sln_k)
'**** Si hay repetición, extraemos la repetición que toca de la posología
                If Mid(sls_tp, 1, 10) = "REPETICION" Then
                    sls_ara = sla_poso(sln_k)
                    If sls_ara <> "" Then
                        sla_ara = Split(sls_ara, "|")
                        sla_poso(sln_k) = sla_ara(sln_pos)
                    Else
                        sla_poso(sln_k) = 0
                    End If
                End If
'**** Convertimos a numérico la posologia
                If Not IsNumeric(sla_poso(sln_k)) Then sla_poso(sln_k) = 0
'Buscamos los posibles registros existentes
                Sql = "select id from controlRegMed with (nolock) where residente='" & sls_id & "' and momento=convert(datetime,'" & sls_fecha & " " & sla_hora(sln_k) & "',103) "
                Sql = Sql & "and pautaId='" & sls_arId & "' and toma='" & sln_k & "'"
                If SGS_ND > 9 Then
                    Loga ("Buscamos los posibles registros existentes en controlRegMed")
                    Loga ("Ejecutando: " & Sql)
                    Loga ("==============================================================================")
                End If
                Set rsVell = sf_rec(Sql)
'**** Modificación o baja
                If Not rsVell.EOF Then
'**** Modificación
                    If sla_poso(sln_k) > 0 Then
                        Sql = "update controlRegMed set cantidad='" & sla_poso(sln_k) & "',articulo='" & sls_arArId & "',tipo='" & sls_arTipo & "',"
                        Sql = Sql & "siPrecisa='" & sls_arSP & "' where residente='" & sls_id & "' and momento=convert(datetime,'" & sls_fecha & " " & sla_hora(sln_k) & "',103) "
                        Sql = Sql & "and pautaId='" & sls_arId & "' and toma='" & sln_k & "'"
                        If SGS_ND > 9 Then
                            Loga ("Modificamos el que correponde si existe y la posologia es mayor que 0")
                            Loga ("Ejecutando: " & Sql)
                            Loga ("==========================================================================")
                        End If
                        Set rsUpd = sf_rec(Sql)
'**** Baja
                    Else
                        Sql = "delete controlRegMed where residente='" & sls_id & "' and momento=convert(datetime,'" & sls_fecha & " " & sla_hora(sln_k) & "',103) and "
                        Sql = Sql & "pautaId='" & sls_arId & "' and toma='" & sln_k & "'"
                        If SGS_ND > 9 Then
                            Loga ("Borramos el que corresponde si existe y la posologia es 0")
                            Loga ("Ejecutando: " & Sql)
                            Loga ("==========================================================================")
                        End If
                        Set rsUpd = sf_rec(Sql)
                    End If
'**** Alta
                Else
                    If sla_poso(sln_k) > 0 Then
                        Sql = "insert into controlRegMed (id,residente,momento,articulo,toma,cantidad,pautaId,siPrecisa,tipo,realizado) values "
                        Sql = Sql & "(newId(),'" & sls_id & "','" & sls_fecha & " " & sla_hora(sln_k) & "','" & sls_arArId & "','" & sln_k & "','" & sla_poso(sln_k) & "',"
                        Sql = Sql & "'" & sls_arId & "','" & sls_arSP & "','" & sls_arTipo & "','0')"
                        If SGS_ND > 9 Then
                            Loga ("Lo damos de alta en controlRegMed si no existe")
                            Loga ("Ejecutando: " & Sql)
                            Loga ("===========================================================================")
                        End If
                        Set rsUpd = sf_rec(Sql)
                    End If
                End If
            Next
'**** Guardamos cadena con articulos que tocan
            sls_cadPautas = sls_cadPautas & ",'" & sls_arId & "'"
        End If
        rsAR.MoveNext
        DoEvents
    Loop
' Ahora damos de baja los posibles registros de articulos que estan en controlRegMed y ya no estan en articuloResidente
    If sls_cadPautas <> "" Then
        sls_cadPautas = Mid(sls_cadPautas, 2)
        Sql = "delete from controlRegMed where residente='" & sls_id & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) "
        Sql = Sql & "and pautaId not in (" & sls_cadPautas & ")"
        If SGS_ND > 9 Then
            Loga ("Borramos los posibles registros que queden en controlRegMed, que ahora no tienen correspondiente en articuloResidente.")
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsUpd = sf_rec(Sql)
    End If
End Function
'************************************************************************************************
'*  Tratamiento de Otros Registros
'************************************************************************************************
Function sf_regOtrosRegistros(ByVal sls_fecha, ByVal sls_id)
'pauta en -> otrosRegistros
'Registros en -> registroAseos
    Dim sls_tp 'Tipo planificacion
    Dim sls_pla 'Planificacion
    Dim sls_fIni ' Fecha inicio
    Dim sls_cadPautas ' Para guardar una cadena con las pautas tratadas
    Dim slb_toca 'flag de si toca o no hoy
    Dim slb_tocaH ' Flag de si toca o no esa hora
    Dim sln_pos ' En caso de repetición, que repetición debemos coger
    Dim sla_ara ' Array para las repeticiones
    sls_cadPautas = ""
    Sql = "select * from otrosRegistros with (nolock) where residente='" & sls_id & "' and fecini<=convert(datetime,'" & sls_fecha & "',103) and (fecFin>=convert(datetime,'" & sls_fecha & " 23:59:59.999',103) or fecFin is null) and regEstado='A'"
    If SGS_ND > 9 Then
        Loga ("Abrimos cursor contra los registros de pautas de otros registros.")
        Loga ("Ejecutando: " & Sql)
        Loga ("=================================================================================")
    End If
    Set rsOR = sf_rec(Sql)
    Do While Not rsOR.EOF
        sls_tp = rsOR("tipoPlanifica")
        sls_pla = rsOR("planificacion")
        sls_fIni = rsOR("fecIni")
        sls_pauta = rsOR("id")
        sls_tipo = rsOR("tipoRegistro") & "-" & rsOR("registro")
        sls_cadPautas = sls_cadPautas & ",'" & sls_pauta & "'"
        sls_ara = sf_calcPlan(sls_tp, sls_pla, sls_fIni, sls_fecha)
'*** Cogemos solo el primer caracter con si toca o no, ya que no hay repeticiones
        sls_ara = Mid(sls_ara, 1, 1)
        If sls_ara = "1" Then
            slb_toca = True
        Else
            slb_toca = False
        End If
'****Si el registro toca hoy
        If slb_toca Then
            sls_ara = rsOR("cuando")
'**** Inicializamos la cadena de las horas
            sls_cadHoras = ""
'**** Seguimos las 24 horas para ver si toca el registro y lo tratamos
            For sln_k = 1 To 24
                sls_hora = Mid(sls_ara, sln_k, 1)
                If sls_hora = "1" Then
                    slb_tocaH = True
                Else
                    slb_tocaH = False
                End If
                If slb_tocaH Then
                    If sln_k < 11 Then
                        sls_hora = sls_fecha & " 0" & (sln_k - 1) & ":00:00"
                    Else
                        sls_hora = sls_fecha & " " & (sln_k - 1) & ":00:00"
                    End If
                    sls_cadHoras = sls_cadHoras & ",convert(datetime,'" & sls_hora & "',103)"
                    Sql = "select * from controlReg with (nolock) where residente='" & sls_id & "' and momento=convert(datetime,'" & sls_hora & "',103) and tipo='" & sls_tipo & "' and pautaId='" & sls_pauta & "'"
                    If SGS_ND > 9 Then
                        Loga ("Miramos si la pauta en otrosRegistros, ya tiene su registro en controlReg")
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    Set rsRC = sf_rec(Sql)
'*** Si no existe el registro de control, lo creamos
                    If rsRC.EOF Then
                        Sql = "insert into controlReg (id,residente,momento,tipo,pautaId,realizado) values("
                        Sql = Sql & "newId(),'" & sls_id & "','" & sls_hora & "','" & sls_tipo & "','" & sls_pauta & "','0')"
                    If SGS_ND > 9 Then
                        Loga ("Si no existe el registro en controlReg, lo creamos. Si ya existe lo dejamos como esté.")
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                        Set rsUp = sf_rec(Sql)
                    End If
                End If
            Next
'*** Si hay una cadena con horas, borramos los que antes tenían registro, y ahora no tienen
            If sls_cadHoras <> "" Then
                sls_cadHoras = Mid(sls_cadHoras, 2)
                If SGS_ND > 9 Then
                    Loga ("Si hemos cambiado la planificación Borramos los registros de controlReg que no corresponden.")
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                Sql = "delete from controlReg where residente='" & sls_id & "' and tipo='" & sls_tipo & "' and pautaId='" & sls_pauta & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and momento not in (" & sls_cadHoras & ")"
                Set rsUp = sf_rec(Sql)
            End If
        End If
        rsOR.MoveNext
        DoEvents
    Loop
'*** Borramos los registros de control cuya pauta ya no existe para esta fecha
    If sls_cadPautas <> "" Then
        sls_cadPautas = Mid(sls_cadPautas, 2)
        Sql = "delete from controlReg where residente='" & sls_id & "' and tipo='" & sls_tipo & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and pautaId not in(" & sls_cadPautas & ")"
        If SGS_ND > 9 Then
            Loga ("Borramos los registros de controlReg que ya no existen en otrosRegistros.")
            Loga ("Ejecutando: " & Sql)
            Loga ("=============================================================================")
        End If
        Set rsUp = sf_rec(Sql)
    End If
End Function

Function sf_rellenaRecetas(sls_file, sls_usu)
'************************************************************************
'* Funcion para llenar tabla farmaciaRecetas a traves de ficheros excel
'1. Conectamos con empresa actual
'2. Cargamos excel
'3. Si el control esta activado se busca si el residentes que hay en el excel existe en alguna residencia
'y si se tiene la autorizacion de farmacia.
'4. Si se ha encontrado residente se busca si ya se ha registrado esa receta
'5. Si no existe la receta se da de alta y se actualiza el stock para ese residente,articulo y fecha.
'************************************************************************'
    Dim objExcel As Excel.Application
    Dim xLibro As Excel.Workbook
    Dim Col As Integer, Fila As Integer
    Dim sls_receta, sls_fecha, sls_art, sls_cant, sls_codigo, sls_cip, sls_ta, sls_inc, sls_residente, sls_db, slb_control, sls_fechaAnt
    Set objExcel = New Excel.Application
    If sls_file = "" Then Exit Function
    Set xLibro = objExcel.Workbooks.Open(sls_file)
    objExcel.Visible = True
    slb_control = 1
    'Conexion con empresa actual
    Sql = "select * from  gdrempresas with (nolock) where regEstado='A' and codigo='" & EmpresaActual & "' "
    Set rsEmp = sf_recGdr(Sql)
    If Not rsEmp.EOF Then
        sls_empId = rsEmp("id")
        SGC_DBGDR = rsEmp("database")
        SGC_SERVER = rsEmp("dbServer")
        Set SGS_CONUSER = CreateObject("ADODB.Connection")
        SGS_CONUSER.Open "WSID=" & SGS_REMOTEADDR & ";UID=" & SGC_DBUSER & ";PWD=" & SGC_DBPASSWORD & ";Database=" & SGC_DBGDR & ";Server=" & SGC_SERVER & ";Driver={SQL Server};DSN='';"
    End If
    With xLibro
      With .Sheets(1)
            For Fila = 2 To ActiveSheet.UsedRange.Rows.Count
                sls_fecha = Cells(Fila, 1)
                sls_codigo = Cells(Fila, 2)
                sls_art = Cells(Fila, 3)
                sls_cant = Cells(Fila, 4)
                sls_receta = Cells(Fila, 5)
                sls_cip = Cells(Fila, 6)
                sls_ta = Cells(Fila, 7)
                sls_inc = Cells(Fila, 10)
                sls_residente = ""
                'Control para comprobar si existe la autorizacion de farmacia para el residente en curso
                If slb_control = 1 Then
                    Sql = "select id,[database] db from gdrempresas where regEstado='A' "
                    Set rsEmp = sf_recGdr(Sql)
                    Do While Not rsEmp.EOF
                       InformaMiss "Tratando " & rsEmp("db") & " - R: " & sls_cip & "a las " & Time
                       'Solo los residentes con autorización de la farmacia
                       Sql = "select r.cip,d.id,r.id idRes,r.centro,r.codigo from residentes r "
                       Sql = Sql & "left join documentos d with (nolock) on (d.idRegistro COLLATE Modern_Spanish_CI_AI=r.id and d.regEstado='A' "
                       Sql = Sql & "and d.tipoRegistro='farmacia' and d.tipoDocumento='controlAutorizacion' ) where r.regEstado='A' and r.cip like '%" & sls_cip & "%' "
                       'sql = "select id,centro,codigo from residentes where regEstado='A' and cip like '%" & sls_cip & "%' "
                       Set rsRes = sf_conexionSQL(rsEmp("id"), rsEmp("db"), Sql)
                       If Not rsRes.EOF Then
                           If rsRes("idRes") <> "" Then
                                sls_empId = rsEmp("id")
                                sls_db = rsEmp("db")
                                sls_residente = rsRes("idRes")
                                sls_centro = rsRes("centro")
                                sls_codResi = rsRes("codigo")
                                'Comprobamos que no exista un registro para esa receta
                                Sql = "select receta from farmaciaRecetas where regEstado='A' and receta like '%" & sls_receta & "%'"
                                Set rsRec = sf_rec(Sql)
                                If rsRec.EOF Then
                                    InformaMiss "Insertando receta " & sls_receta & " - R: " & sls_cip & "a las " & Time
                                    Sql = " select newId() idDoc, newId() id,newId() id2,getDate() ts "
                                    Set rsId = sf_rec(Sql)
                                    sls_idDoc = rsId("idDoc")
                                    sls_id = rsId("id")
                                    sls_id2 = rsId("id2")
                                    sls_timeStamp = rsId("ts")
                                    'Introducimos registro en tabla farmaciaRecetas en empresa actual
                                    Sql = "INSERT into farmaciaRecetas (id,receta,fecha,codigo,articulo,cantidad,residente,cip,ta,incidencias,db,regEstado,usrAlta,usrMod,fecAlta,fecMod) values ( "
                                    Sql = Sql & "newId(),'" & sls_receta & "','" & sls_fecha & "','" & sls_codigo & "','" & sls_art & "','" & sls_cant & "','" & sls_residente & "','" & sls_cip & "', "
                                    Sql = Sql & "'" & sls_ta & "','" & sls_idDoc & "','" & sls_db & "','A','" & sls_usu & "','" & sls_usu & "','" & sls_timeStamp & "','" & sls_timeStamp & "') "
                                    Set rsIns = sf_rec(Sql)
                                    'Aqui hacemos el insert de los campos (incidencia) en la tabla de infoEditor
                                    Sql = "INSERT into infoEditor(id,tabla,campo,valor,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                                    Sql = Sql & "values('" & sls_idDoc & "','datosReceta','incidencia','" & sls_inc & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
                                    Set rsIns = sf_rec(Sql)
                                    'Insert de la receta en stock
                                    sls_saco = sf_actualizaStock(sls_empId, sls_db, sls_codigo, sls_residente, sls_fecha, sls_cant, sls_usu)
                                End If
                           End If
                       End If
                       rsEmp.MoveNext
                    Loop
                Else
                    InformaMiss "Tratando R: " & sls_cip & "a las " & Time
                    'Comprobamos que no exista un registro para esa receta
                    Sql = "select receta from farmaciaRecetas where regEstado='A' and receta like '%" & sls_receta & "%'"
                    Set rsRec = sf_rec(Sql)
                    If rsRec.EOF Then
                        InformaMiss "Insertando receta " & sls_receta & " - R: " & sls_cip & "a las " & Time
                        Sql = " select newId() idDoc, newId() id"
                        Set rsId = sf_rec(Sql)
                        sls_idDoc = rsId("idDoc")
                        sls_id = rsId("id")
                        Sql = "INSERT into farmaciaRecetas (id,receta,fecha,codigo,articulo,cantidad,residente,cip,ta,incidencias,db,regEstado,usrAlta,usrMod,fecAlta,fecMod) values ( "
                        Sql = Sql & "newId(),'" & sls_receta & "','" & sls_fecha & "','" & sls_codigo & "','" & sls_art & "','" & sls_cant & "','" & sls_residente & "','" & sls_cip & "', "
                        Sql = Sql & "'" & sls_ta & "','" & sls_idDoc & "','" & sls_db & "','A','" & sls_usu & "','" & sls_usu & "',getDate(),getDate()) "
                        Set rsIns = sf_rec(Sql)
                        'Aqui hacemos el insert de los campos (incidencia) en la tabla de infoEditor
                        Sql = "INSERT into infoEditor(id,tabla,campo,valor,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                        Sql = Sql & "values('" & sls_idDoc & "','datosReceta','incidencia','" & sls_inc & "','A','GdR','GdR',getDate(),getDate())"
                        Set rsIns = sf_rec(Sql)
                        'Insert de la receta en stock
                        sls_saco = sf_actualizaStock(sls_empId, SGC_DBGDR, sls_codigo, sls_residente, sls_fecha, sls_cant, sls_usu)
                    End If
                End If
            Next
      End With
    End With
    xLibro.Close SaveChanges:=False
    objExcel.Quit
    Set objExcel = Nothing
    Set xLibro = Nothing
End Function


Function sf_asistenteDatos1(sls_file, sls_usu)
'************************************************************************
'* Funcion para llenar tabla residentes,datoasi,contratos a traves de ficheros excel
'1. Conectamos con empresa actual
'2. Cargamos excel
'3. Si el control esta activado se busca si el residentes que hay en el excel existe en alguna residencia
'y si se tiene la autorizacion de farmacia.
'4. Si se ha encontrado residente se busca si ya se ha registrado esa receta
'5. Si no existe la receta se da de alta y se actualiza el stock para ese residente,articulo y fecha.
'************************************************************************'
    Dim objExcel As Excel.Application
    Dim xLibro As Excel.Workbook
    Dim Col As Integer, Fila As Integer
    Dim sls_codigo, sls_nombre, sls_apellido1, sls_apellido2, sls_fechaNac, sls_lugarNac, sls_provinciaNac
    Dim sls_paisNax, sls_estadoCivil, sls_tipoDoc, sls_documento, sls_numSS, sls_numTS, sls_contrato, sls_libroReg, sls_numLibroReg
    Dim sls_fechaIni, sls_fechaFin, sls_residente, sls_db, slb_control, sls_mes, sls_anyo, sls_dia, sls_fechaNacF, sls_fechaNacF2
    Dim sls_agenda, sls_detalle, sls_agenda2, sls_detalle2, sls_ocupacion, sls_idDoc, sls_centro, sls_estado
    Set objExcel = New Excel.Application
    If sls_file = "" Then Exit Function
    Set xLibro = objExcel.Workbooks.Open(sls_file)
    objExcel.Visible = True
    slb_control = 1
    'Conexion con empresa actual
    Sql = "select id,[database],dbServer from  gdrempresas with (nolock) where regEstado='A' and codigo='" & EmpresaActual & "' "
    Set rsEmp = sf_recGdr(Sql)
    If Not rsEmp.EOF Then
        sls_empId = rsEmp("id")
        SGC_DBGDR = rsEmp("database")
        SGC_SERVER = rsEmp("dbServer")
        Set SGS_CONUSER = CreateObject("ADODB.Connection")
        SGS_CONUSER.Open "WSID=" & SGS_REMOTEADDR & ";UID=" & SGC_DBUSER & ";PWD=" & SGC_DBPASSWORD & ";Database=" & SGC_DBGDR & ";Server=" & SGC_SERVER & ";Driver={SQL Server};DSN='';"
    End If
    With xLibro
      With .Sheets(1)
        For Fila = 3 To ActiveSheet.UsedRange.Rows.Count
            'Residente
            sls_codigo = Cells(Fila, 1)
            sls_nombre = Cells(Fila, 2)
            sls_apellido1 = Cells(Fila, 3)
            sls_apellido2 = Cells(Fila, 4)
            sls_sexo = Left(Cells(Fila, 5), 1)
            sls_fechaNac = Cells(Fila, 6)
            sls_lugarNac = Cells(Fila, 7)
            sls_provinciaNac = Cells(Fila, 8)
            sls_paisNac = Cells(Fila, 9)
            sls_estadoCivil = Cells(Fila, 10)
            sls_tipoDoc = Cells(Fila, 11)
            sls_documento = Cells(Fila, 12)
            sls_numSS = Cells(Fila, 13)
            sls_numSS = Replace(sls_numSS, " ", "")
            sls_numTS = Cells(Fila, 14)
            sls_numTS = Replace(sls_numTS, " ", "")
            'Contrato
            sls_libroReg = Cells(Fila, 15)
            sls_numLibroReg = Cells(Fila, 16)
            sls_fechaIni = Cells(Fila, 17)
            sls_fechaFin = Cells(Fila, 18)
            Sql = "select top 1 id from centros where regestado='A' order by fecAlta"
            Set rsCen = sf_rec(Sql)
            If Not rsCen.EOF Then sls_centro = rsCen("id")
            InformaMiss "Tratando " & rsEmp("database") & " - R: " & sls_codigo & " - " & Time
            Sql = "select newId() idResi,newId() idContrato,newId() idOcupacion,newId() idDoc,getDate() timeStamp"
            Set rsTime = sf_rec(Sql)
            sls_residente = rsTime("idResi")
            sls_contrato = rsTime("idContrato")
            sls_contrato = rsTime("idContrato")
            sls_ocupacion = rsTime("idOcupacion")
            sls_idDoc = rsTime("idDoc")
            sls_timeStamp = rsTime("timeStamp")
            sls_estado = "A"
            Select Case sls_estadoCivil
                Case "CASADO"
                    sls_estadoCivil = "CA"
                Case "DIVORCIADO"
                    sls_estadoCivil = "DI"
                Case "SEPARADO"
                    sls_estadoCivil = "SE"
                Case "SOLTERO"
                    sls_estadoCivil = "SO"
                Case "VIUDO"
                    sls_estadoCivil = "VI"
            End Select
            Select Case sls_libroReg
                Case "CENTRO DE DIA"
                    sls_libroReg = "CDIA"
                Case "HOGAR RESIDENCIAL"
                    sls_libroReg = "HRES"
                Case "RESIDENCIA ASISTIDA"
                    sls_libroReg = "RASI"
                Case "COMEDOR SOCIAL"
                    sls_libroReg = "CSOC"
            End Select
            'Alta residente
            Sql = "INSERT into residentes(id,nombre,apellido1,apellido2,fechaNac,lugarNac,paisNac,provinciaNac,sexo,estadoCivil,"
            Sql = Sql & "tipoDocu,documento,centro,cip,ss,estado,codigo,regEstado,usrAlta,usrMod,fecAlta,fecMod) values("
            Sql = Sql & "'" & sls_residente & "','" & sls_nombre & "','" & sls_apellido1 & "','" & sls_apellido2 & "'," & sf_iif(sls_fechaNac = "", "null", "'" & sls_fechaNac & "'") & ","
            Sql = Sql & "'" & sls_lugarNac & "','" & sls_paisNac & "','" & sls_provinciaNac & "','" & sls_sexo & "','" & sls_estadoCivil & "','" & sls_tipoDocu & "','" & sls_documento & "','" & sls_centro & "',"
            Sql = Sql & "'" & sls_numTS & "','" & sls_numSS & "','" & sls_estado & "','" & sls_codigo & "',"
            Sql = Sql & "'A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
            Set rsIns = sf_rec(Sql)
            'Agenda
            sls_anyo = Year(Date)
            sls_mes = Month(sls_fechaNac)
            sls_dia = Day(sls_fechaNac)
            sls_fechaNacF = sls_dia & "/" & sls_mes & "/" & sls_anyo
            sls_fechaNacF = FormatDateTime(sls_fechaNacF, 2)
            sls_fechaNacF2 = sls_dia & "/" & sls_mes & "/" & sls_anyo + 1
            sls_fechaNacF2 = FormatDateTime(sls_fechaNacF2, 2)
            'Registros año actual y siguiente
            Sql = "select newId() idAgenda,newId() idDetalle,newId() idAgenda2,newId() idDetalle2 "
            Set rsId = sf_rec(Sql)
            sls_idAgenda = rsId("idAgenda")
            sls_idDetalle = rsId("idDetalle")
            sls_idAgenda2 = rsId("idAgenda2")
            sls_idDetalle2 = rsId("idDetalle2")
            Sql = "INSERT into agenda(id,centro,momentoIni,momentoFin,tipoAnotacion,idOrigen,resumen,detalle,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
            Sql = Sql & "values('" & sls_idAgenda & "','" & sls_centro & "','" & sls_fechaNacF & "','" & sls_fechaNacF & "',"
            Sql = Sql & "'aniversario','" & sls_id & "','Aniversario','" & sls_idDetalle & "','A',"
            Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
            Set rsIns = sf_rec(Sql)
            Sql = "INSERT into infoEditor(id,tabla,campo,valor,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
            Sql = Sql & "values('" & sls_idDetalle & "','datosAdmin','agenda','Aniversario:" & sls_fechaNacF & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
            Set rsIns = sf_rec(Sql)
            Sql = "INSERT into agendaResidente(id,idAnotacion,residente,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
            Sql = Sql & "values(newId(),'" & sls_idAgenda & "','" & sls_id & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
            Set rsIns = sf_rec(Sql)
            Sql = "INSERT into agenda(id,centro,momentoIni,momentoFin,tipoAnotacion,idOrigen,resumen,detalle,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
            Sql = Sql & "values('" & sls_idAgenda2 & "','" & sls_centro & "','" & sls_fechaNacF2 & "','" & sls_fechaNacF2 & "',"
            Sql = Sql & "'aniversario','" & sls_id & "','Aniversario','" & sls_idDetalle2 & "','A',"
            Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
            Set rsIns = sf_rec(Sql)
            Sql = "INSERT into infoEditor(id,tabla,campo,valor,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
            Sql = Sql & "values('" & sls_idDetalle2 & "','datosAdmin','agenda','Aniversario:" & sls_fechaNacF2 & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
            Set rsIns = sf_rec(Sql)
            Sql = "INSERT into agendaResidente(id,idAnotacion,residente,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
            Sql = Sql & "values(newId(),'" & sls_idAgenda2 & "','" & sls_id & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
            Set rsIns = sf_rec(Sql)
            'Contrato
            Sql = "INSERT into contratosReservas(id,centro,residente,estado,estadoContrato,libroReg,numLibroReg,pactosAd,"
            Sql = Sql & "fechaInicio,fechaFin,fechaFinReal,diasEstancia,regEstado,usrAlta,usrMod,fecAlta,fecMod) values("
            Sql = Sql & "'" & sls_contrato & "','" & sls_centro & "','" & sls_residente & "','A','CONTRATO','" & sls_libroReg & "','" & sls_numLibroReg & "',"
            Sql = Sql & "'" & sls_idDoc & "'," & sf_iif(sls_fechaIni = "", "null", "'" & sls_fechaIni & "'") & "," & sf_iif(sls_fechaFin = "", "null", "'" & sls_fechaFin & "'") & "," & sf_iif(sls_fechaFin = "", "null", "'" & sls_fechaFin & "'") & ","
            Sql = Sql & "'1111111','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
            Set rsIns = sf_rec(Sql)
            'Aqui hacemos el insert de los campos (pactosAd) en la tabla de infoEditor
            Sql = "INSERT into infoEditor(id,tabla,campo,valor,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
            Sql = Sql & "values('" & sls_idDoc & "','datosContratos','pactosAd','" & sls_pactosAd & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
            Set rsIns = sf_rec(Sql)
            'Aqui hacemos el insert de los campos en la tabla de ocupacion
            Sql = "INSERT into ocupacion(id,centro,residente,contrato,estadoContrato,estado,fechaInicio,fechaFin,fechaFinReal,cama,tipoCama,regEstado,usrAlta,usrMod,fecAlta,fecMod) "
            Sql = Sql & "values ('" & sls_ocupacion & "','" & sls_centro & "','" & sls_residente & "','" & sls_contrato & "','CONTRATO','A'," & sf_iif(sls_fechaIni = "", "null", "'" & sls_fechaIni & "'") & ","
            Sql = Sql & "" & sf_iif(sls_fechaFin = "", "null", "'" & sls_fechaFin & "'") & "," & sf_iif(sls_fechaFin = "", "null", "'" & sls_fechaFin & "'") & ","
            Sql = Sql & "NULL,NULL,'A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
            Set rsIns = sf_rec(Sql)
        Next
      End With
    End With
    xLibro.Close SaveChanges:=False
    objExcel.Quit
    Set objExcel = Nothing
    Set xLibro = Nothing
End Function

'************************************************************************
'* Función sf_conexionSQL
'* Abre conexion con empresa escogida
'* Ejecución de un Recordset en la BBDD  de la empresa escogida
'* Devuelve recordSet
'************************************************************************
Function sf_conexionSQL(ByVal sls_empresa, ByVal sls_db, ByVal sls_sql)
    Dim slo_rs
    If sls_empresa <> EmpresaActual Then
        If sls_db = "" Then
            Sql = "select [database] from gdrEmpresas with (nolock) where id='" & sls_empresa & "' and regEstado='A' "
            Set rsDB = sf_recGdr(Sql)
            If Not rsDB.EOF Then
                sls_db = rsDB("database")
            End If
        End If
        Set SGS_CONSQL = CreateObject("ADODB.Connection")
        SGS_CONSQL.Open "WSID=" & sls_empresa & ";UID=" & SGC_DBUSER & ";PWD=" & SGC_DBPASSWORD & ";Database=" & sls_db & ";Server=" & SGC_SERVER & ";Driver={SQL Server};DSN='';"
        
        ultimaSql = sls_sql
    
        Set slo_rs = CreateObject("ADODB.Recordset")
        Set slo_comm = CreateObject("ADODB.Command")
        
        slo_comm.ActiveConnection = SGS_CONSQL
        slo_comm.CommandTimeout = 1500
        slo_comm.CommandText = sls_sql
        Set slo_rs = slo_comm.Execute
    Else
        ultimaSql = sls_sql
        
        Set slo_rs = CreateObject("ADODB.Recordset")
        Set slo_comm = CreateObject("ADODB.Command")
        
        slo_comm.ActiveConnection = SGS_CONUSER
        slo_comm.CommandTimeout = 1500
        slo_comm.CommandText = sls_sql
        Set slo_rs = slo_comm.Execute
    End If
    Set sf_conexionSQL = slo_rs
End Function
'************************************************************************
'* Funcion para sustituir caracteres en base a una expresion regular, para tratar mayusculas y minusculas
'************************************************************************
Function sf_replaceER(ByVal sls_entrada, ByVal sls_in, ByVal sls_out)
    Dim sls_regEx
    Set sls_regEx = New RegExp
    sls_regEx.Pattern = sls_in
    sls_regEx.IgnoreCase = True ' Para que cambie mayusculas y minusculas
    sls_regEx.Global = True
    sf_replaceER = sls_regEx.Replace(sls_entrada, sls_out)
End Function
'************************************************************************************************
'*  Tratamiento Registros Contenciones
'************************************************************************************************
Function sf_regContenciones(ByVal sls_fecha, ByVal sls_id)
'Pauta en -> pautasContenciones
'Registros en ->registroContenciones
    Dim sls_ara, sls_hora, sls_fec
    Dim sls_tipo ' Tipo de registro en contenciones
    Dim sls_pauta 'Para el id de la pauta
    Dim sls_cadHoras ' Para crear una cadena con las horas con curas en cada pauta
    Dim sls_cadPautas ' Para crear una cadena con todas las pautas del residente en una fecha
    sls_tipo = "CONTENCIONES"
    sls_cadPautas = ""
'*** Seguimos las contenciones activas del residente
    Sql = "select * from pautasContenciones with (nolock) where residente='" & sls_id & "' and regEstado='A' and fecIni<=convert(datetime,'" & sls_fecha & "',103) and (fecFin>=convert(datetime,'" & sls_fecha & "',103) or fecFin is null)"
    If SGS_ND > 9 Then
        Loga ("Segumos las pautas de contenciones activas del residente.")
        Loga ("Ejecutando: " & Sql)
        Loga ("=============================================================================")
    End If
    Set rsCon = sf_rec(Sql)
    Do While Not rsCon.EOF
'**** Vamos creando la cadena de las pautas
        sls_pauta = rsCon("id")
        sls_cadPautas = sls_cadPautas & ",'" & sls_pauta & "'"
        sls_cuando = rsCon("cuando")
        slb_ini = 0
        slb_fin = 0
'**** Inicializamos la cadena de las horas
        sls_cadHoras = ""
        If IsNull(sls_cuando) Or sls_cuando = "" Then
            sls_lenCuando = 0
        Else
            sls_lenCuando = Len(sls_cuando)
        End If
        For sln_i = 1 To sls_lenCuando
            sls_ahora = Mid(sls_cuando, sln_i, 1)
            If sls_ahora = 1 Then
                If slb_ini = 0 Then
                    slb_ini = 1
                    slb_fin = 0
                    sls_hora = sln_i - 1
                    If sls_hora = 24 Then sls_hora = 0
                    If Len(sls_hora) = 1 Then sls_hora = "0" & sls_hora
                    sls_hora = sls_fecha & " " & sls_hora & ":00"
                    sls_subtipo = "-INI"
                    sls_cadHoras = sls_cadHoras & ",convert(datetime,'" & sls_hora & "',103)"
                    Sql = "select * from controlReg with (nolock) where residente='" & sls_id & "' and momento=convert(datetime,'" & sls_hora & "',103) and tipo='" & sls_tipo & sls_subtipo & "' and pautaId='" & sls_pauta & "'"
                    If SGS_ND > 9 Then
                        Loga ("Miramos si la pauta en pautasContenciones, ya tiene su registro en controlReg para el INI.")
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    Set rsRC = sf_rec(Sql)
'*** Si no existe el registro de control, lo creamos
                    If rsRC.EOF Then
                        Sql = "insert into controlReg (id,residente,momento,tipo,pautaId,realizado) values("
                        Sql = Sql & "newId(),'" & sls_id & "','" & sls_hora & "','" & sls_tipo & sls_subtipo & "','" & sls_pauta & "','0')"
                        If SGS_ND > 9 Then
                            Loga ("Si no existe el registro INI en controlReg, lo creamos.")
                            Loga ("Ejecutando: " & Sql)
                            Loga ("=============================================================================")
                        End If
                        Set rsUp = sf_rec(Sql)
                    End If
                Else
                    sls_hora = sln_i
                    If sln_i = 24 Then sls_hora = 0
                    If Len(sls_hora) = 1 Then sls_hora = "0" & sls_hora
                    sls_hora = sls_fecha & " " & sls_hora & ":00"
                    sls_cadHoras = sls_cadHoras & ",convert(datetime,'" & sls_hora & "',103)"
                End If
            End If
'*** Miramos siguiente hora
            If sln_i < Len(sls_cuando) Then
                If Mid(sls_cuando, sln_i + 1, 1) = 0 And slb_ini = 1 Then
'**** Tratamos el fin de la contención
                    slb_fin = 1
                    slb_ini = 0
                    sls_hora = sln_i - 1
                    If sls_hora = 24 Then sls_hora = 0
                    If Len(sls_hora) = 1 Then sls_hora = "0" & sls_hora
                    sls_hora = sls_fecha & " " & sls_hora & ":00"
                    sls_subtipo = "-FIN"
                    sls_cadHoras = sls_cadHoras & ",convert(datetime,'" & sls_hora & "',103)"
                    Sql = "select * from controlReg with (nolock) where residente='" & sls_id & "' and momento=convert(datetime,'" & sls_hora & "',103) and tipo='" & sls_tipo & sls_subtipo & "' and pautaId='" & sls_pauta & "'"
                    If SGS_ND > 9 Then
                        Loga ("Miramos si la pauta en pautasContenciones, ya tiene su registro en controlReg para el FIN")
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    Set rsRC = sf_rec(Sql)
'*** Si no existe el registro de control, lo creamos
                    If rsRC.EOF Then
                        Sql = "insert into controlReg (id,residente,momento,tipo,pautaId,realizado) values("
                        Sql = Sql & "newId(),'" & sls_id & "','" & sls_hora & "','" & sls_tipo & sls_subtipo & "','" & sls_pauta & "','0')"
                        If SGS_ND > 9 Then
                            Loga ("Si no existe el registro FIN en controlReg, lo creamos.")
                            Loga ("Ejecutando: " & Sql)
                            Loga ("=============================================================================")
                        End If
                        Set rsUp = sf_rec(Sql)
                    End If
                End If
            End If
        Next
'*** Si hay una cadena con horas, borramos los que antes tenían registro, y ahora no tienen
        If sls_cadHoras <> "" Then
            sls_cadHoras = Mid(sls_cadHoras, 2)
            Sql = "delete from controlReg where residente='" & sls_id & "' and tipo like'" & sls_tipo & "-%' and pautaId='" & sls_pauta & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and momento not in (" & sls_cadHoras & ")"
            If SGS_ND > 9 Then
                Loga ("Si hay una cadena con horas, borramos los que antes tenían registro en controlReg, y ahora no tienen.")
                Loga ("Ejecutando: " & Sql)
                Loga ("=============================================================================")
            End If
            Set rsUp = sf_rec(Sql)
        End If
    rsCon.MoveNext
    DoEvents
    Loop
'*** Borramos los registros de control cuya pauta ya no existe para esta fecha
    If sls_cadPautas <> "" Then
        sls_cadPautas = Mid(sls_cadPautas, 2)
        Sql = "delete from controlReg where residente='" & sls_id & "' and tipo like '" & sls_tipo & "-%' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and pautaId not in(" & sls_cadPautas & ")"
            If SGS_ND > 9 Then
                Loga ("Borramos los registros de control cuya pauta ya no existe para esta fecha.")
                Loga ("Ejecutando: " & Sql)
                Loga ("=============================================================================")
            End If
        Set rsUp = sf_rec(Sql)
    End If
End Function
'************************************************************************************************
'*  Tratamiento Registros CambiosPosturales
'************************************************************************************************
Function sf_regCPosturales(ByVal sls_fecha, ByVal sls_id)
'Pautas en -> pautasCambioPostural
'Registros en -> registroCambioPostural
    Dim sls_ara, sls_hora, sls_fec
    Dim sls_tipo ' Tipo de registro en cambiosposturales
    Dim sls_pauta 'Para el id de la pauta
    Dim sls_cadHoras ' Para crear una cadena con las horas en cada pauta
    Dim sls_cadPautas ' Para crear una cadena con todas las pautas del residente en una fecha
    sls_tipo = "CPOSTURALES"
    sls_cadPautas = ""
'*** Seguimos los cambios posturales activas del residente
    Sql = "select * from pautasCambioPostural with (nolock) where residente='" & sls_id & "' and regEstado='A' and fecIni<=convert(datetime,'" & sls_fecha & "',103) and (fecFin>=convert(datetime,'" & sls_fecha & "',103) or fecFin is null)"
    If SGS_ND > 9 Then
        Loga ("Segumos las pautas de cambioPostural activas del residente.")
        Loga ("Ejecutando: " & Sql)
        Loga ("=============================================================================")
    End If
    Set rsCP = sf_rec(Sql)
    Do While Not rsCP.EOF
'**** Vamos creando la cadena de las pautas
        sls_pauta = rsCP("id")
        sls_cadPautas = sls_cadPautas & ",'" & sls_pauta & "'"
        sls_ara = rsCP("cuando")
'**** Inicializamos la cadena de las horas
        sls_cadHoras = ""
'**** Seguimos las 24 horas para ver si toca Cambio Postural lo tratamos
        For sln_k = 1 To 24
            sls_hora = Mid(sls_ara, sln_k, 1)
            If sls_hora = "1" Then
                slb_toca = True
            Else
                slb_toca = False
            End If
            If slb_toca Then
                If sln_k < 11 Then
                    sls_hora = sls_fecha & " 0" & (sln_k - 1) & ":00:00"
                Else
                    sls_hora = sls_fecha & " " & (sln_k - 1) & ":00:00"
                End If
                sls_cadHoras = sls_cadHoras & ",convert(datetime,'" & sls_hora & "',103)"
                Sql = "select * from controlReg with (nolock) where residente='" & sls_id & "' and momento=convert(datetime,'" & sls_hora & "',103) and tipo='" & sls_tipo & "' and pautaId='" & sls_pauta & "'"
                If SGS_ND > 9 Then
                    Loga ("Miramos si la pauta de cambioPostural esta ya en controlReg.")
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                Set rsRC = sf_rec(Sql)
'*** Si no existe el registro de control, lo creamos
                If rsRC.EOF Then
                    Sql = "insert into controlReg (id,residente,momento,tipo,pautaId,realizado) values("
                    Sql = Sql & "newId(),'" & sls_id & "','" & sls_hora & "','" & sls_tipo & "','" & sls_pauta & "','0')"
                    If SGS_ND > 9 Then
                        Loga ("Si la pauta de cambioPostural no esta en controlReg, la damos de alta.")
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    Set rsUp = sf_rec(Sql)
                End If
            End If
        Next
'*** Si hay una cadena con horas, borramos los que antes tenían registro, y ahora no tienen
        If sls_cadHoras <> "" Then
            sls_cadHoras = Mid(sls_cadHoras, 2)
            Sql = "delete from controlReg where residente='" & sls_id & "' and tipo='" & sls_tipo & "' and pautaId='" & sls_pauta & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and momento not in (" & sls_cadHoras & ")"
            If SGS_ND > 9 Then
                Loga ("Borramos registros en controlReg, para esa pauta, que eran en horarios distintos a los actuales.")
                Loga ("Ejecutando: " & Sql)
                Loga ("=============================================================================")
            End If
            Set rsUp = sf_rec(Sql)
        End If
    rsCP.MoveNext
    DoEvents
    Loop
'*** Borramos los registros de control cuya pauta ya no existe para esta fecha
    If sls_cadPautas <> "" Then
        sls_cadPautas = Mid(sls_cadPautas, 2)
        Sql = "delete from controlReg where residente='" & sls_id & "' and tipo='" & sls_tipo & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and pautaId not in(" & sls_cadPautas & ")"
        If SGS_ND > 9 Then
            Loga ("Borramos las pautas de controlReg, que ya no están en cambioPostural.")
            Loga ("Ejecutando: " & Sql)
            Loga ("=============================================================================")
        End If
        Set rsUp = sf_rec(Sql)
    End If
End Function
'************************************************************************************************
'*  Tratamiento Registros Curas
'************************************************************************************************
Function sf_regCuras(ByVal sls_fecha, ByVal sls_id)
'pautas en -> curas
'Registros en -> registroCuras
    Dim sls_ara, sls_hora, sls_fec
    Dim sls_tipo ' Tipo de registro en curas
    Dim sls_pauta 'Para el id de la pauta
    Dim sls_cadHoras ' Para crear una cadena con las horas con curas en cada pauta
    Dim sls_cadPautas ' Para crear una cadena con todas las pautas del residente en una fecha
    sls_tipo = "CURAS"
    sls_cadPautas = ""
'*** Seguimos las curas activas del residente
    Sql = "select * from curas with (nolock) where residente='" & sls_id & "' and regEstado='A' and ini<=convert(datetime,'" & sls_fecha & "',103) and (fin>=convert(datetime,'" & sls_fecha & "',103) or fin is null)"
    If SGS_ND > 9 Then
        Loga ("Seguimos las curas activas del residente.")
        Loga ("Ejecutando: " & Sql)
        Loga ("=============================================================================")
    End If
    Set rsCu = sf_rec(Sql)
    Do While Not rsCu.EOF
'**** Vamos creando la cadena de las pautas
        sls_pauta = rsCu("id")
        sls_cadPautas = sls_cadPautas & ",'" & sls_pauta & "'"
        sls_tp = rsCu("tipoPlanifica")
        sls_pla = rsCu("planificacion")
        sls_fIni = rsCu("ini")
        sls_tocaPlan = sf_calcPlan(sls_tp, sls_pla, sls_fIni, sls_fecha)
        sls_tocaPlan = Mid(sls_tocaPlan, 1, 1)
        If sls_tocaPlan = "1" Then
            slb_toca = True
        Else
            slb_toca = False
        End If
        If slb_toca Then
            sls_ara = rsCu("cuando")
            'Inicializamos la cadena de las horas
            sls_cadHoras = ""
            'Seguimos las 24 horas para ver si toca cura y la tratamos
            For sln_k = 1 To 24
                sls_hora = Mid(sls_ara, sln_k, 1)
                If sls_hora = "1" Then
                    slb_toca = True
                Else
                    slb_toca = False
                End If
                If slb_toca Then
                    If sln_k < 11 Then
                        sls_hora = sls_fecha & " 0" & (sln_k - 1) & ":00:00"
                    Else
                        sls_hora = sls_fecha & " " & (sln_k - 1) & ":00:00"
                    End If
                    sls_cadHoras = sls_cadHoras & ",convert(datetime,'" & sls_hora & "',103)"
                    Sql = "select * from controlReg with (nolock) where residente='" & sls_id & "' and momento=convert(datetime,'" & sls_hora & "',103) and tipo='" & sls_tipo & "' and pautaId='" & sls_pauta & "'"
                    If SGS_ND > 9 Then
                        Loga ("Miramos si la pauta de curas esta ya en controlReg.")
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    Set rsRC = sf_rec(Sql)
                    'Si no existe el registro de control, lo creamos
                    If rsRC.EOF Then
                        Sql = "insert into controlReg (id,residente,momento,tipo,pautaId,realizado) values("
                        Sql = Sql & "newId(),'" & sls_id & "','" & sls_hora & "','" & sls_tipo & "','" & sls_pauta & "','0')"
                        If SGS_ND > 9 Then
                            Loga ("Si la pauta de curas no esta en controlReg, la creamos.")
                            Loga ("Ejecutando: " & Sql)
                            Loga ("=============================================================================")
                        End If
                        Set rsUp = sf_rec(Sql)
                    End If
                End If
            Next
    '*** Si hay una cadena con horas, borramos los que antes tenían registro, y ahora no tienen
            If sls_cadHoras <> "" Then
                sls_cadHoras = Mid(sls_cadHoras, 2)
                Sql = "delete from controlReg where residente='" & sls_id & "' and tipo='" & sls_tipo & "' and pautaId='" & sls_pauta & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and momento not in (" & sls_cadHoras & ")"
                If SGS_ND > 9 Then
                    Loga ("Borramos de controlReg, las pautas que hayan cambiado de hora de esa pauta.")
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                Set rsUp = sf_rec(Sql)
            End If
        End If
        rsCu.MoveNext
        DoEvents
    Loop
'*** Borramos los registros de control cuya pauta ya no existe para esta fecha
    If sls_cadPautas <> "" Then
        sls_cadPautas = Mid(sls_cadPautas, 2)
        Sql = "delete from controlReg where residente='" & sls_id & "' and tipo='" & sls_tipo & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and pautaId not in(" & sls_cadPautas & ")"
        If SGS_ND > 9 Then
            Loga ("Borramos de controlReg, las pautas que se han borrado de curas.")
            Loga ("Ejecutando: " & Sql)
            Loga ("=============================================================================")
        End If
        Set rsUp = sf_rec(Sql)
    End If
End Function


Sub Loga(Str As String, Optional EsError As Boolean = False)
On Error Resume Next
If Not Left(Str, 5) = "=====" Then Informa Str
    DoEvents
    
    'If slo_fname Is Nothing Then
        'Set slo_fs = CreateObject("Scripting.FileSystemObject")
        'Set slo_fname = slo_fs.CreateTextFile("C:\data\gdr\Batch\log\" & Year(Now) & Month(Now) & Day(Now) & "-" & Hour(Now) & Minute(Now) & ".log", True)
    'End If
    'If Not slo_fname Is Nothing Then
    '    slo_fname.WriteLine Str
    'End If
    
End Sub


'************************************************************************************************
'*  Tratamiento Registros Pañales
'************************************************************************************************
Function sf_regPanales(ByVal sls_fecha, ByVal sls_id)
'Pauta en ->pautasPanales
'Registros en ->registroPanales
    Dim sls_ara, sls_hora, sls_fec
    Dim sls_tipo ' Tipo de registro en pañales
    Dim sls_pauta 'Para el id de la pauta
    Dim sls_cadHoras ' Para crear una cadena con las horas con curas en cada pauta
    Dim sls_cadPautas ' Para crear una cadena con todas las pautas del residente en una fecha
    sls_tipo = "PANAL"
    sls_cadPautas = ""
'*** Seguimos los ambios de pañal activos del residente
    Sql = "select * from pautasPanales with (nolock) where residente='" & sls_id & "' and regEstado='A' and fecIni<=convert(datetime,'" & sls_fecha & "',103) and (fecFin>=convert(datetime,'" & sls_fecha & "',103) or fecFin is null)"
    If SGS_ND > 9 Then
        Loga ("Seguimos las pautas de pañales del residente.")
        Loga ("Ejecutando: " & Sql)
        Loga ("=============================================================================")
    End If
    Set rsPan = sf_rec(Sql)
    Do While Not rsPan.EOF
'**** Vamos creando la cadena de las pautas
        sls_pauta = rsPan("id")
        sls_cadPautas = sls_cadPautas & ",'" & sls_pauta & "'"
        sls_ara = rsPan("cuando")
        sls_ambito = rsPan("ambitoHorario")
        sls_cadHoras = ""
'**** Seguimos las 24 horas para ver si toca Pañal y la tratamos
        For sln_k = 1 To 24
            sls_hora = Mid(sls_ara, sln_k, 1)
            If sls_hora = "1" Then
                slb_toca = True
            Else
                slb_toca = False
            End If
            If slb_toca Then
                If sln_k < 11 Then
                    sls_hora = sls_fecha & " 0" & (sln_k - 1) & ":00:00"
                Else
                    sls_hora = sls_fecha & " " & (sln_k - 1) & ":00:00"
                End If
                sls_cadHoras = sls_cadHoras & ",convert(datetime,'" & sls_hora & "',103)"
                Sql = "select * from controlReg with (nolock) where residente='" & sls_id & "' and momento=convert(datetime,'" & sls_hora & "',103) and tipo='" & sls_tipo & sls_ambito & "' and pautaId='" & sls_pauta & "'"
                If SGS_ND > 9 Then
                    Loga ("Miramos si la pauta de pañales del residente ya esta en controlReg.")
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                Set rsRC = sf_rec(Sql)
'*** Si no existe el registro de control, lo creamos
                If rsRC.EOF Then
                    Sql = "insert into controlReg (id,residente,momento,tipo,pautaId,realizado) values("
                    Sql = Sql & "newId(),'" & sls_id & "','" & sls_hora & "','" & sls_tipo & sls_ambito & "','" & sls_pauta & "','0')"
                    If SGS_ND > 9 Then
                        Loga ("Si la pautas de pañales del residente no está en controlReg la insertamos.")
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    Set rsUp = sf_rec(Sql)
                End If
            End If
        Next
'*** Si hay una cadena con horas, borramos los que antes tenían registro, y ahora no tienen
        If sls_cadHoras <> "" Then
            sls_cadHoras = Mid(sls_cadHoras, 2)
            Sql = "delete from controlReg where residente='" & sls_id & "' and tipo='" & sls_tipo & sls_ambito & "' and pautaId='" & sls_pauta & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and momento not in (" & sls_cadHoras & ")"
            If SGS_ND > 9 Then
                Loga ("Si hay registros en ControlReg para esa pauta en horas que ya no estan, los borramos.")
                Loga ("Ejecutando: " & Sql)
                Loga ("=============================================================================")
            End If
            Set rsUp = sf_rec(Sql)
        End If
    rsPan.MoveNext
    DoEvents
    Loop
'*** Borramos los registros de control cuya pauta ya no existe para esta fecha
    If sls_cadPautas <> "" Then
        sls_cadPautas = Mid(sls_cadPautas, 2)
        Sql = "delete from controlReg where residente='" & sls_id & "' and tipo='" & sls_tipo & sls_ambito & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and pautaId not in(" & sls_cadPautas & ")"
        If SGS_ND > 9 Then
            Loga ("Borramos los registros de controlReg, de pautas de pañales que se han borrado.")
            Loga ("Ejecutando: " & Sql)
            Loga ("=============================================================================")
        End If
        Set rsUp = sf_rec(Sql)
    End If
End Function
'************************************************************************************************
'*  Tratamiento Registros Nutricion Enteral
'************************************************************************************************
Function sf_regNEnteral(ByVal sls_fecha, ByVal sls_id)
'pautas en -> pautasParenteral
'Registros en -> registroNutricionEnteral
    Dim sls_ara, sls_hora, sls_fec
    Dim sls_tipo ' Tipo de registro en curas
    Dim sls_pauta 'Para el id de la pauta
    Dim sls_cadHoras ' Para crear una cadena con las horas con curas en cada pauta
    Dim sls_cadPautas ' Para crear una cadena con todas las pautas del residente en una fecha
    sls_tipo = "NENTE"
    sls_cadPautas = ""
    Dim sla_hora(4)
    Dim sla_tipo(2)
    sla_hora(1) = sf_getParam("NENTETOMA1")
    sla_hora(2) = sf_getParam("NENTETOMA2")
    sla_hora(3) = sf_getParam("NENTETOMA3")
    sla_tipo(1) = "NUT"
    sla_tipo(2) = "HID"
'*** Seguimos las pautas de nutricion del residente
    Sql = "select * from pautasParenteral with (nolock) where residente='" & sls_id & "' and regEstado='A'"
    If SGS_ND > 9 Then
        Loga ("Seguimos los registros activos del residente.")
        Loga ("Ejecutando: " & Sql)
        Loga ("=============================================================================")
    End If
    Set rsNe = sf_rec(Sql)
    Do While Not rsNe.EOF
'**** Vamos creando la cadena de las pautas
        sls_pauta = rsNe("id")
        sls_cadPautas = sls_cadPautas & ",'" & sls_pauta & "'"
'**** Inicializamos la cadena de las horas
        For sln_j = 1 To 2
            For sln_k = 1 To 3
                Sql = "select * from controlReg with (nolock) where residente='" & sls_id & "' and momento=convert(datetime,'" & sls_fecha & " " & sla_hora(sln_k) & "',103) and tipo='" & sls_tipo & sla_tipo(sln_j) & sln_k & "' and pautaId='" & sls_pauta & "'"
                If SGS_ND > 9 Then
                    Loga ("Miramos si la pauta de nutricion esta ya en controlReg.")
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                Set rsRC = sf_rec(Sql)
'*** Si no existe el registro de control, lo creamos
                If rsRC.EOF Then
                    Sql = "insert into controlReg (id,residente,momento,tipo,pautaId,realizado) values("
                    Sql = Sql & "newId(),'" & sls_id & "','" & sls_fecha & " " & sla_hora(sln_k) & "','" & sls_tipo & sla_tipo(sln_j) & sln_k & "','" & sls_pauta & "','0')"
                    If SGS_ND > 9 Then
                        Loga ("Si la pauta de nutricion no esta en controlReg, la creamos.")
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    Set rsUp = sf_rec(Sql)
                End If
            Next
        Next
    rsNe.MoveNext
    DoEvents
    Loop
'*** Borramos los registros de control cuya pauta ya no existe para esta fecha
    If sls_cadPautas <> "" Then
        sls_cadPautas = Mid(sls_cadPautas, 2)
        For sln_j = 1 To 2
            For sln_k = 1 To 3
                Sql = "delete from controlReg where residente='" & sls_id & "' and tipo='" & sls_tipo & sla_tipo(sln_j) & sln_k & "' and momento>=convert(datetime,'" & sls_fecha & "',103) and momento<convert(datetime,'" & sls_fecha & " 23:59:59.999',103) and pautaId not in(" & sls_cadPautas & ")"
                If SGS_ND > 9 Then
                    Loga ("Borramos de controlReg, las pautas que se han borrado de curas.")
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                Set rsUp = sf_rec(Sql)
            Next
        Next
    End If
End Function
'************************************************************************************************
'*  Tratamiento Finalizacion Contratos residente/usuario
'************************************************************************************************
Function sf_finContrato(ByVal sls_fecha, ByVal sls_id, ByVal sls_tipo)
    Dim sls_fechaF, sls_tabla
    sls_fechaF = FormatDateTime(sls_fecha, 2)
    'Parametro de finalizacion de contratos automaticamente
    slb_fin = sf_getParam("FINCONTRATORESIAUT")
    If slb_fin = 1 Or slb_fin = "true" Or slb_fin = True Then
        If sls_tipo = "residente" Then
            sls_tabla = "contratosReservas"
            Sql = "select id from contratosReservas with (nolock) where residente='" & sls_id & "' and estado='A' and regEstado='A' and estadoContrato='CONTRATO' and fechaFinReal is null "
            Sql = Sql & " and fechaFin<=convert(datetime,'" & sls_fechaF & " 23:59:59.999',103) "
        Else
            sls_tabla = "contratosUsr"
            Sql = "select id from contratosUsr with (nolock) where usuario='" & sls_id & "' and estado='A' and regEstado='A' and fechaBaja<=convert(datetime,'" & sls_fechaF & " 23:59:59.999',103) "
        End If
        Set rsCont = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=============================================================================")
        End If
        Do While Not rsCont.EOF
            sls_idC = rsCont("id")
            If sls_tipo = "residente" Then
                'baja registro actual
                Sql = "UPDATE " & sls_tabla & " set regEstado='A',estado='B',causaBaja='FINC',fechaFinReal=convert(datetime,'" & sls_fechaF & "',103),fecMod=getDate(),usrMod='GdR' "
                Sql = Sql & "where id='" & sls_idC & "' and estado='A' and regEstado='A' "
                Set rsUpd = sf_rec(Sql)
                If SGS_ND > 9 Then
                    Loga ("Finalizando contrato" & sls_idC)
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                Sql = "UPDATE ocupacion set estado='B',fecMod=getDate(),usrMod='GdR' where contrato='" & sls_idC & "' and regEstado='A' and estado='A'"
                Set rsUpd = sf_rec(Sql)
                Sql = "UPDATE contratosTarifa set estado='B',fecMod=getDate(),usrMod='GdR' where contrato='" & sls_idC & "' and regEstado='A' and estado='A'"
                Set rsUpd = sf_rec(Sql)
                'Si no existen mas contratos vigentes para este residente, lo damos de baja
                Sql = "select count(id) as num from contratosReservas with (nolock) where residente='" & sls_id & "' and estado='A' and regEstado='A' "
                Set rsContA = sf_rec(Sql)
                If rsContA("num") < 1 Then
                    Sql = "UPDATE residentes set estado='B',fecMod=getDate(),usrMod='GdR' where id='" & sls_id & "' and regEstado='A' "
                    Set rsUpd = sf_rec(Sql)
                    If SGS_ND > 9 Then
                        Loga ("Dando de baja residente" & sls_id)
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    'Cerramos calquier ausencia activa
                    Sql = "select id from ausencias with (nolock) where (fechaFin is null or fechaFin='') and regEstado='A'"
                    Set rsAus = sf_rec(Sql)
                    Do While Not rsAus.EOF
                        Sql = "update ausencias set fechaFin=convert(datetime,'" & sls_fechaF & "',103) where id='" & rsAus("id") & "' and regEstado='A' "
                        Set rsUpd = sf_rec(Sql)
                        rsAus.MoveNext
                    Loop
                End If
            Else
                'baja registro actual
                Sql = "UPDATE " & sls_tabla & " set regEstado='A',estado='B',motivoFinal='FINC',fechaBaja=convert(datetime,'" & sls_fechaF & "',103),fecMod=getDate(),usrMod='GdR' "
                Sql = Sql & "where id='" & sls_idC & "' and estado='A' and regEstado='A' "
                Set rsUpd = sf_rec(Sql)
                If SGS_ND > 9 Then
                    Loga ("Finalizando contrato" & sls_idC)
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                'Si no existen mas contratos vigentes para este usuario, lo damos de baja
                Sql = "select count(id) as num from contratosUsr with (nolock) where usuario='" & sls_id & "' and estado='A' and regEstado='A' "
                Set rsContA = sf_rec(Sql)
                If rsContA("num") < 1 Then
                    Sql = "UPDATE usuarios set estado='B',fecMod=getDate(),usrMod='GdR' where id='" & sls_id & "' and regEstado='A' "
                    Set rsUpd = sf_rec(Sql)
                    If SGS_ND > 9 Then
                        Loga ("Dando de baja usuario" & sls_id)
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                End If
            End If
            rsCont.MoveNext
            DoEvents
        Loop
    End If
End Function

'************************************************************************************************
'*  Tratamiento Altas Contratos residente/usuario
'************************************************************************************************
Function sf_altaContrato(ByVal sls_fecha, ByVal sls_id, ByVal sls_tipo)
    Dim sls_fechaF, sls_tabla
    sls_fechaF = FormatDateTime(sls_fecha, 2)
    'Buscamos contratos en pre-Alta para la fecha indicada
    If sls_tipo = "residente" Then
        sls_tabla = "residentes"
        sls_tablaC = "contratosReservas"
        Sql = "select id,residente from contratosReservas with (nolock) where residente='" & sls_id & "' and estado='P' and regEstado='A' and estadoContrato='CONTRATO' "
        Sql = Sql & " and fechaInicio<=convert(datetime,'" & sls_fechaF & " 23:59:59.999',103) and fechaInicio>=convert(datetime,'" & sls_fechaF & "',103) "
    Else
        sls_tabla = "usuarios"
        sls_tablaC = "contratosUsr"
        Sql = "select id,usuario from contratosUsr with (nolock) where usuario='" & sls_id & "' and estado='P' and regEstado='A' "
        Sql = Sql & " and fechaAlta<=convert(datetime,'" & sls_fechaF & " 23:59:59.99',103) and fechaAlta>=convert(datetime,'" & sls_fechaF & "',103) "
    End If
    Set rsCont = sf_rec(Sql)
    If SGS_ND > 9 Then
        Loga ("Ejecutando: " & Sql)
        Loga ("=============================================================================")
    End If
    'Modificamos contratos
    Do While Not rsCont.EOF
        If sls_tipo = "residente" Then
            sls_id = rsCont("residente")
            sls_idC = rsCont("id")
        Else
            sls_id = rsCont("usuario")
            sls_idC = rsCont("id")
        End If
        Sql = "select getDate() timeStamp"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        'Alta residente/usuario
        Sql = "UPDATE " & sls_tabla & " set estado='A',fecMod='" & sls_timeStamp & "',usrMod='GdR' where id='" & sls_id & "' and regEstado='A' "
        Set rsUpd = sf_rec(Sql)
        'Alta contrato
        Sql = "UPDATE " & sls_tablaC & " set estado='A',fecMod='" & sls_timeStamp & "',usrMod='GdR' where id='" & sls_idC & "' and regEstado='A' "
        Set rsUpd = sf_rec(Sql)
        If sls_tipo = "residente" Then
            If SGS_ND > 9 Then
                Loga ("Alta residente" & sls_id)
                Loga ("Ejecutando: " & Sql)
                Loga ("=============================================================================")
            End If
        Else
            If SGS_ND > 9 Then
                Loga ("Alta usuario" & sls_id)
                Loga ("Ejecutando: " & Sql)
                Loga ("=============================================================================")
            End If
        End If
        rsCont.MoveNext
        DoEvents
    Loop
End Function
'************************************************************************************************
'*  Tratamiento Aniversarios residentes
'************************************************************************************************
Function sf_aniversario(ByVal sls_id)
    If SGS_ND > 5 Then
        Loga ("Tratando aniversario de residente : " & sls_id)
        Loga ("=============================================================================")
    End If
    Sql = "select fechaNac,centro from residentes with (nolock) where id='" & sls_id & "' and regEstado='A' and estado in ('A','AU') "
    Set rsFec = sf_rec(Sql)
    If SGS_ND > 9 Then
        Loga ("Ejecutando: " & Sql)
        Loga ("=============================================================================")
    End If
    If Not rsFec.EOF Then
        sls_centro = rsFec("centro")
        sls_fechaNac = rsFec("fechaNac")
        If sls_fechaNac <> "" And IsNull(sls_fechaNac) = False And sls_fechaNac > "01/01/1900" Then
            sls_anyo = Year(Date)
            sls_mes = Month(sls_fechaNac)
            If Len(sls_mes) = 1 Then sls_mes = "0" & sls_mes
            sls_dia = Day(sls_fechaNac)
            If Len(sls_dia) = 1 Then sls_dia = "0" & sls_dia
            If sls_dia = "29" And sls_mes = "02" Then sls_dia = "28" 'Año bisiesto, restamos un dia siempre
            sls_fechaNacF = CDate(sls_dia & "/" & sls_mes & "/" & sls_anyo)
            sls_fechaNacF = FormatDateTime(sls_fechaNacF, 2)
            sls_fechaNacF2 = CDate(sls_dia & "/" & sls_mes & "/" & CInt(sls_anyo + 1))
            sls_fechaNacF2 = FormatDateTime(sls_fechaNacF2, 2)
            sls_fechaNacF = sls_dia & "/" & sls_mes & "/" & sls_anyo
            sls_fechaNacF2 = DateAdd("yyyy", 1, sls_fechaNac)
            'Registros año actual
            Sql = "select id,detalle,momentoIni,detalle from agenda with (nolock) where idOrigen='" & sls_id & "' and regEstado='A' and tipoAnotacion='aniversario' "
            Sql = Sql & "and momentoIni<'01/01/" & sls_anyo + 1 & "' and momentoIni>='01/01/" & sls_anyo & "' "
            Set rsExist = sf_rec(Sql)
            If SGS_ND > 9 Then
                Loga ("Ejecutando: " & Sql)
                Loga ("=============================================================================")
            End If
            slb_altaH = 0
            If Not rsExist.EOF Then
                sls_idAgenda = rsExist("id")
                sls_momentoIni = rsExist("momentoIni")
                sls_idDetalle = rsExist("detalle")
                'Si existen registros pero con fechas diferentes a las guardadas, damos de baja actuales
                If DateDiff("d", sls_momentoIni, sls_fechaNacF) <> 0 Then
                    Sql = "UPDATE agenda set regEstado='H',fecMod=getDate(),usrMod='GdR' where id='" & sls_idAgenda & "' and regEstado='A' "
                    Set rsUpd = sf_rec(Sql)
                    If SGS_ND > 9 Then
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    If slb_aviso = 1 Then
                        Sql = "select id from agendaAviso with (nolock) where regEstado='A' and idAgenda='" & sls_idAgenda & "' "
                        Set rsAviso = sf_rec(Sql)
                        If Not rsAviso.EOF Then
                            sls_idAviso = rsAviso("id")
                            Sql = "UPDATE agendaAviso set regEstado='H',fecMod=getDate(),usrMod='GdR' where id='" & sls_idAviso & "' and regEstado='A' "
                            Set rsUpd = sf_rec(Sql)
                            If SGS_ND > 9 Then
                                slo_fname.WriteLine ("Ejecutando: " & Sql)
                                slo_fname.WriteLine ("=============================================================================")
                            End If
                        End If
                    End If
                    Sql = "UPDATE agendaResidente set regEstado='H',fecMod=getDate(),usrMod='GdR' where idAnotacion='" & sls_idAgenda & "' and regEstado='A' "
                    Set rsUpd = sf_rec(Sql)
                    If SGS_ND > 9 Then
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    Sql = "UPDATE infoEditor set regEstado='H',fecMod=getDate(),usrMod='GdR' where id='" & sls_idDetalle & "' and regEstado='A' "
                    Set rsUpd = sf_rec(Sql)
                    If SGS_ND > 9 Then
                        Loga ("Ejecutando: " & Sql)
                        Loga ("=============================================================================")
                    End If
                    sls_idAviso = rsAviso("idAviso")
                    slb_altaH = 1
                ElseIf DateDiff("d", sls_momentoIni, sls_fechaNacF) = 0 Then
                    slb_altaH = 0
                End If
            Else
                Sql = "select newId() idAgenda,newId() idDetalle,newId() idAviso"
                Set rsId = sf_rec(Sql)
                sls_idAgenda = rsId("idAgenda")
                sls_idDetalle = rsId("idDetalle")
                sls_idAviso = rsId("idAviso")
                slb_altaH = 1
            End If
            If slb_altaH = 1 Then
                Sql = "select getDate() timeStamp "
                Set rsTime = sf_rec(Sql)
                sls_timeStamp = rsTime("timeStamp")
                Sql = "INSERT into agenda(id,centro,momentoIni,momentoFin,tipoAnotacion,idOrigen,resumen,detalle,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                Sql = Sql & "values('" & sls_idAgenda & "','" & sls_centro & "','" & sls_fechaNacF & "','" & sls_fechaNacF & "',"
                Sql = Sql & "'aniversario','" & sls_id & "','" & sf_limpia(SGD_Literal.Item("aniversario")) & "','" & sls_idDetalle & "','A',"
                Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
                Set rsIns = sf_rec(Sql)
                If SGS_ND > 9 Then
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                If slb_aviso = 1 Then
                    Sql = "INSERT into agendaAviso(id,idAgenda,fecha,avisado,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                    Sql = Sql & "values('" & sls_idAviso & "','" & sls_idAgenda & "','" & sls_fechaNacF & "',0,'A',"
                    Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
                    Set rsIns = sf_rec(Sql)
                    If SGS_ND > 9 Then
                        slo_fname.WriteLine ("Ejecutando: " & Sql)
                        slo_fname.WriteLine ("=============================================================================")
                    End If
                End If
                Sql = "INSERT into infoEditor(id,tabla,campo,valor,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                Sql = Sql & "values('" & sls_idDetalle & "','datosAdmin','agenda','" & sf_limpia(SGD_Literal.Item("aniversario")) & ":" & sls_fechaNacF & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
                Set rsIns = sf_rec(Sql)
                If SGS_ND > 9 Then
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                Sql = "INSERT into agendaResidente(id,idAnotacion,residente,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                Sql = Sql & "values(newId(),'" & sls_idAgenda & "','" & sls_id & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
                Set rsIns = sf_rec(Sql)
                If SGS_ND > 9 Then
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
            End If
            'Registros año siguiente
            Sql = "select id,detalle,momentoIni,detalle from agenda with (nolock) where idOrigen='" & sls_id & "' and regEstado='A' and tipoAnotacion='aniversario' "
            Sql = Sql & "and momentoIni<'01/01/" & sls_anyo + 2 & "' and momentoIni>='01/01/" & sls_anyo + 1 & "' "
            Set rsExist = sf_rec(Sql)
            If SGS_ND > 9 Then
                Loga ("Ejecutando: " & Sql)
                Loga ("=============================================================================")
            End If
            slb_altaF = 0
            If Not rsExist.EOF Then
                sls_idAgenda = rsExist("id")
                sls_momentoIni = rsExist("momentoIni")
                sls_idDetalle = rsExist("detalle")
                'Si existen registros pero con fechas diferentes a las guardadas, damos de baja actuales
                If DateDiff("d", sls_momentoIni, sls_fechaNacF2) <> 0 Then
                    Sql = "UPDATE agenda set regEstado='H',fecMod=getDate(),usrMod='GdR' where id='" & sls_idAgenda & "' and regEstado='A' "
                    Set rsUpd = sf_rec(Sql)
                    If slb_aviso = 1 Then
                        Sql = "select id from agendaAviso with (nolock) where regEstado='A' and idAgenda='" & sls_idAgenda & "' "
                        Set rsAviso = sf_rec(Sql)
                        If Not rsAviso.EOF Then
                            sls_idAviso = rsAviso("id")
                            Sql = "UPDATE agendaAviso set regEstado='H',fecMod=getDate(),usrMod='GdR' where id='" & sls_idAviso & "' and regEstado='A' "
                            Set rsUpd = sf_rec(Sql)
                        End If
                    End If
                    Sql = "UPDATE agendaResidente set regEstado='H',fecMod=getDate(),usrMod='GdR' where idAnotacion='" & sls_idAgenda & "' and regEstado='A' "
                    Set rsUpd = sf_rec(Sql)
                    Sql = "UPDATE infoEditor set regEstado='H',fecMod=getDate(),usrMod='GdR' where id='" & sls_idDetalle & "' and regEstado='A' "
                    Set rsUpd = sf_rec(Sql)
                    slb_altaF = 1
                ElseIf DateDiff("d", sls_momentoIni, sls_fechaNacF) = 0 Then
                    slb_altaF = 0
                End If
            Else
                Sql = "select newId() idAgenda,newId() idDetalle, newId() idAviso"
                Set rsId = sf_rec(Sql)
                sls_idAgenda = rsId("idAgenda")
                sls_idDetalle = rsId("idDetalle")
                sls_idAviso = rsId("idAviso")
                slb_altaF = 1
            End If
            If slb_altaF = 1 Then
                Sql = "select getDate() timeStamp "
                Set rsTime = sf_rec(Sql)
                sls_timeStamp = rsTime("timeStamp")
                Sql = "INSERT into agenda(id,centro,momentoIni,momentoFin,tipoAnotacion,idOrigen,resumen,detalle,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                Sql = Sql & "values('" & sls_idAgenda & "','" & sls_centro & "','" & sls_fechaNacF2 & "','" & sls_fechaNacF2 & "',"
                Sql = Sql & "'aniversario','" & sls_id & "','" & sf_limpia(SGD_Literal.Item("aniversario")) & "','" & sls_idDetalle & "','A',"
                Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
                Set rsIns = sf_rec(Sql)
                If SGS_ND > 9 Then
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                If slb_aviso = 1 Then
                    Sql = "INSERT into agendaAviso(id,idAgenda,fecha,avisado,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                    Sql = Sql & "values('" & sls_idAviso & "','" & sls_idAgenda & "','" & sls_fechaNacF2 & "',0,'A',"
                    Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
                    Set rsIns = sf_rec(Sql)
                    If SGS_ND > 9 Then
                        slo_fname.WriteLine ("Ejecutando: " & Sql)
                        slo_fname.WriteLine ("=============================================================================")
                    End If
                End If
                Sql = "INSERT into infoEditor(id,tabla,campo,valor,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                Sql = Sql & "values('" & sls_idDetalle & "','datosAdmin','agenda','" & sf_limpia(SGD_Literal.Item("aniversario")) & ":" & sls_fechaNacF2 & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
                Set rsIns = sf_rec(Sql)
                If SGS_ND > 9 Then
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
                Sql = "INSERT into agendaResidente(id,idAnotacion,residente,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
                Sql = Sql & "values(newId(),'" & sls_idAgenda & "','" & sls_id & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
                Set rsIns = sf_rec(Sql)
                If SGS_ND > 9 Then
                    Loga ("Ejecutando: " & Sql)
                    Loga ("=============================================================================")
                End If
            End If
        End If
    End If
End Function

'**********************************************************************************
'*Funciones de generacion de alertas segun condiciones determinadas para cada una
'**********************************************************************************
'* sf_getParamAlerta( parametro , tipoAlerta)
'*devuelve el valor del parametro correspondiente al tipo de alerta
'**********************************************************************************
Function sf_getParamAlerta(ByVal sls_param, ByVal sls_tipoAlerta)
    Dim sls_sql, sls_valor, sls_tip
    sls_valor = ""
    sls_tip = ""
    Sql = "select id from gdrAlertas with (nolock) where tipo='" & sls_tipoAlerta & "' and regEstado='A'"
    Set rsIdAlerta = sf_recGdr(Sql)
    If Not rsIdAlerta.EOF Then
        sls_idAlerta = rsIdAlerta("id")
    Else
        sls_idAlerta = ""
    End If
    Sql = "select tipoDato,valor from alertasParam with (nolock) where idAlerta='" & sls_idAlerta & "' and parametro='" & sls_param & "' and regEstado='A'"
    Set rsParametro = sf_rec(Sql)
    If Not rsParametro.EOF Then
        sls_valor = rsParametro("valor")
        sls_tip = rsParametro("tipoDato")
    End If
    If sls_tip = "ENTERO" Then
        If IsNumeric(sls_valor) Then
            sls_valor = CDbl(sls_valor)
        End If
    End If
    If Mid(sls_tip, 1, 3) = "DEC" Then
        If Mid(UCase(SGC_SERVER), 1, 12) = "SOLUNOVA-SRV" Then
            sls_valor = Replace(sls_valor, ".", ",")
        End If
        sls_valor = CDbl(sls_valor)
    End If
    sf_getParamAlerta = sls_valor
End Function
'**********************************************************************************
'* sf_alertaAfeitado( residente,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaAfeitado(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    sls_fec = 0
    'Consultamos ultimo registro alertas
    Sql = "select top 1 control from alertas with (nolock) where idAlerta='" & sls_idAlerta & "'  and  regEstado='A' and residente=" & sls_residente & " order by fecAlta desc"
    Set rsControl = sf_rec(Sql)
    If Not rsControl.EOF Then
        sls_control = rsControl("control")
    Else
        sls_control = ""
    End If
    'Consultamos si el residente ha regresado de una ausencia
    Sql = "select fechaFin from ausencias with (nolock) where regEstado='A' and residente=" & sls_residente & "  "
    Set rsAus = sf_rec(Sql)
    If Not rsAus.EOF Then
        sls_finAus = rsAus("fechaFin")
    Else
        sls_finAus = 0
    End If
    Sql = "select top 1 id,fecAlta from  with (nolock)registroAseos where tipo='AFEITADO' and regEstado='A' and residente=" & sls_residente & "  "
    Sql = Sql & " and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) order by fecAlta desc"
    Set rsAseos = sf_rec(Sql)
    If SGS_ND > 9 Then
        Loga ("Ejecutando: " & Sql)
        Loga ("=================================================================================")
    End If
    'Comparamos la fechaFin de la ausencia con el utlimo registro de deposicion y escojemos el mayor
    If Not rsAseos.EOF Then
        sls_fecAlta = rsAseos("fecAlta")
        sls_idReg = rsAseos("id")
        If sls_idReg = sls_control Then
            sls_fec = 0
        Else
            If sls_finAus <> 0 And sls_fecAlta <> "" Then
                sls_difFec = DateDiff("d", sls_finAus, sls_fecAlta)
                If sls_difFec < 0 Then
                    sls_fec = sls_finAus
                Else
                    sls_fec = sls_fecAlta
                End If
            Else
                sls_fec = sls_fecAlta
            End If
        End If
    Else
        sls_fec = 0
        sls_idReg = ""
    End If
    If sls_fec <> 0 Then
        sls_dif = DateDiff("d", sls_fec, Date)
    Else
        sls_dif = 0
    End If
    If sls_dif > sls_param1 Then
        slb_crearAlerta = 1
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-SupNumMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegAfeitado")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "','" & sls_idReg & "')"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaBano( residente,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaBano(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    sls_fec = 0
    'Consultamos ultimo registro alertas
    Sql = " select top 1 control from alertas with (nolock) where idAlerta='" & sls_idAlerta & "'  and  regEstado='A' and residente=" & sls_residente & " order by fecAlta desc"
    Set rsControl = sf_rec(Sql)
    If Not rsControl.EOF Then
        sls_control = rsControl("control")
    Else
        sls_control = ""
    End If
    'Consultamos si el residente ha regresado de una ausencia
    Sql = "select fechaFin from ausencias with (nolock) where regEstado='A' and residente=" & sls_residente & "  "
    Set rsAus = sf_rec(Sql)
    If SGS_ND > 9 Then
        Loga ("Ejecutando: " & Sql)
        Loga ("=================================================================================")
    End If
    If Not rsAus.EOF Then
        sls_finAus = rsAus("fechaFin")
    Else
        sls_finAus = 0
    End If
    Sql = "select top 1 fecAlta,id from registroAseos with (nolock) where tipo='BANO' and regEstado='A' and residente=" & sls_residente & "  "
    Sql = Sql & " and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) order by fecAlta desc"
    Set rsAseos = sf_rec(Sql)
    If SGS_ND > 9 Then
        Loga ("Ejecutando: " & Sql)
        Loga ("=================================================================================")
    End If
    'Comparamos la fechaFin de la ausencia con el utlimo registro de deposicion y escojemos el mayor
    If Not rsAseos.EOF Then
        sls_fecAlta = rsAseos("fecAlta")
        sls_idReg = rsAseos("id")
        If sls_idReg = sls_control Then
            sls_fec = 0
        Else
            If sls_finAus <> 0 And sls_fecAlta <> "" Then
                sls_difFec = DateDiff("d", sls_finAus, sls_fecAlta)
                If sls_difFec < 0 Then
                    sls_fec = sls_finAus
                Else
                    sls_fec = sls_fecAlta
                End If
            Else
                sls_fec = sls_fecAlta
            End If
        End If
    Else
        sls_fec = 0
    End If
    If sls_fec <> 0 Then
        sls_dif = DateDiff("d", sls_fec, Date)
    Else
        sls_dif = 0
    End If
    If sls_dif > sls_param1 Then
        slb_crearAlerta = 1
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-SupNumMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegBano")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "','" & sls_idReg & "')"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaCaida( residente,parametro1,parametro2,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaCaida(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_param2, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    sls_fecha1 = (FormatDateTime(DateAdd("d", -sls_param2, Date), 2)) 'Dia conseguido al restar periodo de tiempo a la fecha actual
    Sql = "select fecAlta from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and regEstado='A' and residente=" & sls_residente & " "
    Set rsFecAlerta = sf_rec(Sql)
    If Not rsFecAlerta.EOF Then
        sls_fecAlta = rsFecAlerta("fecAlta")
        sls_fecha2 = (FormatDateTime(DateAdd("d", -sls_fecAlta, Date), 2)) 'Dia conseguido al restar periodo de tiempo a la fecha actual
    Else
        sls_fecha2 = ""
    End If
    Sql = "select id from registroCaidas with (nolock) where regEstado='A' and residente=" & sls_residente & " and fecAlta>convert(datetime,'" & sls_fechaAyer & "',103) and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) "
    Set rsHoy = sf_rec(Sql)
    If Not rsHoy.EOF Then
        If sls_fecha2 <> "" Then
            sls_difFec = DateDiff("d", sls_fecha1, sls_fecha2)
            If sls_difFec < 0 Then
                sls_fec = sls_fecha1
            Else
                sls_fec = sls_fecha2
            End If
        Else
            sls_fec = sls_fecha1
        End If
        Sql = "select count(id) as num from registroCaidas with (nolock) where regEstado='A' and residente=" & sls_residente & " and fechaCaida>=convert(datetime,'" & sls_fec & "',103) and fechaCaida<=getDate() "
        Set rsCaidas = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsCaidas.EOF Then
            sls_num = rsCaidas("num")
            If sls_num > sls_param1 Then
                slb_crearAlerta = 1
            End If
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-SupNumMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-CaidasPeriodo")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaCamas( residente,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaCamas(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer1 = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    sls_fec = 0
    'Consultamos ultimo registro alertas
    Sql = "select top 1 control from alertas with (nolock) where idAlerta='" & sls_idAlerta & "'  and  regEstado='A' and residente=" & sls_residente & " order by fecAlta desc"
    Set rsControl = sf_rec(Sql)
    If Not rsControl.EOF Then
        sls_control = rsControl("control")
    Else
        sls_control = ""
    End If
    'Consultamos si el residente ha regresado de una ausencia
    Sql = "select fechaFin from ausencias with (nolock) where regEstado='A' and residente=" & sls_residente & "  "
    Set rsAus = sf_rec(Sql)
    If SGS_ND > 9 Then
        Loga ("Ejecutando: " & Sql)
        Loga ("=================================================================================")
    End If
    If Not rsAus.EOF Then
        sls_finAus = rsAus("fechaFin")
    Else
        sls_finAus = 0
    End If
    Sql = "select top 1 fecAlta,id from registroAseos with (nolock) where tipo='CAMBIOROPACAMA' and regEstado='A' and residente=" & sls_residente & " "
    Sql = Sql & " and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) order by fecAlta desc"
    Set rsAseos = sf_rec(Sql)
    If SGS_ND > 9 Then
        Loga ("Ejecutando: " & Sql)
        Loga ("=================================================================================")
    End If
    'Comparamos la fechaFin de la ausencia con el utlimo registro de deposicion y escojemos el mayor
    If Not rsAseos.EOF Then
        sls_fecAlta = rsAseos("fecAlta")
        sls_idReg = rsAseos("id")
        If sls_idReg = sls_control Then
            sls_fec = 0
        Else
            If sls_finAus <> 0 And sls_fecAlta <> "" Then
                sls_difFec = DateDiff("d", sls_finAus, sls_fecAlta)
                If sls_difFec < 0 Then
                    sls_fec = sls_finAus
                Else
                    sls_fec = sls_fecAlta
                End If
            Else
                sls_fec = sls_fecAlta
            End If
        End If
    Else
        sls_fec = 0
    End If
    If sls_fec <> 0 Then
        sls_dif = DateDiff("d", sls_fec, Date)
    Else
        sls_dif = 0
    End If
    If sls_dif > sls_param1 Then
        slb_crearAlerta = 1
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-SupNumMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegCambioCama")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "','" & sls_idReg & "')"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaDeposicion( residente,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaDeposicion(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    sls_fec = 0
    'Consultamos ultimo registro alertas
    Sql = "select top 1 control from alertas with (nolock) where idAlerta='" & sls_idAlerta & "'  and  regEstado='A' and residente=" & sls_residente & " order by fecAlta desc"
    Set rsControl = sf_rec(Sql)
    If Not rsControl.EOF Then
        sls_control = rsControl("control")
    Else
        sls_control = ""
    End If
    'Consultamos si el residente ha regresado de una ausencia
    Sql = "select fechaFin from ausencias with (nolock) where regEstado='A' and residente=" & sls_residente & "  "
    Set rsAus = sf_rec(Sql)
    If SGS_ND > 9 Then
        Loga ("Ejecutando: " & Sql)
        Loga ("=================================================================================")
    End If
    If Not rsAus.EOF Then
        sls_finAus = rsAus("fechaFin")
    Else
        sls_finAus = 0
    End If
    Sql = "select top 1 momento,id from registroDeposiciones with (nolock) where regEstado='A' and residente=" & sls_residente & "  "
    Sql = Sql & " and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) order by momento desc"
    Set rsDep = sf_rec(Sql)
    If SGS_ND > 9 Then
        Loga ("Ejecutando: " & Sql)
        Loga ("=================================================================================")
    End If
    'Comparamos la fechaFin de la ausencia con el utlimo registro de deposicion y escojemos el mayor
    If Not rsDep.EOF Then
        sls_momento = rsDep("momento")
        sls_idReg = rsDep("id")
        If sls_idReg = sls_control Then
            sls_fec = 0
        Else
            If sls_finAus <> 0 And sls_momento <> "" Then
                sls_difFec = DateDiff("d", sls_finAus, sls_momento)
                If sls_difFec < 0 Then
                    sls_fec = sls_finAus
                Else
                    sls_fec = sls_momento
                End If
            Else
                sls_fec = sls_momento
            End If
        End If
    Else
        sls_fec = 0
    End If
    If sls_fec <> 0 Then
        sls_dif = DateDiff("d", sls_fec, Date)
    Else
        sls_dif = 0
    End If
    If sls_dif > sls_param1 Then
        slb_crearAlerta = 1
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-SupNumMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegDeposi")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "','" & sls_idReg & "')"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaFCardiaca( residente,parametro1,parametro2,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaFCardiaca(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_param2, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    slb_max = 0
    slb_min = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and residente=" & sls_residente & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        Sql = "select top 1 frecuencia from registroFrecuencias with (nolock) where residente=" & sls_residente & " and regEstado='A' and fecAlta>convert(datetime,'" & sls_fechaAyer & "',103) and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) "
        Sql = Sql & " order by momento desc"
        Set rsFC = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsFC.EOF Then
            sls_frec = rsFC("frecuencia")
            If CDbl(sls_frec) > CDbl(sls_param1) Then
                slb_crearAlerta = 1
                slb_max = 1
            ElseIf CDbl(sls_frec) < CDbl(sls_param2) Then
                slb_crearAlerta = 1
                slb_min = 1
            End If
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        If slb_max = 1 Then
            sls_comentario = SGD_Literal.Item(sls_idioma & "-SupMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegFCardiaca")
        Else
            sls_comentario = SGD_Literal.Item(sls_idioma & "-InfMin") & "(" & sls_param2 & ")" & SGD_Literal.Item(sls_idioma & "-RegFCardiaca")
        End If
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaGlicemia( residente,parametro1,parametro2,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaGlicemia(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_param2, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    slb_max = 0
    slb_min = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and residente=" & sls_residente & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        Sql = "select top 1 glicemia from registroGlicemias with (nolock) where residente=" & sls_residente & " and regEstado='A' and fecAlta>convert(datetime,'" & sls_fechaAyer & "',103) and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) "
        Sql = Sql & " order by momento desc"
        Set rsGli = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsGli.EOF Then
            sls_glicemia = rsGli("glicemia")
            If CDbl(sls_glicemia) > CDbl(sls_param1) Then
                slb_crearAlerta = 1
                slb_max = 1
            ElseIf CDbl(sls_glicemia) < CDbl(sls_param2) Then
                slb_crearAlerta = 1
                slb_min = 1
            End If
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        If slb_max = 1 Then
            sls_comentario = SGD_Literal.Item(sls_idioma & "-SupMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegGlicemia")
        Else
            sls_comentario = SGD_Literal.Item(sls_idioma & "-InfMin") & "(" & sls_param2 & ")" & SGD_Literal.Item(sls_idioma & "-RegGlicemia")
        End If
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaO2( residente,parametro1,parametro2,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaO2(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_param2, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    slb_max = 0
    slb_min = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and residente=" & sls_residente & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        Sql = "select top 1 nivel from registroSatO2 with (nolock) where residente=" & sls_residente & " and regEstado='A' and fecAlta>convert(datetime,'" & sls_fechaAyer & "',103) and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) "
        Sql = Sql & " order by momento desc"
        Set rsSat = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsSat.EOF Then
            sls_sat = rsSat("nivel")
            If CDbl(sls_sat) > CDbl(sls_param1) Then
                slb_crearAlerta = 1
                slb_max = 1
            ElseIf CDbl(sls_sat) < CDbl(sls_param2) Then
                slb_crearAlerta = 1
                slb_min = 1
            End If
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        If slb_max = 1 Then
            sls_comentario = SGD_Literal.Item(sls_idioma & "-SupMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegSatO2")
        Else
            sls_comentario = SGD_Literal.Item(sls_idioma & "-InfMin") & "(" & sls_param2 & ")" & SGD_Literal.Item(sls_idioma & "-RegSatO2")
        End If
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaTemperatura( residente,parametro1,parametro2,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaTemperatura(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_param2, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    slb_max = 0
    slb_min = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and residente=" & sls_residente & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        Sql = "select top 1 temperatura from registroTemperatura with (nolock) where residente=" & sls_residente & " and regEstado='A' and fecAlta>convert(datetime,'" & sls_fechaAyer & "',103) and fecAlta<convert(datetime,'" & sls_fechaMan & "',103)"
        Sql = Sql & "order by momento desc"
        Set rsTemp = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsTemp.EOF Then
            sls_temp = rsTemp("temperatura")
            Loga ("Ejecutando: " & CDbl(sls_temp) & ">" & CDbl(sls_param1))
            Loga ("Ejecutando: " & CDbl(sls_temp) & "<" & CDbl(sls_param2))
            Loga ("=================================================================================")
            If CDbl(sls_temp) > CDbl(sls_param1) Then
                slb_crearAlerta = 1
                slb_max = 1
            ElseIf CDbl(sls_temp) < CDbl(sls_param2) Then
                slb_crearAlerta = 1
                slb_min = 1
            End If
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        If slb_max = 1 Then
            sls_comentario = SGD_Literal.Item(sls_idioma & "-SupMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegTemp")
        Else
            sls_comentario = SGD_Literal.Item(sls_idioma & "-InfMin") & "(" & sls_param2 & ")" & SGD_Literal.Item(sls_idioma & "-RegTemp")
        End If
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaTension( residente,parametro1,parametro2,parametro3,parametro4,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaTension(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_param2, ByVal sls_param3, ByVal sls_param4, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    slb_maxMax = 0
    slb_maxMin = 0
    slb_minMax = 0
    slb_minMin = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and residente=" & sls_residente & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        Sql = "select top 1 maxima,minima from registroTension with (nolock) where residente=" & sls_residente & " and regEstado='A' and fecAlta>convert(datetime,'" & sls_fechaAyer & "',103) and fecAlta<convert(datetime,'" & sls_fechaMan & "',103)"
        Sql = Sql & "   order by momento desc"
        Set rsTension = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsTension.EOF Then
            sls_tensionMax = rsTension("maxima")
            sls_tensionMin = rsTension("minima")
            If sls_tensionMax > sls_param1 Then
                slb_crearAlerta = 1
                slb_maxMax = 1
            ElseIf CDbl(sls_tensionMax) < CDbl(sls_param2) Then
                slb_crearAlerta = 1
                slb_maxMin = 1
            ElseIf CDbl(sls_tensionMin) > CDbl(sls_param3) Then
                slb_crearAlerta = 1
                slb_minMax = 1
            ElseIf CDbl(sls_tensionMin) < CDbl(sls_param4) Then
                slb_crearAlerta = 1
                slb_minMin = 1
            End If
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        If slb_maxMax = 1 Then
            sls_comentario = SGD_Literal.Item(sls_idioma & "-SupMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegTensMax")
        ElseIf slb_maxMin = 1 Then
            sls_comentario = SGD_Literal.Item(sls_idioma & "-InfMin") & "(" & sls_param2 & ")" & SGD_Literal.Item(sls_idioma & "-RegTensMax")
        ElseIf slb_minMax = 1 Then
            sls_comentario = SGD_Literal.Item(sls_idioma & "-SupMax") & "(" & sls_param3 & ")" & SGD_Literal.Item(sls_idioma & "-RegTensMin")
        ElseIf slb_minMin = 1 Then
            sls_comentario = SGD_Literal.Item(sls_idioma & "-InfMin") & "(" & sls_param4 & ")" & SGD_Literal.Item(sls_idioma & "-RegTensMin")
        End If
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaUnas( residente,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaUnas(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    sls_fec = 0
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and residente=" & sls_residente & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        'Consultamos ultimo registro alertas
        Sql = "select top 1 control from alertas with (nolock) where idAlerta='" & sls_idAlerta & "'  and  regEstado='A' and residente=" & sls_residente & " order by fecAlta desc"
        Set rsControl = sf_rec(Sql)
        If Not rsControl.EOF Then
            sls_control = rsControl("control")
        Else
            sls_control = ""
        End If
        'Consultamos si el residente ha regresado de una ausencia
        Sql = "select fechaFin from ausencias with (nolock) where regEstado='A' and residente=" & sls_residente & "  "
        Set rsAus = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsAus.EOF Then
            sls_finAus = rsAus("fechaFin")
        Else
            sls_finAus = 0
        End If
        Sql = "select top 1 fecAlta,id from registroAseos with (nolock) where tipo='CORTARUÑAS' and regEstado='A' and residente=" & sls_residente & " "
        Sql = Sql & " and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) order by fecAlta desc"
        Set rsAseos = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        'Comparamos la fechaFin de la ausencia con el utlimo registro de deposicion y escojemos el mayor
        If Not rsAseos.EOF Then
            sls_fecAlta = rsAseos("fecAlta")
            sls_idReg = rsAseos("id")
            If sls_idReg = sls_control Then
                sls_fec = 0
            Else
                If sls_finAus <> 0 And sls_fecAlta <> "" Then
                    sls_difFec = DateDiff("d", sls_finAus, sls_fecAlta)
                    If sls_difFec < 0 Then
                        sls_fec = sls_finAus
                    Else
                        sls_fec = sls_fecAlta
                    End If
                Else
                    sls_fec = sls_fecAlta
                End If
            End If
        Else
            sls_fec = 0
        End If
        If sls_fec <> 0 Then
            sls_dif = DateDiff("d", sls_fec, Date)
        Else
            sls_dif = 0
        End If
        If sls_dif > sls_param1 Then
            slb_crearAlerta = 1
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-SupNumMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegCorteUnas")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaRegresar( residente,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaRegresar(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and residente=" & sls_residente & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        'Consultamos ultimos registros de entrada y salida del residente para la salida autorizada
        Sql = "select top 1 momento from movimientosAutorizSalida with (nolock) where regEstado='A' and residente =" & sls_residente & " and tipoMovimiento='SALIDA' "
        Sql = Sql & " and fecAlta>convert(datetime,'" & sls_fechaAyer & "',103) and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) order by momento desc"
        Set rsMovS = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsMovS.EOF Then
            sls_momentoS = rsMovS("momento")
            Sql = "select top 1 momento from movimientosAutorizSalida with (nolock) where regEstado='A' and residente =" & sls_residente & " and tipoMovimiento='ENTRADA' "
            Sql = Sql & " and fecAlta>convert(datetime,'" & sls_fechaAyer & "',103) and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) order by momento desc "
            Set rsMovE = sf_rec(Sql)
            If SGS_ND > 9 Then
                Loga ("Ejecutando: " & Sql)
                Loga ("=================================================================================")
            End If
            If Not rsMovE.EOF Then
                sls_momentoE = rsMovE("momento")
            Else
                sls_momentoE = ""
            End If
            sls_dif = DateDiff(h, sls_momentoS, sls_momentoE)
            If sls_dif > sls_param1 Then
                slb_crearAlerta = 1
            End If
        Else
            sls_momentoS = ""
            sls_momentoE = ""
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-SupNumMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegAutSalida")
        'sls_comentario=SGD_Literal.item(sls_idioma & "-SupRegreso") & "(" & sls_param1 & ")" & SGD_Literal.item(sls_idioma & "-RegAutSalida")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaPeso( residente,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaPeso(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and residente=" & sls_residente & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        Sql = "select peso,momento from registroPeso with (nolock) where residente=" & sls_residente & " and regEstado='A' and fecAlta>convert(datetime,'" & sls_fechaAyer & "',103) and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) "
        Set rsPesoA = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsPesoA.EOF Then
            sls_pesoAct = rsPesoA("peso")
            sls_momentoAct = rsPesoA("momento")
            Sql = "select top 1 peso,momento from registroPeso with (nolock) where residente=" & sls_residente & " and regEstado='A' and fecAlta<=convert(datetime,'" & sls_fechaAyer & "',103) order by momento desc"
            Set rsPesoUlt = sf_rec(Sql)
            If SGS_ND > 9 Then
                Loga ("Ejecutando: " & Sql)
                Loga ("=================================================================================")
            End If
            If Not rsPesoUlt.EOF Then
                sls_pesoUlt = rsPesoUlt("peso")
                sls_momentoUlt = rsPesoUlt("momento")
            Else
                sls_pesoUlt = 0
                sls_momentoUlt = 0
            End If
            If sls_pesoAct <> "" And sls_pesoUlt <> "" And sls_momentoAct <> "" And sls_momentoUlt <> "" Then
                sls_difT = Abs(DateDiff("d", sls_momentoAct, sls_momentoUlt))
                sls_difP = Abs(sls_pesoAct - sls_pesoUlt)
                If sls_difT <> 0 And sls_difP <> 0 Then
                    sls_prop1 = sls_difP / sls_difT
                    sls_prop2 = sls_param1 / 30
                    If sls_prop1 > sls_prop2 Then
                        slb_crearAlerta = 1
                    End If
                End If
            End If
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-SupVarMes") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-RegPeso")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaDocCaduca empleado,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el empleado seleccionado
'**********************************************************************************
Function sf_alertaDocCaduca(ByVal sls_usuario, ByVal sls_empresa, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdiomaE(sls_empresa)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    sls_fec = 0
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and usuario=" & sls_usuario & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        Sql = "select id,fechaCaducidad from usuarios with (nolock) where regEstado='A' and id=" & sls_usuario & " and (fechaCaducidad<>'' or fechaCaducidad is not null)  "
        Sql = Sql & " and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) order by fecAlta desc"
        Set rsFec = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsFec.EOF Then
            sls_fecCad = rsFec("fechaCaducidad")
            sls_dif = DateDiff("d", Date, sls_fecCad)
        Else
            sls_dif = 0
        End If
        If sls_dif > 0 And sls_dif < sls_param1 Then
            slb_crearAlerta = 1
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-Faltan") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-DocCaduca")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,usuario,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','',''," & sls_usuario & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaContEmpCaduca (empleado,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el empleado seleccionado
'**********************************************************************************
Function sf_alertaContEmpCaduca(ByVal sls_usuario, ByVal sls_empresa, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdiomaE(sls_empresa)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    sls_fec = 0
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and usuario=" & sls_usuario & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        Sql = "select id,fechaBaja from contratosUsr with (nolock) where regEstado='A' and usuario=" & sls_usuario & " and (fechaBaja<>'' or fechaBaja is not null)  "
        Sql = Sql & " and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) order by fecAlta desc"
        Set rsFec = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsFec.EOF Then
            sls_fecCad = rsFec("fechaBaja")
            sls_dif = DateDiff("d", Date, sls_fecCad)
        Else
            sls_dif = 0
        End If
        If sls_dif > 0 And sls_dif < sls_param1 Then
            slb_crearAlerta = 1
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-Faltan") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-ContEmpCaduca")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,usuario,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','',''," & sls_usuario & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaEmpleadoRegResi (empleado,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el empleado seleccionado
'**********************************************************************************
Function sf_alertaEmpleadoRegResi(ByVal sls_usuario, ByVal sls_empresa, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdiomaE(sls_empresa)
    sls_fechaParam = (FormatDateTime(DateAdd("d", CInt("-" & Abs(sls_param1)), Date), 2)) 'Dia segun param
    sls_fec = 0
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and usuario=" & sls_usuario & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        Sql = "select top 1 id from lopd with (nolock) where usrAlta=" & sls_usuario & " and momento>'" & sls_fechaParam & "'"
        Sql = Sql & " and pagina='RESpesResDatosAdm' order by momento desc"
        Set rsFec = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsFec.EOF Then
            slb_crearAlerta = 0
        Else
            slb_crearAlerta = 1
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-SupNumMax") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-EmpleadoRegResi")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,usuario,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','',''," & sls_usuario & ",'" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'* sf_alertaContResiCaduca (empleado,parametro1,idAlerta)
'*Inserta registro en alertas si se cumplen las condiciones para el residente seleccionado
'**********************************************************************************
Function sf_alertaContResiCaduca(ByVal sls_residente, ByVal sls_centro, ByVal sls_param1, ByVal sls_idAlerta, ByVal sls_tipoAlerta)
    slb_crearAlerta = 0
    sls_idioma = sf_getIdioma(sls_centro)
    sls_fechaAyer = (FormatDateTime(DateAdd("d", -1, Date), 2)) 'Dia anterior a hoy
    sls_fechaMan = (FormatDateTime(DateAdd("d", 1, Date), 2)) 'Dia posterior a hoy
    sls_fec = 0
    'Consultamos registro alertas para no duplicar
    Sql = "select id from alertas with (nolock) where idAlerta='" & sls_idAlerta & "' and fecAlta<getDate() and regEstado='A' and residente=" & sls_residente & " "
    Set rsAlertExiste = sf_rec(Sql)
    If Not rsAlertExiste.EOF Then
        slb_crearAlerta = 0
    Else
        Sql = "select id,fechaFin from contratosReservas with (nolock) where regEstado='A' and residente=" & sls_residente & " and "
        Sql = Sql & "(fechaFin<>'' or fechaFin is not null) and (fechaFinReal='' or fechaFinReal is null) "
        Sql = Sql & " and fecAlta<convert(datetime,'" & sls_fechaMan & "',103) and estado='A' and estadoContrato='CONTRATO' order by fecAlta desc"
        Set rsFec = sf_rec(Sql)
        If SGS_ND > 9 Then
            Loga ("Ejecutando: " & Sql)
            Loga ("=================================================================================")
        End If
        If Not rsFec.EOF Then
            sls_fecCad = rsFec("fechaFin")
            sls_dif = DateDiff("d", Date, sls_fecCad)
        Else
            sls_dif = 0
        End If
        If sls_dif > 0 And sls_dif < sls_param1 Then
            slb_crearAlerta = 1
        End If
    End If
    If slb_crearAlerta = 1 Then
        Sql = "select getDate() timeStamp,newId() id"
        Set rsTime = sf_rec(Sql)
        sls_timeStamp = rsTime("timeStamp")
        sls_id = rsTime("id")
        sls_comentario = SGD_Literal.Item(sls_idioma & "-Faltan") & "(" & sls_param1 & ")" & SGD_Literal.Item(sls_idioma & "-ContResiCaduca")
        Sql = "INSERT into alertas(id,idAlerta,centro,residente,usuario,tipo,comentario,actuacion,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod,control) "
        Sql = Sql & "values ('" & sls_id & "','" & sls_idAlerta & "','" & sls_centro & "'," & sls_residente & ",'','" & sls_tipoAlerta & "','" & sf_limpia(sls_comentario) & "','','CREADA','A',"
        Sql = Sql & "'GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "',null)"
        If SGS_ND > 4 Then
            Loga ("!ALERTA generada: " & Sql)
            Loga ("=================================================================================")
        End If
        Set rsIns = sf_rec(Sql)
    End If
End Function
'**********************************************************************************
'************************************************************
'* Funciones necesarias
'************************************************************
Function sf_recGdr(ByVal sls_sql)
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
Function sf_rec(ByVal sls_sql)
    ultimaSql = sls_sql
    
    Set slo_rs = CreateObject("ADODB.Recordset")
    Set slo_comm = CreateObject("ADODB.Command")
    
    slo_comm.ActiveConnection = SGS_CONUSER
    slo_comm.CommandTimeout = 1500
    slo_comm.CommandText = sls_sql
    Set slo_rs = slo_comm.Execute
    Set sf_rec = slo_rs
End Function
Function sf_existeGdr(ByVal sls_tabla)
    Dim Sql
    Sql = "SELECT * FROM sys.objects WHERE name='" & sls_tabla & "' AND type='U'"
    Set rsTbl = sf_recGdr(Sql)
    If Not rsTbl.EOF Then
        sf_existeGdr = True
        rsTbl.Close
    Else
        sf_existeGdr = False
    End If
End Function

Function sf_existeView(ByVal sls_vista)
    Dim Sql
    Sql = "SELECT * FROM sys.objects WHERE name='" & sls_vista & "' AND type='V'"
    Set rsTbl = sf_rec(Sql)
    If Not rsTbl.EOF Then
        sf_existeView = True
        rsTbl.Close
    Else
        sf_existeView = False
    End If
End Function
Function sf_dropView(ByVal sls_vista)
    Dim Sql
    Sql = "DROP VIEW [" & sls_vista & "]"
    Set rsTbl = sf_rec(Sql)
End Function
'************************************************************************************************
'*  Tratamiento Tablas Farmacia Blister
'************************************************************************************************
Function sf_farmaciaBlister(ByVal sls_fecha)
    Dim sls_mes, sls_anyo
    sls_mes = Month(sls_fecha)
    If Len(sls_mes) = 1 Then sls_mes = "0" & sls_mes
    sls_anyo = Year(sls_fecha)
    If sf_existe("farmaciaBlister-" & sls_mes & "-" & sls_anyo) = False Then
        sls_saco = sf_creaFarmaciaBlister(sls_mes, sls_anyo)
    End If
    If sf_existe("farmaciaBlister-" & sls_mes & "-" & sls_anyo & "-firma") = False Then
        sls_saco = sf_creaFarmaciaBlisterFirma(sls_mes, sls_anyo)
    End If
    If sf_existe("farmaciaStock-" & sls_mes & "-" & sls_anyo) = False Then
        sls_saco = sf_creaFarmaciaStock(sls_mes, sls_anyo)
    End If
End Function

'************************************************************************************************
'*  Crea tabla FarmaciaBlister (datos blister)
'************************************************************************************************
Function sf_creaFarmaciaBlister(sls_mes, sls_anyo)
    Sql = " CREATE TABLE [farmaciaBlister-" & sls_mes & "-" & sls_anyo & "]("
    Sql = Sql & "[id]             [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[centro]         [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[codResi]        [nvarchar] (25)    COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[residente]      [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NULL,"
    Sql = Sql & "[numBlister]     [nvarchar] (25)    COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[fecha]          [datetime]                                        NOT NULL,"
    Sql = Sql & "[pautaId]        [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NULL,"
    Sql = Sql & "[regEstado]      [nvarchar] (25)    COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[usrAlta]        [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[usrMod]         [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[fecAlta]        [datetime]                                        NOT NULL,"
    Sql = Sql & "[fecMod]         [datetime]                                        NOT NULL,"
    Sql = Sql & "CONSTRAINT       [PK_farmaciaBlister-" & sls_mes & "-" & sls_anyo & "]  PRIMARY KEY CLUSTERED "
    Sql = Sql & "([id] ASC,[regEstado] ASC,[fecAlta] ASC) WITH (PAD_INDEX=OFF, STATISTICS_NORECOMPUTE=OFF, "
    Sql = Sql & "IGNORE_DUP_KEY=OFF, ALLOW_ROW_LOCKS=ON, ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]"
    Sql = Sql & ") ON [PRIMARY]"
    Set Rs = sf_rec(Sql)
    'Indice
    Sql = "CREATE NONCLUSTERED INDEX [GK_farmaciaBlister-" & sls_mes & "-" & sls_anyo & "] ON [farmaciaBlister-" & sls_mes & "-" & sls_anyo & "] ("
    Sql = Sql & "[residente] ASC,"
    Sql = Sql & "[numBlister] ASC,"
    Sql = Sql & "[fecha] ASC,"
    Sql = Sql & "[regEstado] ASC,"
    Sql = Sql & "[fecAlta] ASC"
    Sql = Sql & " )WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]"
    Set Rs = sf_rec(Sql)
End Function
'************************************************************************************************
'*  Crea tabla FarmaciaBlisterFirma (control de preparacion blister)
'************************************************************************************************'
Function sf_creaFarmaciaBlisterFirma(sls_mes, sls_anyo)
    Sql = "CREATE TABLE [farmaciaBlister-" & sls_mes & "-" & sls_anyo & "-firma]("
    Sql = Sql & "[id]             [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[centro]         [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS  NOT NULL,"
    Sql = Sql & "[codResi]        [nvarchar] (25)    COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[residente]      [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NULL,"
    Sql = Sql & "[numBlister]     [nvarchar] (25)    COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[fechaIni]       [datetime]                                        NOT NULL,"
    Sql = Sql & "[fechaFin]       [datetime]                                        NOT NULL,"
    Sql = Sql & "[firmaProf]      [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NULL,"
    Sql = Sql & "[firmaFarm]      [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NULL,"
    Sql = Sql & "[enviado]        [bit]                                             NULL,"
    Sql = Sql & "[incidencia]     [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NULL,"
    Sql = Sql & "[regEstado]      [nvarchar] (25)    COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[usrAlta]        [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[usrMod]         [nvarchar] (255)   COLLATE Modern_Spanish_CI_AS   NOT NULL,"
    Sql = Sql & "[fecAlta]        [datetime]                                        NOT NULL,"
    Sql = Sql & "[fecMod]         [datetime]                                        NOT NULL,"
    Sql = Sql & "CONSTRAINT       [PK_farmaciaBlister-" & sls_mes & "-" & sls_anyo & "-firma]  PRIMARY KEY CLUSTERED "
    Sql = Sql & "([id] ASC,[regEstado] ASC,[fecAlta] ASC) WITH (PAD_INDEX=OFF, STATISTICS_NORECOMPUTE=OFF, "
    Sql = Sql & "IGNORE_DUP_KEY=OFF, ALLOW_ROW_LOCKS=ON, ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]"
    Sql = Sql & ") ON [PRIMARY]"
    Set Rs = sf_rec(Sql)
    'Indice
    Sql = "CREATE NONCLUSTERED INDEX [GK_farmaciaBlister-" & sls_mes & "-" & sls_anyo & "-firma] ON [farmaciaBlister-" & sls_mes & "-" & sls_anyo & "-firma] ("
    Sql = Sql & "[residente] ASC,"
    Sql = Sql & "[numBlister] ASC,"
    Sql = Sql & "[fechaIni] ASC,"
    Sql = Sql & "[fechaFin] ASC,"
    Sql = Sql & "[regEstado] ASC,"
    Sql = Sql & "[fecAlta] ASC"
    Sql = Sql & " )WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]"
    Set Rs = sf_rec(Sql)
End Function

'************************************************************************************************
'*  Crea tabla FarmaciaStock-MM-YYYY (control de preparacion blister)
'************************************************************************************************''
Function sf_creaFarmaciaStock(sls_mes, sls_anyo)
    Sql = "CREATE TABLE [farmaciaStock-" & sls_mes & "-" & sls_anyo & "]("
    Sql = Sql & "[id]             [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[centro]         [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[codResi]        [nvarchar] (25)    COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[residente]      [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[fechaIni]       [datetime]                                        NULL,"
    Sql = Sql & "[fechaFin]       [datetime]                                        NULL,"
    Sql = Sql & "[tipo]           [nvarchar] (25)    COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[articulo]       [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NULL,"
    Sql = Sql & "[cantidad]       [float]                                           NULL,"
    Sql = Sql & "[pautaId]        [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NULL,"
    Sql = Sql & "[stock]          [float]                                           NULL,"
    Sql = Sql & "[regEstado]      [nvarchar] (25)    COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[usrAlta]        [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[usrMod]         [nvarchar] (255)   COLLATE Modern_Spanish_CI_AI   NOT NULL,"
    Sql = Sql & "[fecAlta]        [datetime]                                        NOT NULL,"
    Sql = Sql & "[fecMod]         [datetime]                                        NOT NULL,"
    Sql = Sql & "CONSTRAINT       [PK_farmaciaStock-" & sls_mes & "-" & sls_anyo & "]  PRIMARY KEY CLUSTERED "
    Sql = Sql & "([id] ASC,[regEstado] ASC,[fecAlta] ASC) WITH (PAD_INDEX=OFF, STATISTICS_NORECOMPUTE=OFF, "
    Sql = Sql & "IGNORE_DUP_KEY=OFF, ALLOW_ROW_LOCKS=ON, ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]"
    Sql = Sql & ") ON [PRIMARY]"
    Set Rs = sf_rec(Sql)
    'Indice
    Sql = "CREATE NONCLUSTERED INDEX [GK_farmaciaStock-" & sls_mes & "-" & sls_anyo & "] ON [farmaciaStock-" & sls_mes & "-" & sls_anyo & "] ("
    Sql = Sql & "[residente] ASC,"
    Sql = Sql & "[articulo] ASC,"
    Sql = Sql & "[fechaIni] ASC,"
    Sql = Sql & "[fechaFin] ASC,"
    Sql = Sql & "[regEstado] ASC,"
    Sql = Sql & "[fecAlta] ASC"
    Sql = Sql & " )WITH (SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, IGNORE_DUP_KEY = OFF, ONLINE = OFF) ON [PRIMARY]"
    Set Rs = sf_rec(Sql)
End Function

Function sf_existe(ByVal sls_tabla)
    Dim Sql
    Sql = "SELECT * FROM sys.objects WHERE name='" & sls_tabla & "' AND type='U'"
    Set rsTbl = sf_rec(Sql)
    If Not rsTbl.EOF Then
        sf_existe = True
        rsTbl.Close
    Else
        sf_existe = False
    End If
End Function
Function sf_dropGdr(ByVal sls_tabla)
    Dim Sql
    Sql = "DROP TABLE [" & sls_tabla & "]"
    Set rsTbl = sf_recGdr(Sql)
End Function
Function sf_drop(ByVal sls_tabla)
    Dim Sql
    Sql = "DROP TABLE [" & sls_tabla & "]"
    Set rsTbl = sf_rec(Sql)
End Function
'************************************************************************
' sf_finalMes(fecha)
' Devuelve el último día del mes de una fecha
'************************************************************************
Function sf_finalMes(ByVal sls_fecha)
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
    sf_finalMes = sln_fm
End Function
'************************************************************************
' sf_getParam(nombre Parámetro)
' Devuelve el valor del parametro solicitado
' Lo convierte al tipo correspondiente
'************************************************************************
Function sf_getParam(ByVal sls_param)
    Dim sls_sql, sls_valor, sls_tip
    sls_valor = ""
    sls_tip = ""
    sls_sql = "select tipoParametro,valor from parametros with (nolock) where codigoParametro='" & sls_param & "' and regEstado='A'"
    Set rsParametro = sf_rec(sls_sql)
    If Not rsParametro.EOF Then
        sls_valor = rsParametro("valor")
        sls_tip = rsParametro("tipoParametro")
    End If
    If sls_tip = "ENTERO" Or sls_tip = "BOLEANO" Then
        If IsNumeric(sls_valor) Then
            sls_valor = CDbl(sls_valor)
        End If
    End If
    If Mid(sls_tip, 1, 3) = "DEC" Then
        sls_valor = CDbl(sls_valor)
    End If
    sf_getParam = sls_valor
End Function
'************************************************************************
'* Funcion para reemplazar caracteres peligrosos en la base de datos
'** Investigar porque se ha puesto el 13 y el 10
'************************************************************************
Function sf_limpia(ByVal sls_texto)
    If IsNull(sls_texto) Then
        sls_texto = ""
    End If
    If sls_texto <> "" Then
        'sls_texto=Replace(sls_texto,chr(13),"")
        'sls_texto=Replace(sls_texto,chr(10),"")
        sls_texto = Replace(sls_texto, " & ", "@@am@@")
        sls_texto = Replace(sls_texto, ";", "@@sc@@")
        sls_texto = Replace(sls_texto, "=", "@@eq@@")
        sls_texto = Replace(sls_texto, "#", "@@al@@")
        sls_texto = Replace(sls_texto, "?", "@@qm@@")
        sls_texto = Replace(sls_texto, "'", "@@qs@@")
        sls_texto = Replace(sls_texto, """", "@@qd@@")
        sls_texto = Replace(sls_texto, "--", "@@dg@@")
        sls_texto = Join(Split(sls_texto, "select"), "@@se@@")
        sls_texto = Join(Split(sls_texto, "drop"), "@@dr@@")
        sls_texto = Join(Split(sls_texto, "insert"), "@@in@@")
        sls_texto = Join(Split(sls_texto, "update"), "@@up@@")
        sls_texto = Join(Split(sls_texto, "delete"), "@@de@@")
    End If
    sf_limpia = sls_texto
End Function
'************************************************************************
'* Funcion para devolver los caracteres peligrosos que habiamos reemplazado anteriormente
'************************************************************************
Function sf_restaura(ByVal sls_texto)
    If IsNull(sls_texto) Then
        sls_texto = ""
    End If
    If sls_texto <> "" Then
        sls_texto = Replace(sls_texto, "@@de@@", "delete")
        sls_texto = Replace(sls_texto, "@@up@@", "update")
        sls_texto = Replace(sls_texto, "@@in@@", "insert")
        sls_texto = Replace(sls_texto, "@@dr@@", "drop")
        sls_texto = Replace(sls_texto, "@@se@@", "select")
        sls_texto = Replace(sls_texto, "@@dg@@", "--")
        sls_texto = Replace(sls_texto, "@@qd@@", """")
        sls_texto = Replace(sls_texto, "@@qs@@", "'")
        sls_texto = Replace(sls_texto, "@@qm@@", "?")
        sls_texto = Replace(sls_texto, "@@al@@", "#")
        sls_texto = Replace(sls_texto, "@@eq@@", "=")
        sls_texto = Replace(sls_texto, "@@sc@@", ";")
        sls_texto = Replace(sls_texto, "@@am@@", " & ")
        sls_texto = Replace(sls_texto, Chr(13), "")
        sls_texto = Replace(sls_texto, Chr(10), "")
    End If
    sf_restaura = sls_texto
End Function


'***************************************************************************************************
'*sf_enviarMailInterno. Funcion para enviar e-mails internos
'* Espera:  Usuario remitente
'*              Residencia destino
'*              Usuario destino
'*              Asunto. Por defecto, "Sin Asunto"
'*              Cuerpo del mensaje en formato HTML
'*              Siempre que envia devuelve true
'***************************************************************************************************
Function sf_enviarMailInterno(ByVal sls_de, ByVal sls_empresa, ByVal sls_a, ByVal sls_asunto, ByVal sls_cuerpo)
    sls_fecha = ""
    Sql = "select newId() id,newId() id2"
    Set rsTime = sf_rec(Sql)
    sls_idMens = rsTime("id")
    sls_idDoc = rsTime("id2")
    Sql = "select getDate() timeStamp"
    Set rsTime = sf_rec(Sql)
    sls_timeStamp = rsTime("timeStamp")
    Sql = "INSERT into mensajes(id,estado,remitente,idRespuesta,idReenvio,asunto,mensaje,fecha,acuseRec,enviado,regEstado,usrAlta,usrMod,fecAlta,fecMod) "
    Sql = Sql & "values ('" & sls_idMens & "','ENVIADO','" & sls_de & "',null,null,'" & sls_asunto & "','" & sls_idDoc & "'," & sf_iif(sls_fecha = "", "getDate()", "'" & sls_fecha & "'") & ",0,1,"
    Sql = Sql & "'A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
    Set rsIns = sf_rec(Sql)
    Sql = "INSERT into infoEditor(id,tabla,campo,valor,regEstado,usrAlta,usrMod,fecAlta,fecMod)"
    Sql = Sql & "values('" & sls_idDoc & "','mensajes','mensaje','" & sls_cuerpo & "','A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
    Set rsIns = sf_rec(Sql)
    'Insertamos registro en mensajeControlE
    Sql = "select getDate() timeStamp,newId() id"
    Set rsTime = sf_rec(Sql)
    sls_timeStamp = rsTime("timeStamp")
    sls_id = rsTime("id")
    Sql = "INSERT into mensajesControlE(id,idMens,tipoEnvio,paraCopia,tipoDestino,destino,fecha,regEstado,usrAlta,usrMod,fecAlta,fecMod) "
    Sql = Sql & "values ('" & sls_id & "','" & sls_idMens & "','ENVIO','PARA','USUARIO','" & sls_a & "'," & sf_iif(sls_fecha = "", "getDate()", "'" & sls_fecha & "'") & ","
    Sql = Sql & "'A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
    Set rsIns = sf_rec(Sql)
    'Insertamos registro en mensajeControl
    Sql = "select getDate() timeStamp,newId() id"
    Set rsTime = sf_rec(Sql)
    sls_idMC = rsTime("id")
    sls_timeStamp = rsTime("timeStamp")
    Sql = "INSERT into mensajesControl(id,idMens,destino,fecha,fecLeido,estado,regEstado,usrAlta,usrMod,fecAlta,fecMod) "
    Sql = Sql & "values ('" & sls_idMC & "','" & sls_idMens & "','" & sls_a & "'," & sf_iif(sls_fecha = "", "getDate()", "'" & sls_fecha & "'") & ",null,'RECIBIDO',"
    Sql = Sql & "'A','GdR','GdR','" & sls_timeStamp & "','" & sls_timeStamp & "')"
    Set rsIns = sf_rec(Sql)
    sf_enviarMailInterno = True
End Function
'*******************************************************************************
'* Funcion sf_existeInd
'* Si existe indice lo elimina antes de crear
'*******************************************************************************
Function sf_existeInd(sls_index)
    Dim Sql
    Sql = "SELECT * FROM sys.indexes WHERE name='" & sls_index & "' "
    Set rsInd = sf_rec(Sql)
    If Not rsInd.EOF Then
        sf_existeInd = True
        rsInd.Close
    Else
        sf_existeInd = False
    End If
End Function
'*******************************************************************************
'*Funcion sf_dropInd
'*Elimina indice
'*******************************************************************************
Function sf_dropInd(ByVal sls_index, ByVal sls_tabla)
    Dim Sql
    Sql = "DROP INDEX [" & sls_index & "] ON [" & sls_tabla & "] WITH ( ONLINE = OFF )"
    Set rsInd = sf_rec(Sql)
End Function
'**********************************************************************************
'* sf_getIdioma(  centro )
'*devuelve el idioma del centro actual
'**********************************************************************************
Function sf_getIdioma(ByVal sls_centro)
    Dim sls_valor
    Sql = "select idioma from centros where id='" & sls_centro & "' and regEstado='A'"
    Set rsIdioma = sf_rec(Sql)
    If Not rsIdioma.EOF Then
        sls_valor = rsIdioma("idioma")
    Else
        sls_valor = ""
    End If
    sf_getIdioma = sls_valor
End Function
'**********************************************************************************
'* sf_getIdiomaE( empresa)
'*devuelve el idioma del centro actual
'**********************************************************************************
Function sf_getIdiomaE(ByVal sls_empresa)
    Dim sls_valor
    Sql = "select idioma from gdrEmpresas with (nolock) where id='" & sls_empresa & "' and regEstado='A'"
    Set rsIdioma = sf_recGdr(Sql)
    If Not rsIdioma.EOF Then
        sls_valor = rsIdioma("idioma")
    Else
        sls_valor = ""
    End If
    sf_getIdiomaE = sls_valor
End Function
'**********************************************************************************
'* Función sf_iif()
'* Equivalente al si() de excel si(condicion,valor si verdadero, valor si falso)
'************************************************************************
Function sf_iif(ByVal slb_cond, ByVal sls_true, ByVal sls_false)
    If slb_cond Then
        sf_iif = sls_true
    Else
        sf_iif = sls_false
    End If
End Function
'**********************************************************************************
'* Función sf_residentes(sls_centro)
'* Devuelve cadena de residentes actuales en la residencia
'* Descarta los ausentes y los residentes de centro de dia que hayan salido
'**********************************************************************************
Function sf_residentes(ByVal sls_centro, ByVal slb_cDia)
    sls_resiCad = ""
    sln_maxTiempoPos = sf_getParam("TACTIEMPOPOS") 'Parametro de maximo tiempo posterior
    If slb_cDia = 1 Then
        slb_regMovCDIA = sf_getParam("REGMOVCDIA") 'Parametro de control de movimientos de CDia
    Else
        slb_regMovCDIA = 0
    End If
    sls_fecha1 = (FormatDateTime(DateAdd("d", 1, sls_hoy), 2) & " 00:00:00.001")   'Dia siguiente
    sls_fecha2 = (FormatDateTime(DateAdd("d", -1, sls_hoy), 2) & " 23:59:59.999")   'Dia anterior
    'Select residentes
    Sql = "select distinct r.id,r.estado,c.libroReg,r.nombre,r.apellido1,r.apellido2 from residentes r with (nolock) left join contratosReservas c with (nolock) on "
    Sql = Sql & " (r.id=c.residente and c.regEstado='A' and c.estadoContrato='CONTRATO')  where (r.estado='A' or r.estado='AU') and r.regEstado='A' "
    If sls_centro <> "" Then
        Sql = Sql & "and r.centro='" & sls_centro & "' "
    End If
    Sql = Sql & "order by r.apellido1,r.apellido2,r.nombre,c.libroReg,r.estado "
    Set rsResi = sf_rec(Sql)
    Do While Not rsResi.EOF
        slb_excluir = 0
        sls_idRes = rsResi("id")
        sls_estado = rsResi("estado")
        sls_libroReg = rsResi("libroReg")
        'Controlamos entradas y salidas si es de CDia
        If sls_libroReg = "CDIA" Then
            If slb_regMovCDIA = 1 Then
                'Si el residente ha entrado al centro hoy
                Sql = "select top 1 id,tipoMovimiento,momento from movimientosCDia with (nolock) where momento>=convert(datetime,'" & sls_fecha2 & "',103) and momento<=getDate() and regEstado='A' "
                Sql = Sql & " and residente='" & sls_idRes & "' and tipoMovimiento='E' order by momento asc"
                Set rsMovE = sf_rec(Sql)
                If Not rsMovE.EOF Then
                    'Comprobamos si el ultimo registro es de salida
                    Sql = "select top 1 id,tipoMovimiento,momento from movimientosCDia with (nolock) where momento>=convert(datetime,'" & sls_fecha2 & "',103) and momento<=getDate() and regEstado='A' "
                    Sql = Sql & " and residente='" & sls_idRes & "' order by momento desc"
                    Set rsMov = sf_rec(Sql)
                    If Not rsMov.EOF Then
                        sls_idMov = rsMov("id")
                        sls_tipoMov = rsMov("tipoMovimiento")
                        sls_momentoMov = rsMov("momento")
                        If sls_tipoMov = "S" Then
                            If DateDiff("n", sls_momentoMov, Date) > sln_maxTiempoPos Then
                                slb_excluir = 1
                            End If
                        End If
                    End If
                Else
                 slb_excluir = 1
                End If
            End If
        End If
        If slb_excluir = 0 Then
            sls_resiCad = sls_resiCad & "'" & sls_idRes & "',"
        End If
        rsResi.MoveNext
        DoEvents
    Loop
    If sls_resiCad <> "" Then
        sls_resiCad = Mid(sls_resiCad, 1, Len(sls_resiCad) - 1)
    End If
    sf_residentes = sls_resiCad
End Function


